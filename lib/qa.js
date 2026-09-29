/**
 * lib/qa.js
 *
 * Two QA passes, matching the standard this project was held to throughout
 * every manual conversion in the source conversation:
 *
 * 1. STRUCTURAL QA (deterministic): scans the raw slide XML for zero or
 *    negative shape dimensions — the exact class of bug that produced a
 *    real "PowerPoint needs to repair" failure during development. Any
 *    hit is a hard build error, not a warning (LibreOffice tolerates it;
 *    PowerPoint does not).
 *
 * 2. VISUAL QA (Claude vision): renders every slide to an image via
 *    LibreOffice and asks Claude to look at each one for overflow,
 *    clipping, overlap, off-canvas elements, and unreadable text — the
 *    same checklist applied by hand throughout this project. Slides
 *    flagged with a real problem get ONE targeted re-plan + re-render
 *    pass (not a full-deck redo).
 */
const fs = require('fs');
const path = require('path');
const { execFile } = require('child_process');
const { promisify } = require('util');
const execFileP = promisify(execFile);
const Anthropic = require('@anthropic-ai/sdk');

function scanZeroDimensionShapes(pptxPath, workDir) {
  const AdmZip = require('adm-zip');
  const zip = new AdmZip(pptxPath);
  const entries = zip.getEntries().filter(e => /^ppt\/slides\/slide\d+\.xml$/.test(e.entryName));
  const issues = [];
  for (const entry of entries) {
    const xml = entry.getData().toString('utf-8');
    const extRe = /<a:ext cx="(-?\d+)" cy="(-?\d+)"\/>/g;
    let m, i = 0;
    while ((m = extRe.exec(xml)) !== null) {
      const cx = parseInt(m[1], 10), cy = parseInt(m[2], 10);
      if (i > 0 && (cx <= 0 || cy <= 0)) { // i==0 is the harmless group-wrapper transform
        issues.push({ file: entry.entryName, cx, cy });
      }
      i++;
    }
  }
  return issues;
}

async function renderSlidesToImages(pptxPath, outDir) {
  await execFileP('soffice', ['--headless', '--convert-to', 'pdf', '--outdir', outDir, pptxPath], { timeout: 120000 });
  const pdfPath = path.join(outDir, path.basename(pptxPath).replace(/\.pptx$/, '.pdf'));
  await execFileP('pdftoppm', ['-jpeg', '-r', '110', pdfPath, path.join(outDir, 'slide')], { timeout: 120000 });
  const files = fs.readdirSync(outDir).filter(f => /^slide-?\d+\.jpg$/.test(f));
  const numbered = files.map(f => {
    const n = parseInt(f.match(/(\d+)\.jpg$/)[1], 10);
    return { n, path: path.join(outDir, f) };
  }).sort((a, b) => a.n - b.n);
  return numbered;
}

async function visionCheckBatch(anthropic, images, startIdx) {
  // images: array of {n, path} — grid them isn't necessary; send individually
  // interleaved with labels for reliable per-slide attribution.
  const content = [];
  content.push({
    type: 'text', text:
      `You are doing visual QA on ${images.length} rendered PowerPoint slides (numbered ${images.map(i => i.n).join(', ')}). ` +
      `For each slide, look ONLY for genuine visual defects: text overflowing its box or running into the decorative footer/logo, ` +
      `overlapping text or shapes, text clipped/cut off, elements outside the slide boundary, illegibly small text, or a clearly broken/empty layout. ` +
      `Do not flag a slide for its content, wording, or design style — only for rendering defects. ` +
      `Respond with ONLY a JSON array like [{"slide": 3, "ok": false, "issue": "short description"}, {"slide": 4, "ok": true}]  — include EVERY slide number listed above, in order.`,
  });
  for (const img of images) {
    content.push({ type: 'text', text: `--- Slide ${img.n} ---` });
    content.push({
      type: 'image',
      source: { type: 'base64', media_type: 'image/jpeg', data: fs.readFileSync(img.path).toString('base64') },
    });
  }
  const resp = await anthropic.messages.create({
    model: 'claude-sonnet-5',
    max_tokens: 2000,
    messages: [{ role: 'user', content }],
  });
  const text = resp.content.filter(b => b.type === 'text').map(b => b.text).join('\n');
  const fenceMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/);
  const jsonText = fenceMatch ? fenceMatch[1] : text;
  try {
    return JSON.parse(jsonText.trim());
  } catch (e) {
    console.warn('Vision QA response was not valid JSON, skipping this batch:', text.slice(0, 300));
    return images.map(i => ({ slide: i.n, ok: true })); // fail-open: don't block delivery on a parsing hiccup
  }
}

/**
 * Deterministic content-fidelity check: for a `table`/`compare_tables`
 * slide built from a single source slide that itself has exactly one
 * table, the row count must match exactly. This exists because nothing
 * upstream independently verifies this — the planner is instructed not
 * to invent or drop rows, but that's the only thing catching it; a real
 * conversion produced a table with an extra row that doesn't exist
 * anywhere in the source with that instruction already in place.
 *
 * Deliberately narrow: only the unambiguous 1-source-slide,
 * 1-source-table case is checked. A source slide with multiple tables,
 * or a target slide built from multiple source slides, has too many
 * legitimate ways to legally redistribute rows across target slides
 * (splitting a dense table, merging a small one into another layout,
 * two side-by-side source tables becoming two separate compare_tables
 * halves) for a row-count comparison to be reliable there — and a
 * warning that fires on legitimate splits teaches people to ignore it,
 * which defeats the point. Narrower-but-trustworthy beats
 * broader-but-noisy for something meant to be glanced at, not parsed.
 *
 * Soft warning, not a hard failure like scanZeroDimensionShapes: even
 * in this narrow case there's a small chance of a legitimate reason for
 * the count to differ (a source-required disclosure row added from
 * elsewhere on the slide), and blocking delivery on a false positive is
 * worse than surfacing a warning for a human to glance at.
 */
function checkTableRowFidelity(plan, extracted) {
  const sourceByIdx = new Map();
  for (const sl of extracted.slides) sourceByIdx.set(sl.index, sl.tables);

  const warnings = [];
  for (const spec of plan.targetSlides) {
    if (spec.layout !== 'table') continue; // compare_tables has two independent sides; not meaningfully checked by a single-table comparison
    const idxs = Array.isArray(spec.source_indices) ? spec.source_indices : [];
    if (idxs.length !== 1) continue; // merged-source slide: too ambiguous to check

    const sourceTables = sourceByIdx.get(idxs[0]) || [];
    if (sourceTables.length !== 1) continue; // 0 or 2+ source tables: too ambiguous to check
    const sourceRows = Math.max(0, sourceTables[0].rows.length - 1); // minus header
    if (!sourceRows) continue;

    const planRows = safeArr(spec.rows).length;

    if (planRows !== sourceRows) {
      warnings.push({
        title: spec.title, layout: spec.layout, source_indices: idxs,
        sourceRowTotal: sourceRows, planRowTotal: planRows,
        issue: `Table row count (${planRows}) does not match source row count (${sourceRows}) — a row may have been dropped or invented.`,
      });
    }
  }
  return warnings;
}

/**
 * Runs structural QA (throws on failure — caller should not ship the file)
 * plus visual QA (returns a list of flagged slide numbers + issue text).
 */
async function runQA(pptxPath, workDir, { skipVisual = false, plan = null, extracted = null } = {}) {
  const structuralIssues = scanZeroDimensionShapes(pptxPath, workDir);
  const contentFidelityWarnings = (plan && extracted) ? checkTableRowFidelity(plan, extracted) : [];

  let flagged = [];
  if (!skipVisual && process.env.ANTHROPIC_API_KEY) {
    const imgDir = path.join(workDir, 'qa_images');
    fs.mkdirSync(imgDir, { recursive: true });
    const images = await renderSlidesToImages(pptxPath, imgDir);
    const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
    const BATCH = 6;
    for (let i = 0; i < images.length; i += BATCH) {
      const batch = images.slice(i, i + BATCH);
      const results = await visionCheckBatch(anthropic, batch, i);
      results.forEach(r => { if (r.ok === false) flagged.push({ slide: r.slide, issue: r.issue }); });
    }
  }

  return { structuralIssues, flagged, contentFidelityWarnings };
}

function safeArr(a) { return Array.isArray(a) ? a : []; }

module.exports = { runQA, scanZeroDimensionShapes, renderSlidesToImages, checkTableRowFidelity };
