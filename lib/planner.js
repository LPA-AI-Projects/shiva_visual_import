/**
 * lib/planner.js
 *
 * Sends extracted source-slide content to Claude in batches, using the
 * brand system prompt, and gets back a structured slide plan (JSON) that
 * the renderer executes deterministically.
 */
const fs = require('fs');
const path = require('path');
const Anthropic = require('@anthropic-ai/sdk');

const SYSTEM_PROMPT = fs.readFileSync(path.join(__dirname, 'system-prompt.md'), 'utf-8');
const BATCH_SIZE = 12; // source slides per Claude call — keeps output reliable & fast
const MODEL = 'claude-sonnet-5';

function client() {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY is not set');
  return new Anthropic({ apiKey });
}

function trimSlideForPrompt(slide) {
  // Drop bulky/irrelevant fields before sending to the model; keep what it needs to reason.
  return {
    index: slide.index,
    title: slide.title,
    text_blocks: slide.text_blocks.map(b => ({
      paragraphs: b.paragraphs,
      column_x: b.x,
    })),
    tables: slide.tables.map(t => ({ rows: t.rows })),
    has_images: slide.images.length > 0,
    image_refs: slide.images.map(im => im.path),
    column_hint: slide.column_hint,
    dominant_image_sparse_text: slide.dominant_image_sparse_text || false,
  };
}

async function planBatch(anthropic, batchSlides, batchStartIdx, batchEndIdx, totalSlides, onProgress) {
  const userContent = [
    `Source deck has ${totalSlides} slides total. This batch covers source slides ${batchStartIdx}-${batchEndIdx}.`,
    `Produce the slide plan for ONLY these slides (source_indices must be within this range).`,
    `Extracted content for slides ${batchStartIdx}-${batchEndIdx}:`,
    '```json',
    JSON.stringify(batchSlides.map(trimSlideForPrompt), null, 1),
    '```',
  ].join('\n');

  const resp = await anthropic.messages.create({
    model: MODEL,
    max_tokens: 8000,
    system: SYSTEM_PROMPT,
    messages: [{ role: 'user', content: userContent }],
  });

  const text = resp.content.filter(b => b.type === 'text').map(b => b.text).join('\n');
  let jsonText = text.trim();
  // Strip markdown code fences if present
  const fenceMatch = jsonText.match(/```(?:json)?\s*([\s\S]*?)```/);
  if (fenceMatch) jsonText = fenceMatch[1].trim();

  let parsed;
  try {
    parsed = JSON.parse(jsonText);
  } catch (e) {
    throw new Error(`Planner returned invalid JSON for slides ${batchStartIdx}-${batchEndIdx}: ${e.message}\n---\n${jsonText.slice(0, 2000)}`);
  }
  if (!parsed.slides || !Array.isArray(parsed.slides)) {
    throw new Error(`Planner response missing "slides" array for batch ${batchStartIdx}-${batchEndIdx}`);
  }
  return parsed.slides;
}

/**
 * Full plan for an extracted deck. Batches through Claude, returns the
 * concatenated slide plan array (one entry per TARGET slide, in order).
 */
async function planDeck(extracted, onProgress) {
  const anthropic = client();
  const slides = extracted.slides;
  const total = slides.length;
  const allPlanned = [];

  for (let start = 0; start < total; start += BATCH_SIZE) {
    const batch = slides.slice(start, start + BATCH_SIZE);
    const batchStartIdx = batch[0].index;
    const batchEndIdx = batch[batch.length - 1].index;
    if (onProgress) onProgress({ stage: 'planning', batchStart: batchStartIdx, batchEnd: batchEndIdx, total });
    const planned = await planBatch(anthropic, batch, batchStartIdx, batchEndIdx, total, onProgress);
    allPlanned.push(...planned);
  }

  // Sanity check: every source index 1..total should appear somewhere.
  const covered = new Set();
  allPlanned.forEach(s => (s.source_indices || []).forEach(i => covered.add(i)));
  const missing = [];
  for (let i = 1; i <= total; i++) if (!covered.has(i)) missing.push(i);
  if (missing.length) {
    console.warn('WARNING: planner did not cover source slides:', missing);
  }

  return { targetSlides: allPlanned, missingSourceIndices: missing };
}

module.exports = { planDeck, SYSTEM_PROMPT };
