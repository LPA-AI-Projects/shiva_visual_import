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
const BATCH_SIZE = 6;
const MODEL = process.env.PLANNER_MODEL || 'claude-sonnet-5';
const MAX_TOKENS = Number(process.env.PLANNER_MAX_TOKENS) || 16_384;

function client() {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) throw new Error('ANTHROPIC_API_KEY is not set');
  return new Anthropic({ apiKey });
}

function trimSlideForPrompt(slide) {
  return {
    index: slide.index,
    title: slide.title,
    text_blocks: slide.text_blocks.map(b => ({
      paragraphs: b.paragraphs,
      column_x: b.x,
      // y is worth keeping even though nothing downstream currently
      // reasons about it explicitly: it's what would let a future
      // heuristic (or Claude itself) recognize bottom-of-slide
      // boilerplate that extract.py's deck-wide pass didn't quite
      // qualify (e.g. present on fewer than 25% of slides). Previously
      // dropped here, which meant that signal never reached the
      // planner at all.
      row_y: b.y,
    })),
    tables: slide.tables.map(t => ({ rows: t.rows })),
    has_images: slide.images.length > 0,
    image_refs: slide.images.map(im => im.path),
    column_hint: slide.column_hint,
    dominant_image_sparse_text: slide.dominant_image_sparse_text || false,
    // Verbatim speaker notes, if the source slide had any. The system
    // prompt instructs the planner to copy this into the target slide's
    // `notes` field unchanged — never summarized, never left behind.
    notes: slide.notes || null,
  };
}

function extractJsonText(text) {
  let jsonText = (text || '').trim();
  if (!jsonText) return jsonText;

  const fenceMatch = jsonText.match(/```(?:json)?\s*([\s\S]*?)```/);
  if (fenceMatch) return fenceMatch[1].trim();

  if (jsonText.startsWith('```')) {
    return jsonText.replace(/^```(?:json)?\s*\n?/, '').trim();
  }
  return jsonText;
}

function parsePlannerResponse(jsonText, batchStartIdx, batchEndIdx) {
  if (!jsonText) {
    const err = new Error(
      `Planner returned empty response for slides ${batchStartIdx}-${batchEndIdx}`
    );
    err.retriable = false;
    throw err;
  }

  let parsed;
  try {
    parsed = JSON.parse(jsonText);
  } catch (e) {
    const err = new Error(
      `Planner returned invalid JSON for slides ${batchStartIdx}-${batchEndIdx}: ${e.message}\n---\n${jsonText.slice(0, 2000)}`
    );
    err.retriable = true;
    throw err;
  }

  if (!parsed.slides || !Array.isArray(parsed.slides)) {
    const err = new Error(
      `Planner response missing "slides" array for batch ${batchStartIdx}-${batchEndIdx}`
    );
    err.retriable = true;
    throw err;
  }
  return parsed.slides;
}

async function planBatchOnce(anthropic, batchSlides, batchStartIdx, batchEndIdx, totalSlides) {
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
    max_tokens: MAX_TOKENS,
    system: SYSTEM_PROMPT,
    messages: [{ role: 'user', content: userContent }],
  });

  const text = resp.content.filter(b => b.type === 'text').map(b => b.text).join('\n');
  const jsonText = extractJsonText(text);

  if (resp.stop_reason === 'max_tokens') {
    const err = new Error(
      `Planner output truncated (max_tokens) for slides ${batchStartIdx}-${batchEndIdx}`
    );
    err.retriable = true;
    throw err;
  }

  return parsePlannerResponse(jsonText, batchStartIdx, batchEndIdx);
}

async function planBatch(anthropic, batchSlides, totalSlides, onProgress) {
  const batchStartIdx = batchSlides[0].index;
  const batchEndIdx = batchSlides[batchSlides.length - 1].index;

  if (onProgress) {
    onProgress({
      stage: 'planning',
      batchStart: batchStartIdx,
      batchEnd: batchEndIdx,
      total: totalSlides,
    });
  }

  try {
    return await planBatchOnce(
      anthropic,
      batchSlides,
      batchStartIdx,
      batchEndIdx,
      totalSlides
    );
  } catch (err) {
    if (batchSlides.length <= 1 || err.retriable === false) throw err;

    console.warn(
      `Planner batch ${batchStartIdx}-${batchEndIdx} failed (${err.message}); retrying as smaller batches`
    );
    const mid = Math.ceil(batchSlides.length / 2);
    const first = await planBatch(
      anthropic,
      batchSlides.slice(0, mid),
      totalSlides,
      onProgress
    );
    const second = await planBatch(
      anthropic,
      batchSlides.slice(mid),
      totalSlides,
      onProgress
    );
    return first.concat(second);
  }
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
    const planned = await planBatch(anthropic, batch, total, onProgress);
    allPlanned.push(...planned);
  }

  const covered = new Set();
  allPlanned.forEach(s => (s.source_indices || []).forEach(i => covered.add(i)));
  const missing = [];
  for (let i = 1; i <= total; i++) if (!covered.has(i)) missing.push(i);
  if (missing.length) {
    console.warn('WARNING: planner did not cover source slides:', missing);
  }

  return {
    targetSlides: allPlanned,
    missingSourceIndices: missing,
    // Extracted once, deck-wide, by extract.py's boilerplate-footer
    // detection — never sent to Claude, never subject to its per-slide
    // judgment. render.js stamps it onto every rendered slide uniformly.
    commonFooter: extracted.common_footer || null,
  };
}

module.exports = { planDeck, SYSTEM_PROMPT };
