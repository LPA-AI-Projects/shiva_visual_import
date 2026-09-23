/**
 * lib/render.js
 *
 * Takes the planner's JSON slide plan and deterministically builds the
 * actual PPTX using the brand component library. This is the "unified
 * layout dispatcher": each `layout` value maps to a fixed sequence of
 * component calls, so the LLM never has to know pptxgenjs calling
 * conventions — only content + layout choice.
 */
const { createDeck } = require('./components');

function safeArr(a) { return Array.isArray(a) ? a : []; }

function renderSlide(deck, spec, imagesDir) {
  const path = require('path');
  switch (spec.layout) {
    case 'cover': {
      deck.coverSlide({
        eyebrow: spec.eyebrow, title: spec.title, titleAccent: spec.titleAccent,
        tagline: spec.tagline, byline: spec.byline, pills: safeArr(spec.pills),
      });
      break;
    }
    case 'toc': {
      deck.tocSlide(spec.title, safeArr(spec.items));
      break;
    }
    case 'module_divider': {
      deck.moduleDivider(spec.moduleNum, spec.title, spec.subtitle);
      break;
    }
    case 'quote': {
      deck.quoteSlide(spec.quote, spec.caption);
      break;
    }
    case 'title_body': {
      const s = deck.newSlide();
      if (spec.eyebrow) deck.eyebrow(s, spec.eyebrow);
      deck.titleOnly(s, spec.title);
      let y = 1.5;
      if (spec.subtitle) y = deck.block(s, 0.6, y, 8.75, spec.subtitle, { h: 0.32, size: 11, italic: true, color: deck.pres ? undefined : undefined });
      safeArr(spec.paragraphs).forEach(p => {
        const approxLines = Math.max(1, Math.ceil(p.length / 100));
        const h = approxLines * 0.25 + 0.15;
        y = deck.block(s, 0.6, y, 8.75, p, { h, size: 11.5 }) + 0.08;
      });
      break;
    }
    case 'checklist': {
      const s = deck.newSlide();
      if (spec.eyebrow) deck.eyebrow(s, spec.eyebrow);
      if (spec.tag) deck.tag(s, spec.tag, colorByName(spec.tagColor));
      deck.titleOnly(s, spec.title);
      let y = 1.5;
      if (spec.subtitle) { y = deck.block(s, 0.6, y, 8.75, spec.subtitle, { h: 0.3, size: 11, italic: true, color: '6E6E6E' }); y += 0.15; }
      deck.checklistAuto(s, 0.6, y, 8.5, safeArr(spec.items), { size: 11.5, gap: 0.22 });
      break;
    }
    case 'numbered_list': {
      const s = deck.newSlide();
      if (spec.eyebrow) deck.eyebrow(s, spec.eyebrow);
      deck.titleOnly(s, spec.title);
      let y = 1.5;
      if (spec.subtitle) { y = deck.block(s, 0.6, y, 8.75, spec.subtitle, { h: 0.3, size: 11, italic: true, color: '6E6E6E' }); y += 0.15; }
      deck.numberedList(s, safeArr(spec.items), { y, size: 11.5 });
      break;
    }
    case 'card_grid': {
      const s = deck.newSlide();
      if (spec.eyebrow) deck.eyebrow(s, spec.eyebrow);
      deck.titleOnly(s, spec.title);
      let y = 1.7;
      if (spec.subtitle) { deck.block(s, 0.6, 1.45, 8.75, spec.subtitle, { h: 0.3, size: 11, italic: true, color: '6E6E6E' }); y = 1.85; }
      deck.cardGrid(s, safeArr(spec.cards), { y, h: 5.15 - y - 0.35 });
      break;
    }
    case 'two_col_cards': {
      const s = deck.newSlide();
      if (spec.eyebrow) deck.eyebrow(s, spec.eyebrow);
      deck.titleOnly(s, spec.title);
      let y = 1.6;
      if (spec.subtitle) { deck.block(s, 0.6, 1.45, 8.75, spec.subtitle, { h: 0.3, size: 11, italic: true, color: '6E6E6E' }); y = 1.85; }
      const items = safeArr(spec.items).map(it => [it.title, it.desc, colorByName(it.color)]);
      const rows = Math.ceil(items.length / 2);
      const ch = Math.max(1.1, (5.15 - y - 0.25 * (rows - 1) - 0.3) / rows);
      deck.twoColCards(s, items, { y, ch });
      break;
    }
    case 'compare_columns': {
      const s = deck.newSlide();
      if (spec.eyebrow) deck.eyebrow(s, spec.eyebrow);
      deck.titleOnly(s, spec.title);
      deck.compareColumns(s, {
        left: { heading: spec.left.heading, intro: spec.left.intro, rows: safeArr(spec.left.rows), color: colorByName(spec.left.color) },
        right: { heading: spec.right.heading, intro: spec.right.intro, rows: safeArr(spec.right.rows), color: colorByName(spec.right.color) },
        y: 1.55, h: 3.75,
      });
      break;
    }
    case 'weak_strong': {
      const s = deck.newSlide();
      if (spec.eyebrow) deck.eyebrow(s, spec.eyebrow);
      deck.titleOnly(s, spec.title);
      deck.weakStrong(s, spec.weakText, spec.strongText, { y: 1.7, h: 3.3, labels: spec.labels });
      break;
    }
    case 'table': {
      const s = deck.newSlide();
      if (spec.eyebrow) deck.eyebrow(s, spec.eyebrow);
      deck.titleOnly(s, spec.title);
      const headers = safeArr(spec.headers);
      const n = headers.length || 1;
      const colW = headers.map(() => 8.75 / n);
      deck.customTable(s, headers, safeArr(spec.rows), colW, { y: 1.6 });
      break;
    }
    case 'flow_diagram': {
      const s = deck.newSlide();
      if (spec.eyebrow) deck.eyebrow(s, spec.eyebrow);
      deck.titleOnly(s, spec.title);
      const steps = safeArr(spec.steps).map(st => ({ label: st.label, caption: st.caption }));
      deck.flowDiagram(s, { steps, y: 2.0 });
      break;
    }
    case 'kpi_stats': {
      deck.kpiStatGrid(spec.title, safeArr(spec.stats), { subtitle: spec.subtitle });
      break;
    }
    case 'donut_stat': {
      deck.donutStat(spec.title, safeArr(spec.segments), { subtitle: spec.subtitle });
      break;
    }
    case 'use_case': {
      deck.useCaseSlide(spec.title, safeArr(spec.parts));
      break;
    }
    case 'activity': {
      deck.activitySlide(spec.title, spec.contextLine, safeArr(spec.steps), spec.expectedOutput);
      break;
    }
    case 'maturity_levels': {
      const s = deck.newSlide();
      deck.titleOnly(s, spec.title);
      deck.maturityLevels(s, safeArr(spec.levels), { y: 1.6 });
      break;
    }
    case 'lab': {
      deck.labSlide(spec.title, spec.objective, spec.scenario, spec.tools, safeArr(spec.instructions), safeArr(spec.deliverables), safeArr(spec.evalCriteria));
      break;
    }
    case 'solution_guide': {
      deck.solutionGuideSlide(spec.title, safeArr(spec.items));
      break;
    }
    case 'knowledge_check': {
      deck.knowledgeCheckSlide(spec.title, spec.scenario, spec.question, safeArr(spec.options));
      break;
    }
    case 'answer_rationale': {
      deck.answerRationaleSlide(spec.title, spec.correctAnswer, spec.why, safeArr(spec.wrongAnswers));
      break;
    }
    case 'resources': {
      deck.resourcesSlide(spec.title, safeArr(spec.resources));
      break;
    }
    case 'formula': {
      deck.formulaSlide(spec.title, spec.formula, spec.plainEnglish, safeArr(spec.notes));
      break;
    }
    case 'image_full': {
      if (spec.source_image_ref && imagesDir) {
        const imgPath = path.join(imagesDir, spec.source_image_ref);
        deck.imageSlide(spec.title, imgPath);
      } else {
        // No image available — fall back to a plain title slide rather than crash.
        const s = deck.newSlide();
        deck.titleOnly(s, spec.title || 'Untitled');
      }
      break;
    }
    default: {
      // Unknown layout from the model — fail soft with a checklist of whatever text we have.
      const s = deck.newSlide();
      deck.titleOnly(s, spec.title || 'Untitled');
      const fallbackItems = safeArr(spec.items).length ? spec.items : ['(Content could not be placed automatically — please review source slide ' + safeArr(spec.source_indices).join(', ') + ')'];
      deck.checklistAuto(s, 0.6, 1.6, 8.5, fallbackItems, { size: 11 });
    }
  }
}

function colorByName(name) {
  const { TEAL, BLUE, ORANGE, PURPLE } = require('./components');
  if (!name) return undefined;
  const key = String(name).toLowerCase();
  if (key.includes('teal')) return TEAL;
  if (key.includes('blue')) return BLUE;
  if (key.includes('orange')) return ORANGE;
  if (key.includes('purple')) return PURPLE;
  if (/^[0-9a-f]{6}$/i.test(name)) return name;
  return undefined;
}

/**
 * Renders a full plan to a pptx file on disk. Returns { outPath, bulletQueuePath }.
 */
async function renderPlan(plan, outDir, imagesDir) {
  const fs = require('fs');
  const path = require('path');
  const deck = createDeck();

  for (const spec of plan.targetSlides) {
    try {
      renderSlide(deck, spec, imagesDir);
    } catch (e) {
      // Never let one bad slide spec kill the whole deck — render a visible
      // placeholder instead so the rest of the conversion still completes.
      console.error('Render error on slide spec', JSON.stringify(spec).slice(0, 300), e);
      const s = deck.newSlide();
      deck.titleOnly(s, (spec && spec.title) || 'Slide');
      deck.checklistAuto(s, 0.6, 1.6, 8.5, ['(This slide could not be rendered automatically: ' + e.message + ')'], { size: 10 });
    }
  }

  const rawPath = path.join(outDir, 'raw.pptx');
  const bulletQueuePath = path.join(outDir, 'bullets.json');
  await deck.pres.writeFile({ fileName: rawPath });
  deck.writeBulletColorQueue(bulletQueuePath);
  return { rawPath, bulletQueuePath, slideCount: plan.targetSlides.length };
}

module.exports = { renderPlan, renderSlide };
