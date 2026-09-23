/**
 * lib/components.js
 *
 * THE LEARNERS POINT BRAND COMPONENT LIBRARY.
 *
 * This is not a generic template engine. Every function in this file
 * reproduces a specific, battle-tested visual pattern that was developed
 * and refined across dozens of real PPTX brand-conversion projects.
 *
 * Canvas: 10in x 5.625in (16:9). Background gradient + logo + footer decor
 * are applied to every slide automatically via newSlide().
 *
 * IMPORTANT: unlike the original conversational scripts (which used a
 * single module-level `pres`), this library is a FACTORY: call
 * createDeck() once per conversion job to get an isolated deck instance
 * (pres + newSlide + every component, all bound to that instance). This
 * makes it safe to run multiple conversions concurrently on a server.
 */
const pptxgen = require('pptxgenjs');
const path = require('path');
const fs = require('fs');

const ASSETS_DIR = path.join(__dirname, '..', 'assets');
const A = p => path.join(ASSETS_DIR, p);

// ---- Brand constants ----
const TEAL = '2A9D8F';
const BLUE = '3D6CB9';
const ORANGE = 'E8783A';
const PURPLE = '8E6BC7';
const TITLE_C = '1A3A5C';   // navy title
const BODY_C = '333333';
const GRAY = '6E6E6E';
const WHITE = 'FFFFFF';
const CARD_BORDER = 'E7E7E7';
const CARD_BG = 'FAFAFC';

const F_TITLE = 'Helvetica';          // titles
const F_BODY = 'Helvetica Neue';      // body text
const F_MED = 'Helvetica Neue Medium'; // labels/emphasis

const SLIDE_W = 10;
const SLIDE_H = 5.625;

const ACCENT_CYCLE = [TEAL, BLUE, ORANGE, PURPLE];

/**
 * Create an isolated deck. Returns { pres, newSlide, ...allComponents }.
 */
function createDeck() {
  const pres = new pptxgen();
  pres.defineLayout({ name: 'LP', width: SLIDE_W, height: SLIDE_H });
  pres.layout = 'LP';

  const bulletColorQueue = [];
  let curSlideBulletLog = null;
  function beginSlideBulletLog() { curSlideBulletLog = []; bulletColorQueue.push(curSlideBulletLog); }
  function logBullets(count, color) { if (curSlideBulletLog) curSlideBulletLog.push({ count, color }); }

  function addBackground(slide) {
    slide.addImage({ path: A('bg_gradient.png'), x: -0.01, y: -0.01, w: SLIDE_W + 0.02, h: SLIDE_H + 0.02 });
  }
  function addHeader(slide) {
    const w = 1.461, h = 0.297;
    const x = SLIDE_W - 0.336 - w, y = 0.258;
    slide.addImage({ path: A('lp_logo.png'), x, y, w, h });
  }
  function addFooterDecor(slide) {
    const w = 2.68, h = 0.72;
    slide.addImage({ path: A('decor_blobs_right_cropped.png'), x: SLIDE_W - w + 0.02, y: SLIDE_H - h, w, h });
  }

  function newSlide() {
    const s = pres.addSlide();
    beginSlideBulletLog();
    addBackground(s);
    addHeader(s);
    addFooterDecor(s);
    return s;
  }

  function titleOnly(slide, title) {
    slide.addText(title, { x: 0.6, y: 0.72, w: 8.3, h: 0.5, fontFace: F_TITLE, bold: true, fontSize: 22, color: TITLE_C, align: 'left' });
  }

  function eyebrow(slide, text) {
    slide.addText(String(text).toUpperCase(), { x: 0.6, y: 0.38, w: 5.2, h: 0.26, fontFace: F_MED, bold: true, fontSize: 10, color: TEAL, align: 'left' });
  }

  function tag(slide, text, color) {
    const w = 0.11 * text.length + 0.3;
    const x = 6.0 - w;
    slide.addShape('roundRect', { x, y: 0.36, w, h: 0.26, rectRadius: 0.04, fill: { color }, line: { type: 'none' } });
    slide.addText(String(text).toUpperCase(), { x, y: 0.36, w, h: 0.26, fontFace: F_MED, bold: true, fontSize: 8, color: WHITE, align: 'center', valign: 'middle' });
  }

  function block(slide, x, y, w, text, opts = {}) {
    slide.addText(text, {
      x, y, w, h: opts.h || 0.5, fontFace: opts.bold ? F_MED : F_BODY, bold: !!opts.bold, italic: !!opts.italic,
      fontSize: opts.size || 11, color: opts.color || BODY_C, valign: opts.valign || 'top', align: opts.align || 'left',
      lineSpacingMultiple: opts.ls || 1.15,
    });
    return y + (opts.h || 0.5);
  }

  function checklistAuto(slide, x, y, w, items, opts = {}) {
    const size = opts.size || 10.5;
    const charsPerLine = (w - 0.26) * 118 / size;
    const gap = opts.gap != null ? opts.gap : 0.12;
    items.forEach(it => {
      const label = typeof it === 'string' ? it : it.label;
      const m = label.match(/^([A-Z][A-Za-z0-9 &()\/'\u2019\u2011-]{2,60}):\s*(.+)$/s);
      const lines = Math.max(1, Math.ceil(label.length / charsPerLine));
      const rowH = lines * (size / 72) * 1.3 + gap;
      slide.addShape('roundRect', { x, y: y + 0.03, w: 0.14, h: 0.14, rectRadius: 0.03, fill: { color: WHITE }, line: { color: TEAL, width: 1 } });
      if (m && m[2].length > 3) {
        slide.addText([{ text: m[1] + ':  ', options: { bold: true, fontFace: F_MED, fontSize: size, color: TITLE_C } },
                        { text: m[2], options: { fontFace: F_BODY, fontSize: size, color: BODY_C } }],
          { x: x + 0.24, y, w: w - 0.24, h: rowH, valign: 'top', lineSpacingMultiple: 1.15 });
      } else {
        slide.addText(label, { x: x + 0.24, y, w: w - 0.24, h: rowH, fontFace: F_BODY, fontSize: size, color: BODY_C, valign: 'top', lineSpacingMultiple: 1.15 });
      }
      y += rowH;
    });
    return y;
  }

  function checkItemsAuto(slide, x, y, w, items, opts = {}) {
    const size = opts.size || 10.5;
    const charsPerLine = (w - 0.32) * 118 / size;
    const gap = opts.gap != null ? opts.gap : 0.13;
    items.forEach(it => {
      const lines = Math.max(1, Math.ceil(it.length / charsPerLine));
      const rowH = lines * (size / 72) * 1.3 + gap;
      slide.addText('\u2713', { x, y: y - 0.02, w: 0.28, h: rowH, fontFace: F_MED, bold: true, fontSize: size + 1, color: TEAL, valign: 'top' });
      slide.addText(it, { x: x + 0.3, y, w: w - 0.3, h: rowH, fontFace: F_BODY, fontSize: size, color: BODY_C, valign: 'top', lineSpacingMultiple: 1.15 });
      y += rowH;
    });
    return y;
  }

  function numberedList(slide, items, opts = {}) {
    const x = opts.x != null ? opts.x : 0.6;
    let y = opts.y != null ? opts.y : 1.5;
    const w = opts.w != null ? opts.w : 8.75;
    const size = opts.size || 11;
    const dia = opts.dia || 0.3;
    const textX = x + dia + 0.2;
    const textW = w - dia - 0.2;
    const charsPerLine = textW * 118 / size;
    const gap = opts.gap != null ? opts.gap : 0.14;
    items.forEach((it, i) => {
      const term = typeof it === 'object' ? it.term : null;
      const def = typeof it === 'object' ? it.def : it;
      const fullLen = (term ? term.length + 3 : 0) + def.length;
      const lines = Math.max(1, Math.ceil(fullLen / charsPerLine));
      const rowH = Math.max(dia, lines * (size / 72) * 1.3 + gap);
      slide.addShape('ellipse', { x, y, w: dia, h: dia, fill: { color: TEAL }, line: { type: 'none' } });
      slide.addText(String(i + 1), { x, y, w: dia, h: dia, fontFace: F_MED, bold: true, fontSize: size - 1.5, color: WHITE, align: 'center', valign: 'middle', wrap: false });
      const runs = term
        ? [{ text: term + '  \u2014  ', options: { bold: true, fontFace: F_MED, fontSize: size, color: TITLE_C } }, { text: def, options: { fontFace: F_BODY, fontSize: size, color: BODY_C } }]
        : [{ text: def, options: { fontFace: F_BODY, fontSize: size, color: BODY_C } }];
      slide.addText(runs, { x: textX, y: y - 0.03, w: textW, h: rowH, valign: 'top', lineSpacingMultiple: 1.18 });
      y += rowH;
    });
    return y;
  }

  function bulletCard(slide, { title, items, x = 0.6, y = 1.35, w = 8.75, h = 3.55, fontSize = 13, space = 9 }) {
    slide.addShape('roundRect', { x, y, w, h, rectRadius: 0.07, fill: { color: WHITE }, line: { color: CARD_BORDER, width: 0.75 },
      shadow: { type: 'outer', color: '9A9A9A', opacity: 0.15, blur: 6, offset: 1.5, angle: 90 } });
    const paras = items.map(t => ({ text: t, options: { bullet: { code: '25CF', indent: 14 }, fontFace: F_BODY, fontSize, color: BODY_C, breakLine: true, paraSpaceAfter: space } }));
    slide.addText(paras, { x: x + 0.32, y: y + 0.22, w: w - 0.64, h: h - 0.4, valign: 'top', lineSpacingMultiple: 1.08 });
    logBullets(items.length, TEAL);
  }

  function cardGrid(slide, cards, opts = {}) {
    const n = cards.length;
    const gap = opts.gap != null ? opts.gap : 0.22;
    const cw = (8.75 - gap * (n - 1)) / n;
    const cy = opts.y || 1.7, ch = opts.h || 2.9;
    const titleSize = opts.titleSize || 12.5;
    const titleCharsPerLine = (cw - 0.36) * 90 / titleSize;
    cards.forEach((card, i) => {
      const cx = 0.6 + i * (cw + gap);
      slide.addShape('roundRect', { x: cx, y: cy, w: cw, h: ch, rectRadius: 0.07, fill: { color: WHITE }, line: { color: 'E5E5E5', width: 0.75 },
        shadow: { type: 'outer', color: '9A9A9A', opacity: 0.13, blur: 5, offset: 1.5, angle: 90 } });
      const color = (opts.colors || ACCENT_CYCLE)[i % (opts.colors || ACCENT_CYCLE).length];
      if (opts.numbered !== false) {
        slide.addShape('ellipse', { x: cx + 0.18, y: cy + 0.18, w: 0.34, h: 0.34, fill: { color }, line: { type: 'none' } });
        slide.addText(String(i + 1), { x: cx + 0.18, y: cy + 0.18, w: 0.34, h: 0.34, fontFace: F_MED, bold: true, fontSize: 13, color: WHITE, align: 'center', valign: 'middle' });
      } else {
        slide.addShape('rect', { x: cx, y: cy, w: cw, h: 0.07, fill: { color }, line: { type: 'none' } });
      }
      const titleY = opts.numbered !== false ? cy + 0.62 : cy + 0.22;
      const titleLines = Math.max(1, Math.ceil(card.title.length / titleCharsPerLine));
      const titleH = titleLines * (titleSize / 72) * 1.25 + 0.1;
      slide.addText(card.title, { x: cx + 0.18, y: titleY, w: cw - 0.36, h: titleH, fontFace: F_MED, bold: true, fontSize: titleSize, color: TITLE_C, valign: 'top', lineSpacingMultiple: 1.1 });
      const descY = titleY + titleH + 0.08;
      slide.addText(card.desc, { x: cx + 0.18, y: descY, w: cw - 0.36, h: cy + ch - descY - 0.15, fontFace: F_BODY, fontSize: opts.descSize || 10, color: BODY_C, valign: 'top', lineSpacingMultiple: 1.2 });
    });
  }

  function twoColCards(slide, items, opts = {}) {
    const cw = (8.75 - 0.3) / 2;
    const ch = opts.ch || 1.45;
    const y0 = opts.y || 1.6;
    items.forEach(([title, desc, color], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const cx = 0.6 + col * (cw + 0.3);
      const cy = y0 + row * (ch + 0.25);
      slide.addShape('roundRect', { x: cx, y: cy, w: cw, h: ch, rectRadius: 0.06, fill: { color: WHITE }, line: { color: 'E5E5E5', width: 0.75 },
        shadow: { type: 'outer', color: '9A9A9A', opacity: 0.12, blur: 4, offset: 1.5, angle: 90 } });
      slide.addShape('rect', { x: cx, y: cy, w: 0.07, h: ch, fill: { color: color || TEAL }, line: { type: 'none' } });
      slide.addText(title, { x: cx + 0.22, y: cy + 0.14, w: cw - 0.4, h: 0.32, fontFace: F_MED, bold: true, fontSize: 12.5, color: TITLE_C });
      slide.addText(desc, { x: cx + 0.22, y: cy + 0.5, w: cw - 0.4, h: ch - 0.62, fontFace: F_BODY, fontSize: 9.3, color: BODY_C, valign: 'top', lineSpacingMultiple: 1.2 });
    });
  }

  function numberedStepCards(slide, steps, opts = {}) {
    let y = opts.y || 1.55;
    const gap = opts.gap != null ? opts.gap : 0.08;
    const descSize = opts.descSize || 11;
    const descW = opts.descW || 7.9;
    const charsPerLine = descW * 118 / descSize;
    steps.forEach((step, i) => {
      const lines = Math.max(1, Math.ceil(step.desc.length / charsPerLine));
      const cardH = Math.max(opts.minCardH || 0.62, 0.38 + lines * (descSize / 72) * 1.25 + 0.12);
      slide.addShape('roundRect', { x: 0.6, y, w: 8.75, h: cardH, rectRadius: 0.05, fill: { color: WHITE }, line: { type: 'none' },
        shadow: { type: 'outer', color: '9A9A9A', opacity: 0.1, blur: 4, offset: 1, angle: 90 } });
      slide.addShape('rect', { x: 0.6, y, w: 0.06, h: cardH, fill: { color: BLUE }, line: { type: 'none' } });
      slide.addShape('ellipse', { x: 0.82, y: y + 0.13, w: 0.32, h: 0.32, fill: { color: BLUE }, line: { type: 'none' } });
      slide.addText(String(i + 1), { x: 0.82, y: y + 0.13, w: 0.32, h: 0.32, fontFace: F_MED, bold: true, fontSize: 12, color: WHITE, align: 'center', valign: 'middle' });
      slide.addText(step.title, { x: 1.28, y: y + 0.08, w: descW, h: 0.3, fontFace: F_MED, bold: true, fontSize: 13.5, color: TITLE_C, valign: 'top' });
      slide.addText(step.desc, { x: 1.28, y: y + 0.38, w: descW, h: cardH - 0.4, fontFace: F_BODY, fontSize: descSize, color: GRAY, valign: 'top', lineSpacingMultiple: 1.1 });
      y += cardH + gap;
    });
    return y;
  }

  function gridCards(slide, { cards, x = 0.6, y = 1.35, w = 8.75, h = 3.55, cols = 2, rowGap = 0.22, colGap = 0.24, fontSize = 12 }) {
    const rows = Math.ceil(cards.length / cols);
    const cw = (w - colGap * (cols - 1)) / cols;
    const ch = (h - rowGap * (rows - 1)) / rows;
    cards.forEach((c, i) => {
      const col = i % cols, row = Math.floor(i / cols);
      const cx = x + col * (cw + colGap), cy = y + row * (ch + rowGap);
      slide.addShape('roundRect', { x: cx, y: cy, w: cw, h: ch, rectRadius: 0.06, fill: { color: WHITE }, line: { color: CARD_BORDER, width: 0.75 },
        shadow: { type: 'outer', color: '9A9A9A', opacity: 0.13, blur: 5, offset: 1.2, angle: 90 } });
      slide.addShape('rect', { x: cx + 0.06, y: cy, w: cw - 0.12, h: 0.045, fill: { color: c.color || TEAL }, line: { type: 'none' } });
      slide.addText(c.heading, { x: cx + 0.2, y: cy + 0.13, w: cw - 0.4, h: 0.3, fontFace: F_MED, bold: true, fontSize: 12.5, color: TITLE_C });
      slide.addText(c.body, { x: cx + 0.2, y: cy + 0.46, w: cw - 0.4, h: ch - 0.6, fontFace: F_BODY, fontSize, color: BODY_C, valign: 'top', lineSpacingMultiple: 1.08 });
    });
  }

  function customTable(slide, headers, rows, colW, opts = {}) {
    const x0 = opts.x || 0.6, y0 = opts.y || 1.55;
    let y = y0, cx = x0;
    const hH = opts.headerH || 0.36;
    const headerSize = opts.headerSize || 10.5;
    headers.forEach((h, i) => {
      slide.addShape('rect', { x: cx, y, w: colW[i], h: hH, fill: { color: TITLE_C }, line: { color: WHITE, width: 0.5 } });
      slide.addText(h, { x: cx + 0.08, y, w: colW[i] - 0.16, h: hH, fontFace: F_MED, bold: true, fontSize: headerSize, color: WHITE, valign: 'middle' });
      cx += colW[i];
    });
    y += hH;
    const bodySize = opts.bodySize || 10;
    rows.forEach((row, ri) => {
      const cellLines = row.map((cell, ci) => {
        const charsPerLine = (colW[ci] - 0.16) * 118 / bodySize;
        return Math.max(1, Math.ceil(String(cell).length / charsPerLine));
      });
      const maxLines = Math.max(...cellLines);
      const rh = Math.max(opts.minRowH || 0.34, maxLines * (bodySize / 72) * 1.25 + 0.12);
      let cx2 = x0;
      const fill = ri % 2 === 0 ? 'F5F5F5' : WHITE;
      row.forEach((cell, ci) => {
        slide.addShape('rect', { x: cx2, y, w: colW[ci], h: rh, fill: { color: fill }, line: { color: 'E0E0E0', width: 0.5 } });
        slide.addText(String(cell), { x: cx2 + 0.08, y, w: colW[ci] - 0.16, h: rh, fontFace: (opts.boldCol0 && ci === 0) ? F_MED : F_BODY, bold: !!(opts.boldCol0 && ci === 0), fontSize: bodySize, color: (opts.boldCol0 && ci === 0) ? TITLE_C : BODY_C, valign: 'middle', lineSpacingMultiple: 1.05 });
        cx2 += colW[ci];
      });
      y += rh;
    });
    return y;
  }

  function glossaryTable(slide, { rows, x = 0.5, y = 1.15, w = 9.35, termW = 2.15, fontSize = 9.3, termFontSize = 9.6 }) {
    const defW = w - termW;
    const defCharsPerLine = defW * 118 / fontSize;
    const termCharsPerLine = (termW - 0.15) * 82 / termFontSize;
    let curY = y;
    rows.forEach((r, ri) => {
      const defLines = Math.max(1, Math.ceil(r[1].length / defCharsPerLine));
      const termLines = Math.max(1, Math.ceil(r[0].length / termCharsPerLine));
      const lines = Math.max(defLines, termLines);
      const rowH = Math.max(0.26, lines * (fontSize / 72) * 1.42 + 0.09);
      const fill = ri % 2 === 0 ? 'F5F5F5' : 'FFFFFF';
      slide.addShape('rect', { x, y: curY, w, h: rowH, fill: { color: fill }, line: { type: 'none' } });
      slide.addText(r[0], { x: x + 0.07, y: curY, w: termW - 0.12, h: rowH, fontFace: F_MED, bold: true, fontSize: termFontSize, color: TEAL, valign: 'middle', lineSpacingMultiple: 0.95 });
      slide.addText(r[1], { x: x + termW, y: curY, w: defW - 0.1, h: rowH, fontFace: F_BODY, fontSize, color: BODY_C, valign: 'middle', lineSpacingMultiple: 1.0 });
      curY += rowH;
    });
    return curY;
  }

  function compareColumns(slide, { left, right, x = 0.6, y = 1.35, w = 8.75, h = 3.9, gap = 0.25 }) {
    const cw = (w - gap) / 2;
    [[left, x, TEAL], [right, x + cw + gap, BLUE]].forEach(([col, cx, defaultColor]) => {
      const color = col.color || defaultColor;
      slide.addShape('roundRect', { x: cx, y, w: cw, h, rectRadius: 0.06, fill: { color: WHITE }, line: { color: CARD_BORDER, width: 0.75 },
        shadow: { type: 'outer', color: '9A9A9A', opacity: 0.13, blur: 5, offset: 1.2, angle: 90 } });
      slide.addShape('roundRect', { x: cx, y, w: cw, h: 0.42, rectRadius: 0.05, fill: { color }, line: { type: 'none' } });
      slide.addShape('rect', { x: cx, y: y + 0.21, w: cw, h: 0.21, fill: { color }, line: { type: 'none' } });
      slide.addText(col.heading, { x: cx + 0.15, y, w: cw - 0.3, h: 0.42, fontFace: F_MED, bold: true, fontSize: 12, color: WHITE, valign: 'middle', lineSpacingMultiple: 1.0 });
      let cy = y + 0.58;
      if (col.intro) {
        slide.addText(col.intro, { x: cx + 0.15, y: cy, w: cw - 0.3, h: 0.9, fontFace: F_BODY, fontSize: 9.3, color: BODY_C, valign: 'top', lineSpacingMultiple: 1.12 });
        cy += 0.95;
      }
      const textW = cw - 0.3;
      const fontSize = 9.7;
      const charsPerLine = textW * 118 / fontSize;
      (col.rows || []).forEach(r => {
        const hasLabel = r[0] && r[0].length > 0;
        const fullLen = (hasLabel ? r[0].length + 2 : 0) + r[1].length;
        const lines = Math.max(1, Math.ceil(fullLen / charsPerLine));
        const rh = lines * (fontSize / 72) * 1.35 + 0.14;
        const paras = hasLabel
          ? [{ text: r[0] + ': ', options: { bold: true, fontFace: F_MED, fontSize, color: TITLE_C } }, { text: r[1], options: { fontFace: F_BODY, fontSize, color: BODY_C } }]
          : [{ text: r[1], options: { fontFace: F_BODY, fontSize, color: BODY_C } }];
        slide.addText(paras, { x: cx + 0.15, y: cy, w: textW, h: rh, valign: 'top', lineSpacingMultiple: 1.08 });
        cy += rh;
      });
    });
  }

  function weakStrong(slide, weakText, strongText, opts = {}) {
    const y = opts.y || 2.1;
    const cw = (8.75 - 0.3) / 2;
    const h = opts.h || 1.6;
    const labels = opts.labels || ['Weak', 'Strong'];
    slide.addShape('roundRect', { x: 0.6, y, w: cw, h, rectRadius: 0.06, fill: { color: 'FDEDED' }, line: { color: 'D9534F', width: 1 } });
    slide.addText('\u2715  ' + labels[0], { x: 0.8, y: y + 0.15, w: cw - 0.4, h: 0.3, fontFace: F_MED, bold: true, fontSize: 12, color: 'B03A3A' });
    slide.addText(weakText, { x: 0.8, y: y + 0.55, w: cw - 0.4, h: h - 0.7, fontFace: F_BODY, fontSize: 10.5, color: BODY_C, valign: 'top', lineSpacingMultiple: 1.2 });
    const cx2 = 0.6 + cw + 0.3;
    slide.addShape('roundRect', { x: cx2, y, w: cw, h, rectRadius: 0.06, fill: { color: 'EAF5F0' }, line: { color: '5B9E4E', width: 1 } });
    slide.addText('\u2713  ' + labels[1], { x: cx2 + 0.2, y: y + 0.15, w: cw - 0.4, h: 0.3, fontFace: F_MED, bold: true, fontSize: 12, color: '2E7D32' });
    slide.addText(strongText, { x: cx2 + 0.2, y: y + 0.55, w: cw - 0.4, h: h - 0.7, fontFace: F_BODY, fontSize: 10.5, color: BODY_C, valign: 'top', lineSpacingMultiple: 1.2 });
  }

  function flowDiagram(slide, { steps, x = 0.5, y = 1.9, boxW = 1.55, boxH = 0.85, gap = 0.25, colors = null, fontSize = 11.5, captionFontSize = 9.5 }) {
    const n = steps.length;
    const totalW = boxW * n + gap * (n - 1);
    let startX = x + (SLIDE_W - 1.0 - totalW) / 2;
    if (startX < 0.3) startX = 0.3;
    steps.forEach((step, i) => {
      const bx = startX + i * (boxW + gap);
      const color = colors ? colors[i % colors.length] : ACCENT_CYCLE[i % 3];
      slide.addShape('roundRect', { x: bx, y, w: boxW, h: boxH, rectRadius: 0.06, fill: { color: WHITE }, line: { color, width: 1.5 },
        shadow: { type: 'outer', color: '9A9A9A', opacity: 0.13, blur: 5, offset: 1.2, angle: 90 } });
      slide.addText(step.label, { x: bx + 0.06, y: y + 0.06, w: boxW - 0.12, h: boxH - 0.12, fontFace: F_MED, bold: true, fontSize, color: TITLE_C, align: 'center', valign: 'middle', lineSpacingMultiple: 1.0 });
      if (step.caption) {
        slide.addText(step.caption, { x: bx - 0.15, y: y + boxH + 0.08, w: boxW + 0.3, h: 0.55, fontFace: F_BODY, fontSize: captionFontSize, color: GRAY, align: 'center', valign: 'top', lineSpacingMultiple: 1.05 });
      }
      if (i < n - 1) {
        slide.addShape('rightArrow', { x: bx + boxW + 0.03, y: y + boxH / 2 - 0.09, w: gap - 0.06, h: 0.18, fill: { color: 'C9C9C9' }, line: { type: 'none' } });
      }
    });
  }

  // Safe line-shape helper: OOXML forbids zero/negative width or height on a
  // shape (PowerPoint flags the file as needing repair even though
  // LibreOffice silently tolerates it). ALWAYS use this instead of a raw
  // addShape('line', ...) call when either endpoint may vary at runtime.
  function safeLine(slide, x1, y1, x2, y2, lineOpts) {
    const w = Math.max(Math.abs(x2 - x1), 0.01);
    const h = Math.max(Math.abs(y2 - y1), 0.01);
    slide.addShape('line', { x: Math.min(x1, x2), y: Math.min(y1, y2), w, h, line: lineOpts, flipV: y2 < y1 });
  }
  function makeBranchDrawer(slide) {
    return function branch(x1, y1, x2, y2, color, width) { safeLine(slide, x1, y1, x2, y2, { color, width: width || 2 }); };
  }

  function moduleDivider(modNum, title, subtitle) {
    const s = newSlide();
    s.addText(modNum ? `Module ${modNum}` : '', { x: 0.6, y: 2.15, w: 8.75, h: 0.4, fontFace: F_MED, bold: true, fontSize: 14, color: TEAL, align: 'center' });
    s.addText(title, { x: 0.9, y: 2.6, w: 8.15, h: 0.95, fontFace: F_TITLE, bold: true, fontSize: 25, color: TITLE_C, align: 'center', lineSpacingMultiple: 1.08 });
    if (subtitle) s.addText(subtitle, { x: 1.2, y: 3.65, w: 7.6, h: 0.5, fontFace: F_BODY, italic: true, fontSize: 12.5, color: GRAY, align: 'center' });
    return s;
  }

  function quoteSlide(quoteText, caption) {
    const s = newSlide();
    s.addText('\u201c', { x: 1.0, y: 1.6, w: 1.0, h: 0.9, fontFace: F_TITLE, bold: true, fontSize: 60, color: TEAL });
    block(s, 1.3, 2.05, 7.4, quoteText, { h: 1.1, size: 21, bold: true, color: TITLE_C, align: 'left' });
    if (caption) block(s, 1.3, 3.35, 7.4, caption, { h: 0.4, size: 12, italic: true, color: GRAY });
    return s;
  }

  function useCaseSlide(title, parts, opts = {}) {
    const s = newSlide();
    tag(s, opts.tagText || 'Use Case', opts.tagColor || BLUE);
    titleOnly(s, title);
    const cw = (8.75 - 0.2 * 2) / 3;
    const labels = opts.labels || ['BACKGROUND', 'CHALLENGE', 'APPROACH', 'RECOMMENDED SOLUTION', 'EXPECTED OUTCOMES', 'LESSONS LEARNED'];
    const colColors = [TEAL, ORANGE, BLUE];
    labels.forEach((label, i) => {
      const col = i % 3, row = Math.floor(i / 3);
      const cx = 0.6 + col * (cw + 0.2);
      const cy = 1.55 + row * 1.95;
      const ch = 1.8;
      s.addShape('roundRect', { x: cx, y: cy, w: cw, h: ch, rectRadius: 0.05, fill: { color: CARD_BG }, line: { type: 'none' } });
      s.addShape('rect', { x: cx, y: cy, w: cw, h: 0.06, fill: { color: colColors[col] }, line: { type: 'none' } });
      s.addText(label, { x: cx + 0.12, y: cy + 0.1, w: cw - 0.24, h: 0.24, fontFace: F_MED, bold: true, fontSize: 8.5, color: colColors[col] });
      s.addText(parts[i] || '', { x: cx + 0.12, y: cy + 0.36, w: cw - 0.24, h: ch - 0.48, fontFace: F_BODY, fontSize: 8, color: BODY_C, valign: 'top', lineSpacingMultiple: 1.12 });
    });
    return s;
  }

  function activitySlide(title, contextLine, steps, expectedOutput, opts = {}) {
    const s = newSlide();
    tag(s, opts.tagText || 'Activity', opts.tagColor || ORANGE);
    titleOnly(s, title);
    if (contextLine) block(s, 0.6, 1.5, 8.75, contextLine, { h: 0.55, size: 10.5, italic: true, color: GRAY });
    const cw = (8.75 - 0.2 * 3) / 4;
    steps.forEach((step, i) => {
      const cx = 0.6 + i * (cw + 0.2);
      const cy = contextLine ? 2.2 : 1.7, ch = 2.0;
      s.addShape('roundRect', { x: cx, y: cy, w: cw, h: ch, rectRadius: 0.05, fill: { color: CARD_BG }, line: { color: 'E5E5E5', width: 0.75 } });
      s.addText(`STEP ${i + 1}`, { x: cx + 0.12, y: cy + 0.12, w: cw - 0.24, h: 0.24, fontFace: F_MED, bold: true, fontSize: 9, color: ORANGE });
      s.addText(step, { x: cx + 0.12, y: cy + 0.4, w: cw - 0.24, h: ch - 0.52, fontFace: F_BODY, fontSize: 9, color: BODY_C, valign: 'top', lineSpacingMultiple: 1.15 });
    });
    if (expectedOutput) {
      const y2 = 4.35;
      s.addShape('roundRect', { x: 0.6, y: y2, w: 8.75, h: 0.75, rectRadius: 0.06, fill: { color: 'F0F7F6' }, line: { color: TEAL, width: 0.75 } });
      s.addText([{ text: 'Expected output:  ', options: { bold: true, fontFace: F_MED, fontSize: 10, color: TITLE_C } }, { text: expectedOutput, options: { fontFace: F_BODY, fontSize: 10, color: BODY_C } }],
        { x: 0.8, y: y2, w: 8.35, h: 0.75, valign: 'middle', lineSpacingMultiple: 1.15 });
    }
    return s;
  }

  function maturityLevels(slide, levels, opts = {}) {
    let y = opts.y || 1.55;
    const rowH = opts.rowH || 0.62;
    const gap = 0.06;
    const colors = ['2A9D8F', '4FA88A', 'C9A227', 'D98A3D', 'C0524A'];
    levels.forEach((lvl, i) => {
      const color = colors[i % colors.length];
      slide.addShape('rect', { x: 0.6, y, w: 0.08, h: rowH, fill: { color }, line: { type: 'none' } });
      slide.addShape('rect', { x: 0.68, y, w: 8.67, h: rowH, fill: { color: CARD_BG }, line: { type: 'none' } });
      slide.addText(lvl.label, { x: 0.85, y: y + 0.06, w: 2.6, h: rowH - 0.12, fontFace: F_MED, bold: true, fontSize: 11.5, color: TITLE_C, valign: 'middle' });
      slide.addText(lvl.desc, { x: 3.5, y: y + 0.06, w: 5.75, h: rowH - 0.12, fontFace: F_BODY, fontSize: 10, color: BODY_C, valign: 'middle', lineSpacingMultiple: 1.1 });
      y += rowH + gap;
    });
    return y;
  }

  function labSlide(title, objective, scenario, tools, instructions, deliverables, evalCriteria, opts = {}) {
    const s = newSlide();
    tag(s, opts.tagText || 'Practical Lab', opts.tagColor || TEAL);
    titleOnly(s, title);
    let y = 1.5;
    const leftW = 4.35;
    let ly = block(s, 0.6, y, leftW, 'OBJECTIVE', { h: 0.22, size: 8.5, bold: true, color: TEAL });
    ly = block(s, 0.6, ly + 0.02, leftW, objective, { h: 0.55, size: 9, color: BODY_C }) + 0.08;
    ly = block(s, 0.6, ly, leftW, 'SCENARIO', { h: 0.22, size: 8.5, bold: true, color: TEAL });
    ly = block(s, 0.6, ly + 0.02, leftW, scenario, { h: 0.6, size: 9, color: BODY_C }) + 0.08;
    ly = block(s, 0.6, ly, leftW, 'TOOLS', { h: 0.22, size: 8.5, bold: true, color: TEAL });
    ly = checklistAuto(s, 0.6, ly + 0.06, leftW, [tools], { size: 9, gap: 0.06 });
    let ry = block(s, 5.2, y, 4.15, 'INSTRUCTIONS & ACTIVITIES', { h: 0.22, size: 8.5, bold: true, color: ORANGE });
    ry = numberedList(s, instructions, { x: 5.2, y: ry + 0.04, w: 4.15, size: 8.3, dia: 0.2, gap: 0.05 });
    const y2 = Math.max(ly, ry) + 0.1;
    const halfW = 4.15;
    let dy = block(s, 0.6, y2, halfW, 'DELIVERABLES', { h: 0.22, size: 8.5, bold: true, color: BLUE });
    checklistAuto(s, 0.6, dy + 0.04, halfW, deliverables, { size: 8.3, gap: 0.06 });
    let ey = block(s, 5.2, y2, halfW, 'EVALUATION CRITERIA', { h: 0.22, size: 8.5, bold: true, color: BLUE });
    checklistAuto(s, 5.2, ey + 0.04, halfW, evalCriteria, { size: 8.3, gap: 0.06 });
    return s;
  }

  function solutionGuideSlide(title, items, opts = {}) {
    const s = newSlide();
    tag(s, opts.tagText || 'Solution Guide', opts.tagColor || TEAL);
    titleOnly(s, title);
    checkItemsAuto(s, 0.6, 1.65, 8.75, items, { size: 11.5, gap: 0.2 });
    return s;
  }

  function kpiStatGrid(title, stats, opts = {}) {
    const s = newSlide();
    titleOnly(s, title);
    if (opts.subtitle) block(s, 0.6, 1.45, 8.75, opts.subtitle, { h: 0.3, size: 11, italic: true, color: GRAY });
    const n = stats.length;
    const gap = 0.25;
    const cw = (8.75 - gap * (n - 1)) / n;
    const cy = opts.subtitle ? 1.95 : 1.7, ch = 1.9;
    stats.forEach((st, i) => {
      const cx = 0.6 + i * (cw + gap);
      s.addShape('roundRect', { x: cx, y: cy, w: cw, h: ch, rectRadius: 0.07, fill: { color: TITLE_C }, line: { type: 'none' } });
      s.addText(st.value, { x: cx + 0.1, y: cy + 0.25, w: cw - 0.2, h: 0.8, fontFace: F_TITLE, bold: true, fontSize: st.value.length > 5 ? 22 : 30, color: WHITE, align: 'center', valign: 'middle' });
      s.addText(st.label, { x: cx + 0.15, y: cy + 1.05, w: cw - 0.3, h: ch - 1.15, fontFace: F_BODY, fontSize: 9.5, color: 'CFE3E0', align: 'center', valign: 'top', lineSpacingMultiple: 1.15 });
    });
    return s;
  }

  function knowledgeCheckSlide(title, scenario, question, options, opts = {}) {
    const s = newSlide();
    tag(s, opts.tagText || 'Knowledge Check', opts.tagColor || ORANGE);
    titleOnly(s, title);
    let y = 1.5;
    if (scenario) {
      s.addShape('roundRect', { x: 0.6, y, w: 8.75, h: 0.85, rectRadius: 0.05, fill: { color: CARD_BG }, line: { type: 'none' } });
      s.addText([{ text: 'Scenario: ', options: { bold: true, fontFace: F_MED, fontSize: 9.5, color: TITLE_C } }, { text: scenario, options: { fontFace: F_BODY, fontSize: 9.5, color: BODY_C } }],
        { x: 0.8, y: y + 0.08, w: 8.35, h: 0.7, valign: 'top', lineSpacingMultiple: 1.2 });
      y += 1.0;
    }
    y = block(s, 0.6, y, 8.75, question, { h: 0.45, size: 12.5, bold: true, color: TITLE_C });
    y += 0.1;
    checklistAuto(s, 0.6, y, 8.5, options, { size: 10.5, gap: 0.16 });
    return s;
  }

  function answerRationaleSlide(title, correctAnswer, why, wrongAnswers, opts = {}) {
    const s = newSlide();
    tag(s, opts.tagText || 'Answer & Rationale', opts.tagColor || TEAL);
    titleOnly(s, title);
    let y = 1.55;
    s.addShape('roundRect', { x: 0.6, y, w: 8.75, h: 0.55, rectRadius: 0.06, fill: { color: 'EAF5F0' }, line: { color: '5B9E4E', width: 0.75 } });
    s.addText('\u2713  ' + correctAnswer, { x: 0.8, y, w: 8.35, h: 0.55, fontFace: F_MED, bold: true, fontSize: 11, color: '2E7D32', valign: 'middle' });
    y += 0.7;
    s.addText([{ text: 'Why it is correct: ', options: { bold: true, fontFace: F_MED, fontSize: 10.5, color: TITLE_C } }, { text: why, options: { fontFace: F_BODY, fontSize: 10.5, color: BODY_C } }],
      { x: 0.6, y, w: 8.75, h: 0.85, valign: 'top', lineSpacingMultiple: 1.2 });
    y += 0.95;
    (wrongAnswers || []).forEach(w => {
      const size = 9.7;
      const charsPerLine = 8.5 * 118 / size;
      const lines = Math.max(1, Math.ceil(w.length / charsPerLine));
      const rowH = lines * (size / 72) * 1.3 + 0.1;
      s.addText('\u2715', { x: 0.6, y, w: 0.25, h: rowH, fontFace: F_MED, bold: true, fontSize: size, color: 'C0524A', valign: 'top' });
      s.addText(w, { x: 0.85, y, w: 8.4, h: rowH, fontFace: F_BODY, fontSize: size, color: BODY_C, valign: 'top', lineSpacingMultiple: 1.15 });
      y += rowH;
    });
    return s;
  }

  function resourcesSlide(title, resources, opts = {}) {
    const s = newSlide();
    tag(s, opts.tagText || 'Curated Resources', opts.tagColor || BLUE);
    titleOnly(s, title);
    let y = 1.6;
    resources.forEach(r => {
      s.addShape('roundRect', { x: 0.6, y, w: 8.75, h: 0.92, rectRadius: 0.06, fill: { color: CARD_BG }, line: { color: 'E5E5E5', width: 0.75 } });
      s.addShape('ellipse', { x: 0.8, y: y + 0.28, w: 0.36, h: 0.36, fill: { color: BLUE }, line: { type: 'none' } });
      s.addText('\u25B6', { x: 0.8, y: y + 0.28, w: 0.36, h: 0.36, fontFace: F_BODY, fontSize: 12, color: WHITE, align: 'center', valign: 'middle' });
      s.addText(r.title, { x: 1.35, y: y + 0.1, w: 7.85, h: 0.3, fontFace: F_MED, bold: true, fontSize: 11.5, color: TITLE_C });
      s.addText(r.source || '', { x: 1.35, y: y + 0.4, w: 7.85, h: 0.24, fontFace: F_BODY, fontSize: 9, color: GRAY });
      s.addText(r.url || '', { x: 1.35, y: y + 0.63, w: 7.85, h: 0.22, fontFace: F_BODY, fontSize: 8.5, color: BLUE });
      y += 1.08;
    });
    return s;
  }

  function formulaSlide(title, formula, plainEnglish, notes, opts = {}) {
    const s = newSlide();
    titleOnly(s, title);
    s.addShape('roundRect', { x: 1.0, y: 1.6, w: 7.75, h: 1.0, rectRadius: 0.07, fill: { color: TITLE_C }, line: { type: 'none' } });
    s.addText(formula, { x: 1.2, y: 1.6, w: 7.35, h: 1.0, fontFace: F_TITLE, bold: true, fontSize: 18, color: WHITE, align: 'center', valign: 'middle' });
    let y = 2.85;
    s.addShape('roundRect', { x: 0.6, y, w: 8.75, h: 0.6, rectRadius: 0.06, fill: { color: 'F0F7F6' }, line: { color: TEAL, width: 0.75 } });
    s.addText([{ text: 'In plain English:  ', options: { bold: true, fontFace: F_MED, fontSize: 10.5, color: TITLE_C } }, { text: plainEnglish, options: { italic: true, fontFace: F_BODY, fontSize: 10.5, color: BODY_C } }],
      { x: 0.8, y, w: 8.35, h: 0.6, valign: 'middle', lineSpacingMultiple: 1.15 });
    y += 0.8;
    if (notes && notes.length) checklistAuto(s, 0.6, y, 8.5, notes, { size: 10.5, gap: 0.16 });
    return s;
  }

  function mcqSlide(slide, { title, question, options }) {
    titleOnly(slide, title);
    slide.addShape('roundRect', { x: 0.6, y: 1.35, w: 8.75, h: 1.15, rectRadius: 0.06, fill: { color: 'F4F0E7' }, line: { type: 'none' } });
    slide.addText(question, { x: 0.85, y: 1.48, w: 8.25, h: 0.95, fontFace: F_MED, bold: true, fontSize: 13, color: TITLE_C, valign: 'top', lineSpacingMultiple: 1.15 });
    const letters = ['A', 'B', 'C', 'D', 'E', 'F'];
    let y = 2.7;
    options.forEach((opt, i) => {
      const d = 0.32;
      slide.addShape('ellipse', { x: 0.6, y, w: d, h: d, fill: { color: WHITE }, line: { color: TEAL, width: 1.5 } });
      slide.addText(letters[i], { x: 0.6, y, w: d, h: d, fontFace: F_MED, bold: true, fontSize: 12, color: TEAL, align: 'center', valign: 'middle' });
      slide.addText(opt, { x: 1.08, y: y - 0.03, w: 8.25, h: 0.55, fontFace: F_BODY, fontSize: 11.5, color: BODY_C, valign: 'top', lineSpacingMultiple: 1.1 });
      y += 0.58;
    });
  }

  function caseStudySlide(slide, { title = 'Case Study', context, twist, questions }) {
    titleOnly(slide, title);
    const y0 = 1.35;
    slide.addShape('roundRect', { x: 0.6, y: y0, w: 8.75, h: 1.15, rectRadius: 0.06, fill: { color: 'E3F3F1' }, line: { type: 'none' } });
    slide.addText([{ text: 'CONTEXT   ', options: { bold: true, color: TEAL, fontFace: F_MED, fontSize: 10.5 } }, { text: context, options: { color: BODY_C, fontFace: F_BODY, fontSize: 11.5 } }],
      { x: 0.85, y: y0 + 0.12, w: 8.25, h: 0.95, valign: 'top', lineSpacingMultiple: 1.15 });
    const y1 = y0 + 1.3;
    slide.addShape('roundRect', { x: 0.6, y: y1, w: 8.75, h: 1.15, rectRadius: 0.06, fill: { color: 'FDE8E1' }, line: { type: 'none' } });
    slide.addText([{ text: 'THE TWIST   ', options: { bold: true, color: ORANGE, fontFace: F_MED, fontSize: 10.5 } }, { text: twist, options: { color: BODY_C, fontFace: F_BODY, fontSize: 11.5 } }],
      { x: 0.85, y: y1 + 0.12, w: 8.25, h: 0.95, valign: 'top', lineSpacingMultiple: 1.15 });
    const y2 = y1 + 1.32;
    const qText = questions.map(q => ({ text: q, options: { bullet: { code: '25B8', indent: 14 }, bold: true, fontFace: F_MED, fontSize: 12, color: TITLE_C, breakLine: true, paraSpaceAfter: 5 } }));
    slide.addText(qText, { x: 0.6, y: y2, w: 8.75, h: 1.0, valign: 'top' });
    logBullets(questions.length, TITLE_C);
  }

  function recapSlide(slide, { title = 'Recap', takeaways, action }) {
    titleOnly(slide, title);
    slide.addText('KEY TAKEAWAYS', { x: 0.6, y: 1.35, w: 8, h: 0.3, fontFace: F_MED, bold: true, fontSize: 11, color: TEAL, charSpacing: 0.8 });
    const items = takeaways.map(t => ({ text: t, options: { bullet: { code: '25CF', indent: 14 }, fontFace: F_BODY, fontSize: 13, color: BODY_C, breakLine: true, paraSpaceAfter: 12 } }));
    slide.addText(items, { x: 0.6, y: 1.75, w: 8.75, h: 1.8, valign: 'top', lineSpacingMultiple: 1.15 });
    logBullets(takeaways.length, TEAL);
    if (action) {
      slide.addShape('roundRect', { x: 0.6, y: 3.65, w: 8.75, h: 1.15, rectRadius: 0.06, fill: { color: 'E9ECFC' }, line: { type: 'none' } });
      slide.addText([{ text: 'ACTION   ', options: { bold: true, color: BLUE, fontFace: F_MED, fontSize: 10.5 } }, { text: action, options: { color: BODY_C, fontFace: F_BODY, fontSize: 12 } }],
        { x: 0.85, y: 3.78, w: 8.25, h: 0.9, valign: 'top', lineSpacingMultiple: 1.15 });
    }
  }

  function cycleWheel(slide, { center, outer, cx = 4.9, cy = 3.55, radius = 1.35, outerR = 0.95, centerR = 0.72 }) {
    const n = outer.length;
    const positions = outer.map((_, i) => {
      const angle = (-90 + i * (360 / n)) * Math.PI / 180;
      return [cx + radius * Math.cos(angle), cy + radius * Math.sin(angle)];
    });
    outer.forEach((item, i) => {
      const [px, py] = positions[i];
      slide.addShape('ellipse', { x: px - outerR / 2, y: py - outerR / 2, w: outerR, h: outerR, fill: { color: ACCENT_CYCLE[i % ACCENT_CYCLE.length], transparency: 15 }, line: { color: WHITE, width: 2 } });
      slide.addText([{ text: item.heading + '\n', options: { bold: true, fontFace: F_MED, fontSize: 11, color: '262626' } }, { text: item.body || '', options: { fontFace: F_BODY, fontSize: 8, color: '262626' } }],
        { x: px - outerR / 2 + 0.1, y: py - outerR / 2, w: outerR - 0.2, h: outerR, align: 'center', valign: 'middle', lineSpacingMultiple: 0.98 });
    });
    slide.addShape('ellipse', { x: cx - centerR / 2, y: cy - centerR / 2, w: centerR, h: centerR, fill: { color: 'F0B429' }, line: { color: WHITE, width: 2.5 },
      shadow: { type: 'outer', color: '9A9A9A', opacity: 0.25, blur: 6, offset: 2, angle: 90 } });
    slide.addText(center, { x: cx - centerR / 2, y: cy - centerR / 2, w: centerR, h: centerR, fontFace: F_MED, bold: true, fontSize: 9, color: '262626', align: 'center', valign: 'middle', lineSpacingMultiple: 0.95 });
  }

  function coverSlide({ eyebrow: eb, title, titleAccent, tagline, byline, pills }) {
    const s = newSlide();
    let y = 1.9;
    if (eb) { s.addText(eb, { x: 0.6, y: 1.55, w: 8.75, h: 0.3, fontFace: F_MED, bold: true, fontSize: 10.5, color: TEAL, align: 'center' }); }
    s.addText(title, { x: 0.6, y, w: 8.75, h: 0.55, fontFace: F_TITLE, bold: true, fontSize: 28, color: TITLE_C, align: 'center' });
    y += 0.52;
    if (titleAccent) { s.addText(titleAccent, { x: 0.6, y, w: 8.75, h: 0.45, fontFace: F_TITLE, bold: true, fontSize: 22, color: TEAL, align: 'center' }); y += 0.5; }
    if (tagline) { s.addText(tagline, { x: 0.9, y: y + 0.1, w: 8.15, h: 0.4, fontFace: F_BODY, italic: true, fontSize: 13, color: GRAY, align: 'center' }); y += 0.6; }
    if (byline) { s.addText(byline, { x: 0.9, y: y + 0.1, w: 8.15, h: 0.35, fontFace: F_MED, bold: true, fontSize: 11.5, color: BODY_C, align: 'center' }); y += 0.5; }
    if (pills && pills.length) {
      const tw = 1.55, gap = 0.2;
      let tx = (10 - (tw * pills.length + gap * (pills.length - 1))) / 2;
      pills.forEach(p => {
        s.addShape('roundRect', { x: tx, y: y + 0.15, w: tw, h: 0.42, rectRadius: 0.21, fill: { color: WHITE }, line: { color: TEAL, width: 1 } });
        s.addText(p, { x: tx, y: y + 0.15, w: tw, h: 0.42, fontFace: F_MED, bold: true, fontSize: 10.5, color: TEAL, align: 'center', valign: 'middle' });
        tx += tw + gap;
      });
    }
    return s;
  }

  function tocSlide(title, items, opts = {}) {
    const s = newSlide();
    titleOnly(s, title);
    const cols = opts.cols || (items.length > 5 ? 2 : 1);
    const colW = cols === 2 ? 4.25 : 8.75;
    const gap = 0.25;
    const rowsPerCol = Math.ceil(items.length / cols);
    items.forEach((it, i) => {
      const col = Math.floor(i / rowsPerCol), row = i % rowsPerCol;
      const cx = 0.6 + col * (colW + gap);
      const cy = 1.7 + row * 0.85;
      s.addShape('ellipse', { x: cx, y: cy, w: 0.42, h: 0.42, fill: { color: TEAL }, line: { type: 'none' } });
      s.addText(String(i + 1).padStart(2, '0'), { x: cx, y: cy, w: 0.42, h: 0.42, fontFace: F_MED, bold: true, fontSize: 12, color: WHITE, align: 'center', valign: 'middle' });
      s.addText(it, { x: cx + 0.55, y: cy, w: colW - 0.55, h: 0.42, fontFace: F_BODY, fontSize: 11.5, color: BODY_C, valign: 'middle', lineSpacingMultiple: 1.1 });
    });
    return s;
  }

  function bulletSlide(title, items, opts = {}) {
    const s = newSlide();
    titleOnly(s, title);
    if (opts.subtitle) block(s, 0.6, 1.45, 8.75, opts.subtitle, { h: 0.3, size: 11, italic: true, color: GRAY });
    checklistAuto(s, 0.6, opts.subtitle ? 1.95 : 1.75, 8.5, items, { size: opts.size || 13, gap: opts.gap != null ? opts.gap : 0.3 });
    return s;
  }

  function imageSlide(title, imagePath, opts = {}) {
    const s = newSlide();
    if (title) titleOnly(s, title);
    const y = title ? 1.4 : 0.9;
    const maxW = 8.75, maxH = SLIDE_H - y - 0.5;
    s.addImage({ path: imagePath, x: 0.6, y, w: maxW, h: maxH, sizing: { type: 'contain', w: maxW, h: maxH } });
    return s;
  }

  return {
    pres, newSlide, block, titleOnly, eyebrow, tag,
    checklistAuto, checkItemsAuto, numberedList,
    bulletCard, cardGrid, twoColCards, numberedStepCards, gridCards,
    customTable, glossaryTable, compareColumns, weakStrong,
    flowDiagram, safeLine, makeBranchDrawer,
    moduleDivider, quoteSlide, useCaseSlide, activitySlide, maturityLevels,
    labSlide, solutionGuideSlide, kpiStatGrid, knowledgeCheckSlide, answerRationaleSlide,
    resourcesSlide, formulaSlide, mcqSlide, caseStudySlide, recapSlide, cycleWheel,
    coverSlide, tocSlide, bulletSlide, imageSlide,
    writeBulletColorQueue: (jsonPath) => fs.writeFileSync(jsonPath, JSON.stringify(bulletColorQueue, null, 2)),
    getBulletColorQueue: () => bulletColorQueue,
  };
}

module.exports = {
  createDeck,
  TEAL, BLUE, ORANGE, PURPLE, TITLE_C, BODY_C, GRAY, WHITE, CARD_BORDER, CARD_BG,
  F_TITLE, F_BODY, F_MED, SLIDE_W, SLIDE_H, ACCENT_CYCLE,
};
