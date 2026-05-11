'use strict';
const PptxGenJS = require('pptxgenjs');
const fs        = require('fs');
const path      = require('path');

// ── Paths ─────────────────────────────────────────────────────────────────────
const ROOT   = process.cwd();
const DATA   = path.join(ROOT, '.tmp', 'data');
const CHARTS = path.join(ROOT, '.tmp', 'charts');
const THUMBS = path.join(ROOT, '.tmp', 'thumbnails');
const OUTPUT = path.join(ROOT, 'output');

// ── Load data ─────────────────────────────────────────────────────────────────
const videoData = JSON.parse(fs.readFileSync(path.join(DATA, 'videos.json'),     'utf8'));
const channels  = JSON.parse(fs.readFileSync(path.join(DATA, 'channels.json'),   'utf8')).channels;
const insights  = JSON.parse(fs.readFileSync(path.join(DATA, 'insights.json'),   'utf8'));
const thumbData = JSON.parse(fs.readFileSync(path.join(DATA, 'thumbnails.json'), 'utf8')).thumbnails;

const videos = videoData.videos;
const topic  = videoData.topic;

// ── Colors (no # prefix) ─────────────────────────────────────────────────────
const C = {
  BG:     '0F1117',
  CARD:   '1A1D27',
  CARD2:  '151821',
  BORDER: '2E3343',
  TEXT:   'E8E9ED',
  MUTED:  '8B8FA3',
  VIOLET: '6C5CE7',
  CYAN:   '00D2D3',
  CORAL:  'FF6B81',
  GOLD:   'FECA57',
  GREEN:  '1DD1A1',
  BLACK:  '000000',
};

// ── Layout: LAYOUT_16x9 = 10" x 5.625" ───────────────────────────────────────
const W = 10;
const H = 5.625;

// ── Helpers ───────────────────────────────────────────────────────────────────
const fmt = n => {
  if (n >= 1e9) return (n / 1e9).toFixed(1) + 'B';
  if (n >= 1e6) return (n / 1e6).toFixed(1) + 'M';
  if (n >= 1e3) return Math.round(n / 1e3) + 'K';
  return String(n);
};

const trunc = (s, n) => s && s.length > n ? s.slice(0, n) + '…' : (s || '');

const todayStr = () => new Date().toLocaleDateString('en-US', {
  year: 'numeric', month: 'long', day: 'numeric'
});

const slugDate = () => {
  const d = new Date();
  return [
    d.getFullYear(),
    String(d.getMonth() + 1).padStart(2, '0'),
    String(d.getDate()).padStart(2, '0'),
    String(d.getHours()).padStart(2, '0'),
    String(d.getMinutes()).padStart(2, '0'),
  ].join('');
};

const fill  = color             => ({ color });
const fillT = (color, pct)      => ({ color, transparency: pct });
const line  = (color, w = 0.75) => ({ color, width: w });
const noLine = ()               => ({ color: '000000', width: 0, transparency: 100 });

// ── Primitives ────────────────────────────────────────────────────────────────

function setBg(slide) {
  slide.background = { color: C.BG };
}

function card(slide, x, y, w, h, opts = {}) {
  const { bg = C.CARD, border = C.BORDER, accent = null } = opts;
  slide.addShape('rect', { x, y, w, h, fill: fill(bg), line: line(border, 0.75) });
  if (accent) {
    slide.addShape('rect', { x, y, w: 0.06, h, fill: fill(accent), line: noLine() });
  }
}

function topBar(slide, x, y, w, color) {
  slide.addShape('rect', { x, y, w, h: 0.055, fill: fill(color), line: noLine() });
}

function tag(slide, x, y, label, color = C.VIOLET) {
  const tw = label.length * 0.072 + 0.22;
  slide.addShape('rect', { x, y, w: tw, h: 0.19, fill: fillT(color, 80), line: line(color, 0.5) });
  slide.addText(label.toUpperCase(), {
    x, y, w: tw, h: 0.19,
    fontSize: 6.5, bold: true, color, align: 'center', valign: 'middle',
    fontFace: 'Trebuchet MS',
  });
}

function heading(slide, x, y, w, text, size = 20) {
  slide.addText(text, {
    x, y, w, h: 0.38,
    fontSize: size, bold: true, color: C.TEXT,
    fontFace: 'Trebuchet MS', valign: 'middle',
  });
}

function sublabel(slide, x, y, w, text) {
  slide.addText(text.toUpperCase(), {
    x, y, w, h: 0.18,
    fontSize: 6.5, bold: true, color: C.MUTED,
    fontFace: 'Trebuchet MS',
  });
}

function body(slide, x, y, w, h, text, opts = {}) {
  const { size = 10, color = C.TEXT, align = 'left', bold = false,
          italic = false, mono = false } = opts;
  slide.addText(text || '', {
    x, y, w, h,
    fontSize: size, bold, italic, color, align,
    fontFace: mono ? 'Consolas' : 'Calibri',
    valign: 'top', wrap: true,
  });
}

function statBlock(slide, x, y, value, label, color = C.CYAN) {
  slide.addText(value, {
    x, y, w: 2.2, h: 0.42,
    fontSize: 22, bold: true, color,
    fontFace: 'Consolas',
  });
  slide.addText(label, {
    x, y: y + 0.4, w: 2.2, h: 0.2,
    fontSize: 8, color: C.MUTED, fontFace: 'Calibri',
  });
}

function chartImg(slide, filename, x, y, w, h) {
  const p = path.join(CHARTS, filename);
  if (fs.existsSync(p)) slide.addImage({ path: p, x, y, w, h });
}

function thumbImg(slide, thumbPath, x, y, w, h) {
  if (thumbPath && fs.existsSync(thumbPath)) {
    slide.addImage({ path: thumbPath, x, y, w, h });
  }
}

// ── Slides ────────────────────────────────────────────────────────────────────

// 1. Title
function slideTitle(pptx) {
  const s = pptx.addSlide();
  setBg(s);

  s.addShape('ellipse', { x: 6.8, y: -1.1, w: 4.2, h: 4.2,
    fill: fillT(C.VIOLET, 91), line: noLine() });
  s.addShape('ellipse', { x: -0.8, y: 3.8, w: 2.4, h: 2.4,
    fill: fillT(C.CYAN, 92), line: noLine() });
  s.addShape('rect', { x: 7.5, y: 2.8, w: 2.8, h: 1.0, rotate: 18,
    fill: fillT(C.GOLD, 93), line: noLine() });

  tag(s, 0.55, 0.95, 'YOUTUBE SPACE ANALYSIS', C.VIOLET);

  s.addText(topic.toUpperCase(), {
    x: 0.5, y: 1.25, w: 8.2, h: 1.5,
    fontSize: 40, bold: true, color: C.TEXT,
    fontFace: 'Trebuchet MS', wrap: true,
  });

  s.addShape('rect', { x: 0.5, y: 2.9, w: 1.1, h: 0.05,
    fill: fill(C.VIOLET), line: noLine() });

  s.addText(todayStr(), {
    x: 0.5, y: 3.05, w: 5, h: 0.25,
    fontSize: 10.5, color: C.MUTED, fontFace: 'Calibri',
  });

  const totalSubs = channels.reduce((sum, c) => sum + (c.subscribers || 0), 0);
  const stats = [
    { v: fmt(videos.length),    l: 'Videos Analyzed',  c: C.CYAN  },
    { v: fmt(channels.length),  l: 'Channels Profiled', c: C.GOLD  },
    { v: fmt(totalSubs),        l: 'Combined Subs',     c: C.GREEN },
  ];
  stats.forEach((st, i) => statBlock(s, 0.5 + i * 2.6, 3.75, st.v, st.l, st.c));
}

// 2. Market Opportunity
function slideMarketOpportunity(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'MARKET OPPORTUNITY', C.GOLD);
  heading(s, 0.4, 0.38, 9, 'Is This Space Worth Entering?', 20);

  const summ    = insights.executive_summary || {};
  const entry   = insights.entry_strategy    || {};
  const verdict = entry.market_verdict       || summ.space_overview || '';
  const bullets = summ.executive_summary_bullets || [];
  const accs    = [C.CYAN, C.VIOLET, C.GOLD, C.GREEN];

  // Verdict card — big and prominent
  card(s, 0.4, 0.86, 9.25, 0.82, { accent: C.GOLD, bg: '1C1A10' });
  body(s, 0.62, 0.94, 8.9, 0.66, verdict, { size: 11.5, bold: true, color: C.GOLD });

  // 4 punchy insight bullets
  bullets.slice(0, 4).forEach((b, i) => {
    const y = 1.82 + i * 0.88;
    card(s, 0.4, y, 9.25, 0.74, { accent: accs[i] });
    body(s, 0.62, y + 0.1, 8.9, 0.58, b, { size: 10 });
  });
}

// 3. Who's Winning & Why
function slideWhoIsWinning(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'COMPETITIVE LANDSCAPE', C.CORAL);
  heading(s, 0.4, 0.38, 9, "Who's Winning & Why", 20);

  chartImg(s, 'top_channels.png', 0.25, 0.84, 5.0, 3.0);

  const ca         = insights.competitive_analysis || {};
  const archetypes = ca.channel_archetypes || [];
  const archColors = [C.VIOLET, C.CYAN, C.GOLD];

  // Market concentration strip
  card(s, 0.25, 3.98, 4.8, 1.45, { accent: C.GREEN });
  sublabel(s, 0.42, 4.08, 4.4, 'Market Concentration');
  body(s, 0.42, 4.3, 4.5, 1.05, ca.market_concentration || '', { size: 9.5 });

  // 3 archetype cards on right
  archetypes.slice(0, 3).forEach((a, i) => {
    const y = 0.84 + i * 1.58;
    card(s, 5.5, y, 4.25, 1.44, { accent: archColors[i] });
    s.addText(a.archetype || '', {
      x: 5.66, y: y + 0.08, w: 3.9, h: 0.28,
      fontSize: 10, bold: true, color: archColors[i], fontFace: 'Trebuchet MS',
    });
    body(s, 5.66, y + 0.38, 3.9, 0.32,
      (a.channels || []).join(', '), { size: 8.5, color: C.MUTED });
    body(s, 5.66, y + 0.7, 3.9, 0.3, a.strategy || '', { size: 9 });
    body(s, 5.66, y + 1.04, 3.9, 0.3,
      `Exploit: ${a.weakness || ''}`, { size: 8.5, color: C.CORAL });
  });
}

// 4. What Content Works
function slideWhatContentWorks(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'CONTENT MAP', C.CYAN);
  heading(s, 0.4, 0.38, 9, 'What Content Works', 20);

  chartImg(s, 'video_length.png', 0.25, 0.84, 5.9, 3.2);

  const themes    = (insights.title_analysis || {}).dominant_content_themes || [];
  const themeAccs = [C.CYAN, C.CORAL, C.GOLD];

  sublabel(s, 6.38, 0.84, 3.37, 'Top Content Themes');
  themes.slice(0, 3).forEach((t, i) => {
    const y = 1.06 + i * 1.05;
    card(s, 6.38, y, 3.37, 0.9, { accent: themeAccs[i] });
    s.addText(String(t.percentage || '?') + '%', {
      x: 6.52, y: y + 0.1, w: 0.75, h: 0.38,
      fontSize: 16, bold: true, color: themeAccs[i], fontFace: 'Consolas',
    });
    body(s, 7.36, y + 0.1, 2.22, 0.28, t.theme || '', { size: 10, bold: true });
    body(s, 7.36, y + 0.4, 2.22, 0.32,
      `Avg: ${t.avg_views || '?'} · e.g. "${trunc(t.top_example || '', 28)}"`,
      { size: 8, color: C.MUTED });
  });

  // Key takeaway strip at bottom
  const corr = (insights.title_analysis || {}).view_correlation_insights || '';
  card(s, 0.25, 4.18, 9.5, 1.22, { bg: '121820', border: C.CYAN, accent: C.CYAN });
  sublabel(s, 0.42, 4.28, 9.1, 'Key Takeaway');
  body(s, 0.42, 4.48, 9.1, 0.84, corr, { size: 9.5, italic: true });
}

// 5. Title & Hook Playbook
function slideTitleHooks(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'STRATEGY', C.VIOLET);
  heading(s, 0.4, 0.38, 9, 'Title & Hook Playbook', 20);

  const ta       = insights.title_analysis || {};
  const hooks    = ta.top_hook_patterns        || [];
  const keywords = ta.high_performing_keywords || [];
  const overused = ta.overused_patterns        || [];

  card(s, 0.4, 0.86, 4.45, 3.38);
  sublabel(s, 0.56, 0.96, 4.1, 'Top Hook Patterns');
  hooks.slice(0, 5).forEach((h, i) => {
    s.addText(`${i + 1}.`, {
      x: 0.56, y: 1.2 + i * 0.48, w: 0.28, h: 0.36,
      fontSize: 10, bold: true, color: C.VIOLET, fontFace: 'Consolas',
    });
    body(s, 0.86, 1.2 + i * 0.48, 3.82, 0.4, h, { size: 10 });
  });

  card(s, 5.1, 0.86, 4.55, 3.38);
  sublabel(s, 5.26, 0.96, 4.2, 'High-Performing Keywords');
  let cx = 5.26, cy = 1.2;
  keywords.slice(0, 12).forEach(kw => {
    const cw = Math.max(0.65, kw.length * 0.076 + 0.22);
    if (cx + cw > 9.4) { cx = 5.26; cy += 0.42; }
    s.addShape('rect', { x: cx, y: cy, w: cw, h: 0.28,
      fill: fillT(C.CYAN, 82), line: line(C.CYAN, 0.5) });
    s.addText(kw, { x: cx, y: cy, w: cw, h: 0.28,
      fontSize: 8.5, color: C.CYAN, align: 'center', valign: 'middle', fontFace: 'Calibri' });
    cx += cw + 0.1;
  });

  if (overused.length) {
    sublabel(s, 5.26, 2.88, 4.2, 'Oversaturated — Avoid');
    overused.slice(0, 3).forEach((o, i) => {
      body(s, 5.26, 3.1 + i * 0.38, 4.2, 0.32, `× ${o}`, { size: 9.5, color: C.CORAL });
    });
  }

  // Bottom posting insight
  const posting = (insights.executive_summary || {}).best_posting_insight || '';
  card(s, 0.4, 4.38, 9.25, 0.98, { bg: '1A1227', border: C.VIOLET, accent: C.VIOLET });
  sublabel(s, 0.62, 4.48, 9.0, 'Posting Insight');
  body(s, 0.62, 4.68, 8.9, 0.62, posting, { size: 9.5, italic: true });
}

// 6. Thumbnail Playbook
function slideThumbnailPlaybook(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'VISUAL PLAYBOOK', C.CORAL);
  heading(s, 0.4, 0.38, 9, 'Thumbnail Playbook', 20);

  const shown = thumbData.filter(t => fs.existsSync(t.path)).slice(0, 6);
  const TW = 1.82, TH = 1.1;
  shown.forEach((t, i) => {
    const row = Math.floor(i / 3), col = i % 3;
    const ix = 0.28 + col * (TW + 0.1);
    const iy = 0.84 + row * (TH + 0.1);
    card(s, ix - 0.05, iy - 0.05, TW + 0.1, TH + 0.1);
    thumbImg(s, t.path, ix, iy, TW, TH);
  });

  const ta   = insights.thumbnail_analysis || {};
  const recs = ta.design_recommendations  || [];
  const recColors = [C.CORAL, C.GOLD, C.GREEN];

  sublabel(s, 6.38, 0.84, 3.37, 'Design Rules');

  // Visual style card
  card(s, 6.38, 1.06, 3.37, 0.72, { accent: C.CYAN });
  sublabel(s, 6.55, 1.16, 3.1, 'Visual Style');
  body(s, 6.55, 1.36, 3.1, 0.34, ta.dominant_visual_style || '', { size: 9 });

  // Color trends card
  card(s, 6.38, 1.86, 3.37, 0.72, { accent: C.GOLD });
  sublabel(s, 6.55, 1.96, 3.1, 'Color Trends');
  body(s, 6.55, 2.16, 3.1, 0.34,
    (ta.color_palette_trends || []).slice(0, 2).join(' · '), { size: 9 });

  // Face/character usage card
  card(s, 6.38, 2.66, 3.37, 0.72, { accent: C.VIOLET });
  sublabel(s, 6.55, 2.76, 3.1, 'Faces & Characters');
  body(s, 6.55, 2.96, 3.1, 0.34, ta.face_character_usage || '', { size: 9 });

  // Design recs
  recs.slice(0, 2).forEach((r, i) => {
    card(s, 6.38, 3.46 + i * 0.9, 3.37, 0.76, { accent: recColors[i] });
    body(s, 6.55, 3.56 + i * 0.9, 3.1, 0.58, r, { size: 9 });
  });
}

// 7. Upload Strategy
function slideUploadStrategy(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'TIMING', C.GREEN);
  heading(s, 0.4, 0.38, 9, 'Upload Strategy', 20);

  chartImg(s, 'upload_cadence.png', 0.25, 0.84, 6.1, 4.62);

  const posting = (insights.executive_summary || {}).best_posting_insight || '';
  card(s, 6.55, 0.84, 3.2, 4.62, { accent: C.GREEN });
  sublabel(s, 6.72, 0.96, 2.9, 'Posting Strategy');
  body(s, 6.72, 1.2, 2.9, 4.1, posting, { size: 10.5 });
}

// 8. Content Gaps
function slideContentGaps(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'OPPORTUNITY', C.GOLD);
  heading(s, 0.4, 0.38, 9, 'Content Gaps & Opportunities', 20);

  const gaps    = (insights.executive_summary || {}).content_gaps || [];
  const gColors = [C.CYAN, C.GOLD, C.GREEN];
  const gNums   = ['01', '02', '03'];

  gaps.slice(0, 3).forEach((g, i) => {
    const y = 0.96 + i * 1.52;
    card(s, 0.4, y, 9.25, 1.32, { accent: gColors[i] });
    s.addText(gNums[i], {
      x: 0.6, y: y + 0.06, w: 0.56, h: 0.55,
      fontSize: 22, bold: true, color: gColors[i], fontFace: 'Consolas',
    });
    body(s, 1.28, y + 0.1,  8.1, 0.36, g.gap || '',         { size: 11, bold: true });
    body(s, 1.28, y + 0.52, 8.1, 0.58, g.opportunity || '',  { size: 10, color: C.MUTED });
  });
}

// 9. Your 90-Day Entry Strategy
function slideEntryStrategy(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'ACTION PLAN', C.GREEN);
  heading(s, 0.4, 0.38, 9, 'Your 90-Day Entry Strategy', 20);

  const entry   = insights.entry_strategy || {};
  const phases  = [
    { label: 'Days 1–30',  actions: entry.days_1_to_30  || [], color: C.VIOLET },
    { label: 'Days 31–60', actions: entry.days_31_to_60 || [], color: C.CYAN   },
    { label: 'Days 61–90', actions: entry.days_61_to_90 || [], color: C.GREEN  },
  ];

  const colW = 2.92;
  phases.forEach((ph, i) => {
    const x = 0.4 + i * (colW + 0.09);

    // Column header
    slide_addColHeader(s, x, 0.86, colW, ph.label, ph.color);

    // Action items
    ph.actions.slice(0, 3).forEach((action, j) => {
      const y = 1.42 + j * 0.96;
      card(s, x, y, colW, 0.86, { accent: ph.color });
      body(s, x + 0.14, y + 0.1, colW - 0.2, 0.7, action, { size: 9.5 });
    });
  });

  // Positioning callout
  const pos = entry.your_positioning || '';
  card(s, 0.4, 4.32, 9.25, 1.1, { bg: '111D1A', border: C.GREEN, accent: C.GREEN });
  sublabel(s, 0.62, 4.42, 9.0, 'Your Positioning Angle');
  body(s, 0.62, 4.62, 8.9, 0.72, pos, { size: 10, italic: true, color: C.GREEN });
}

function slide_addColHeader(s, x, y, w, label, color) {
  s.addShape('rect', { x, y, w, h: 0.44, fill: fillT(color, 75), line: line(color, 0.75) });
  s.addText(label, {
    x: x + 0.1, y, w: w - 0.1, h: 0.44,
    fontSize: 11, bold: true, color, fontFace: 'Trebuchet MS', valign: 'middle',
  });
}

// 10. Methodology
function slideMethodology(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'METHODOLOGY', C.MUTED);
  heading(s, 0.4, 0.38, 9, 'Data Sources & Methodology', 20);

  const statCards = [
    { v: String(videos.length),   l: 'Videos Analyzed',  c: C.CYAN   },
    { v: String(channels.length), l: 'Channels Profiled', c: C.VIOLET },
    { v: 'YouTube v3',            l: 'Data Source',       c: C.GREEN  },
    { v: 'GPT-4o',                l: 'AI Engine',         c: C.GOLD   },
  ];
  statCards.forEach((sc, i) => {
    const x = 0.4 + i * 2.32;
    card(s, x, 0.88, 2.18, 1.0, { accent: sc.c });
    s.addText(sc.v, {
      x: x + 0.18, y: 0.96, w: 1.88, h: 0.38,
      fontSize: 16, bold: true, color: sc.c, fontFace: 'Consolas',
    });
    s.addText(sc.l, {
      x: x + 0.18, y: 1.36, w: 1.88, h: 0.2,
      fontSize: 8, color: C.MUTED, fontFace: 'Calibri',
    });
  });

  card(s, 0.4, 2.05, 9.25, 2.62);
  sublabel(s, 0.58, 2.15, 8.9, 'Analysis Pipeline');
  const steps = [
    'YouTube Data API v3  →  ' + videos.length + ' videos + ' + channels.length + ' channels fetched',
    'youtube-transcript-api  →  Transcript extraction (free, no quota)',
    'GPT-4o  →  Comprehensive text, competitive, and entry strategy analysis',
    'GPT-4o-mini  →  Thumbnail visual analysis (20 images)',
    'matplotlib  →  7 charts generated from raw data',
    'LibreOffice  →  PPTX converted to PDF for universal compatibility',
  ];
  steps.forEach((step, i) => {
    s.addText(`${i + 1}.`, {
      x: 0.58, y: 2.38 + i * 0.35, w: 0.28, h: 0.28,
      fontSize: 9, bold: true, color: C.VIOLET, fontFace: 'Consolas',
    });
    body(s, 0.88, 2.38 + i * 0.35, 8.8, 0.28, step, { size: 9.5 });
  });

  s.addText(`Generated: ${todayStr()}`, {
    x: 0.4, y: 4.78, w: 9.25, h: 0.22,
    fontSize: 9, color: C.MUTED, fontFace: 'Calibri', align: 'right',
  });
}

// ── Build ─────────────────────────────────────────────────────────────────────
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9';

const slideSteps = [
  ['Title',                    () => slideTitle(pptx)],
  ['Market Opportunity',       () => slideMarketOpportunity(pptx)],
  ["Who's Winning & Why",      () => slideWhoIsWinning(pptx)],
  ['What Content Works',       () => slideWhatContentWorks(pptx)],
  ['Title & Hook Playbook',    () => slideTitleHooks(pptx)],
  ['Thumbnail Playbook',       () => slideThumbnailPlaybook(pptx)],
  ['Upload Strategy',          () => slideUploadStrategy(pptx)],
  ['Content Gaps',             () => slideContentGaps(pptx)],
  ['90-Day Entry Strategy',    () => slideEntryStrategy(pptx)],
  ['Methodology',              () => slideMethodology(pptx)],
];

slideSteps.forEach(([name, fn], i) => {
  process.stdout.write(`[${String(i + 1).padStart(2, '0')}/${slideSteps.length}] ${name}\n`);
  fn();
});

if (!fs.existsSync(OUTPUT)) fs.mkdirSync(OUTPUT, { recursive: true });
const slug  = topic.replace(/\s+/g, '_');
const fname = `${slug}_${slugDate()}.pptx`;
const out   = path.join(OUTPUT, fname);

pptx.writeFile({ fileName: out }).then(() => {
  process.stdout.write(`\n[OK] Saved -> ${out}\n`);
}).catch(err => {
  process.stderr.write(`Error: ${err.message}\n`);
  process.exit(1);
});
