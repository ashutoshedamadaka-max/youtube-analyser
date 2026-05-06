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

// Fresh fill object factory — never reuse across calls
const fill  = color             => ({ color });
const fillT = (color, pct)      => ({ color, transparency: pct });
const line  = (color, w = 0.75) => ({ color, width: w });
const noLine = ()               => ({ color: '000000', width: 0, transparency: 100 });

// ── Primitives ────────────────────────────────────────────────────────────────

function setBg(slide) {
  slide.background = { color: C.BG };
}

/** Dark card with optional left accent bar */
function card(slide, x, y, w, h, opts = {}) {
  const { bg = C.CARD, border = C.BORDER, accent = null } = opts;
  slide.addShape('rect', {
    x, y, w, h,
    fill: fill(bg),
    line: line(border, 0.75),
  });
  if (accent) {
    slide.addShape('rect', {
      x, y, w: 0.06, h,
      fill: fill(accent),
      line: noLine(),
    });
  }
}

/** Top accent bar on a card */
function topBar(slide, x, y, w, color) {
  slide.addShape('rect', {
    x, y, w, h: 0.055,
    fill: fill(color),
    line: noLine(),
  });
}

/** Small uppercase pill tag */
function tag(slide, x, y, label, color = C.VIOLET) {
  const tw = label.length * 0.072 + 0.22;
  slide.addShape('rect', {
    x, y, w: tw, h: 0.19,
    fill: fillT(color, 80),
    line: line(color, 0.5),
  });
  slide.addText(label.toUpperCase(), {
    x, y, w: tw, h: 0.19,
    fontSize: 6.5, bold: true, color: color,
    align: 'center', valign: 'middle',
    fontFace: 'Trebuchet MS',
  });
}

/** Section heading */
function heading(slide, x, y, w, text, size = 20) {
  slide.addText(text, {
    x, y, w, h: 0.38,
    fontSize: size, bold: true, color: C.TEXT,
    fontFace: 'Trebuchet MS', valign: 'middle',
  });
}

/** Muted uppercase sub-label */
function sublabel(slide, x, y, w, text) {
  slide.addText(text.toUpperCase(), {
    x, y, w, h: 0.18,
    fontSize: 6.5, bold: true, color: C.MUTED,
    fontFace: 'Trebuchet MS',
  });
}

/** Body copy */
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

/** Big monospace number + small label below */
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

/** Add chart image if it exists */
function chartImg(slide, filename, x, y, w, h) {
  const p = path.join(CHARTS, filename);
  if (fs.existsSync(p)) slide.addImage({ path: p, x, y, w, h });
}

/** Add thumbnail image if it exists */
function thumbImg(slide, thumbPath, x, y, w, h) {
  if (thumbPath && fs.existsSync(thumbPath)) {
    slide.addImage({ path: thumbPath, x, y, w, h });
  }
}

// ── Slides ────────────────────────────────────────────────────────────────────

function slideTitle(pptx) {
  const s = pptx.addSlide();
  setBg(s);

  // Decorative background shapes
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

  // Short accent line
  s.addShape('rect', { x: 0.5, y: 2.9, w: 1.1, h: 0.05,
    fill: fill(C.VIOLET), line: noLine() });

  s.addText(todayStr(), {
    x: 0.5, y: 3.05, w: 5, h: 0.25,
    fontSize: 10.5, color: C.MUTED, fontFace: 'Calibri',
  });

  // Stat row
  const stats = [
    { v: fmt(videos.length),          l: 'Videos Analyzed',  c: C.CYAN   },
    { v: fmt(channels.length),         l: 'Channels Profiled', c: C.GOLD   },
    { v: fmt(videos[0]?.views || 0),   l: 'Top Video Views',  c: C.GREEN  },
  ];
  stats.forEach((st, i) => {
    statBlock(s, 0.5 + i * 2.6, 3.75, st.v, st.l, st.c);
  });
}

function slideExecSummary(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.18, 'OVERVIEW', C.VIOLET);
  heading(s, 0.4, 0.44, 9, 'Executive Summary', 20);

  const summ     = insights.executive_summary || {};
  const overview = summ.space_overview || '';
  const blist    = summ.executive_summary_bullets || [];
  const accents  = [C.CYAN, C.VIOLET, C.GOLD, C.GREEN];

  card(s, 0.4, 0.9, 9.2, 0.68, { accent: C.VIOLET });
  body(s, 0.62, 0.96, 8.85, 0.58, overview, { size: 10 });

  blist.slice(0, 4).forEach((b, i) => {
    const y = 1.72 + i * 0.8;
    card(s, 0.4, y, 9.2, 0.68, { accent: accents[i] });
    body(s, 0.62, y + 0.08, 8.85, 0.55, b, { size: 11 });
  });
}

function slideTopVideos(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'PERFORMANCE', C.CYAN);
  heading(s, 0.4, 0.38, 9, 'Top 10 Videos by Views', 20);

  chartImg(s, 'top_videos.png', 0.25, 0.84, 5.9, 4.62);

  const rankColors  = [C.GOLD, C.MUTED, C.CORAL];
  const rankLabels  = ['1ST', '2ND', '3RD'];

  videos.slice(0, 3).forEach((v, i) => {
    const y = 0.84 + i * 1.54;
    card(s, 6.3, y, 3.45, 1.36, { bg: C.CARD });
    topBar(s, 6.3, y, 3.45, rankColors[i]);

    s.addText(rankLabels[i], {
      x: 6.46, y: y + 0.12, w: 0.6, h: 0.3,
      fontSize: 9.5, bold: true, color: rankColors[i], fontFace: 'Consolas',
    });
    body(s, 6.46, y + 0.42, 3.1, 0.4, trunc(v.title, 46), { size: 9 });
    body(s, 6.46, y + 0.88, 3.1, 0.28,
      `${fmt(v.views)} views  ·  ${trunc(v.channel_title, 20)}`,
      { size: 8.5, color: rankColors[i], bold: true, mono: true });
  });
}

function slideTopChannels(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'CHANNELS', C.GREEN);
  heading(s, 0.4, 0.38, 9, 'Top Channels in the Space', 20);

  chartImg(s, 'top_channels.png', 0.25, 0.84, 5.9, 4.62);

  const accs = [C.CYAN, C.VIOLET, C.GOLD, C.GREEN, C.CORAL];
  channels.slice(0, 5).forEach((c, i) => {
    const y = 0.84 + i * 0.93;
    card(s, 6.3, y, 3.45, 0.8, { accent: accs[i] });
    body(s, 6.52, y + 0.06, 3.0, 0.3, trunc(c.title, 28), { size: 10.5, bold: true });
    body(s, 6.52, y + 0.4,  3.0, 0.24,
      `${fmt(c.subscribers)} subs  ·  ${fmt(c.video_count)} videos`,
      { size: 8.5, color: C.MUTED });
  });
}

function slideEngagement(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'ENGAGEMENT', C.CORAL);
  heading(s, 0.4, 0.38, 9, 'Engagement Deep Dive', 20);

  chartImg(s, 'engagement.png', 0.25, 0.84, 6.1, 3.9);

  const tips = [
    { text: 'High views + low engagement = viral but not sticky', color: C.GOLD  },
    { text: 'High engagement + moderate views = loyal niche audience', color: C.GREEN },
    { text: 'Target the upper-right: big reach AND resonance', color: C.CYAN  },
  ];

  sublabel(s, 6.55, 0.84, 3.2, 'Interpretation');
  tips.forEach((tip, i) => {
    const y = 1.08 + i * 1.05;
    card(s, 6.55, y, 3.2, 0.88, { accent: tip.color });
    body(s, 6.74, y + 0.1, 2.85, 0.7, tip.text, { size: 10 });
  });

  chartImg(s, 'views_distribution.png', 6.55, 4.3, 3.2, 1.18);
}

function slideContentThemes(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'CONTENT MAP', C.GOLD);
  heading(s, 0.4, 0.38, 9, 'Content Themes', 20);

  chartImg(s, 'content_themes.png', 0.25, 0.78, 5.0, 4.68);

  const themes     = (insights.title_analysis || {}).dominant_content_themes || [];
  const themeAccs  = [C.CYAN, C.CORAL, C.GOLD, C.VIOLET, C.GREEN];

  sublabel(s, 5.5, 0.84, 4.2, 'Theme Breakdown');
  themes.slice(0, 5).forEach((t, i) => {
    const y = 1.08 + i * 0.88;
    card(s, 5.5, y, 4.2, 0.74, { accent: themeAccs[i] });
    s.addText(String(t.percentage || '?') + '%', {
      x: 5.65, y: y + 0.08, w: 0.82, h: 0.42,
      fontSize: 16, bold: true, color: themeAccs[i], fontFace: 'Consolas',
    });
    body(s, 6.58, y + 0.14, 2.95, 0.44, t.theme || '', { size: 10.5 });
  });
}

function slideTitleHooks(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'STRATEGY', C.VIOLET);
  heading(s, 0.4, 0.38, 9, 'Title & Hook Patterns', 20);

  const ta        = insights.title_analysis || {};
  const hooks     = ta.top_hook_patterns        || [];
  const keywords  = ta.high_performing_keywords || [];
  const overused  = ta.overused_patterns        || [];
  const corr      = ta.view_correlation_insights || '';

  // Left card — hook patterns
  card(s, 0.4, 0.86, 4.45, 3.12);
  sublabel(s, 0.56, 0.96, 4.1, 'Top Hook Patterns');
  hooks.slice(0, 5).forEach((h, i) => {
    s.addText(`${i + 1}.`, {
      x: 0.56, y: 1.2 + i * 0.44, w: 0.28, h: 0.34,
      fontSize: 10, bold: true, color: C.VIOLET, fontFace: 'Consolas',
    });
    body(s, 0.86, 1.2 + i * 0.44, 3.82, 0.36, h, { size: 10 });
  });

  // Right card — keyword chips
  card(s, 5.1, 0.86, 4.55, 3.12);
  sublabel(s, 5.26, 0.96, 4.2, 'High-Performing Keywords');
  let cx = 5.26, cy = 1.2;
  keywords.slice(0, 10).forEach(kw => {
    const cw = Math.max(0.65, kw.length * 0.076 + 0.22);
    if (cx + cw > 9.4) { cx = 5.26; cy += 0.4; }
    s.addShape('rect', {
      x: cx, y: cy, w: cw, h: 0.28,
      fill: fillT(C.CYAN, 82), line: line(C.CYAN, 0.5),
    });
    s.addText(kw, {
      x: cx, y: cy, w: cw, h: 0.28,
      fontSize: 8.5, color: C.CYAN, align: 'center', valign: 'middle',
      fontFace: 'Calibri',
    });
    cx += cw + 0.1;
  });

  // Insight quote card
  card(s, 0.4, 4.1, 9.25, 0.7, { bg: '1A1227', border: C.VIOLET, accent: C.VIOLET });
  body(s, 0.62, 4.18, 8.9, 0.54, `"${corr}"`, { size: 10, italic: true });

  // Overused warning
  if (overused.length) {
    s.addText('OVERSATURATED — AVOID:  ' + overused.join('  ·  '), {
      x: 0.4, y: 4.93, w: 9.25, h: 0.22,
      fontSize: 8.5, color: C.CORAL, fontFace: 'Calibri', italic: true,
    });
  }
}

function slideThumbnailTrends(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'VISUAL ANALYSIS', C.CORAL);
  heading(s, 0.4, 0.38, 9, 'Thumbnail Visual Trends', 20);

  // 2 x 3 thumbnail grid
  const shown = thumbData.filter(t => fs.existsSync(t.path)).slice(0, 6);
  const TW = 1.82, TH = 1.1;
  shown.forEach((t, i) => {
    const row = Math.floor(i / 3), col = i % 3;
    const ix = 0.28 + col * (TW + 0.1);
    const iy = 0.84 + row * (TH + 0.1);
    card(s, ix - 0.05, iy - 0.05, TW + 0.1, TH + 0.1);
    thumbImg(s, t.path, ix, iy, TW, TH);
  });

  // Three stacked analysis cards on right
  const ta   = insights.thumbnail_analysis || {};
  const rcds = [
    { label: 'VISUAL STYLE', text: ta.dominant_visual_style || '',
      color: C.CORAL },
    { label: 'COLOR TRENDS', text: (ta.color_palette_trends || []).join('  ·  '),
      color: C.GOLD },
    { label: 'DESIGN RECS',  text: (ta.design_recommendations || []).join('\n'),
      color: C.GREEN },
  ];
  rcds.forEach((rc, i) => {
    const y = 0.84 + i * 1.58;
    card(s, 6.38, y, 3.37, 1.4, { accent: rc.color });
    sublabel(s, 6.55, y + 0.1, 3.1, rc.label);
    body(s, 6.55, y + 0.34, 3.1, 0.98, rc.text, { size: 9.5 });
  });
}

function slideUploadCadence(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'TIMING', C.GREEN);
  heading(s, 0.4, 0.38, 9, 'Upload Cadence & Timing', 20);

  chartImg(s, 'upload_cadence.png', 0.25, 0.84, 6.1, 4.62);

  const posting = (insights.executive_summary || {}).best_posting_insight || '';
  card(s, 6.55, 0.84, 3.2, 4.62, { accent: C.GREEN });
  sublabel(s, 6.72, 0.96, 2.9, 'Posting Strategy');
  body(s, 6.72, 1.2, 2.9, 4.0, posting, { size: 10.5 });
}

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
    body(s, 1.28, y + 0.1,  8.1, 0.36, g.gap         || '', { size: 11, bold: true });
    body(s, 1.28, y + 0.52, 8.1, 0.58, g.opportunity  || '', { size: 10, color: C.MUTED });
  });
}

function slideRecommendations(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'ACTION PLAN', C.GREEN);
  heading(s, 0.4, 0.38, 9, 'Content Recommendations', 20);

  const recs    = (insights.executive_summary || {}).recommendations || [];
  const rColors = [C.VIOLET, C.CYAN, C.CORAL, C.GOLD, C.GREEN];

  const grid = [
    { x: 0.4,  y: 0.86, w: 4.55, h: 2.15 },
    { x: 5.1,  y: 0.86, w: 4.55, h: 2.15 },
    { x: 0.4,  y: 3.14, w: 4.55, h: 2.15 },
    { x: 5.1,  y: 3.14, w: 4.55, h: 2.15 },
  ];

  recs.slice(0, 4).forEach((rec, i) => {
    const { x, y, w, h } = grid[i];
    const color = rColors[i];
    card(s, x, y, w, h);
    topBar(s, x, y, w, color);
    s.addText(`#${i + 1}`, {
      x: x + 0.15, y: y + 0.1, w: 0.5, h: 0.3,
      fontSize: 13, bold: true, color, fontFace: 'Consolas',
    });
    body(s, x + 0.15, y + 0.44, w - 0.25, 0.36, rec.title        || '', { size: 11, bold: true });
    body(s, x + 0.15, y + 0.84, w - 0.25, 0.72, rec.description  || '', { size: 9.5 });
    body(s, x + 0.15, y + 1.6,  w - 0.25, 0.45, `Why: ${rec.rationale || ''}`,
      { size: 8.5, color: C.MUTED, italic: true });
  });
}

function slideMethodology(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'METHODOLOGY', C.MUTED);
  heading(s, 0.4, 0.38, 9, 'Data Sources & Methodology', 20);

  // 4 stat cards in a row
  const statCards = [
    { v: String(videos.length),   l: 'Videos Analyzed',   c: C.CYAN   },
    { v: String(channels.length), l: 'Channels Profiled',  c: C.VIOLET },
    { v: 'YouTube v3',            l: 'Data Source',        c: C.GREEN  },
    { v: 'GPT-4o',                l: 'AI Engine',          c: C.GOLD   },
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

  // Pipeline steps
  card(s, 0.4, 2.05, 9.25, 2.62);
  sublabel(s, 0.58, 2.15, 8.9, 'Analysis Pipeline');
  const steps = [
    'YouTube Data API v3  →  ' + videos.length + ' videos + ' + channels.length + ' channels fetched',
    'youtube-transcript-api  →  Transcript extraction (free, no quota)',
    'GPT-4o  →  Title, description, and pattern analysis',
    'GPT-4o-mini  →  Thumbnail visual analysis (20 images)',
    'GPT-4o  →  Executive summary + recommendations synthesis',
    'matplotlib  →  6 charts generated from raw data',
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

function slideCompetitive(pptx) {
  const s = pptx.addSlide();
  setBg(s);
  tag(s, 0.4, 0.12, 'COMPETITIVE LANDSCAPE', C.CORAL);
  heading(s, 0.4, 0.38, 9, 'Competitive Analysis', 20);

  const ca = insights.competitive_analysis || {};

  // Market concentration card (top strip)
  card(s, 0.4, 0.86, 9.25, 0.72, { accent: C.CORAL });
  sublabel(s, 0.58, 0.94, 1.6, 'Market Concentration');
  body(s, 0.58, 1.1, 9.0, 0.38, ca.market_concentration || '', { size: 10 });

  // Channel archetypes — up to 3 cards in a row
  const archetypes = ca.channel_archetypes || [];
  const archColors = [C.VIOLET, C.CYAN, C.GOLD];
  const archW = 2.98;
  sublabel(s, 0.4, 1.72, 9.0, 'Channel Archetypes');

  archetypes.slice(0, 3).forEach((a, i) => {
    const x = 0.4 + i * (archW + 0.055);
    card(s, x, 1.95, archW, 2.02, { accent: archColors[i] });
    s.addText(a.archetype || '', {
      x: x + 0.16, y: 2.02, w: archW - 0.22, h: 0.3,
      fontSize: 10, bold: true, color: archColors[i], fontFace: 'Trebuchet MS',
    });
    body(s, x + 0.16, 2.35, archW - 0.22, 0.5,
      (a.channels || []).join(', '), { size: 8.5, color: C.MUTED });
    body(s, x + 0.16, 2.72, archW - 0.22, 0.55,
      a.strategy || '', { size: 9 });
    body(s, x + 0.16, 3.3, archW - 0.22, 0.52,
      `⚡ ${a.weakness || ''}`, { size: 8.5, color: C.CORAL });
  });

  // Rising vs established + differentiation map
  const risingText = ca.rising_vs_established || '';
  const diffText   = ca.differentiation_map   || '';

  card(s, 0.4, 4.1, 4.55, 1.35, { accent: C.GREEN });
  sublabel(s, 0.58, 4.2, 4.2, 'Rising vs Established');
  body(s, 0.58, 4.4, 4.22, 0.95, risingText, { size: 9.5 });

  card(s, 5.1, 4.1, 4.55, 1.35, { accent: C.GOLD });
  sublabel(s, 5.28, 4.2, 4.2, 'Differentiation Map');
  body(s, 5.28, 4.4, 4.22, 0.95, diffText, { size: 9.5 });
}

// ── Build ─────────────────────────────────────────────────────────────────────
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_16x9';

const slideSteps = [
  ['Title',                () => slideTitle(pptx)],
  ['Executive Summary',    () => slideExecSummary(pptx)],
  ['Top Videos',           () => slideTopVideos(pptx)],
  ['Top Channels',         () => slideTopChannels(pptx)],
  ['Competitive Analysis', () => slideCompetitive(pptx)],
  ['Engagement',           () => slideEngagement(pptx)],
  ['Content Themes',       () => slideContentThemes(pptx)],
  ['Title Hooks',          () => slideTitleHooks(pptx)],
  ['Thumbnail Trends',     () => slideThumbnailTrends(pptx)],
  ['Upload Cadence',     () => slideUploadCadence(pptx)],
  ['Content Gaps',       () => slideContentGaps(pptx)],
  ['Recommendations',    () => slideRecommendations(pptx)],
  ['Methodology',        () => slideMethodology(pptx)],
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
