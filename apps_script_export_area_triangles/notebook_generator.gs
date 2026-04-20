/******************************************************************
 * notebook_generator.gs — EduWonderLab Flagship Render Engine v5.2
 *
 * Changes over v5.1:
 *   - Cover: card height computed to fill slide, no dead zone
 *   - Be Curious: vocabH correctly reserves discussion panel space
 *   - Vocab cards: visual inserted BEFORE text (z-order fix)
 *   - Pills: RECTANGLE background — LibreOffice clips ROUND_RECTANGLE text
 *   - Summary: stem height estimated, bullets offset dynamically
 *   - Discussion panel: LOCKED permanent structured collaboration panel
 *     (2 content-specific questions + write lines) on every task slide
 *   - New slide types: formulaDerivation, collaborateDeep, extensionChallenge
 *   - Entry points: runSession1Only() + runSession2Only() to avoid timeout
 *   - Text rendering fix: use transparent RECTANGLE text containers
 *     to avoid Slides TEXT_BOX instability
 ******************************************************************/

var NB_CONFIG = {
  OUTPUT_FOLDER_ID: '1htito8U51szyw6g9sLw8MJAgKar4jQ3k'
};

var NB_THEME = {
  W: 720, H: 405,
  M: 18, GAP: 8,
  HDR_H: 38,
  FTR_H: 20,
  PANEL_H: 72,

  FONT: 'Calibri',
  FONT_HERO: 'Georgia',
  SZ_HERO: 28,
  SZ_KICKER: 12,
  SZ_TITLE: 17,
  SZ_PROBLEM: 15,
  SZ_BODY: 13,
  SZ_SMALL: 13,
  SZ_TINY: 12,

  WHITE: '#FFFFFF',
  CREAM: '#FAFAF7',
  BODY: '#1A2332',
  MUTED: '#64748B',
  LINE: '#CBD5E1',

  S1: {
    primary: '#0F3554',
    accent: '#0D9488',
    accents: ['#0D9488', '#0891B2', '#1D6FA4', '#D97706'],
    light: '#E0F7F5',
    surface: '#F4FBFC',
    border: '#8BB8CC',
    warm: '#FEF3C7',
    stripe: '#0D9488',
    panel: '#0A2840'
  },
  S2: {
    primary: '#0C2D4E',
    accent: '#0369A1',
    accents: ['#0369A1', '#0891B2', '#059669', '#B45309'],
    light: '#E0F2FE',
    surface: '#F0F9FF',
    border: '#7BB3CC',
    warm: '#FFF7ED',
    stripe: '#0369A1',
    panel: '#061E33'
  }
};

var NB_LAYOUT = {
  CW: NB_THEME.W - NB_THEME.M * 2,
  CY: NB_THEME.HDR_H + NB_THEME.M,
  CH: NB_THEME.H - NB_THEME.HDR_H - NB_THEME.FTR_H - NB_THEME.M * 2,
  CH_PANEL: NB_THEME.H - NB_THEME.HDR_H - NB_THEME.FTR_H - NB_THEME.M * 2 - NB_THEME.PANEL_H
};

/* ================================================================
   ENTRY POINTS
   ================================================================ */

function runNotebookGeneratorFlagship() {
  if (typeof extractNotebookPackageFromDeck_ !== 'function') {
    if (typeof createFlagshipStudentNotebook === 'function') {
      return createFlagshipStudentNotebook();
    }
    throw new Error('Missing NotebookExtractors_Precision.gs. If you are using the standalone notebook file, run createFlagshipStudentNotebook() from Code.gs instead.');
  }
  var pkg = extractNotebookPackageFromDeck_(SlidesApp.getActivePresentation());
  validateNotebookPackage_(pkg);
  return buildFlagshipDecks_(pkg);
}

function runSession1Only() {
  if (typeof extractNotebookPackageFromDeck_ !== 'function') {
    if (typeof createSession1Notebook_ === 'function') {
      return createSession1Notebook_();
    }
    throw new Error('Missing NotebookExtractors_Precision.gs. If you are using the standalone notebook file, run createSession1Notebook_() from Code.gs instead.');
  }
  var pkg = extractNotebookPackageFromDeck_(SlidesApp.getActivePresentation());
  validateNotebookPackage_(pkg);
  return buildSessionDeck_(pkg.meta, pkg.sessions[0], 1);
}

function runSession2Only() {
  if (typeof extractNotebookPackageFromDeck_ !== 'function') {
    if (typeof createSession2Notebook_ === 'function') {
      return createSession2Notebook_();
    }
    throw new Error('Missing NotebookExtractors_Precision.gs. If you are using the standalone notebook file, run createSession2Notebook_() from Code.gs instead.');
  }
  var pkg = extractNotebookPackageFromDeck_(SlidesApp.getActivePresentation());
  validateNotebookPackage_(pkg);
  return buildSessionDeck_(pkg.meta, pkg.sessions[1], 2);
}

function buildFlagshipDecks_(pkg) {
  var outputs = [];
  for (var i = 0; i < pkg.sessions.length; i++) {
    outputs.push(buildSessionDeck_(pkg.meta, pkg.sessions[i], i + 1));
  }
  return outputs;
}

function validateNotebookPackage_(pkg) {
  if (!pkg || !pkg.sessions || !pkg.sessions.length) {
    throw new Error('No notebook sessions extracted.');
  }
}

/* ================================================================
   DECK BUILDER
   ================================================================ */

function buildSessionDeck_(meta, sessionSpec, idx) {
  var pal = idx === 1 ? NB_THEME.S1 : NB_THEME.S2;
  var name = [safeStr_(meta.title), 'Student Notebook',
              safeStr_(sessionSpec.lessonLabel || sessionSpec.sessionLabel)].join(' - ');
  var pres = SlidesApp.create(name);
  try {
    var file = DriveApp.getFileById(pres.getId());
    DriveApp.getFolderById(NB_CONFIG.OUTPUT_FOLDER_ID).addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  } catch (e) {}

  var existing = pres.getSlides();
  for (var i = 0; i < existing.length; i++) existing[i].remove();

  var sections = sessionSpec.sections || [];
  for (var page = 0; page < sections.length; page++) {
    renderSlide_(pres, sections[page] || {}, meta, sessionSpec, pal, page + 1);
  }

  pres.saveAndClose();
  return SlidesApp.openById(pres.getId());
}

/* ================================================================
   SLIDE ROUTER
   ================================================================ */

var PANEL_TYPES_ = {
  'beCurious': true, 'vocabActivity': true, 'formulaDerivation': true,
  'sourceTask': true, 'collaborateDeep': true, 'extensionChallenge': true,
  'summary': true
};

function renderSlide_(pres, section, meta, sessionSpec, pal, pageNum) {
  var isCover = section.type === 'cover';
  var bg = isCover ? pal.primary :
           (pageNum % 2 === 0 ? pal.surface : mixHex_(pal.light, NB_THEME.WHITE, 0.5));
  var slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill(bg);
  if (!isCover) renderHeader_(slide, safeStr_(section.title || ''), safeStr_(section.kicker || ''), pal, pageNum);

  switch (section.type) {
    case 'cover':              renderCover_(slide, section, sessionSpec, pal, meta);              break;
    case 'beCurious':          renderBeCurious_(slide, section, pal, meta, sessionSpec);          break;
    case 'vocabulary':         renderVocabulary_(slide, section, pal, meta, sessionSpec);         break;
    case 'vocabActivity':      renderVocabActivity_(slide, section, pal, meta, sessionSpec);      break;
    case 'formulaDerivation':  renderFormulaDerivation_(slide, section, pal, meta, sessionSpec);  break;
    case 'collaborateDeep':    renderCollaborateDeep_(slide, section, pal, meta, sessionSpec);    break;
    case 'extensionChallenge': renderExtensionChallenge_(slide, section, pal, meta, sessionSpec); break;
    case 'summary':            renderSummary_(slide, section, pal, meta, sessionSpec);            break;
    case 'reflection':         renderReflection_(slide, section, pal, meta, sessionSpec);         break;
    default:                   renderSourceTask_(slide, section, pal, meta, sessionSpec);         break;
  }
}

/* ================================================================
   HEADER
   ================================================================ */

function renderHeader_(slide, title, subtitle, pal, pageNum) {
  var accent = pal.accents[(pageNum - 1) % pal.accents.length];
  addRect_(slide, 0, 0, NB_THEME.W, NB_THEME.HDR_H, NB_THEME.WHITE);
  addRect_(slide, 0, 0, NB_THEME.W, 3, accent);
  addRect_(slide, 0, 3, 6, NB_THEME.HDR_H - 3, accent);
  addText_(slide, title, NB_THEME.M + 6, 5, 450, 18, { size: NB_THEME.SZ_TITLE, color: pal.primary, bold: true });
  if (subtitle) addText_(slide, subtitle, NB_THEME.M + 6, 23, 390, 12, { size: NB_THEME.SZ_TINY, color: NB_THEME.MUTED });
  addRect_(slide, NB_THEME.W - NB_THEME.M - 46, 8, 46, 22, accent);
  addText_(slide, String(pageNum), NB_THEME.W - NB_THEME.M - 46, 11, 46, 16,
    { size: NB_THEME.SZ_SMALL, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
}

/* ================================================================
   FOOTER
   ================================================================ */

function renderFooter_(slide, meta, sessionSpec, pal) {
  var fy = NB_THEME.H - NB_THEME.FTR_H;
  addRect_(slide, 0, fy, NB_THEME.W, NB_THEME.FTR_H, NB_THEME.WHITE);
  addRect_(slide, 0, fy, NB_THEME.W, 1, pal.border);
  addText_(slide, safeStr_(meta.title), NB_THEME.M, fy + 4, 280, 12, { size: NB_THEME.SZ_TINY, color: NB_THEME.MUTED });
  addText_(slide, safeStr_(sessionSpec.lessonLabel || ''), 300, fy + 4, 200, 12, { size: NB_THEME.SZ_TINY, color: NB_THEME.MUTED, align: 'CENTER' });
  if (meta.standard) addText_(slide, safeStr_(meta.standard), NB_THEME.W - NB_THEME.M - 80, fy + 4, 80, 12, { size: NB_THEME.SZ_TINY, color: pal.accent, align: 'RIGHT' });
}

/* ================================================================
   PERMANENT COLLABORATION PANEL
   ================================================================ */

function renderCollaborationPanel_(slide, discussionQs, pal) {
  var panY = NB_THEME.H - NB_THEME.FTR_H - NB_THEME.PANEL_H;
  var panW = NB_THEME.W;
  var qs = (discussionQs || []).slice(0, 2);
  while (qs.length < 2) qs.push('');

  addRect_(slide, 0, panY, panW, NB_THEME.PANEL_H, pal.panel || pal.primary);
  addRect_(slide, 0, panY, 4, NB_THEME.PANEL_H, pal.accent);

  addRect_(slide, NB_THEME.M + 4, panY + 6, 110, 16, pal.accent);
  addText_(slide, '\uD83D\uDDE3 Collaborate', NB_THEME.M + 4, panY + 7, 110, 14,
    { size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });

  var colW = Math.floor((panW - NB_THEME.M * 2 - 120 - NB_THEME.GAP * 2) / 2);
  var q1x = NB_THEME.M + 120;
  var q2x = q1x + colW + NB_THEME.GAP;

  for (var i = 0; i < 2; i++) {
    var qx = i === 0 ? q1x : q2x;
    var qText = qs[i] || '';
    if (!qText) continue;

    addText_(slide, (i + 1) + '. ' + qText, qx, panY + 6, colW, 30,
      { size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE });

    var line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT,
      qx, panY + 40, qx + colW - 4, panY + 40);
    line.getLineFill().setSolidFill(mixHex_(pal.accent, NB_THEME.WHITE, 0.55));
    line.setWeight(0.75);

    var line2 = slide.insertLine(SlidesApp.LineCategory.STRAIGHT,
      qx, panY + 54, qx + colW - 4, panY + 54);
    line2.getLineFill().setSolidFill(mixHex_(pal.accent, NB_THEME.WHITE, 0.35));
    line2.setWeight(0.75);
  }
}

/* ================================================================
   COVER
   ================================================================ */

function renderCover_(slide, section, sessionSpec, pal, meta) {
  addEllipse_(slide, NB_THEME.W - 160, -50, 220, 220, pal.accent, 0.12);
  addEllipse_(slide, NB_THEME.W - 100, 200, 180, 180, pal.accent, 0.07);
  addEllipse_(slide, -40, NB_THEME.H - 100, 160, 140, NB_THEME.WHITE, 0.06);
  addRect_(slide, 0, 0, NB_THEME.W, 4, pal.accent);

  addRect_(slide, NB_THEME.M, 14, 120, 22, pal.accent);
  addText_(slide, safeStr_(sessionSpec.sessionLabel || 'Session'), NB_THEME.M, 15, 120, 20,
    { size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
  if (meta.standard) {
    addRect_(slide, NB_THEME.W - NB_THEME.M - 78, 14, 78, 22, pal.stripe);
    addText_(slide, safeStr_(meta.standard), NB_THEME.W - NB_THEME.M - 78, 15, 78, 20,
      { size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
  }

  addText_(slide, safeStr_(section.kicker || ''), NB_THEME.M, 44, NB_LAYOUT.CW, 14,
    { font: NB_THEME.FONT, size: NB_THEME.SZ_KICKER, color: '#B8D4E8' });
  addText_(slide, safeStr_(section.title || meta.title), NB_THEME.M, 60, NB_LAYOUT.CW - 60, 62,
    { font: NB_THEME.FONT_HERO, size: NB_THEME.SZ_HERO, color: NB_THEME.WHITE, bold: true });
  addRect_(slide, NB_THEME.M, 126, 80, 3, pal.accent);

  var cardY = 136;
  var cardH = NB_THEME.H - NB_THEME.FTR_H - cardY - 2;
  addRoundRectWithStripe_(slide, NB_THEME.M, cardY, NB_LAYOUT.CW, cardH, NB_THEME.WHITE, pal.accent);
  addRect_(slide, NB_THEME.M + 14, cardY + 8, 130, 16, pal.primary);
  addText_(slide, 'LEARNING TARGETS', NB_THEME.M + 14, cardY + 9, 130, 14,
    { size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });

  var objH = Math.floor((cardH - 50) / 2) - 4;
  renderObjectiveRow_(slide, NB_THEME.M + 14, cardY + 28, 656, objH, 'Content Objective',
    ensureICan_(safeStr_(section.contentObjective || (sessionSpec.objectives || {}).content || '')),
    pal.primary, pal.light);
  renderObjectiveRow_(slide, NB_THEME.M + 14, cardY + 28 + objH + 6, 656, objH, 'Language Objective',
    ensureICan_(safeStr_(section.languageObjective || (sessionSpec.objectives || {}).language || '')),
    pal.accent, mixHex_(pal.warm, NB_THEME.WHITE, 0.3));

  var msY = cardY + cardH - 28;
  addRect_(slide, NB_THEME.M, msY, NB_LAYOUT.CW, 26, mixHex_(pal.accent, NB_THEME.WHITE, 0.15));
  addText_(slide, '\uD83D\uDCA1 ' + safeStr_(section.mindset || sessionSpec.mindset || 'I can explain my reasoning and learn from mistakes.'),
    NB_THEME.M + 10, msY + 5, NB_LAYOUT.CW - 20, 16, { size: NB_THEME.SZ_SMALL, color: NB_THEME.WHITE, bold: true });

  renderFooter_(slide, meta, sessionSpec, pal);
}

/* ================================================================
   BE CURIOUS
   ================================================================ */

function renderBeCurious_(slide, section, pal, meta, sessionSpec) {
  var colY = NB_LAYOUT.CY;
  var colH = NB_LAYOUT.CH;
  var IMG_W = 286;
  var IMG_H = 172;

  addRoundRectWithStripe_(slide, NB_THEME.M, colY, IMG_W, colH, NB_THEME.WHITE, pal.accent);

  var imgPlaced = false;
  if (section.imageDataUri) {
    imgPlaced = insertDataUriImage_(slide, section.imageDataUri, NB_THEME.M + 6, colY + 6, IMG_W - 12, IMG_H);
  }
  if (!imgPlaced) {
    addRect_(slide, NB_THEME.M + 6, colY + 6, IMG_W - 12, IMG_H, pal.light);
    addText_(slide, 'Source image from teacher deck', NB_THEME.M + 14, colY + 50, IMG_W - 28, 40,
      { size: NB_THEME.SZ_SMALL, color: pal.primary, align: 'CENTER' });
  }

  addText_(slide, safeStr_(section.imageCaption || 'Study this image carefully.'),
    NB_THEME.M + 8, colY + IMG_H + 10, IMG_W - 16, 28, { size: NB_THEME.SZ_SMALL, color: pal.primary, bold: true });
  if (typeof drawWritingLines_ === 'function') {
    drawWritingLines_(slide, NB_THEME.M + 8, colY + IMG_H + 44, IMG_W - 16, colH - IMG_H - 52, 4, NB_THEME.LINE);
  }

  var rightX = NB_THEME.M + IMG_W + NB_THEME.GAP;
  var rightW = NB_LAYOUT.CW - IMG_W - NB_THEME.GAP;
  var kernelH = 72;

  addRoundRectWithStripe_(slide, rightX, colY, rightW, kernelH, NB_THEME.WHITE, pal.primary);
  addText_(slide, 'Notice Sentence Kernels', rightX + 12, colY + 8, rightW - 24, 16,
    { size: NB_THEME.SZ_SMALL, color: pal.primary, bold: true });
  renderKernelStrip_(slide, '\u{1F441} Notice:', section.noticeKernels || [],
    rightX + 12, colY + 28, rightW - 24, pal.primary, pal);

  var wonderY = colY + kernelH + NB_THEME.GAP;
  addRoundRectWithStripe_(slide, rightX, wonderY, rightW, kernelH, NB_THEME.WHITE, pal.accent);
  addText_(slide, 'Wonder Sentence Kernels', rightX + 12, wonderY + 8, rightW - 24, 16,
    { size: NB_THEME.SZ_SMALL, color: pal.primary, bold: true });
  renderKernelStrip_(slide, '? Wonder:', section.wonderKernels || [],
    rightX + 12, wonderY + 28, rightW - 24, pal.accent, pal);

  var vocabY = wonderY + kernelH + NB_THEME.GAP;
  var vocabH = colH - (kernelH * 2) - (NB_THEME.GAP * 2);
  addRoundRectWithStripe_(slide, rightX, vocabY, rightW, vocabH, pal.surface, pal.border);
  addText_(slide, 'Academic Vocabulary', rightX + 12, vocabY + 8, rightW - 24, 16,
    { size: NB_THEME.SZ_SMALL, color: pal.primary, bold: true });

  var vocab = section.vocabBox || [];
  var rowH = Math.floor((vocabH - 30) / Math.max(1, Math.min(4, vocab.length || 1)));
  for (var v = 0; v < Math.min(4, vocab.length); v++) {
    var vy = vocabY + 28 + v * rowH;
    addRect_(slide, rightX + 12, vy, 96, 16, pal.primary);
    addText_(slide, safeStr_(vocab[v].term || ''), rightX + 12, vy + 1, 96, 14,
      { size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
    addText_(slide, safeStr_(vocab[v].definition || ''), rightX + 116, vy, rightW - 128, rowH - 4,
      { size: NB_THEME.SZ_TINY, color: NB_THEME.BODY });
  }

  renderFooter_(slide, meta, sessionSpec, pal);
}

/* ================================================================
   VOCABULARY
   ================================================================ */

function renderVocabulary_(slide, section, pal, meta, sessionSpec) {
  var terms = section.terms || [];
  var rows = Math.ceil(terms.length / 2);
  var cardH = rows > 1 ? Math.floor((NB_LAYOUT.CH - NB_THEME.GAP * (rows - 1)) / rows) : NB_LAYOUT.CH;
  var cardW = Math.floor((NB_LAYOUT.CW - NB_THEME.GAP) / 2);

  for (var i = 0; i < terms.length; i++) {
    var col = i % 2, row = Math.floor(i / 2);
    renderVocabCard_(slide, terms[i],
      NB_THEME.M + col * (cardW + NB_THEME.GAP),
      NB_LAYOUT.CY + row * (cardH + NB_THEME.GAP),
      cardW, cardH, pal, i);
  }
  renderFooter_(slide, meta, sessionSpec, pal);
}

function renderVocabCard_(slide, term, x, y, w, h, pal, idx) {
  var accent = pal.accents[idx % pal.accents.length];
  addRoundRectWithStripe_(slide, x, y, w, h, NB_THEME.WHITE, accent);

  addRect_(slide, x + 14, y + 10, w - 28, 20, accent);
  addText_(slide, safeStr_(term.term || ''), x + 14, y + 11, w - 28, 18,
    { size: NB_THEME.SZ_SMALL, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });

  var visW = 86, visH = h - 44;
  addRoundRect_(slide, x + 14, y + 36, visW, visH, mixHex_(pal.light, NB_THEME.WHITE, 0.4));
  if (typeof drawVocabVisual_ === 'function') {
    drawVocabVisual_(slide, safeStr_(term.visualKey || term.term || ''),
      x + 18, y + 40, visW - 8, visH - 8, pal);
  }

  var textX = x + 14 + visW + 8;
  var textW = w - 14 - visW - 22;

  addText_(slide, safeStr_(term.definition || ''), textX, y + 36, textW, 34,
    { size: NB_THEME.SZ_BODY, color: NB_THEME.BODY });

  if (term.studentFriendly) {
    addText_(slide, 'In your words:', textX, y + 74, textW, 13,
      { size: NB_THEME.SZ_TINY, color: pal.primary, bold: true });
    addText_(slide, safeStr_(term.studentFriendly), textX, y + 87, textW, 13,
      { size: NB_THEME.SZ_TINY, color: NB_THEME.MUTED });
  }
  if (term.example) {
    var exY = term.studentFriendly ? y + 103 : y + 74;
    addText_(slide, '\u2192 ' + safeStr_(term.example), textX, exY, textW, h - (exY - y) - 6,
      { size: NB_THEME.SZ_TINY, color: NB_THEME.BODY });
  }
}

/* ================================================================
   VOCABULARY ACTIVITY
   ================================================================ */

function renderVocabActivity_(slide, section, pal, meta, sessionSpec) {
  if (section.activityFamily) {
    renderInteractiveFamilyActivity_(slide, section, pal, meta, sessionSpec);
    return;
  }

  var items = section.activities || [];
  var colH = NB_LAYOUT.CH_PANEL;
  var cardW = Math.floor((NB_LAYOUT.CW - NB_THEME.GAP) / 2);
  var cardH = Math.floor((colH - NB_THEME.GAP) / 2);

  for (var i = 0; i < Math.min(4, items.length); i++) {
    var col = i % 2, row = Math.floor(i / 2);
    var x = NB_THEME.M + col * (cardW + NB_THEME.GAP);
    var y = NB_LAYOUT.CY + row * (cardH + NB_THEME.GAP);
    var ac = pal.accents[i % pal.accents.length];
    addRoundRectWithStripe_(slide, x, y, cardW, cardH, NB_THEME.WHITE, ac);
    addRect_(slide, x + 12, y + 10, 88, 17, ac);
    addText_(slide, safeStr_(items[i].label || ('Task ' + (i + 1))), x + 12, y + 11, 88, 15,
      { size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
    addText_(slide, safeStr_(items[i].prompt || ''), x + 12, y + 34, cardW - 24, cardH - 46,
      { size: NB_THEME.SZ_BODY, color: NB_THEME.BODY });
  }

  renderCollaborationPanel_(slide, section.discussionQs || [section.talkPrompt || ''], pal);
  renderFooter_(slide, meta, sessionSpec, pal);
}

function renderInteractiveFamilyActivity_(slide, section, pal, meta, sessionSpec) {
  switch (safeStr_(section.activityFamily)) {
    case 'build_construct':
      renderBuildConstructActivity_(slide, section, pal, meta, sessionSpec);
      return;
    case 'match_pair':
      renderMatchPairActivity_(slide, section, pal, meta, sessionSpec);
      return;
    case 'sort_classify':
      renderSortClassifyActivity_(slide, section, pal, meta, sessionSpec);
      return;
  }

  renderGenericInteractiveFamilyActivity_(slide, section, pal, meta, sessionSpec);
}

function renderGenericInteractiveFamilyActivity_(slide, section, pal, meta, sessionSpec) {
  var colY = NB_LAYOUT.CY;
  var colH = NB_LAYOUT.CH_PANEL;
  var leftW = 422;
  var rightW = NB_LAYOUT.CW - leftW - NB_THEME.GAP;
  var leftX = NB_THEME.M;
  var rightX = NB_THEME.M + leftW + NB_THEME.GAP;
  var topH = 48;
  var gridY = colY + topH + NB_THEME.GAP;
  var gridH = colH - topH - NB_THEME.GAP;
  var items = (section.activities || []).slice(0, 4);
  var isTriad = items.length <= 3;
  var cardW = isTriad ? leftW : Math.floor((leftW - NB_THEME.GAP) / 2);
  var cardH = isTriad ?
    Math.floor((gridH - NB_THEME.GAP * Math.max(0, items.length - 1)) / Math.max(1, items.length)) :
    Math.floor((gridH - NB_THEME.GAP) / 2);

  addRoundRectWithStripe_(slide, leftX, colY, leftW, topH, NB_THEME.WHITE, pal.accent);
  addText_(slide, safeStr_(section.activityInstructions || defaultActivityInstructions_(section.activityFamily)),
    leftX + 12, colY + 10, leftW - 24, topH - 16,
    { size: NB_THEME.SZ_SMALL, color: NB_THEME.BODY, bold: true });

  for (var i = 0; i < items.length; i++) {
    var col = isTriad ? 0 : i % 2;
    var row = isTriad ? i : Math.floor(i / 2);
    var x = isTriad ? leftX : leftX + col * (cardW + NB_THEME.GAP);
    var y = gridY + row * (cardH + NB_THEME.GAP);
    var accent = pal.accents[i % pal.accents.length];
    var prompt = safeStr_(items[i].prompt || '');
    var labelW = isTriad ? 118 : 92;
    var bodyY = isTriad ? y + 34 : y + 32;
    var bodyH = isTriad ? cardH - 58 : cardH - 64;

    addRoundRectWithStripe_(slide, x, y, cardW, cardH, NB_THEME.WHITE, accent);
    addTextShape_(slide, SlidesApp.ShapeType.RECTANGLE, safeStr_(items[i].label || ('Box ' + (i + 1))),
      x + 10, y + 8, labelW, 18,
      { fill: accent, borderColor: accent, size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
    addText_(slide, prompt, x + 10, bodyY, cardW - 20, bodyH,
      { size: NB_THEME.SZ_TINY, color: NB_THEME.BODY });
    addTextShape_(slide, SlidesApp.ShapeType.RECTANGLE, safeStr_(section.dropzoneHint || 'Drag matching cards here'),
      x + 10, y + cardH - 24, cardW - 20, 16,
      { fill: mixHex_(accent, NB_THEME.WHITE, 0.88), borderColor: accent, size: NB_THEME.SZ_TINY, color: accent, align: 'CENTER' });
  }

  renderInteractivePieceBank_(slide, section, pal, rightX, colY, rightW, colH);
  renderCollaborationPanel_(slide, section.discussionQs || [section.talkPrompt || ''], pal);
  renderFooter_(slide, meta, sessionSpec, pal);
}

function renderBuildConstructActivity_(slide, section, pal, meta, sessionSpec) {
  var colY = NB_LAYOUT.CY;
  var colH = NB_LAYOUT.CH_PANEL;
  var leftW = 428;
  var rightW = NB_LAYOUT.CW - leftW - NB_THEME.GAP;
  var leftX = NB_THEME.M;
  var rightX = NB_THEME.M + leftW + NB_THEME.GAP;
  var topH = 42;
  var items = (section.activities || []).slice(0, 3);
  var gridY = colY + topH + NB_THEME.GAP;
  var cardH = Math.floor((colH - topH - NB_THEME.GAP * Math.max(1, items.length)) / Math.max(1, items.length));

  addRoundRectWithStripe_(slide, leftX, colY, leftW, topH, NB_THEME.WHITE, pal.accent);
  addText_(slide, safeStr_(section.activityInstructions || defaultActivityInstructions_(section.activityFamily)),
    leftX + 12, colY + 10, leftW - 24, 20,
    { size: NB_THEME.SZ_SMALL, color: NB_THEME.BODY, bold: true });
  renderSourceRefChip_(slide, section.sourceRefLabel, pal, leftX + leftW - 102, colY + 8, 90);

  for (var i = 0; i < items.length; i++) {
    var y = gridY + i * (cardH + NB_THEME.GAP);
    var accent = pal.accents[i % pal.accents.length];
    addRoundRectWithStripe_(slide, leftX, y, leftW, cardH, NB_THEME.WHITE, accent);
    addTextShape_(slide, SlidesApp.ShapeType.RECTANGLE, safeStr_(items[i].label || ('Step ' + (i + 1))),
      leftX + 12, y + 8, 118, 18,
      { fill: accent, borderColor: accent, size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
    addText_(slide, safeStr_(items[i].prompt || ''), leftX + 12, y + 32, leftW - 24, 28,
      { size: NB_THEME.SZ_TINY, color: NB_THEME.BODY });
    addTextShape_(slide, SlidesApp.ShapeType.RECTANGLE, safeStr_(section.dropzoneHint || 'Place two matching cards here'),
      leftX + 12, y + cardH - 28, leftW - 24, 18,
      { fill: mixHex_(accent, NB_THEME.WHITE, 0.9), borderColor: accent, size: NB_THEME.SZ_TINY, color: accent, align: 'CENTER' });
  }

  renderInteractivePieceBank_(slide, section, pal, rightX, colY, rightW, colH);
  renderCollaborationPanel_(slide, section.discussionQs || [section.talkPrompt || ''], pal);
  renderFooter_(slide, meta, sessionSpec, pal);
}

function renderMatchPairActivity_(slide, section, pal, meta, sessionSpec) {
  var colY = NB_LAYOUT.CY;
  var colH = NB_LAYOUT.CH_PANEL;
  var leftW = 430;
  var rightW = NB_LAYOUT.CW - leftW - NB_THEME.GAP;
  var leftX = NB_THEME.M;
  var rightX = NB_THEME.M + leftW + NB_THEME.GAP;
  var topH = 42;
  var rows = (section.activities || []).slice(0, 4);
  var rowY = colY + topH + NB_THEME.GAP;
  var rowH = Math.floor((colH - topH - NB_THEME.GAP * Math.max(1, rows.length)) / Math.max(1, rows.length));

  addRoundRectWithStripe_(slide, leftX, colY, leftW, topH, NB_THEME.WHITE, pal.accent);
  addText_(slide, safeStr_(section.activityInstructions || defaultActivityInstructions_(section.activityFamily)),
    leftX + 12, colY + 10, leftW - 24, 20,
    { size: NB_THEME.SZ_SMALL, color: NB_THEME.BODY, bold: true });
  renderSourceRefChip_(slide, section.sourceRefLabel, pal, leftX + leftW - 102, colY + 8, 90);

  for (var i = 0; i < rows.length; i++) {
    var y = rowY + i * (rowH + NB_THEME.GAP);
    var accent = pal.accents[i % pal.accents.length];
    addRoundRectWithStripe_(slide, leftX, y, leftW, rowH, NB_THEME.WHITE, accent);
    addTextShape_(slide, SlidesApp.ShapeType.RECTANGLE, safeStr_(rows[i].label || ('Match ' + (i + 1))),
      leftX + 12, y + 8, 92, 18,
      { fill: accent, borderColor: accent, size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
    addText_(slide, safeStr_(rows[i].prompt || ''), leftX + 12, y + 32, 246, rowH - 40,
      { size: NB_THEME.SZ_TINY, color: NB_THEME.BODY });
    addTextShape_(slide, SlidesApp.ShapeType.RECTANGLE, safeStr_(section.dropzoneHint || 'Place the matching card here'),
      leftX + leftW - 156, y + 16, 144, rowH - 28,
      { fill: mixHex_(accent, NB_THEME.WHITE, 0.9), borderColor: accent, size: NB_THEME.SZ_TINY, color: accent, align: 'CENTER' });
  }

  renderInteractivePieceBank_(slide, section, pal, rightX, colY, rightW, colH);
  renderCollaborationPanel_(slide, section.discussionQs || [section.talkPrompt || ''], pal);
  renderFooter_(slide, meta, sessionSpec, pal);
}

function renderSortClassifyActivity_(slide, section, pal, meta, sessionSpec) {
  var colY = NB_LAYOUT.CY;
  var colH = NB_LAYOUT.CH_PANEL;
  var leftW = 432;
  var rightW = NB_LAYOUT.CW - leftW - NB_THEME.GAP;
  var leftX = NB_THEME.M;
  var rightX = NB_THEME.M + leftW + NB_THEME.GAP;
  var topH = 42;
  var cols = (section.activities || []).slice(0, 2);
  var boxY = colY + topH + NB_THEME.GAP;
  var boxH = colH - topH - NB_THEME.GAP;
  var boxW = Math.floor((leftW - NB_THEME.GAP) / 2);

  addRoundRectWithStripe_(slide, leftX, colY, leftW, topH, NB_THEME.WHITE, pal.accent);
  addText_(slide, safeStr_(section.activityInstructions || defaultActivityInstructions_(section.activityFamily)),
    leftX + 12, colY + 10, leftW - 24, 20,
    { size: NB_THEME.SZ_SMALL, color: NB_THEME.BODY, bold: true });
  renderSourceRefChip_(slide, section.sourceRefLabel, pal, leftX + leftW - 102, colY + 8, 90);

  for (var i = 0; i < cols.length; i++) {
    var x = leftX + i * (boxW + NB_THEME.GAP);
    var accent = pal.accents[i % pal.accents.length];
    addRoundRectWithStripe_(slide, x, boxY, boxW, boxH, NB_THEME.WHITE, accent);
    addTextShape_(slide, SlidesApp.ShapeType.RECTANGLE, safeStr_(cols[i].label || ('Category ' + (i + 1))),
      x + 12, boxY + 8, boxW - 24, 18,
      { fill: accent, borderColor: accent, size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
    addText_(slide, safeStr_(cols[i].prompt || ''), x + 12, boxY + 32, boxW - 24, 42,
      { size: NB_THEME.SZ_TINY, color: NB_THEME.BODY });
    addTextShape_(slide, SlidesApp.ShapeType.RECTANGLE, safeStr_(section.dropzoneHint || 'Sort the matching cards here'),
      x + 12, boxY + boxH - 28, boxW - 24, 18,
      { fill: mixHex_(accent, NB_THEME.WHITE, 0.9), borderColor: accent, size: NB_THEME.SZ_TINY, color: accent, align: 'CENTER' });
  }

  renderInteractivePieceBank_(slide, section, pal, rightX, colY, rightW, colH);
  renderCollaborationPanel_(slide, section.discussionQs || [section.talkPrompt || ''], pal);
  renderFooter_(slide, meta, sessionSpec, pal);
}

function renderInteractivePieceBank_(slide, section, pal, x, y, w, h) {
  var pieces = (section.movablePieces || []).slice(0, 10);
  var answerCheck = safeStr_(section.answerCheck || '');
  var bankH = answerCheck ? h - 64 : h;
  var singleColumn = pieces.length <= 6;
  var chipW = singleColumn ? (w - 20) : Math.floor((w - 26) / 2);
  var chipH = singleColumn ? 24 : 18;
  var bankTitle = safeStr_(section.activityBankTitle || interactiveFamilyTitle_(section.activityFamily));
  var bankIntro = safeStr_(section.bankIntro || 'Move the cards into the matching boxes, then explain your choices.');

  addRoundRectWithStripe_(slide, x, y, w, bankH, pal.surface, pal.stripe);
  addTextShape_(slide, SlidesApp.ShapeType.RECTANGLE, bankTitle,
    x + 10, y + 8, w - 20, 18,
    { fill: pal.primary, borderColor: pal.primary, size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
  addText_(slide, bankIntro,
    x + 10, y + 30, w - 20, 24,
    { size: NB_THEME.SZ_TINY, color: NB_THEME.BODY });

  for (var i = 0; i < pieces.length; i++) {
    var col = singleColumn ? 0 : i % 2;
    var row = singleColumn ? i : Math.floor(i / 2);
    var chipX = x + 10 + col * (chipW + 6);
    var chipY = y + 58 + row * (chipH + 6);
    if (chipY + chipH > y + bankH - 8) break;
    addTextShape_(slide, SlidesApp.ShapeType.RECTANGLE, safeStr_(pieces[i]),
      chipX, chipY, chipW, chipH,
      { fill: NB_THEME.WHITE, borderColor: pal.accent, size: NB_THEME.SZ_TINY, color: pal.primary, align: 'CENTER' });
  }

  if (answerCheck) {
    addRoundRectWithStripe_(slide, x, y + bankH + NB_THEME.GAP, w, h - bankH - NB_THEME.GAP, mixHex_(pal.light, NB_THEME.WHITE, 0.45), pal.accent);
    addText_(slide, 'Teacher Check', x + 10, y + bankH + 12, w - 20, 14,
      { size: NB_THEME.SZ_TINY, color: pal.primary, bold: true });
    addText_(slide, answerCheck, x + 10, y + bankH + 28, w - 20, h - bankH - 24,
      { size: NB_THEME.SZ_TINY, color: NB_THEME.BODY });
  }
}

function defaultActivityInstructions_(family) {
  switch (safeStr_(family)) {
    case 'sort_classify':
      return 'Sort each movable card into the matching category box, then explain one sort.';
    case 'match_pair':
      return 'Match the movable cards to the correct box, then explain one pairing.';
    case 'sequence_order':
      return 'Place the movable cards in the correct order, then explain the sequence.';
    case 'build_construct':
      return 'Use the movable cards to build the model, formula, or explanation, then justify it.';
    case 'detect_justify':
      return 'Decide which cards fit the claim, then justify your choice with evidence.';
    case 'compare_rank':
      return 'Compare the cards, place them in the best box, and explain your ranking.';
    default:
      return 'Move the cards into the matching boxes and explain one of your choices.';
  }
}

function interactiveFamilyTitle_(family) {
  switch (safeStr_(family)) {
    case 'sort_classify':
      return 'Interactive Sort';
    case 'match_pair':
      return 'Interactive Match';
    case 'sequence_order':
      return 'Sequence Builder';
    case 'build_construct':
      return 'Build + Explain';
    case 'detect_justify':
      return 'Decide + Justify';
    case 'compare_rank':
      return 'Compare + Rank';
    default:
      return 'Interactive Activity';
  }
}

/* ================================================================
   FORMULA DERIVATION
   ================================================================ */

function renderFormulaDerivation_(slide, section, pal, meta, sessionSpec) {
  var colH = NB_LAYOUT.CH_PANEL;
  var colY = NB_LAYOUT.CY;
  var stepW = Math.floor((NB_LAYOUT.CW - NB_THEME.GAP * 2) / 3);
  var steps = section.steps || [];

  renderSourceRefChip_(slide, section.sourceRefLabel, pal, NB_THEME.W - NB_THEME.M - 96, colY - 2, 86);

  for (var i = 0; i < Math.min(3, steps.length); i++) {
    var x = NB_THEME.M + i * (stepW + NB_THEME.GAP);
    var accent = pal.accents[i % pal.accents.length];
    addRoundRectWithStripe_(slide, x, colY, stepW, colH, NB_THEME.WHITE, accent);

    addRect_(slide, x + 12, colY + 8, 24, 24, accent);
    addText_(slide, String(i + 1), x + 12, colY + 9, 24, 22,
      { size: NB_THEME.SZ_SMALL, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });

    addText_(slide, safeStr_(steps[i].title || ''), x + 42, colY + 12, stepW - 54, 18,
      { size: NB_THEME.SZ_SMALL, color: pal.primary, bold: true });

    var vizH = Math.floor(colH * 0.45);
    addRoundRect_(slide, x + 12, colY + 38, stepW - 24, vizH, pal.light);
    if (typeof drawWorkspaceVisual_ === 'function') {
      drawWorkspaceVisual_(slide, safeStr_(steps[i].visualKey || 'triangle'),
        x + 18, colY + 44, stepW - 36, vizH - 12, pal);
    }

    addText_(slide, safeStr_(steps[i].explanation || ''), x + 12, colY + 40 + vizH, stepW - 24, colH - vizH - 50,
      { size: NB_THEME.SZ_BODY, color: NB_THEME.BODY });
  }

  renderCollaborationPanel_(slide, section.discussionQs || [], pal);
  renderFooter_(slide, meta, sessionSpec, pal);
}

/* ================================================================
   SOURCE TASK
   ================================================================ */

function renderSourceTask_(slide, section, pal, meta, sessionSpec) {
  var leftW = 414, rightW = NB_LAYOUT.CW - leftW - NB_THEME.GAP;
  var leftX = NB_THEME.M, rightX = NB_THEME.M + leftW + NB_THEME.GAP;
  var colY = NB_LAYOUT.CY;
  var colH = NB_LAYOUT.CH_PANEL;

  addRoundRectWithStripe_(slide, leftX, colY, leftW, colH, NB_THEME.WHITE, pal.accent);

  var curY = colY + 8;

  if (section.stageLabel) {
    addRect_(slide, leftX + 12, curY, 142, 17, pal.accent);
    addText_(slide, safeStr_(section.stageLabel), leftX + 12, curY + 1, 142, 15,
      { size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
    renderSourceRefChip_(slide, section.sourceRefLabel, pal, leftX + leftW - 108, curY, 96);
    curY += 23;
  } else {
    renderSourceRefChip_(slide, section.sourceRefLabel, pal, leftX + leftW - 108, curY, 96);
    curY += section.sourceRefLabel ? 22 : 4;
  }

  var probText = safeStr_(section.problem || '');
  var probH = estimateTextHeight_(probText, leftW - 32);
  addRoundRect_(slide, leftX + 12, curY, leftW - 24, probH, pal.light);
  addRect_(slide, leftX + 12, curY, 4, probH, pal.accent);
  addText_(slide, probText, leftX + 20, curY + 6, leftW - 36, probH - 10,
    { size: NB_THEME.SZ_PROBLEM, color: pal.primary, bold: true });
  curY += probH + 5;

  if (section.supportQuestion) {
    var suppH = Math.max(38, estimateTextHeight_(section.supportQuestion, leftW - 44) + 24);
    renderThinkBox_(slide, leftX + 12, curY, leftW - 24, suppH,
      safeStr_(section.supportLabel || 'Think About It'),
      safeStr_(section.supportQuestion), pal.light, pal.primary);
    curY += suppH + 5;
  }

  var revealH = section.reveal ? 52 : 0;
  var maxBulletY = colY + colH - revealH - 8;
  var lines = section.sourceLines || [];
  for (var i = 0; i < lines.length; i++) {
    var ly = curY + i * 18;
    if (ly + 17 > maxBulletY) break;
    addText_(slide, '\u2022 ' + safeStr_(lines[i]), leftX + 16, ly, leftW - 32, 18,
      { size: NB_THEME.SZ_BODY, color: NB_THEME.BODY });
  }

  if (section.reveal) {
    renderRevealBox_(slide, leftX + 12, colY + colH - revealH, leftW - 24, revealH,
      safeStr_(section.revealLabel || 'Reveal'), safeStr_(section.reveal), pal);
  }

  var imgH = Math.floor(colH * 0.50);
  addRoundRectWithStripe_(slide, rightX, colY, rightW, imgH, NB_THEME.WHITE, pal.border);
  var imgPlaced = false;
  if (section.imageDataUri) {
    imgPlaced = insertDataUriImage_(slide, section.imageDataUri, rightX + 6, colY + 6, rightW - 12, imgH - 12);
  }
  if (!imgPlaced && typeof drawWorkspaceVisual_ === 'function') {
    drawWorkspaceVisual_(slide, safeStr_(section.visualKey || 'formula'),
      rightX + 8, colY + 8, rightW - 16, imgH - 16, pal);
  }

  var wsY = colY + imgH + NB_THEME.GAP;
  var wsH = colH - imgH - NB_THEME.GAP;
  addRoundRectWithStripe_(slide, rightX, wsY, rightW, wsH, pal.surface, pal.stripe);
  if (section.strategyTitle) {
    addRect_(slide, rightX + 10, wsY + 8, rightW - 20, 17, pal.accent);
    addText_(slide, safeStr_(section.strategyTitle), rightX + 10, wsY + 9, rightW - 20, 15,
      { size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
  }
  if (section.factCards && section.factCards.length) {
    renderFactCardGrid_(slide, section.factCards, pal, rightX + 10, wsY + 32, rightW - 20, wsH - 40);
  } else if (typeof drawWritingLines_ === 'function') {
    drawWritingLines_(slide, rightX + 10, wsY + 32, rightW - 20, wsH - 40, 3, NB_THEME.LINE);
  }

  renderCollaborationPanel_(slide, section.discussionQs || [section.talkPrompt || ''], pal);
  renderFooter_(slide, meta, sessionSpec, pal);
}

function renderFactCardGrid_(slide, factCards, pal, x, y, w, h) {
  var cards = (factCards || []).slice(0, 4);
  if (!cards.length) return;

  var cols = cards.length <= 2 ? 1 : 2;
  var rows = Math.ceil(cards.length / cols);
  var gap = 6;
  var cardW = cols === 1 ? w : Math.floor((w - gap) / 2);
  var cardH = Math.min(40, Math.floor((h - gap * Math.max(0, rows - 1) - 28) / rows));

  addText_(slide, 'Keep these source facts visible as you solve.', x, y, w, 16,
    { size: NB_THEME.SZ_TINY, color: pal.primary, bold: true });

  for (var i = 0; i < cards.length; i++) {
    var col = cols === 1 ? 0 : i % cols;
    var row = cols === 1 ? i : Math.floor(i / cols);
    var cx = x + col * (cardW + gap);
    var cy = y + 20 + row * (cardH + gap);
    addTextShape_(slide, SlidesApp.ShapeType.RECTANGLE, safeStr_(cards[i]),
      cx, cy, cardW, cardH,
      {
        fill: NB_THEME.WHITE,
        borderColor: pal.accent,
        size: NB_THEME.SZ_TINY,
        color: pal.primary,
        bold: true,
        align: 'CENTER'
      });
  }

  if (typeof drawWritingLines_ === 'function') {
    var linesY = y + 24 + rows * (cardH + gap);
    var linesH = h - (linesY - y);
    if (linesH > 20) {
      drawWritingLines_(slide, x, linesY, w, linesH, 2, NB_THEME.LINE);
    }
  }
}

/* ================================================================
   COLLABORATE DEEP
   ================================================================ */

function renderCollaborateDeep_(slide, section, pal, meta, sessionSpec) {
  var colH = NB_LAYOUT.CH_PANEL;
  var colY = NB_LAYOUT.CY;

  renderSourceRefChip_(slide, section.sourceRefLabel, pal, NB_THEME.W - NB_THEME.M - 96, colY - 2, 86);

  var probText = safeStr_(section.problem || '');
  var probH = Math.max(44, estimateTextHeight_(probText, NB_LAYOUT.CW - 28));
  addRoundRectWithStripe_(slide, NB_THEME.M, colY, NB_LAYOUT.CW, probH, pal.light, pal.accent);
  addText_(slide, probText, NB_THEME.M + 16, colY + 8, NB_LAYOUT.CW - 28, probH - 14,
    { size: NB_THEME.SZ_PROBLEM, color: pal.primary, bold: true });

  var partH = colH - probH - NB_THEME.GAP;
  var partW = Math.floor((NB_LAYOUT.CW - NB_THEME.GAP) / 2);
  var partY = colY + probH + NB_THEME.GAP;
  var roles = section.partnerRoles || ['Partner A', 'Partner B'];
  var tasks = section.partnerTasks || ['', ''];

  for (var i = 0; i < 2; i++) {
    var px = NB_THEME.M + i * (partW + NB_THEME.GAP);
    var ac = pal.accents[i % 2];
    var taskText = safeStr_(tasks[i] || '');
    var taskH = Math.min(76, Math.max(46, estimateTextHeight_(taskText, partW - 24) - 8));
    addRoundRectWithStripe_(slide, px, partY, partW, partH, NB_THEME.WHITE, ac);
    addRect_(slide, px + 12, partY + 8, partW - 24, 18, ac);
    addText_(slide, safeStr_(roles[i] || ''), px + 12, partY + 9, partW - 24, 16,
      { size: NB_THEME.SZ_SMALL, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
    addText_(slide, taskText, px + 12, partY + 32, partW - 24, taskH,
      { size: NB_THEME.SZ_BODY, color: NB_THEME.BODY });
    if (typeof drawWritingLines_ === 'function') {
      drawWritingLines_(slide, px + 12, partY + 36 + taskH, partW - 24, partH - taskH - 44, 4, NB_THEME.LINE);
    }
  }

  renderCollaborationPanel_(slide, section.discussionQs || [], pal);
  renderFooter_(slide, meta, sessionSpec, pal);
}

/* ================================================================
   EXTENSION CHALLENGE
   ================================================================ */

function renderExtensionChallenge_(slide, section, pal, meta, sessionSpec) {
  var colH = NB_LAYOUT.CH_PANEL;
  var colY = NB_LAYOUT.CY;
  var leftW = 388, rightW = NB_LAYOUT.CW - leftW - NB_THEME.GAP;

  renderSourceRefChip_(slide, section.sourceRefLabel, pal, NB_THEME.W - NB_THEME.M - 96, colY - 2, 86);

  addRoundRectWithStripe_(slide, NB_THEME.M, colY, leftW, colH, NB_THEME.WHITE, pal.accent);
  addRect_(slide, NB_THEME.M + 12, colY + 8, 130, 17, pal.accent);
  addText_(slide, '\u2605 Extension Challenge', NB_THEME.M + 12, colY + 9, 130, 15,
    { size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });

  var probText = safeStr_(section.problem || '');
  var probH = estimateTextHeight_(probText, leftW - 32);
  addRoundRect_(slide, NB_THEME.M + 12, colY + 30, leftW - 24, probH, pal.light);
  addRect_(slide, NB_THEME.M + 12, colY + 30, 4, probH, pal.accent);
  addText_(slide, probText, NB_THEME.M + 20, colY + 36, leftW - 36, probH - 10,
    { size: NB_THEME.SZ_PROBLEM, color: pal.primary, bold: true });

  var hintY = colY + 30 + probH + 6;
  if (section.hint) {
    var hintH = Math.max(34, estimateTextHeight_(section.hint, leftW - 44) + 20);
    renderThinkBox_(slide, NB_THEME.M + 12, hintY, leftW - 24, hintH, 'Hint', section.hint, pal.light, pal.primary);
    hintY += hintH + 6;
  }

  if (typeof drawWritingLines_ === 'function') {
    drawWritingLines_(slide, NB_THEME.M + 16, hintY, leftW - 28, colH - (hintY - colY) - 8, 4, NB_THEME.LINE);
  }

  var rx = NB_THEME.M + leftW + NB_THEME.GAP;
  var halfH = Math.floor((colH - NB_THEME.GAP) / 2);

  addRoundRectWithStripe_(slide, rx, colY, rightW, halfH, pal.light, pal.stripe);
  addText_(slide, 'My Conjecture:', rx + 12, colY + 8, rightW - 24, 14,
    { size: NB_THEME.SZ_SMALL, color: pal.primary, bold: true });
  addText_(slide, 'I think\u2026', rx + 12, colY + 26, rightW - 24, 14,
    { size: NB_THEME.SZ_TINY, color: NB_THEME.MUTED });
  if (typeof drawWritingLines_ === 'function') {
    drawWritingLines_(slide, rx + 12, colY + 44, rightW - 24, halfH - 52, 3, NB_THEME.LINE);
  }

  var proofY = colY + halfH + NB_THEME.GAP;
  addRoundRectWithStripe_(slide, rx, proofY, rightW, halfH, pal.surface, pal.stripe);
  addText_(slide, 'My Proof / Justification:', rx + 12, proofY + 8, rightW - 24, 14,
    { size: NB_THEME.SZ_SMALL, color: pal.primary, bold: true });
  if (typeof drawWritingLines_ === 'function') {
    drawWritingLines_(slide, rx + 12, proofY + 26, rightW - 24, halfH - 34, 4, NB_THEME.LINE);
  }

  renderCollaborationPanel_(slide, section.discussionQs || [], pal);
  renderFooter_(slide, meta, sessionSpec, pal);
}

/* ================================================================
   SUMMARY
   ================================================================ */

function renderSummary_(slide, section, pal, meta, sessionSpec) {
  var leftW = 346, rightW = NB_LAYOUT.CW - leftW - NB_THEME.GAP;
  var colY = NB_LAYOUT.CY;
  var colH = NB_LAYOUT.CH_PANEL;

  addRoundRectWithStripe_(slide, NB_THEME.M, colY, leftW, colH, NB_THEME.WHITE, pal.accent);

  var stemText = safeStr_(section.summaryStem || 'Today I learned\u2026');
  var stemH = estimateTextHeight_(stemText, leftW - 24) + 4;
  addText_(slide, stemText, NB_THEME.M + 12, colY + 8, leftW - 24, stemH,
    { size: NB_THEME.SZ_SMALL, color: pal.primary, bold: true });

  var bulletY = colY + 8 + stemH + 4;
  var bullets = section.bullets || [];
  for (var i = 0; i < bullets.length; i++) {
    addText_(slide, '\u2022 ' + safeStr_(bullets[i]), NB_THEME.M + 12, bulletY + i * 21, leftW - 24, 20,
      { size: NB_THEME.SZ_BODY, color: NB_THEME.BODY });
  }
  var writeY = bulletY + bullets.length * 21 + 4;
  if (typeof drawWritingLines_ === 'function') {
    drawWritingLines_(slide, NB_THEME.M + 12, writeY, leftW - 24, colH - (writeY - colY) - 6, 5, NB_THEME.LINE);
  }

  var rx = NB_THEME.M + leftW + NB_THEME.GAP;
  addRoundRectWithStripe_(slide, rx, colY, rightW, colH, pal.surface, pal.stripe);
  addRect_(slide, rx + 12, colY + 8, rightW - 24, 17, pal.primary);
  addText_(slide, safeStr_(section.formulaTitle || 'Key Structure'), rx + 12, colY + 9, rightW - 24, 15,
    { size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });

  var mathLines = section.mathLines || [];
  for (var j = 0; j < mathLines.length; j++) {
    var isFirst = j === 0;
    addText_(slide, safeStr_(mathLines[j]), rx + 12, colY + 32 + j * 25, rightW - 24, 23,
      { font: isFirst ? NB_THEME.FONT_HERO : NB_THEME.FONT,
        size: isFirst ? NB_THEME.SZ_PROBLEM : NB_THEME.SZ_BODY,
        color: isFirst ? pal.primary : NB_THEME.BODY, bold: isFirst });
  }
  if (section.comparePrompt) {
    var cmpY = colY + 32 + mathLines.length * 25 + 6;
    renderThinkBox_(slide, rx + 12, cmpY, rightW - 24, 52,
      'Compare', safeStr_(section.comparePrompt), pal.light, pal.primary);
  }

  renderCollaborationPanel_(slide, section.discussionQs || [section.talkPrompt || ''], pal);
  renderFooter_(slide, meta, sessionSpec, pal);
}

/* ================================================================
   REFLECTION
   ================================================================ */

function renderReflection_(slide, section, pal, meta, sessionSpec) {
  var prompts = section.prompts || [];
  var n = Math.min(3, prompts.length);
  var cardH = Math.floor((NB_LAYOUT.CH - NB_THEME.GAP * (n - 1)) / Math.max(1, n));

  for (var i = 0; i < n; i++) {
    var y = NB_LAYOUT.CY + i * (cardH + NB_THEME.GAP);
    var ac = pal.accents[i % pal.accents.length];
    addRoundRectWithStripe_(slide, NB_THEME.M, y, NB_LAYOUT.CW, cardH, NB_THEME.WHITE, ac);
    addRect_(slide, NB_THEME.M + 10, y + Math.floor(cardH / 2) - 11, 24, 24, ac);
    addText_(slide, String(i + 1), NB_THEME.M + 10, y + Math.floor(cardH / 2) - 10, 24, 22,
      { size: NB_THEME.SZ_SMALL, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
    addText_(slide, safeStr_(prompts[i]), NB_THEME.M + 42, y + 8, NB_LAYOUT.CW - 54, 22,
      { size: NB_THEME.SZ_PROBLEM, color: pal.primary, bold: true });
    if (typeof drawWritingLines_ === 'function') {
      drawWritingLines_(slide, NB_THEME.M + 42, y + 36, NB_LAYOUT.CW - 54, cardH - 44, 3, NB_THEME.LINE);
    }
  }
  renderFooter_(slide, meta, sessionSpec, pal);
}

/* ================================================================
   SUB-RENDERERS
   ================================================================ */

function renderThinkBox_(slide, x, y, w, h, label, text, fill, accent) {
  addRoundRect_(slide, x, y, w, h, fill);
  addRect_(slide, x, y, 3, h, accent);
  var pillW = Math.min(148, w - 16);
  addRect_(slide, x + 8, y + 6, pillW, 15, accent);
  addText_(slide, label, x + 8, y + 7, pillW, 13,
    { size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
  addText_(slide, text, x + 8, y + 25, w - 16, h - 29, { size: NB_THEME.SZ_SMALL, color: NB_THEME.BODY });
}

function renderSourceRefChip_(slide, label, pal, x, y, w) {
  var text = safeStr_(label);
  if (!text) return;
  addTextShape_(slide, SlidesApp.ShapeType.RECTANGLE, text, x, y, w, 16,
    {
      fill: mixHex_(pal.primary, NB_THEME.WHITE, 0.18),
      borderColor: pal.primary,
      size: NB_THEME.SZ_TINY,
      color: pal.primary,
      bold: true,
      align: 'CENTER'
    });
}

function renderRevealBox_(slide, x, y, w, h, label, text, pal) {
  addRoundRect_(slide, x, y, w, h, pal.warm);
  addRect_(slide, x, y, 3, h, pal.accent);
  addRect_(slide, x + 8, y + 6, 70, 15, pal.accent);
  addText_(slide, label, x + 8, y + 7, 70, 13,
    { size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
  addText_(slide, text, x + 8, y + 25, w - 16, h - 29,
    { size: NB_THEME.SZ_BODY, color: NB_THEME.BODY, bold: true });
}

function renderObjectiveRow_(slide, x, y, w, h, label, text, labelColor, fill) {
  addRoundRect_(slide, x, y, w, h, fill);
  addRect_(slide, x, y, 3, h, labelColor);
  addRect_(slide, x + 8, y + 6, 140, 15, labelColor);
  addText_(slide, label, x + 8, y + 7, 140, 13,
    { size: NB_THEME.SZ_TINY, color: NB_THEME.WHITE, bold: true, align: 'CENTER' });
  addText_(slide, text, x + 8, y + 25, w - 16, h - 29, { size: NB_THEME.SZ_SMALL, color: NB_THEME.BODY });
}

function renderKernelStrip_(slide, title, kernels, x, y, w, accent, pal) {
  addRoundRect_(slide, x, y, w, 44, mixHex_(accent, NB_THEME.WHITE, 0.88));
  addText_(slide, title, x + 6, y + 3, w - 12, 13, { size: NB_THEME.SZ_TINY, color: accent, bold: true });
  for (var i = 0; i < Math.min(2, kernels.length); i++) {
    addText_(slide, safeStr_(kernels[i]), x + 6, y + 16 + i * 13, w - 12, 12,
      { size: NB_THEME.SZ_TINY, color: NB_THEME.BODY });
  }
}

/* ================================================================
   PRIMITIVES
   ================================================================ */

function addRect_(slide, x, y, w, h, fill, opacity) {
  w = Math.max(2, Number(w || 0)); h = Math.max(2, Number(h || 0));
  var s = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  s.getFill().setSolidFill(fill, typeof opacity === 'number' ? opacity : 1);
  s.getBorder().setTransparent();
  return s;
}

function addRoundRect_(slide, x, y, w, h, fill, border, opacity) {
  w = Math.max(2, Number(w || 0)); h = Math.max(2, Number(h || 0));
  var s = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, w, h);
  if (fill) s.getFill().setSolidFill(fill, typeof opacity === 'number' ? opacity : 1);
  else s.getFill().setTransparent();
  if (border) { s.getBorder().getLineFill().setSolidFill(border); s.getBorder().setWeight(0.75); }
  else s.getBorder().setTransparent();
  return s;
}

function addRoundRectWithStripe_(slide, x, y, w, h, fill, stripeColor) {
  addRoundRect_(slide, x, y, w, h, fill || NB_THEME.WHITE);
  if (stripeColor) addRect_(slide, x, y, 5, h, stripeColor);
}

function addEllipse_(slide, x, y, w, h, fill, opacity) {
  w = Math.max(2, Number(w || 0)); h = Math.max(2, Number(h || 0));
  var s = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, x, y, w, h);
  s.getFill().setSolidFill(fill, typeof opacity === 'number' ? opacity : 1);
  s.getBorder().setTransparent();
  return s;
}

function addTextShape_(slide, shapeType, text, x, y, w, h, opts) {
  opts = opts || {};
  w = Math.max(4, Number(w || 0));
  h = Math.max(4, Number(h || 0));

  var shape = slide.insertShape(shapeType || SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  if (opts.fill) shape.getFill().setSolidFill(opts.fill, typeof opts.opacity === 'number' ? opts.opacity : 1);
  else shape.getFill().setTransparent();

  if (opts.borderColor) {
    shape.getBorder().getLineFill().setSolidFill(opts.borderColor);
    shape.getBorder().setWeight(opts.borderWeight || 0.75);
  } else {
    shape.getBorder().setTransparent();
  }

  var tf = shape.getText();
  tf.setText(safeStr_(text));
  var style = tf.getTextStyle();
  style.setFontFamily(opts.font || NB_THEME.FONT);
  style.setFontSize(opts.size || NB_THEME.SZ_BODY);
  style.setForegroundColor(opts.color || NB_THEME.BODY);
  style.setBold(!!opts.bold);

  tf.getParagraphStyle().setParagraphAlignment(
    opts.align === 'CENTER' ? SlidesApp.ParagraphAlignment.CENTER :
    opts.align === 'RIGHT'  ? SlidesApp.ParagraphAlignment.END :
    SlidesApp.ParagraphAlignment.START
  );

  return shape;
}

function addText_(slide, text, x, y, w, h, opts) {
  opts = opts || {};
  w = Math.max(4, Number(w || 0));
  h = Math.max(4, Number(h || 0));

  // Use a transparent rectangle as the text container because it is a
  // reliable text-bearing shape in Slides and avoids TEXT_BOX instability.
  var shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  shape.getFill().setTransparent();
  shape.getBorder().setTransparent();

  var tf = shape.getText();
  tf.setText(safeStr_(text));
  var style = tf.getTextStyle();
  style.setFontFamily(opts.font || NB_THEME.FONT);
  style.setFontSize(opts.size || NB_THEME.SZ_BODY);
  style.setForegroundColor(opts.color || NB_THEME.BODY);
  style.setBold(!!opts.bold);

  tf.getParagraphStyle().setParagraphAlignment(
    opts.align === 'CENTER' ? SlidesApp.ParagraphAlignment.CENTER :
    opts.align === 'RIGHT'  ? SlidesApp.ParagraphAlignment.END :
    SlidesApp.ParagraphAlignment.START
  );

  return shape;
}

function insertDataUriImage_(slide, dataUri, x, y, w, h) {
  if (!dataUri) return false;
  var src = String(dataUri).trim();
  try {
    var blob;
    if (src.indexOf('data:') === 0) {
      var parts = src.split(',');
      if (parts.length < 2) return false;
      var mimeType = parts[0].indexOf('jpeg') > -1 ? 'image/jpeg' : 'image/png';
      blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), mimeType, 'nb_img');
    } else {
      blob = DriveApp.getFileById(src).getBlob();
    }
    var img = slide.insertImage(blob);
    img.setLeft(x); img.setTop(y); img.setWidth(w); img.setHeight(h);
    return true;
  } catch (e) {
    Logger.log('[insertDataUriImage_] ' + e.message);
    return false;
  }
}

/* ================================================================
   UTILITIES
   ================================================================ */

function ensureICan_(text) {
  var t = safeStr_(text).trim();
  if (!t) return 'I can explain my mathematical thinking clearly.';
  if (!/^i can\b/i.test(t)) return 'I can ' + t.charAt(0).toLowerCase() + t.slice(1);
  return t;
}

function estimateTextHeight_(text, containerW) {
  if (!text || !containerW) return 48;
  var charsPerLine = Math.floor(containerW / 7.2);
  var lines = Math.ceil(String(text).length / Math.max(1, charsPerLine));
  return Math.max(44, lines * 24 + 18);
}

function safeStr_(v) { return (v === null || v === undefined) ? '' : String(v); }

function mixHex_(h1, h2, r) {
  r = typeof r === 'number' ? r : 0.5;
  var c1 = hexToRgb_(h1 || '#FFFFFF'), c2 = hexToRgb_(h2 || '#FFFFFF');
  return rgbToHex_(Math.round(c1.r * (1 - r) + c2.r * r), Math.round(c1.g * (1 - r) + c2.g * r), Math.round(c1.b * (1 - r) + c2.b * r));
}

function hexToRgb_(hex) {
  hex = String(hex).replace('#', '');
  return { r: parseInt(hex.substring(0, 2), 16), g: parseInt(hex.substring(2, 4), 16), b: parseInt(hex.substring(4, 6), 16) };
}

function rgbToHex_(r, g, b) {
  function pad(v) { var s = Number(v).toString(16); return s.length === 1 ? '0' + s : s; }
  return '#' + pad(r) + pad(g) + pad(b);
}
