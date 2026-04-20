var NOTEBOOK_CONFIG = {
  TITLE: 'Determine the Area of Triangles',
  GRADE_BAND: 'Grade 6',
  SOURCE_NOTE: 'Source deck standard not listed; notebook follows the lesson learning targets from the slides.',
  SOURCE_FILE: '5.2 Area of Triangles.pptx',
  SLIDE_W: 960,
  SLIDE_H: 540,
  MARGIN: 28,
  HEADER_H: 56,
  FOOTER_H: 24,
  BODY_FONT: 'Arial',
  DISPLAY_FONT: 'Georgia'
};

function createFlagshipStudentNotebook() {
  selfTestNotebook_();
  var session1 = createSession1Notebook_();
  var session2 = createSession2Notebook_();
  var outputs = [session1, session2];

  for (var i = 0; i < outputs.length; i++) {
    Logger.log(outputs[i].title + ': ' + outputs[i].url);
  }

  return outputs;
}

function createSession1Notebook_() {
  return buildNotebookPresentation_(getSession1NotebookSpec_());
}

function createSession2Notebook_() {
  return buildNotebookPresentation_(getSession2NotebookSpec_());
}

function selfTestNotebook_() {
  if (typeof createFlagshipStudentNotebook !== 'function') {
    throw new Error('Missing createFlagshipStudentNotebook().');
  }
  if (typeof createSession1Notebook_ !== 'function') {
    throw new Error('Missing createSession1Notebook_().');
  }
  if (typeof createSession2Notebook_ !== 'function') {
    throw new Error('Missing createSession2Notebook_().');
  }
  if (NOTEBOOK_CONFIG.SLIDE_W <= 0 || NOTEBOOK_CONFIG.SLIDE_H <= 0) {
    throw new Error('Slide dimensions must be positive.');
  }

  var session1 = getSession1NotebookSpec_();
  var session2 = getSession2NotebookSpec_();
  if (!session1.slides || !session1.slides.length) {
    throw new Error('Session 1 notebook slides are missing.');
  }
  if (!session2.slides || !session2.slides.length) {
    throw new Error('Session 2 notebook slides are missing.');
  }

  var types = {};
  for (var i = 0; i < session1.slides.length; i++) {
    types[session1.slides[i].type] = true;
  }
  for (var j = 0; j < session2.slides.length; j++) {
    types[session2.slides[j].type] = true;
  }
  if (!types.cover || !types.vocabulary || !types.exitTicket) {
    throw new Error('Critical slide types are missing from the notebook specs.');
  }

  return true;
}

function buildNotebookPresentation_(deckSpec) {
  var presentation = SlidesApp.create(deckSpec.presentationTitle);
  var slides = presentation.getSlides();
  var firstSlide = slides && slides.length ? slides[0] : presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  clearSlide_(firstSlide);

  for (var i = 0; i < deckSpec.slides.length; i++) {
    var slide = i === 0 ? firstSlide : presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    renderNotebookSlide_(slide, deckSpec, deckSpec.slides[i], i + 1);
  }

  presentation.saveAndClose();
  var reopened = SlidesApp.openById(presentation.getId());
  return {
    id: reopened.getId(),
    title: reopened.getName(),
    url: reopened.getUrl()
  };
}

function renderNotebookSlide_(slide, deckSpec, slideSpec, pageNum) {
  var theme = deckSpec.theme;
  setSlideBackground_(slide, slideSpec.type === 'cover' ? theme.coverBg : theme.background);

  if (slideSpec.type !== 'cover') {
    addHeader_(slide, deckSpec, slideSpec, pageNum);
    addFooter_(slide, deckSpec, pageNum);
  }

  switch (slideSpec.type) {
    case 'cover':
      renderCoverSlide_(slide, deckSpec, slideSpec);
      break;
    case 'curiosity':
      renderCuriositySlide_(slide, deckSpec, slideSpec);
      break;
    case 'targets':
      renderTargetsSlide_(slide, deckSpec, slideSpec);
      break;
    case 'vocabulary':
      renderVocabularySlide_(slide, deckSpec, slideSpec);
      break;
    case 'notes':
      renderNotesSlide_(slide, deckSpec, slideSpec);
      break;
    case 'workedExample':
      renderWorkedExampleSlide_(slide, deckSpec, slideSpec);
      break;
    case 'formula':
      renderFormulaSlide_(slide, deckSpec, slideSpec);
      break;
    case 'practice':
      renderPracticeSlide_(slide, deckSpec, slideSpec);
      break;
    case 'review':
      renderReviewSlide_(slide, deckSpec, slideSpec);
      break;
    case 'application':
      renderApplicationSlide_(slide, deckSpec, slideSpec);
      break;
    case 'summary':
      renderSummarySlide_(slide, deckSpec, slideSpec);
      break;
    case 'reflection':
      renderReflectionSlide_(slide, deckSpec, slideSpec);
      break;
    case 'exitTicket':
      renderExitTicketSlide_(slide, deckSpec, slideSpec);
      break;
    default:
      throw new Error('Unsupported slide type: ' + slideSpec.type);
  }
}

function renderCoverSlide_(slide, deckSpec, slideSpec) {
  var theme = deckSpec.theme;
  addAccentOrb_(slide, 760, -40, 220, theme.accentSoft, 0.18);
  addAccentOrb_(slide, 804, 296, 160, theme.accentSoft, 0.1);
  addAccentOrb_(slide, -52, 388, 180, '#FFFFFF', 0.08);
  addRectBand_(slide, 0, 0, NOTEBOOK_CONFIG.SLIDE_W, 6, theme.accent);

  addChip_(slide, 38, 34, 156, 28, slideSpec.kicker, theme.accent, '#FFFFFF');

  addTextBlock_(slide, 40, 92, 610, 52, slideSpec.title, {
    font: NOTEBOOK_CONFIG.DISPLAY_FONT,
    size: 28,
    color: '#FFFFFF',
    bold: true
  });
  addTextBlock_(slide, 40, 150, 620, 44, slideSpec.subtitle, {
    size: 15,
    color: '#E6EEF8',
    lineSpacing: 1.15
  });

  addRoundedBox_(slide, 40, 216, 560, 174, {
    fill: theme.coverCard,
    border: theme.coverBorder,
    radius: 22
  });
  addTextBlock_(slide, 64, 236, 240, 24, 'Today\'s Focus', {
    font: NOTEBOOK_CONFIG.DISPLAY_FONT,
    size: 16,
    color: theme.primary,
    bold: true
  });
  addBulletList_(slide, 64, 272, 500, 104, slideSpec.focusBullets, {
    size: 13.5,
    color: theme.ink
  });

  addRoundedBox_(slide, 648, 86, 252, 162, {
    fill: '#FFFFFF',
    border: theme.coverBorder,
    radius: 22
  });
  addTextBlock_(slide, 672, 108, 200, 22, 'Notebook Details', {
    font: NOTEBOOK_CONFIG.DISPLAY_FONT,
    size: 15,
    color: theme.primary,
    bold: true
  });
  addMiniFact_(slide, 672, 140, 'Lesson', NOTEBOOK_CONFIG.TITLE, theme);
  addMiniFact_(slide, 672, 172, 'Session', deckSpec.sessionLabel, theme);
  addMiniFact_(slide, 672, 204, 'Source', NOTEBOOK_CONFIG.SOURCE_FILE, theme);

  addRoundedBox_(slide, 648, 268, 252, 122, {
    fill: theme.surface,
    border: theme.border,
    radius: 18
  });
  addTextBlock_(slide, 670, 286, 210, 18, 'Design Notes', {
    size: 13,
    color: theme.primary,
    bold: true
  });
  addBulletList_(slide, 670, 312, 204, 66, slideSpec.designNotes, {
    size: 11.5,
    color: theme.ink
  });

  addNameLine_(slide, 40, 448, 420, theme, 'Student');
  addNameLine_(slide, 484, 448, 184, theme, 'Date');
  addNameLine_(slide, 692, 448, 208, theme, 'Class');

  addTextBlock_(slide, 40, 500, 860, 18, NOTEBOOK_CONFIG.SOURCE_NOTE, {
    size: 10.5,
    color: '#D5E3F4',
    italic: true
  });
}

function renderCuriositySlide_(slide, deckSpec, slideSpec) {
  var theme = deckSpec.theme;
  addPageTitle_(slide, deckSpec, slideSpec);
  addDirectionsBox_(slide, 32, 86, 896, 56, 'Launch Prompt', slideSpec.prompt, theme);

  addPromptBox_(slide, 32, 158, 294, 188, 'I Notice', slideSpec.noticePrompt, theme, {
    fill: '#FFFFFF'
  });
  addLinedResponseArea_(slide, 48, 222, 262, 104, theme, 5);

  addPromptBox_(slide, 342, 158, 294, 188, 'I Wonder', slideSpec.wonderPrompt, theme, {
    fill: '#FFFFFF'
  });
  addLinedResponseArea_(slide, 358, 222, 262, 104, theme, 5);

  addPromptBox_(slide, 652, 158, 276, 90, 'Mindset Move', slideSpec.mindsetPrompt, theme, {
    fill: theme.surface
  });
  addSentenceStarterBox_(slide, 652, 262, 276, 132, slideSpec.starters, theme);

  addPromptBox_(slide, 32, 364, 604, 118, 'Reason It Out', slideSpec.reasoningPrompt, theme, {
    fill: theme.surface
  });
  addLinedResponseArea_(slide, 48, 410, 572, 58, theme, 3);

  addRoundedBox_(slide, 652, 410, 276, 58, {
    fill: '#FFFFFF',
    border: theme.border,
    radius: 16
  });
  addTextBlock_(slide, 670, 426, 240, 16, 'Talk Move', {
    size: 11,
    color: theme.primary,
    bold: true
  });
  addTextBlock_(slide, 670, 444, 240, 16, slideSpec.partnerPrompt, {
    size: 11,
    color: theme.ink
  });
}

function renderTargetsSlide_(slide, deckSpec, slideSpec) {
  var theme = deckSpec.theme;
  addPageTitle_(slide, deckSpec, slideSpec);

  addRoundedBox_(slide, 32, 92, 532, 364, {
    fill: '#FFFFFF',
    border: theme.border,
    radius: 22
  });
  addTextBlock_(slide, 54, 114, 220, 22, 'I Can Statements', {
    font: NOTEBOOK_CONFIG.DISPLAY_FONT,
    size: 16,
    color: theme.primary,
    bold: true
  });
  addBulletList_(slide, 54, 150, 486, 154, slideSpec.bullets, {
    size: 14,
    color: theme.ink
  });
  addPromptBox_(slide, 54, 320, 486, 114, 'Bridge From Prior Learning', slideSpec.bridgeText, theme, {
    fill: theme.surface
  });

  addPromptBox_(slide, 588, 92, 340, 142, 'Formula Focus', slideSpec.formulaFocus, theme, {
    fill: theme.softAccent
  });
  addEquationStrip_(slide, 608, 150, 300, 64, slideSpec.equationLines, theme);

  addPromptBox_(slide, 588, 252, 340, 108, 'Success Criteria', slideSpec.successPrompt, theme, {
    fill: '#FFFFFF'
  });
  addLinedResponseArea_(slide, 606, 302, 304, 44, theme, 2);

  addPromptBox_(slide, 588, 374, 340, 82, 'Source Note', slideSpec.sourceNote, theme, {
    fill: theme.surface
  });
}

function renderVocabularySlide_(slide, deckSpec, slideSpec) {
  var theme = deckSpec.theme;
  addPageTitle_(slide, deckSpec, slideSpec);
  addDirectionsBox_(slide, 32, 90, 896, 44, 'How to Use This Page', slideSpec.prompt, theme);
  addVocabularyTable_(slide, 32, 150, 896, 306, slideSpec.entries, theme);
}

function renderNotesSlide_(slide, deckSpec, slideSpec) {
  var theme = deckSpec.theme;
  addPageTitle_(slide, deckSpec, slideSpec);

  addPromptBox_(slide, 32, 90, 360, 74, 'Think First', slideSpec.prompt, theme, {
    fill: '#FFFFFF'
  });
  addStackedFactCards_(slide, 32, 176, 360, slideSpec.factCards, theme);
  addDirectionsBox_(slide, 32, 360, 360, 96, 'Directions', joinBulletsAsLines_(slideSpec.directions), theme);

  addRoundedBox_(slide, 420, 90, 508, 212, {
    fill: '#FFFFFF',
    border: theme.border,
    radius: 22
  });
  renderDiagramByName_(slide, slideSpec.diagram, 448, 120, 450, 160, theme);

  addPromptBox_(slide, 420, 316, 508, 66, 'Write About the Diagram', slideSpec.diagramPrompt, theme, {
    fill: theme.surface
  });
  addLinedResponseArea_(slide, 436, 392, 476, 64, theme, 3);
}

function renderWorkedExampleSlide_(slide, deckSpec, slideSpec) {
  var theme = deckSpec.theme;
  addPageTitle_(slide, deckSpec, slideSpec);
  addDirectionsBox_(slide, 32, 90, 896, 54, 'Worked Example', slideSpec.prompt, theme);

  addRoundedBox_(slide, 32, 160, 430, 296, {
    fill: '#FFFFFF',
    border: theme.border,
    radius: 22
  });
  addTextBlock_(slide, 54, 180, 176, 18, 'Givens', {
    size: 13,
    color: theme.primary,
    bold: true
  });
  addBulletList_(slide, 54, 204, 384, 82, slideSpec.givens, {
    size: 12.5,
    color: theme.ink
  });
  addTextBlock_(slide, 54, 296, 176, 18, 'Steps', {
    size: 13,
    color: theme.primary,
    bold: true
  });
  addNumberedList_(slide, 54, 320, 384, 96, slideSpec.steps, {
    size: 12.5,
    color: theme.ink
  });
  addPromptBox_(slide, 54, 420, 384, 24, 'Check Your Work', slideSpec.checkPrompt || 'Write the final labeled answer only after you finish the steps above.', theme, {
    fill: theme.softAccent
  });

  addRoundedBox_(slide, 486, 160, 442, 166, {
    fill: '#FFFFFF',
    border: theme.border,
    radius: 22
  });
  renderDiagramByName_(slide, slideSpec.diagram, 512, 186, 390, 116, theme);

  addSentenceStarterBox_(slide, 486, 340, 210, 116, slideSpec.starters, theme);
  addPromptBox_(slide, 714, 340, 214, 116, 'Explain It', slideSpec.explainPrompt, theme, {
    fill: '#FFFFFF'
  });
  addLinedResponseArea_(slide, 728, 384, 184, 56, theme, 3);
}

function renderFormulaSlide_(slide, deckSpec, slideSpec) {
  var theme = deckSpec.theme;
  addPageTitle_(slide, deckSpec, slideSpec);
  addDirectionsBox_(slide, 32, 88, 896, 48, 'Pattern to Rule', slideSpec.intro, theme);

  addEquationCard_(slide, 44, 158, 264, 118, slideSpec.equationCards[0], theme, theme.softAccent);
  addEquationCard_(slide, 348, 158, 264, 118, slideSpec.equationCards[1], theme, '#FFFFFF');
  addEquationCard_(slide, 652, 158, 264, 118, slideSpec.equationCards[2], theme, theme.surface);

  addPromptBox_(slide, 44, 298, 568, 158, 'Talk Through the Formula', slideSpec.prompt, theme, {
    fill: '#FFFFFF'
  });
  addLinedResponseArea_(slide, 60, 344, 536, 96, theme, 5);

  addPromptBox_(slide, 636, 298, 280, 78, 'Checkpoints', joinBulletsAsLines_(slideSpec.checkpoints), theme, {
    fill: '#FFFFFF'
  });
  addSentenceStarterBox_(slide, 636, 390, 280, 66, slideSpec.starters, theme);
}

function renderPracticeSlide_(slide, deckSpec, slideSpec) {
  var theme = deckSpec.theme;
  addPageTitle_(slide, deckSpec, slideSpec);
  addDirectionsBox_(slide, 32, 88, 896, 44, 'Your Turn', slideSpec.prompt, theme);

  var cards = slideSpec.problems || [];
  var cardY = 150;
  if (cards.length === 3) {
    for (var i = 0; i < cards.length; i++) {
      addProblemCard_(slide, 32 + i * 300, cardY, 276, 148, cards[i], theme);
    }
    addSentenceStarterBox_(slide, 32, 318, 274, 118, slideSpec.support, theme);
    addPromptBox_(slide, 324, 318, 604, 118, 'Show Your Thinking', slideSpec.workspacePrompt, theme, {
      fill: '#FFFFFF'
    });
    addLinedResponseArea_(slide, 340, 364, 572, 58, theme, 3);
  } else {
    for (var j = 0; j < cards.length; j++) {
      addProblemCard_(slide, 32 + j * 448, cardY, 416, 164, cards[j], theme);
    }
    addSentenceStarterBox_(slide, 32, 334, 290, 118, slideSpec.support, theme);
    addPromptBox_(slide, 340, 334, 588, 118, 'Write It Clearly', slideSpec.workspacePrompt, theme, {
      fill: '#FFFFFF'
    });
    addLinedResponseArea_(slide, 356, 380, 556, 56, theme, 3);
  }
}

function renderReviewSlide_(slide, deckSpec, slideSpec) {
  var theme = deckSpec.theme;
  addPageTitle_(slide, deckSpec, slideSpec);
  addDirectionsBox_(slide, 32, 88, 896, 50, 'Quick Review', slideSpec.prompt, theme);

  addPromptBox_(slide, 32, 154, 420, 152, 'What Could the Question Be?', slideSpec.openResponse, theme, {
    fill: '#FFFFFF'
  });
  addLinedResponseArea_(slide, 48, 210, 388, 82, theme, 4);

  addPromptBox_(slide, 472, 154, 214, 152, 'Partner Mindset', slideSpec.partnerPrompt, theme, {
    fill: theme.surface
  });
  addSentenceStarterBox_(slide, 706, 154, 222, 152, slideSpec.starters, theme);

  for (var i = 0; i < slideSpec.cards.length; i++) {
    addReviewCard_(slide, 32 + i * 224, 328, 204, 126, slideSpec.cards[i], theme);
  }
}

function renderApplicationSlide_(slide, deckSpec, slideSpec) {
  var theme = deckSpec.theme;
  addPageTitle_(slide, deckSpec, slideSpec);

  addPromptBox_(slide, 32, 88, 468, 162, slideSpec.contextTitle, slideSpec.context, theme, {
    fill: '#FFFFFF'
  });
  addPromptBox_(slide, 32, 266, 468, 94, 'Knowns and Unknowns', joinBulletsAsLines_(slideSpec.knowns), theme, {
    fill: theme.surface
  });
  addPromptBox_(slide, 32, 376, 468, 80, 'Precision Check', slideSpec.precisionPrompt, theme, {
    fill: '#FFFFFF'
  });

  addRoundedBox_(slide, 524, 88, 404, 206, {
    fill: '#FFFFFF',
    border: theme.border,
    radius: 22
  });
  renderDiagramByName_(slide, slideSpec.diagram, 548, 116, 356, 156, theme);

  addPromptBox_(slide, 524, 308, 404, 62, 'Solve', slideSpec.solvePrompt, theme, {
    fill: theme.surface
  });
  addLinedResponseArea_(slide, 540, 382, 372, 74, theme, 4);
}

function renderSummarySlide_(slide, deckSpec, slideSpec) {
  var theme = deckSpec.theme;
  addPageTitle_(slide, deckSpec, slideSpec);

  for (var i = 0; i < slideSpec.cards.length; i++) {
    addSummaryCard_(slide, 32 + i * 224, 100, 204, 160, slideSpec.cards[i], theme);
  }

  addPromptBox_(slide, 32, 280, 430, 176, 'Write the Big Idea', slideSpec.bigIdeaPrompt, theme, {
    fill: '#FFFFFF'
  });
  addLinedResponseArea_(slide, 48, 334, 398, 108, theme, 5);

  addPromptBox_(slide, 486, 280, 442, 82, 'Learning Targets Revisited', joinBulletsAsLines_(slideSpec.targets), theme, {
    fill: theme.surface
  });
  addSentenceStarterBox_(slide, 486, 376, 442, 80, slideSpec.starters, theme);
}

function renderReflectionSlide_(slide, deckSpec, slideSpec) {
  var theme = deckSpec.theme;
  addPageTitle_(slide, deckSpec, slideSpec);

  addPromptBox_(slide, 32, 90, 560, 118, 'Reflect', slideSpec.prompt, theme, {
    fill: '#FFFFFF'
  });
  addLinedResponseArea_(slide, 48, 144, 528, 50, theme, 3);

  addRoundedBox_(slide, 616, 90, 312, 176, {
    fill: theme.surface,
    border: theme.border,
    radius: 22
  });
  addTextBlock_(slide, 638, 112, 180, 20, 'Success Checklist', {
    font: NOTEBOOK_CONFIG.DISPLAY_FONT,
    size: 15,
    color: theme.primary,
    bold: true
  });
  addCheckList_(slide, 638, 146, 256, 96, slideSpec.checklist, theme);

  addPromptBox_(slide, 32, 228, 896, 84, 'Sentence Starter', slideSpec.starterPrompt, theme, {
    fill: theme.softAccent
  });

  addPromptBox_(slide, 32, 330, 430, 126, 'What I Understand Now', slideSpec.nowPrompt, theme, {
    fill: '#FFFFFF'
  });
  addLinedResponseArea_(slide, 48, 382, 398, 60, theme, 3);

  addPromptBox_(slide, 486, 330, 442, 126, 'What I Still Need to Practice', slideSpec.nextPrompt, theme, {
    fill: '#FFFFFF'
  });
  addLinedResponseArea_(slide, 502, 382, 410, 60, theme, 3);
}

function renderExitTicketSlide_(slide, deckSpec, slideSpec) {
  var theme = deckSpec.theme;
  addPageTitle_(slide, deckSpec, slideSpec);
  addDirectionsBox_(slide, 32, 90, 896, 42, 'Exit Ticket', slideSpec.prompt, theme);

  for (var i = 0; i < slideSpec.questions.length; i++) {
    addExitQuestionCard_(slide, 32 + i * 300, 152, 276, 154, i + 1, slideSpec.questions[i], theme);
  }

  addPromptBox_(slide, 32, 326, 428, 130, 'Precision Reflection', slideSpec.precisionReflection, theme, {
    fill: '#FFFFFF'
  });
  addLinedResponseArea_(slide, 48, 378, 396, 64, theme, 3);

  addSentenceStarterBox_(slide, 484, 326, 208, 130, slideSpec.starters, theme);
  addPromptBox_(slide, 712, 326, 216, 130, 'Check Before You Go', joinBulletsAsLines_(slideSpec.checks), theme, {
    fill: theme.surface
  });
}

function addHeader_(slide, deckSpec, slideSpec, pageNum) {
  var theme = deckSpec.theme;
  addRectBand_(slide, 0, 0, NOTEBOOK_CONFIG.SLIDE_W, NOTEBOOK_CONFIG.HEADER_H, '#FFFFFF');
  addRectBand_(slide, 0, 0, NOTEBOOK_CONFIG.SLIDE_W, 4, theme.accent);
  addRectBand_(slide, 0, 4, 8, NOTEBOOK_CONFIG.HEADER_H - 4, theme.accent);

  addTextBlock_(slide, 28, 12, 510, 20, slideSpec.title, {
    size: 18,
    color: theme.primary,
    bold: true
  });
  addTextBlock_(slide, 28, 31, 520, 14, slideSpec.kicker || deckSpec.sessionFocus, {
    size: 10.5,
    color: theme.muted
  });

  addChip_(slide, 760, 12, 172, 28, deckSpec.sessionLabel + ' | ' + pageNum, theme.accent, '#FFFFFF');
}

function addFooter_(slide, deckSpec) {
  var theme = deckSpec.theme;
  var y = NOTEBOOK_CONFIG.SLIDE_H - NOTEBOOK_CONFIG.FOOTER_H;
  addRectBand_(slide, 0, y, NOTEBOOK_CONFIG.SLIDE_W, NOTEBOOK_CONFIG.FOOTER_H, '#FFFFFF');
  addRectBand_(slide, 0, y, NOTEBOOK_CONFIG.SLIDE_W, 1, theme.border);
  addTextBlock_(slide, 28, y + 6, 390, 12, NOTEBOOK_CONFIG.TITLE + ' | ' + NOTEBOOK_CONFIG.GRADE_BAND, {
    size: 9.5,
    color: theme.muted
  });
  addTextBlock_(slide, 468, y + 6, 464, 12, NOTEBOOK_CONFIG.SOURCE_NOTE, {
    size: 8.5,
    color: theme.muted,
    align: SlidesApp.ParagraphAlignment.RIGHT
  });
}

function addPageTitle_(slide, deckSpec, slideSpec) {
  var theme = deckSpec.theme;
  addTextBlock_(slide, 32, 88, 520, 26, slideSpec.title, {
    font: NOTEBOOK_CONFIG.DISPLAY_FONT,
    size: 22,
    color: theme.primary,
    bold: true
  });
}

function addRoundedBox_(slide, x, y, w, h, options) {
  var shape = safeAddShape_(slide, SlidesApp.ShapeType.ROUNDED_RECTANGLE, x, y, w, h);
  try {
    shape.getAdjustments()[0] = options.radius ? Math.min(0.6, options.radius / 100) : 0.22;
  } catch (e) {}
  applyFillAndBorder_(shape, options.fill, options.border, options.borderWeight || 1.2, options.transparency || 0);
  return shape;
}

function addRectBand_(slide, x, y, w, h, fill) {
  var shape = safeAddShape_(slide, SlidesApp.ShapeType.RECTANGLE, x, y, w, h);
  applyFillAndBorder_(shape, fill, fill, 0, 0);
  return shape;
}

function addChip_(slide, x, y, w, h, text, fill, color) {
  addRoundedBox_(slide, x, y, w, h, {
    fill: fill,
    border: fill,
    radius: 16
  });
  addTextBlock_(slide, x + 8, y + 6, w - 16, h - 12, text, {
    size: 11,
    color: color,
    bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
}

function addMiniFact_(slide, x, y, label, value, theme) {
  addTextBlock_(slide, x, y, 64, 14, label + ':', {
    size: 10,
    color: theme.muted,
    bold: true
  });
  addTextBlock_(slide, x + 68, y, 150, 14, value, {
    size: 10,
    color: theme.ink
  });
}

function addPromptBox_(slide, x, y, w, h, label, body, theme, options) {
  options = options || {};
  addRoundedBox_(slide, x, y, w, h, {
    fill: options.fill || '#FFFFFF',
    border: options.border || theme.border,
    radius: 18
  });
  if (h <= 70) {
    var labelW = Math.min(138, Math.max(84, Math.round(w * 0.3)));
    addTextBlock_(slide, x + 16, y + 10, labelW, h - 20, label, {
      size: 11.5,
      color: theme.primary,
      bold: true
    });
    addTextBlock_(slide, x + 16 + labelW, y + 10, w - 32 - labelW, h - 20, body, {
      size: 11.5,
      color: theme.ink,
      lineSpacing: 1.08
    });
    return;
  }

  addTextBlock_(slide, x + 16, y + 14, w - 32, 18, label, {
    size: 13,
    color: theme.primary,
    bold: true
  });
  addTextBlock_(slide, x + 16, y + 38, w - 32, h - 52, body, {
    size: 12,
    color: theme.ink,
    lineSpacing: 1.12
  });
}

function addDirectionsBox_(slide, x, y, w, h, label, body, theme) {
  addPromptBox_(slide, x, y, w, h, label, body, theme, {
    fill: theme.surface,
    border: theme.border
  });
}

function addSentenceStarterBox_(slide, x, y, w, h, starters, theme) {
  addRoundedBox_(slide, x, y, w, h, {
    fill: theme.softAccent,
    border: theme.border,
    radius: 18
  });
  addTextBlock_(slide, x + 16, y + 12, w - 32, 18, 'Sentence Starters', {
    size: 12.5,
    color: theme.primary,
    bold: true
  });
  addBulletList_(slide, x + 16, y + 36, w - 32, h - 48, starters, {
    size: 11.5,
    color: theme.ink
  });
}

function addVocabularyTable_(slide, x, y, w, h, entries, theme) {
  addRoundedBox_(slide, x, y, w, h, {
    fill: '#FFFFFF',
    border: theme.border,
    radius: 20
  });
  var col1 = 148;
  var col2 = 252;
  var col3 = 276;
  var col4 = w - col1 - col2 - col3 - 24;
  var rowH = (h - 58) / entries.length;

  addHeaderCell_(slide, x + 12, y + 14, col1, 24, 'Word', theme);
  addHeaderCell_(slide, x + 12 + col1, y + 14, col2, 24, 'Student-Friendly Definition', theme);
  addHeaderCell_(slide, x + 12 + col1 + col2, y + 14, col3, 24, 'Lesson Example', theme);
  addHeaderCell_(slide, x + 12 + col1 + col2 + col3, y + 14, col4, 24, 'Visual Cue', theme);

  for (var i = 0; i < entries.length; i++) {
    var rowY = y + 44 + i * rowH;
    addRowDivider_(slide, x + 12, rowY, w - 24, theme.border);
    addTextBlock_(slide, x + 18, rowY + 10, col1 - 18, rowH - 12, entries[i].word, {
      size: 11.5,
      color: theme.primary,
      bold: true
    });
    addTextBlock_(slide, x + 12 + col1 + 8, rowY + 10, col2 - 16, rowH - 12, entries[i].definition, {
      size: 10.5,
      color: theme.ink
    });
    addTextBlock_(slide, x + 12 + col1 + col2 + 8, rowY + 10, col3 - 16, rowH - 12, entries[i].example, {
      size: 10.5,
      color: theme.ink
    });
    addTextBlock_(slide, x + 12 + col1 + col2 + col3 + 8, rowY + 10, col4 - 16, rowH - 12, entries[i].visual, {
      size: 10.5,
      color: theme.ink
    });
  }
}

function addHeaderCell_(slide, x, y, w, h, text, theme) {
  addRoundedBox_(slide, x, y, w, h, {
    fill: theme.primary,
    border: theme.primary,
    radius: 14
  });
  addTextBlock_(slide, x + 8, y + 4, w - 16, h - 8, text, {
    size: 10.5,
    color: '#FFFFFF',
    bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
}

function addStackedFactCards_(slide, x, y, w, facts, theme) {
  for (var i = 0; i < facts.length; i++) {
    addRoundedBox_(slide, x, y + i * 56, w, 46, {
      fill: i % 2 === 0 ? '#FFFFFF' : theme.surface,
      border: theme.border,
      radius: 16
    });
    addTextBlock_(slide, x + 16, y + 13 + i * 56, w - 32, 18, facts[i], {
      size: 12,
      color: theme.ink
    });
  }
}

function addEquationStrip_(slide, x, y, w, h, lines, theme) {
  for (var i = 0; i < lines.length; i++) {
    addRoundedBox_(slide, x + i * 94, y, 82, h, {
      fill: '#FFFFFF',
      border: theme.border,
      radius: 14
    });
    addTextBlock_(slide, x + i * 94 + 8, y + 20, 66, 26, lines[i], {
      size: 15,
      color: theme.primary,
      bold: true,
      align: SlidesApp.ParagraphAlignment.CENTER
    });
  }
}

function addEquationCard_(slide, x, y, w, h, card, theme, fill) {
  addRoundedBox_(slide, x, y, w, h, {
    fill: fill,
    border: theme.border,
    radius: 20
  });
  addTextBlock_(slide, x + 18, y + 18, w - 36, 18, card.label, {
    size: 12,
    color: theme.muted,
    bold: true
  });
  addTextBlock_(slide, x + 18, y + 46, w - 36, 28, card.equation, {
    font: NOTEBOOK_CONFIG.DISPLAY_FONT,
    size: 22,
    color: theme.primary,
    bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
  addTextBlock_(slide, x + 18, y + 84, w - 36, 20, card.note, {
    size: 11,
    color: theme.ink,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
}

function addProblemCard_(slide, x, y, w, h, problem, theme) {
  addRoundedBox_(slide, x, y, w, h, {
    fill: '#FFFFFF',
    border: theme.border,
    radius: 18
  });
  addChip_(slide, x + 16, y + 14, 86, 24, problem.label, theme.accent, '#FFFFFF');
  addTextBlock_(slide, x + 16, y + 52, w - 32, h - 66, problem.prompt, {
    size: 11.5,
    color: theme.ink,
    lineSpacing: 1.08
  });
}

function addReviewCard_(slide, x, y, w, h, card, theme) {
  addRoundedBox_(slide, x, y, w, h, {
    fill: '#FFFFFF',
    border: theme.border,
    radius: 18
  });
  addTextBlock_(slide, x + 14, y + 14, w - 28, 18, card.title, {
    size: 12,
    color: theme.primary,
    bold: true
  });
  addTextBlock_(slide, x + 14, y + 40, w - 28, h - 54, card.body, {
    size: 11,
    color: theme.ink,
    lineSpacing: 1.1
  });
}

function addSummaryCard_(slide, x, y, w, h, card, theme) {
  addRoundedBox_(slide, x, y, w, h, {
    fill: card.fill || '#FFFFFF',
    border: theme.border,
    radius: 18
  });
  addTextBlock_(slide, x + 14, y + 14, w - 28, 18, card.title, {
    size: 12,
    color: theme.primary,
    bold: true
  });
  addTextBlock_(slide, x + 14, y + 40, w - 28, h - 52, card.body, {
    size: 11,
    color: theme.ink,
    lineSpacing: 1.08
  });
}

function addExitQuestionCard_(slide, x, y, w, h, num, question, theme) {
  addRoundedBox_(slide, x, y, w, h, {
    fill: '#FFFFFF',
    border: theme.border,
    radius: 18
  });
  addChip_(slide, x + 16, y + 14, 34, 24, String(num), theme.accent, '#FFFFFF');
  addTextBlock_(slide, x + 60, y + 16, w - 76, 34, question, {
    size: 11.5,
    color: theme.ink,
    lineSpacing: 1.08
  });
  addLinedResponseArea_(slide, x + 16, y + 68, w - 32, 70, theme, 4);
}

function addNameLine_(slide, x, y, w, theme, label) {
  addTextBlock_(slide, x, y, 64, 16, label + ':', {
    size: 11,
    color: '#FFFFFF',
    bold: true
  });
  var line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x + 64, y + 12, x + w, y + 12);
  line.getLineFill().setSolidFill('#FFFFFF');
  line.setWeight(1);
}

function addLinedResponseArea_(slide, x, y, w, h, theme, lineCount) {
  addRoundedBox_(slide, x, y, w, h, {
    fill: '#FFFFFF',
    border: theme.border,
    radius: 14
  });
  var inset = 12;
  var top = y + 18;
  var usable = h - 26;
  var gap = lineCount > 1 ? usable / (lineCount - 1) : usable;
  for (var i = 0; i < lineCount; i++) {
    var ly = top + i * gap;
    var line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x + inset, ly, x + w - inset, ly);
    line.getLineFill().setSolidFill(theme.line);
    line.setWeight(0.9);
  }
}

function addCheckList_(slide, x, y, w, h, items, theme) {
  var rowH = h / items.length;
  for (var i = 0; i < items.length; i++) {
    var box = safeAddShape_(slide, SlidesApp.ShapeType.RECTANGLE, x, y + i * rowH + 2, 10, 10);
    applyFillAndBorder_(box, '#FFFFFF', theme.border, 1, 0);
    addTextBlock_(slide, x + 18, y + i * rowH - 2, w - 18, rowH + 6, items[i], {
      size: 11,
      color: theme.ink
    });
  }
}

function addBulletList_(slide, x, y, w, h, items, options) {
  var box = safeAddTextBox_(slide, x, y, w, h, '');
  var tf = box.getText();
  tf.setText('');
  for (var i = 0; i < items.length; i++) {
    if (i > 0) {
      tf.appendText('\n');
    }
    var start = tf.asString().length;
    tf.appendText('• ' + items[i]);
    var end = tf.asString().length - 1;
    applyTextStyle_(tf, start, end, options);
  }
}

function addNumberedList_(slide, x, y, w, h, items, options) {
  var box = safeAddTextBox_(slide, x, y, w, h, '');
  var tf = box.getText();
  tf.setText('');
  for (var i = 0; i < items.length; i++) {
    if (i > 0) {
      tf.appendText('\n');
    }
    var start = tf.asString().length;
    tf.appendText((i + 1) + '. ' + items[i]);
    var end = tf.asString().length - 1;
    applyTextStyle_(tf, start, end, options);
  }
}

function addTextBlock_(slide, x, y, w, h, text, options) {
  var box = safeAddTextBox_(slide, x, y, w, h, text || '');
  var range = box.getText();
  applyTextStyle_(range, 0, Math.max(0, range.asString().length - 1), options || {});
  if (options && options.align) {
    try {
      range.getParagraphStyle().setParagraphAlignment(options.align);
    } catch (e) {}
  }
  return box;
}

function applyTextStyle_(textRange, start, end, options) {
  if (!textRange) return;
  var actualEnd = typeof end === 'number' ? end : Math.max(0, textRange.asString().length - 1);
  if (actualEnd < start) return;
  var style = textRange.getRange(start, actualEnd).getTextStyle();
  style.setFontFamily(options.font || NOTEBOOK_CONFIG.BODY_FONT);
  style.setFontSize(options.size || 12);
  style.setForegroundColor(options.color || '#1F2937');
  style.setBold(Boolean(options.bold));
  style.setItalic(Boolean(options.italic));
  if (options.underline) {
    style.setUnderline(true);
  }
}

function applyFillAndBorder_(shape, fill, border, borderWeight, transparency) {
  if (!shape) return;
  shape.getFill().setSolidFill(fill || '#FFFFFF', transparency || 0);
  if (border) {
    shape.getBorder().getLineFill().setSolidFill(border);
    shape.getBorder().setWeight(borderWeight || 1);
  } else {
    shape.getBorder().getLineFill().setSolidFill(fill || '#FFFFFF', 100);
  }
}

function safeAddShape_(slide, shapeType, x, y, w, h) {
  return slide.insertShape(shapeType, x, y, w, h);
}

function safeAddTextBox_(slide, x, y, w, h, text) {
  var box = slide.insertTextBox(text || '', x, y, w, h);
  var textRange = box.getText();
  try {
    textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.LEFT);
  } catch (e) {}
  return box;
}

function clearSlide_(slide) {
  var elements = slide.getPageElements();
  for (var i = elements.length - 1; i >= 0; i--) {
    elements[i].remove();
  }
}

function setSlideBackground_(slide, color) {
  slide.getBackground().setSolidFill(color);
}

function addAccentOrb_(slide, x, y, size, fill, transparency) {
  var orb = safeAddShape_(slide, SlidesApp.ShapeType.ELLIPSE, x, y, size, size);
  applyFillAndBorder_(orb, fill, fill, 0, Math.round((transparency || 0) * 100));
}

function addRowDivider_(slide, x, y, w, color) {
  var line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x, y, x + w, y);
  line.getLineFill().setSolidFill(color);
  line.setWeight(0.75);
}

function renderDiagramByName_(slide, name, x, y, w, h, theme) {
  switch (name) {
    case 'triangleA':
      drawTriangleADiagram_(slide, x, y, w, h, theme);
      break;
    case 'parallelogramBridge':
      drawParallelogramBridgeDiagram_(slide, x, y, w, h, theme);
      break;
    case 'triangleB':
      drawTriangleBDiagram_(slide, x, y, w, h, theme);
      break;
    case 'bermuda':
      drawBermudaDiagram_(slide, x, y, w, h, theme);
      break;
    case 'cabin':
      drawCabinDiagram_(slide, x, y, w, h, theme);
      break;
    default:
      addPromptBox_(slide, x, y, w, h, 'Diagram Space', 'Use this space to sketch the figure from the source lesson.', theme, {
        fill: '#FFFFFF'
      });
  }
}

function drawTriangleADiagram_(slide, x, y, w, h, theme) {
  drawTriangleFrame_(slide, x, y, w, h, theme);
  var left = x + 46;
  var bottom = y + h - 30;
  var top = y + 24;
  var right = x + w - 36;
  var heightX = x + 152;
  drawLine_(slide, left, bottom, right, bottom, theme.primary, 2);
  drawLine_(slide, left, bottom, heightX, top + 18, theme.primary, 2);
  drawLine_(slide, heightX, top + 18, right, bottom, theme.primary, 2);
  drawDashedLine_(slide, heightX, top + 18, heightX, bottom, theme.accent, 1.4);
  drawRightAngle_(slide, heightX, bottom, theme.accent);
  addTextBlock_(slide, heightX - 22, y + 48, 54, 16, '15 cm', {
    size: 11,
    color: theme.primary,
    bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
  addTextBlock_(slide, x + 170, bottom + 4, 90, 16, '30 cm', {
    size: 11,
    color: theme.primary,
    bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
  addChip_(slide, x + w - 102, y + 10, 82, 22, 'Triangle A', theme.accent, '#FFFFFF');
}

function drawParallelogramBridgeDiagram_(slide, x, y, w, h, theme) {
  drawTriangleFrame_(slide, x, y, w, h, theme);
  var left = x + 74;
  var bottom = y + h - 26;
  var top = y + 30;
  var right = x + w - 80;
  var shift = 70;

  drawLine_(slide, left + shift, top, right + shift, top, theme.primary, 2);
  drawLine_(slide, left, bottom, right, bottom, theme.primary, 2);
  drawLine_(slide, left, bottom, left + shift, top, theme.primary, 2);
  drawLine_(slide, right, bottom, right + shift, top, theme.primary, 2);
  drawLine_(slide, left + shift, top, right, bottom, theme.accent, 1.5);
  drawDashedLine_(slide, left + shift + 42, top, left + shift + 42, bottom, theme.accent, 1.3);
  drawRightAngle_(slide, left + shift + 42, bottom, theme.accent);

  addTextBlock_(slide, x + 16, y + 12, 120, 16, '2 copies make a parallelogram', {
    size: 10.5,
    color: theme.primary,
    bold: true
  });
  addTextBlock_(slide, left + shift + 22, y + 52, 60, 16, '15 cm', {
    size: 10.5,
    color: theme.primary,
    bold: true
  });
  addTextBlock_(slide, x + 180, bottom + 4, 96, 16, 'base = 30 cm', {
    size: 10.5,
    color: theme.primary,
    bold: true
  });
  addTextBlock_(slide, x + 20, y + h - 24, w - 40, 14, 'Parallelogram area = 30 × 15 = 450; one triangle is half of 450.', {
    size: 10.5,
    color: theme.ink,
    italic: true,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
}

function drawTriangleBDiagram_(slide, x, y, w, h, theme) {
  drawTriangleFrame_(slide, x, y, w, h, theme);
  var left = x + 86;
  var bottom = y + h - 28;
  var top = y + 18;
  var right = x + w - 56;
  drawLine_(slide, left, bottom, left, top, theme.primary, 2);
  drawLine_(slide, left, bottom, right, bottom, theme.primary, 2);
  drawLine_(slide, left, top, right, bottom, theme.primary, 2);
  drawRightAngle_(slide, left, bottom, theme.accent);
  addTextBlock_(slide, left - 10, y + 44, 54, 16, '14 in.', {
    size: 11,
    color: theme.primary,
    bold: true
  });
  addTextBlock_(slide, x + 188, bottom + 4, 72, 16, '10 in.', {
    size: 11,
    color: theme.primary,
    bold: true
  });
  addChip_(slide, x + w - 102, y + 10, 82, 22, 'Triangle B', theme.accent, '#FFFFFF');
}

function drawBermudaDiagram_(slide, x, y, w, h, theme) {
  drawTriangleFrame_(slide, x, y, w, h, theme);
  var bermudaX = x + w - 76;
  var bermudaY = y + 26;
  var miamiX = x + 80;
  var miamiY = y + 88;
  var sanJuanX = x + w - 96;
  var sanJuanY = y + h - 24;
  drawLine_(slide, miamiX, miamiY, bermudaX, bermudaY, theme.primary, 1.8);
  drawLine_(slide, bermudaX, bermudaY, sanJuanX, sanJuanY, theme.primary, 1.8);
  drawLine_(slide, sanJuanX, sanJuanY, miamiX, miamiY, theme.primary, 1.8);
  drawDashedLine_(slide, miamiX, miamiY, x + w - 160, y + h / 2, theme.accent, 1.4);
  drawDashedLine_(slide, x + w - 160, y + h / 2, x + w - 126, y + 56, theme.accent, 1.4);
  drawRightAngle_(slide, x + w - 126, y + 56, theme.accent);
  addLabelTag_(slide, miamiX - 28, miamiY - 12, 56, 20, 'Miami', theme);
  addLabelTag_(slide, bermudaX - 34, bermudaY - 20, 68, 20, 'Bermuda', theme);
  addLabelTag_(slide, sanJuanX - 34, sanJuanY + 2, 72, 20, 'San Juan', theme);
  addLabelTag_(slide, x + w - 92, y + 12, 92, 22, 'A = 437,409', theme);
  addLabelTag_(slide, x + w - 86, y + h - 58, 72, 22, 'h = 917', theme);
  addTextBlock_(slide, x + 120, y + 96, 86, 16, 'b = ?', {
    size: 12,
    color: theme.accent,
    bold: true
  });
}

function drawCabinDiagram_(slide, x, y, w, h, theme) {
  drawTriangleFrame_(slide, x, y, w, h, theme);
  var left = x + 46;
  var right = x + w - 42;
  var bottom = y + h - 26;
  var peakX = x + w / 2;
  var peakY = y + 24;
  drawLine_(slide, left, bottom, peakX, peakY, theme.accent, 2.4);
  drawLine_(slide, peakX, peakY, right, bottom, theme.accent, 2.4);
  drawLine_(slide, left, bottom, right, bottom, theme.accent, 2.4);
  drawDashedLine_(slide, peakX, peakY, peakX, bottom, theme.primary, 1.3);
  drawRightAngle_(slide, peakX, bottom, theme.primary);
  addLabelTag_(slide, peakX + 14, y + 30, 82, 22, '23 3/4 ft', theme);
  addLabelTag_(slide, x + w / 2 - 40, bottom + 4, 80, 22, '34 ft', theme);
}

function drawTriangleFrame_(slide, x, y, w, h, theme) {
  addRoundedBox_(slide, x, y, w, h, {
    fill: theme.diagramBg,
    border: theme.border,
    radius: 18
  });
}

function drawLine_(slide, x1, y1, x2, y2, color, weight) {
  var line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x1, y1, x2, y2);
  line.getLineFill().setSolidFill(color);
  line.setWeight(weight || 1.2);
  return line;
}

function drawDashedLine_(slide, x1, y1, x2, y2, color, weight) {
  var line = drawLine_(slide, x1, y1, x2, y2, color, weight);
  try {
    line.setDashStyle(SlidesApp.DashStyle.DASH);
  } catch (e) {}
  return line;
}

function drawRightAngle_(slide, x, y, color) {
  drawLine_(slide, x, y, x + 12, y, color, 1.2);
  drawLine_(slide, x + 12, y, x + 12, y - 12, color, 1.2);
  drawLine_(slide, x + 12, y - 12, x, y - 12, color, 1.2);
}

function addLabelTag_(slide, x, y, w, h, text, theme) {
  addRoundedBox_(slide, x, y, w, h, {
    fill: '#FFFFFF',
    border: theme.accent,
    radius: 12
  });
  addTextBlock_(slide, x + 4, y + 4, w - 8, h - 8, text, {
    size: 10,
    color: theme.primary,
    bold: true,
    align: SlidesApp.ParagraphAlignment.CENTER
  });
}

function joinBulletsAsLines_(items) {
  return items.join('\n');
}

function getSession1NotebookSpec_() {
  return {
    presentationTitle: 'Determine the Area of Triangles - Session 1 Student Notebook',
    sessionLabel: 'Session 1',
    sessionFocus: 'Launch, build the triangle area formula, and practice with precision.',
    theme: createTheme_(1),
    slides: [
      {
        type: 'cover',
        kicker: 'Session 1 Student Notebook',
        title: 'Determine the Area of Triangles',
        subtitle: 'Compose triangles into parallelograms, use base and height with precision, and explain why the formula A = 1/2 bh works.',
        focusBullets: [
          'Start with a launch that gets you noticing and wondering about triangle area.',
          'Use Triangle A to connect triangle area to the area of a parallelogram.',
          'Use Triangle B to apply the formula to a right triangle and name the correct square units.'
        ],
        designNotes: [
          'Visible writing space on every page',
          'Sentence supports built in',
          'Source examples preserved from slides 1–25'
        ]
      },
      {
        type: 'curiosity',
        title: 'Launch: This or That?',
        kicker: 'Be Curious',
        prompt: 'The source lesson begins by asking, “Would you rather have THIS or THAT box of gold coins?” Use the launch image and the triangle examples from the deck to compare what seems more reasonable.',
        noticePrompt: 'What do you notice about the triangle figures, labels, or measurements shown in the lesson visuals?',
        wonderPrompt: 'What do you wonder about base, height, and how a triangle’s area compares to a rectangle or parallelogram?',
        mindsetPrompt: 'How do you establish a positive relationship with a classmate in math class while you disagree or revise an idea?',
        reasoningPrompt: 'Explain how you could decide whether a triangle-area answer is reasonable before you calculate it exactly.',
        partnerPrompt: 'I can listen, restate, and build on my partner’s idea.',
        starters: [
          'I notice that ...',
          'I wonder why ...',
          'A reasonable answer might be ... because ...'
        ]
      },
      {
        type: 'targets',
        title: 'Learning Target',
        kicker: 'Session 1 Goals',
        bullets: [
          'I can find the area of a triangle by composing it into a parallelogram.',
          'I can use the formula A = 1/2 bh to calculate the area of a triangle.',
          'I can attend to precision by naming the correct square units in my answer.'
        ],
        bridgeText: 'We already know how to find the area of a parallelogram. Today we use that structure to discover why the area of a triangle is half of a related parallelogram.',
        formulaFocus: 'The same base and height can describe both the triangle and the parallelogram built from two copies of that triangle.',
        equationLines: ['b × h', '÷ 2', 'A = 1/2 bh'],
        successPrompt: 'Write one thing you must look for before you use the triangle area formula.',
        sourceNote: 'The source deck ends with the same learning targets and uses the Triangle A and Triangle B examples to build toward them.'
      },
      {
        type: 'vocabulary',
        title: 'Academic Vocabulary',
        kicker: 'Words That Support the Math',
        prompt: 'Use the definitions and examples to help you explain your reasoning in complete, precise sentences.',
        entries: [
          {
            word: 'area',
            definition: 'the amount of surface inside a figure',
            example: 'Triangle A has an area of 225 square centimeters.',
            visual: 'Count or imagine square units filling the inside.'
          },
          {
            word: 'base',
            definition: 'the side used with the height in the area formula',
            example: 'Triangle A uses a base of 30 cm.',
            visual: 'The bottom segment in the diagram.'
          },
          {
            word: 'height',
            definition: 'the perpendicular distance from the base to the opposite vertex',
            example: 'Triangle A has a height of 15 cm.',
            visual: 'A dashed segment meeting the base at a right angle.'
          },
          {
            word: 'parallelogram',
            definition: 'a four-sided figure with both pairs of opposite sides parallel',
            example: 'Two copies of Triangle A can form a parallelogram.',
            visual: 'A slanted rectangle-like shape built from 2 congruent triangles.'
          },
          {
            word: 'square units',
            definition: 'units used to measure area',
            example: 'Square centimeters and square inches are both area units.',
            visual: 'Units are written with a squared label, such as cm² or in².'
          }
        ]
      },
      {
        type: 'notes',
        title: 'Guided Notes: Triangle A',
        kicker: 'Base, Height, and What We Know',
        prompt: 'What information do you know about Triangle A from the lesson diagram?',
        factCards: [
          'Triangle A has a base of 30 centimeters.',
          'Triangle A has a perpendicular height of 15 centimeters.',
          'The height meets the base at a right angle, so it is the correct measurement for the formula.'
        ],
        directions: [
          'Label the base and the height in the diagram.',
          'Underline the measurement that tells you how tall the triangle is.',
          'Use words and numbers to explain why 15 cm is a height and not a side length along the edge.'
        ],
        diagramPrompt: 'How do you know which segment is the height of Triangle A?',
        diagram: 'triangleA'
      },
      {
        type: 'workedExample',
        title: 'Worked Example: Compose Triangle A',
        kicker: 'Parallelogram Connection',
        prompt: 'How can you use what you know about finding the area of a parallelogram to find the area of Triangle A?',
        givens: [
          'Use 2 copies of Triangle A to compose a parallelogram.',
          'The composed parallelogram keeps the same base of 30 cm and height of 15 cm.',
          'Triangle A is exactly half of that parallelogram.'
        ],
        steps: [
          'Find the area of the parallelogram: 30 × 15 = 450 square centimeters.',
          'Notice that the parallelogram is made of 2 congruent triangles.',
          'Divide by 2: 450 ÷ 2 = 225 square centimeters.'
        ],
        answer: 'Triangle A has an area of 225 square centimeters.',
        explainPrompt: 'What relationship do you notice between the area of the triangle and the area of the parallelogram?',
        starters: [
          'The parallelogram area is ...',
          'One triangle is half because ...',
          'So Triangle A has area ...'
        ],
        diagram: 'parallelogramBridge'
      },
      {
        type: 'formula',
        title: 'Build the Formula',
        kicker: 'From Pattern to Rule',
        intro: 'The lesson reveal explains that we can write a formula that reflects half the area of a parallelogram.',
        equationCards: [
          {
            label: 'Parallelogram',
            equation: 'A = bh',
            note: 'Base times height'
          },
          {
            label: 'Triangle',
            equation: 'A = 1/2 bh',
            note: 'Half of the related parallelogram'
          },
          {
            label: 'Precision',
            equation: 'units²',
            note: 'Area is written in square units'
          }
        ],
        prompt: 'Write the story of the formula in your own words. Explain how the parallelogram model helps the formula make sense.',
        checkpoints: [
          'Circle the base and height in a diagram before multiplying.',
          'Multiply base × height first, then take half.',
          'Write the answer with square units.'
        ],
        starters: [
          'First, I find ...',
          'Next, I divide by 2 because ...',
          'The formula works because ...'
        ]
      },
      {
        type: 'workedExample',
        title: 'Worked Example: Triangle B',
        kicker: 'Right Triangles',
        prompt: 'What type of triangle is Triangle B? What does this tell you about the height of the triangle?',
        givens: [
          'Triangle B is a right triangle.',
          'The base is 10 inches.',
          'The vertical side is the height, and it measures 14 inches.'
        ],
        steps: [
          'Substitute into the formula: A = 1/2 × 10 × 14.',
          'Multiply 10 × 14 to get 140.',
          'Take half of 140 to get 70.'
        ],
        answer: 'Triangle B has an area of 70 square inches.',
        explainPrompt: 'Why can the vertical side be used as the height in this right triangle?',
        starters: [
          'Because the triangle is right, ...',
          'The height is the side that ...',
          'The area is ... square inches.'
        ],
        diagram: 'triangleB'
      },
      {
        type: 'practice',
        title: 'Precision Practice',
        kicker: 'Using Symbols Appropriately',
        prompt: 'Use the source examples to respond with precise mathematical language.',
        problems: [
          {
            label: 'Triangle A',
            prompt: 'State the area of Triangle A and name the correct unit.'
          },
          {
            label: 'Triangle B',
            prompt: 'State the area of Triangle B and name the correct unit.'
          },
          {
            label: 'Precision',
            prompt: 'Explain why Triangle A uses square centimeters while Triangle B uses square inches.'
          }
        ],
        support: [
          'Use the words square centimeters or square inches.',
          'Match the unit to the original side lengths.',
          'Explain what the squared unit means.'
        ],
        workspacePrompt: 'Write a precise answer and explain why the unit must be squared.'
      },
      {
        type: 'practice',
        title: 'Let’s Explore More',
        kicker: 'Lucy’s Strategy',
        prompt: 'The source task asks whether Lucy’s method is acceptable. Use the given measurements to justify your reasoning.',
        problems: [
          {
            label: 'Reasoning',
            prompt: 'Lucy took half of the height and then multiplied by the base. Is that an acceptable way to find the area of a triangle? Explain.'
          },
          {
            label: 'Evidence',
            prompt: 'Use the source measurements b = 58.3 cm and h = 20 cm to support your explanation.'
          }
        ],
        support: [
          'The order of multiplication does not change the product.',
          'Half of 20 is 10, so Lucy is still finding 1/2 × b × h.',
          'Use numbers and words in your explanation.'
        ],
        workspacePrompt: 'Write a complete explanation that proves whether Lucy’s strategy is equivalent to the triangle area formula.'
      },
      {
        type: 'reflection',
        title: 'Session 1 Reflection',
        kicker: 'Make the Learning Stick',
        prompt: 'How can composing a triangle into a parallelogram help you remember the formula for the area of a triangle?',
        checklist: [
          'I can identify a base and a perpendicular height.',
          'I can explain why a triangle is half of a related parallelogram.',
          'I can calculate area and include correct square units.'
        ],
        starterPrompt: 'This shows that ...',
        nowPrompt: 'Write one idea that makes more sense to you now than it did at the start of the lesson.',
        nextPrompt: 'Name one question or skill you still want to practice tomorrow.'
      }
    ]
  };
}

function getSession2NotebookSpec_() {
  return {
    presentationTitle: 'Determine the Area of Triangles - Session 2 Student Notebook',
    sessionLabel: 'Session 2',
    sessionFocus: 'Review the formula, solve for a missing base, and apply triangle area in context.',
    theme: createTheme_(2),
    slides: [
      {
        type: 'cover',
        kicker: 'Session 2 Student Notebook',
        title: 'Determine the Area of Triangles',
        subtitle: 'Use the area formula with precision to solve for a missing dimension, explain what the answer means, and apply triangle area in a real-world context.',
        focusBullets: [
          'Revisit the formula and the meaning of base, height, and square units.',
          'Use the Bermuda Triangle problem to solve for a missing base when area and height are known.',
          'Apply your understanding to summary, challenge, and cabin-painting tasks from the source lesson.'
        ],
        designNotes: [
          'Session 2 stands alone but connects to Session 1',
          'Application and reflection built in',
          'Source examples preserved from slides 26–45'
        ]
      },
      {
        type: 'review',
        title: 'Quick Review: What Could the Question Be?',
        kicker: 'Be Curious',
        prompt: 'The source deck opens Session 2 by asking, “What could the question be?” Use what you remember from Session 1 to anticipate what a triangle-area problem might ask you to find.',
        openResponse: 'Write a possible question that could match a triangle diagram. What information would you need before you could answer it?',
        partnerPrompt: 'How are you establishing a positive relationship with a classmate in math class while you compare possible questions?',
        starters: [
          'A likely question is ...',
          'I would need to know ...',
          'My partner helped me see ...'
        ],
        cards: [
          {
            title: 'Triangle A',
            body: 'Base = 30 cm, height = 15 cm, area = 225 square centimeters.'
          },
          {
            title: 'Triangle B',
            body: 'Right triangle with base = 10 in, height = 14 in, area = 70 square inches.'
          },
          {
            title: 'Formula',
            body: 'Triangle area uses A = 1/2 bh.'
          },
          {
            title: 'Precision',
            body: 'A context may ask for an estimate or for an exact value. Read carefully.'
          }
        ]
      },
      {
        type: 'targets',
        title: 'Session 2 Learning Targets',
        kicker: 'Focus for Today',
        bullets: [
          'I can attend to precision to find a missing dimension of a triangle by using the area formula.',
          'I can explain what a missing measure represents in the context of a problem.',
          'I can decide whether a situation calls for an estimate or an exact solution.'
        ],
        bridgeText: 'Session 2 keeps the same formula but uses inverse thinking. Instead of finding area from base and height, we may know area and one dimension and need to solve for the missing one.',
        formulaFocus: 'If A and h are known, then 2A = bh and b = 2A ÷ h.',
        equationLines: ['A = 1/2 bh', '2A = bh', 'b = 2A ÷ h'],
        successPrompt: 'How do you know whether the number you solve for represents a base, a height, or an area?',
        sourceNote: 'The source lesson names precision as a focus and uses the Bermuda Triangle context to solve for a missing base.'
      },
      {
        type: 'notes',
        title: 'Guided Notes: Solving for a Missing Base',
        kicker: 'Use the Formula in Reverse',
        prompt: 'What changes when the problem gives you the area and the height instead of the area and the base?',
        factCards: [
          'Start with A = 1/2 bh.',
          'Multiply the area by 2 to undo the one-half.',
          'Divide by the known dimension to isolate the missing base.'
        ],
        directions: [
          'Box the formula you start with.',
          'Underline the value that represents the area.',
          'Circle the dimension you still need to solve for.'
        ],
        diagramPrompt: 'Use the diagram to connect each number to the part of the triangle it describes.',
        diagram: 'bermuda'
      },
      {
        type: 'application',
        title: 'Application Setup: The Bermuda Triangle',
        kicker: 'Represent the Situation',
        contextTitle: 'Source Context',
        context: 'The Bermuda Triangle is an area in the Atlantic Ocean between Miami, Florida, Bermuda, and San Juan, Puerto Rico. The triangle covers about 437,409 square miles. The source question asks: What is the approximate distance from Bermuda to San Juan?',
        knowns: [
          'Area A = 437,409 square miles',
          'Height h = 917 miles',
          'Unknown base b = distance from Bermuda to San Juan'
        ],
        precisionPrompt: 'Does the problem call for an estimate or an exact solution? Explain how the word about affects your thinking.',
        solvePrompt: 'What tool can you use to represent the problem, and what does each label mean in the diagram?',
        diagram: 'bermuda'
      },
      {
        type: 'workedExample',
        title: 'Worked Example: Solve for the Missing Base',
        kicker: 'The Bermuda Triangle',
        prompt: 'Let’s substitute the given values for the area and the height into the formula, then solve for b.',
        givens: [
          'A = 437,409 square miles',
          'h = 917 miles',
          'b is the distance from Bermuda to San Juan'
        ],
        steps: [
          'Start with A = 1/2 bh.',
          'Substitute and solve: 437,409 = 1/2 × b × 917.',
          'Multiply the area by 2, then divide by 917: b = 2(437,409) ÷ 917 = 954.'
        ],
        answer: 'The distance from Bermuda to San Juan is about 954 miles.',
        explainPrompt: 'What does the 954-mile answer represent in the context of the triangle?',
        starters: [
          'The missing measure is ...',
          'This value represents ...',
          'The estimate makes sense because ...'
        ],
        diagram: 'bermuda'
      },
      {
        type: 'practice',
        title: 'Collaborate and Connect',
        kicker: 'Explain the Missing Measure',
        prompt: 'Use the worked example to write and talk about what the answer means.',
        problems: [
          {
            label: 'Interpret',
            prompt: 'What is the missing measure in the Bermuda Triangle problem?'
          },
          {
            label: 'Meaning',
            prompt: 'What does 954 miles represent in the context of the map?'
          },
          {
            label: 'Precision',
            prompt: 'Why is it appropriate to say about 954 miles instead of simply 954 miles?'
          }
        ],
        support: [
          'Name the side of the triangle the number belongs to.',
          'Connect the number to the cities or locations in the context.',
          'Use the word about because the area in the source problem is approximate.'
        ],
        workspacePrompt: 'Write a response that names the missing measure, explains the meaning, and uses precise language.'
      },
      {
        type: 'practice',
        title: 'Challenge / Extension',
        kicker: 'Let’s Explore More',
        prompt: 'The source extension asks: What other information would you need to calculate the distance from San Juan to Miami using the area formula?',
        problems: [
          {
            label: 'Needed Fact',
            prompt: 'Describe the extra measurement you would need before you could find the side from San Juan to Miami.'
          },
          {
            label: 'Why',
            prompt: 'Explain how that additional measurement would connect to A = 1/2 bh.'
          }
        ],
        support: [
          'A different side needs a different perpendicular height.',
          'The area can stay the same while the chosen base changes.',
          'Name which point would create the perpendicular height.'
        ],
        workspacePrompt: 'Explain why a new base requires a matching perpendicular height if you want to keep using the area formula.'
      },
      {
        type: 'summary',
        title: 'Summarize: Determine the Area of Triangles',
        kicker: 'Pull the Ideas Together',
        cards: [
          {
            title: 'Structure',
            body: 'A triangle can be understood as half of a related parallelogram.',
            fill: '#FFFFFF'
          },
          {
            title: 'Formula',
            body: 'Triangle area uses A = 1/2 bh.',
            fill: '#FFFFFF'
          },
          {
            title: 'Precision',
            body: 'Area answers use square units and match the context.',
            fill: '#FFFFFF'
          },
          {
            title: 'Missing Measure',
            body: 'When a dimension is unknown, use inverse operations to solve for it.',
            fill: '#FFFFFF'
          }
        ],
        bigIdeaPrompt: 'Explain the biggest idea you want to remember from this lesson about the area of triangles.',
        targets: [
          'Use composing or the formula to find area.',
          'Attend to precision when writing units.',
          'Solve for a missing dimension when the context gives you enough information.'
        ],
        starters: [
          'The pattern suggests that ...',
          'I know this because ...',
          'The relationship between the variables is ...'
        ]
      },
      {
        type: 'application',
        title: 'Apply: Painting a Cabin',
        kicker: 'Real-World Triangle Area',
        contextTitle: 'Source Context',
        context: 'The source lesson ends with a cabin image whose triangular front has a base of 34 feet and a height of 23 3/4 feet. Use the triangle area formula to decide how much triangular wall space could be painted.',
        knowns: [
          'Base b = 34 feet',
          'Height h = 23 3/4 feet',
          'Unknown area A = ? square feet'
        ],
        precisionPrompt: 'Decide whether you should report an exact area or a rounded estimate, and explain why.',
        solvePrompt: 'Set up the expression you would use, show your work, and explain what the resulting area means in this context.',
        diagram: 'cabin'
      },
      {
        type: 'exitTicket',
        title: 'Exit Ticket',
        kicker: 'Show What You Know',
        prompt: 'Use the lesson examples from both sessions to answer each prompt clearly and precisely.',
        questions: [
          'Write the area formula for a triangle and describe what each variable represents.',
          'Why is 14 inches the correct height for Triangle B?',
          'What does the 954-mile answer represent in the Bermuda Triangle problem?'
        ],
        precisionReflection: 'How did you attend to precision today when you solved for an area or a missing dimension?',
        starters: [
          'I used precision when ...',
          'My answer represented ...',
          'I checked my units by ...'
        ],
        checks: [
          'Did I use a perpendicular height?',
          'Did I include square units for area?',
          'Did I explain what the number means in context?'
        ]
      }
    ]
  };
}

function createTheme_(sessionNumber) {
  if (sessionNumber === 1) {
    return {
      primary: '#163A5F',
      accent: '#0F8B8D',
      background: '#F5F8FC',
      coverBg: '#163A5F',
      coverCard: '#FFFFFF',
      coverBorder: '#C9D8EA',
      surface: '#EAF5F5',
      softAccent: '#DDF2F0',
      border: '#C8D5E5',
      line: '#C9D5E2',
      muted: '#5F7187',
      ink: '#1F2937',
      accentSoft: '#5CC7C8',
      diagramBg: '#F8FBFE'
    };
  }
  return {
    primary: '#16324F',
    accent: '#2A7BCB',
    background: '#F6F9FD',
    coverBg: '#16324F',
    coverCard: '#FFFFFF',
    coverBorder: '#C7D7EB',
    surface: '#EBF4FF',
    softAccent: '#DDECFD',
    border: '#C5D4E4',
    line: '#C9D5E2',
    muted: '#5C7188',
    ink: '#1F2937',
    accentSoft: '#8AB7E8',
    diagramBg: '#F7FAFE'
  };
}
