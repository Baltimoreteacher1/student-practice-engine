/******************************************************************
 * Notebookgrids.gs  —  EduWonderLab Premium Notebook Grids v2.1
 *
 * Added: triangle, rightTriangle, base, parallelogram, compose,
 *        decompose, formula visuals for triangle lessons
 * Fixed: drawWritingLines_ spacing
 * Kept:  all legacy trapezoid visuals for backward compat
 * Fixed: formula text uses insertTextBox() for stable text rendering
 ******************************************************************/

/* ================================================================
   WRITING LINES
   ================================================================ */
function drawWritingLines_(slide, x, y, w, h, count, colorHex) {
  var lines = Math.max(1, count || 4);
  var usable = Math.max(lines * 10, h - 4);
  var gap = Math.floor(usable / lines);
  var color = colorHex || '#CBD5E1';
  for (var i = 0; i < lines; i++) {
    var ly = y + i * gap + gap - 3;
    var line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x, ly, x + w, ly);
    line.getLineFill().setSolidFill(color);
    line.setWeight(0.75);
  }
}

/* ================================================================
   VOCAB VISUAL ROUTER
   ================================================================ */
function drawVocabVisual_(slide, key, x, y, w, h, pal) {
  var k = safeStr_(key).toLowerCase().replace(/[^a-z0-9]/g, '');
  switch (k) {
    case 'triangle':
      drawTriangleOutline_(slide, x, y, w, h, pal, true); break;
    case 'righttriangle':
    case 'height':
      drawRightTriangleVisual_(slide, x, y, w, h, pal); break;
    case 'base':
    case 'bases':
      drawTriangleBasesVisual_(slide, x, y, w, h, pal); break;
    case 'parallelogram':
    case 'compose':
    case 'trianglecompose':
      drawTriangleComposeVisual_(slide, x, y, w, h, pal); break;
    case 'decompose':
      drawTriangleDecomposeVisual_(slide, x, y, w, h, pal); break;
    case 'formula':
    case 'areaformula':
    case 'approximatearea':
    case 'precision':
    case 'missingdimension':
    case 'approximate':
      drawTriangleFormulaVisual_(slide, x, y, w, h, pal); break;
    case 'trapezoid':
    case 'isoscelestrapezoid':
      drawTrapezoidOutline_(slide, x, y, w, h, pal, true); break;
    case 'righttrapezoid':
      drawRightTrapezoid_(slide, x, y, w, h, pal); break;
    default:
      drawTriangleFormulaVisual_(slide, x, y, w, h, pal); break;
  }
}

function drawWorkspaceVisual_(slide, key, x, y, w, h, pal) {
  drawVocabVisual_(slide, key, x, y, w, h, pal);
}

/* ================================================================
   TRIANGLE VISUALS
   ================================================================ */

function drawTriangleOutline_(slide, x, y, w, h, pal, labeled) {
  var pad = 10;
  var bx1 = x + pad, by = y + h - pad;
  var bx2 = x + w - pad;
  var tx = x + Math.floor(w / 2), ty = y + pad;

  drawLine_(slide, bx1, by, bx2, by, pal.primary, 2.5);
  drawLine_(slide, bx1, by, tx, ty, pal.primary, 2.5);
  drawLine_(slide, bx2, by, tx, ty, pal.primary, 2.5);

  if (labeled) {
    drawMiniLabel_(slide, 'b', x + pad + 2, by - 14, 18, 12, pal);
    drawHighlightLine_(slide, tx, ty + 2, tx, by - 2, pal.accent, 1.5, true);
    drawMiniLabel_(slide, 'h', tx + 4, y + Math.floor(h / 2) - 6, 16, 12, pal);
  }
}

function drawRightTriangleVisual_(slide, x, y, w, h, pal) {
  var pad = 10;
  var bx1 = x + pad, by = y + h - pad;
  var bx2 = x + w - pad;
  var tx = x + pad, ty = y + pad;

  drawLine_(slide, bx1, by, bx2, by, pal.primary, 2.5);
  drawLine_(slide, bx1, by, tx, ty, pal.primary, 2.5);
  drawLine_(slide, bx2, by, tx, ty, pal.primary, 2.5);
  drawRightAngleMarker_(slide, bx1, by, 9, pal.accent);
  drawMiniLabel_(slide, 'h', x + pad + 4, y + Math.floor(h / 2) - 6, 16, 12, pal);
  drawMiniLabel_(slide, 'b', x + Math.floor(w / 2) - 4, by - 14, 16, 12, pal);
}

function drawTriangleBasesVisual_(slide, x, y, w, h, pal) {
  drawTriangleOutline_(slide, x, y, w, h, pal, false);
  var pad = 10, by = y + h - pad;
  drawHighlightLine_(slide, x + pad + 2, by, x + w - pad - 2, by, pal.accent, 3.5);
  drawMiniLabel_(slide, 'base', x + pad, by - 14, 28, 12, pal);
}

function drawTriangleComposeVisual_(slide, x, y, w, h, pal) {
  var safeW = Math.max(60, w), safeH = Math.max(44, h);
  var halfW = Math.floor(safeW * 0.42);

  var b1x1 = x + 4, b1y = y + safeH - 12;
  var b1x2 = x + halfW, t1x = x + Math.floor(halfW / 2), t1y = y + 10;
  drawLine_(slide, b1x1, b1y, b1x2, b1y, pal.primary, 2);
  drawLine_(slide, b1x1, b1y, t1x, t1y, pal.primary, 2);
  drawLine_(slide, b1x2, b1y, t1x, t1y, pal.primary, 2);

  var arrowX = x + halfW + 6, arrowY = y + Math.floor(safeH / 2);
  drawLine_(slide, arrowX, arrowY, arrowX + 12, arrowY, pal.accent, 2);
  drawLine_(slide, arrowX + 8, arrowY - 4, arrowX + 12, arrowY, pal.accent, 1.5);
  drawLine_(slide, arrowX + 8, arrowY + 4, arrowX + 12, arrowY, pal.accent, 1.5);

  var px = arrowX + 16, py = y + 8;
  var prW = Math.max(22, Math.floor(safeW * 0.36)), prH = safeH - 16, slant = 7;
  drawLine_(slide, px + slant, py, px + prW + slant, py, pal.accent, 1.5);
  drawLine_(slide, px, py + prH, px + prW, py + prH, pal.accent, 1.5);
  drawLine_(slide, px + slant, py, px, py + prH, pal.accent, 1.5);
  drawLine_(slide, px + prW + slant, py, px + prW, py + prH, pal.accent, 1.5);
}

function drawTriangleDecomposeVisual_(slide, x, y, w, h, pal) {
  drawTriangleOutline_(slide, x, y, w, h, pal, false);
  var tx = x + Math.floor(w / 2), ty = y + 10, by = y + h - 10;
  drawHighlightLine_(slide, tx, ty, tx, by, pal.accent, 1.5, true);
  drawMiniLabel_(slide, 'h', tx + 4, y + Math.floor(h / 2) - 6, 16, 12, pal);
}

function drawTriangleFormulaVisual_(slide, x, y, w, h, pal) {
  var card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x + 6, y + 10, w - 12, h - 20);
  card.getFill().setSolidFill(pal.light);
  card.getBorder().getLineFill().setSolidFill(pal.accent);
  card.getBorder().setWeight(1);

  var line1 = slide.insertTextBox('A = \u00bd \u00d7 b \u00d7 h', x + 10, y + 16, w - 20, 20);
  line1.getFill().setTransparent();
  line1.getBorder().setTransparent();
  var t1 = line1.getText();
  t1.getTextStyle().setFontFamily('Georgia').setFontSize(14).setBold(true).setForegroundColor(pal.primary);
  t1.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  var line2 = slide.insertTextBox('b = base,  h = height', x + 10, y + 38, w - 20, 13);
  line2.getFill().setTransparent();
  line2.getBorder().setTransparent();
  var t2 = line2.getText();
  t2.getTextStyle().setFontFamily('Calibri').setFontSize(9).setBold(false).setForegroundColor(pal.primary);
  t2.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

/* ================================================================
   LEGACY TRAPEZOID VISUALS
   ================================================================ */

function drawTrapezoidOutline_(slide, x, y, w, h, pal, isIso) {
  var pad = 10;
  var lbx = x + pad, lby = y + h - pad;
  var rbx = x + w - pad, rby = y + h - pad;
  var ltx = x + Math.floor(w * 0.32), lty = y + pad;
  var rtx = x + w - Math.floor(w * 0.32), rty = y + pad;

  drawLine_(slide, lbx, lby, rbx, rby, pal.primary, 2.5);
  drawLine_(slide, ltx, lty, rtx, rty, pal.primary, 2.5);
  drawLine_(slide, lbx, lby, ltx, lty, pal.primary, 2.5);
  drawLine_(slide, rbx, rby, rtx, rty, pal.primary, 2.5);
  drawMiniLabel_(slide, 'b\u2081', x + pad, lby - 14, 20, 12, pal);
  drawMiniLabel_(slide, 'b\u2082', x + Math.floor(w * 0.37), lty - 14, 20, 12, pal);

  if (isIso) {
    var midL = { x: Math.floor((lbx + ltx) / 2), y: Math.floor((lby + lty) / 2) };
    var midR = { x: Math.floor((rbx + rtx) / 2), y: Math.floor((rby + rty) / 2) };
    drawTickMark_(slide, midL.x, midL.y, pal);
    drawTickMark_(slide, midR.x, midR.y, pal);
  }
}

function drawRightTrapezoid_(slide, x, y, w, h, pal) {
  var pad = 10;
  var lbx = x + pad, lby = y + h - pad;
  var rbx = x + w - pad, rby = y + h - pad;
  var ltx = x + pad, lty = y + pad;
  var rtx = x + w - Math.floor(w * 0.38), rty = y + pad;

  drawLine_(slide, lbx, lby, rbx, rby, pal.primary, 2.5);
  drawLine_(slide, ltx, lty, rtx, rty, pal.primary, 2.5);
  drawLine_(slide, lbx, lby, ltx, lty, pal.primary, 2.5);
  drawLine_(slide, rbx, rby, rtx, rty, pal.primary, 2.5);
  drawRightAngleMarker_(slide, lbx, lby, 10, pal.accent);
  drawRightAngleMarker_(slide, ltx, lty, 10, pal.accent);
}

/* ================================================================
   LOW-LEVEL HELPERS
   ================================================================ */

function drawLine_(slide, x1, y1, x2, y2, hex, weight) {
  var line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, x1, y1, x2, y2);
  line.getLineFill().setSolidFill(hex || '#0F172A');
  line.setWeight(weight || 1.5);
  return line;
}

function drawHighlightLine_(slide, x1, y1, x2, y2, hex, weight, dashed) {
  var line = drawLine_(slide, x1, y1, x2, y2, hex, weight);
  if (dashed) line.setDashStyle(SlidesApp.DashStyle.DASH);
  return line;
}

function drawMiniLabel_(slide, text, x, y, w, h, pal) {
  var s = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, y, w, h);
  s.getFill().setSolidFill(pal.primary);
  s.getBorder().setTransparent();
  var tr = s.getText();
  tr.setText(text);
  tr.getTextStyle().setFontFamily('Calibri').setFontSize(8).setBold(true).setForegroundColor('#FFFFFF');
  tr.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

function drawTickMark_(slide, x, y, pal) {
  drawLine_(slide, x - 5, y - 5, x + 5, y + 5, pal.accent, 1.5);
}

function drawRightAngleMarker_(slide, x, y, size, color) {
  drawLine_(slide, x, y - size, x, y, color, 1.2);
  drawLine_(slide, x, y, x + size, y, color, 1.2);
}

function safeStr_(v) {
  return (v === null || v === undefined) ? '' : String(v);
}
