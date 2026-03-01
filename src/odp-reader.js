/**
 * odp-reader.js — ODP (OpenDocument Presentation) file loader.
 *
 * Converts ODP files into the same internal structure that PptxRenderer
 * uses for PPTX files, so the existing rendering pipeline works unchanged.
 *
 * Strategy: parse ODP content.xml/styles.xml and generate synthetic OOXML
 * slide documents + relationships, injecting them into the _files map.
 */

import { parseXml } from './utils.js';

const enc = new TextEncoder();
const EMU_PER_CM = 360000;  // 914400 EMU/inch ÷ 2.54 cm/inch
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// ── ODP XML helpers ──────────────────────────────────────────────────────────

function _g1(node, localName) {
  if (!node) return null;
  const all = node.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) {
    if (all[i].localName === localName) return all[i];
  }
  return null;
}

function _gtn(node, localName) {
  if (!node) return [];
  const r = [];
  const all = node.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) {
    if (all[i].localName === localName) r.push(all[i]);
  }
  return r;
}

/** Direct children with a given local name. */
function _children(node, localName) {
  if (!node) return [];
  const r = [];
  for (const c of node.children) {
    if (c.localName === localName) r.push(c);
  }
  return r;
}

/** Parse an ODP length value (e.g. "12.345cm", "10mm", "5in", "36pt") to EMU. */
function parseLength(val) {
  if (!val) return 0;
  val = val.trim();
  if (val.endsWith('cm'))  return Math.round(parseFloat(val) * EMU_PER_CM);
  if (val.endsWith('mm'))  return Math.round(parseFloat(val) * EMU_PER_CM / 10);
  if (val.endsWith('in'))  return Math.round(parseFloat(val) * 914400);
  if (val.endsWith('pt'))  return Math.round(parseFloat(val) * 12700);
  if (val.endsWith('px'))  return Math.round(parseFloat(val) * 9525);
  return Math.round(parseFloat(val) * EMU_PER_CM); // default to cm
}

/** Parse an ODP font size (e.g. "18pt") to OOXML hundredths of a point. */
function parseFontSize(val) {
  if (!val) return 1800;
  val = val.trim();
  if (val.endsWith('pt')) return Math.round(parseFloat(val) * 100);
  if (val.endsWith('cm')) return Math.round(parseFloat(val) / 2.54 * 7200);
  return Math.round(parseFloat(val) * 100);
}

/** Parse an ODP color (#RRGGBB) to 6-digit hex. */
function parseColor(val) {
  if (!val) return null;
  return val.replace('#', '').toUpperCase();
}

/** Parse angle from ODP rotate transform string, return degrees. */
function parseRotation(transform) {
  if (!transform) return 0;
  const m = transform.match(/rotate\(([^)]+)\)/);
  if (!m) return 0;
  // ODP stores rotation in radians (negative for clockwise)
  const rad = parseFloat(m[1]);
  // Convert to OOXML 60000ths of a degree (clockwise positive)
  return Math.round(-rad * 180 / Math.PI * 60000);
}

function escXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ── Style resolver ─────────────────────────────────────────────────────────

/**
 * Build a map of style:name → style properties by parsing
 * office:automatic-styles and office:styles from content.xml and styles.xml.
 */
function buildStyleMap(contentDoc, stylesDoc) {
  const map = {};

  for (const doc of [stylesDoc, contentDoc]) {
    if (!doc) continue;
    for (const container of [..._gtn(doc, 'automatic-styles'), ..._gtn(doc, 'styles')]) {
      for (const style of _children(container, 'style')) {
        const name = style.getAttribute('style:name') || style.getAttribute('name');
        if (!name) continue;
        map[name] = style;
      }
    }
  }
  return map;
}

/**
 * Resolve graphic properties for a shape by walking the style chain.
 */
function resolveGraphicProps(styleName, styleMap) {
  const props = { fill: null, fillColor: null, stroke: null, strokeColor: null, strokeWidth: null };
  let name = styleName;
  const visited = new Set();
  while (name && !visited.has(name)) {
    visited.add(name);
    const style = styleMap[name];
    if (!style) break;
    const gp = _g1(style, 'graphic-properties');
    if (gp) {
      if (!props.fill) props.fill = gp.getAttribute('draw:fill');
      if (!props.fillColor) props.fillColor = parseColor(gp.getAttribute('draw:fill-color'));
      if (!props.stroke) props.stroke = gp.getAttribute('draw:stroke');
      if (!props.strokeColor) props.strokeColor = parseColor(gp.getAttribute('svg:stroke-color'));
      if (!props.strokeWidth) props.strokeWidth = gp.getAttribute('svg:stroke-width');
    }
    name = style.getAttribute('style:parent-style-name');
  }
  return props;
}

/**
 * Resolve text properties for a text style.
 */
function resolveTextProps(styleName, styleMap) {
  const props = { fontSize: null, fontFamily: null, color: null, bold: false, italic: false, underline: false, strikethrough: false };
  let name = styleName;
  const visited = new Set();
  while (name && !visited.has(name)) {
    visited.add(name);
    const style = styleMap[name];
    if (!style) break;
    const tp = _g1(style, 'text-properties');
    if (tp) {
      if (!props.fontSize) props.fontSize = tp.getAttribute('fo:font-size');
      if (!props.fontFamily) props.fontFamily = tp.getAttribute('style:font-name') || tp.getAttribute('fo:font-family');
      if (!props.color) props.color = parseColor(tp.getAttribute('fo:color'));
      if (tp.getAttribute('fo:font-weight') === 'bold') props.bold = true;
      if (tp.getAttribute('fo:font-style') === 'italic') props.italic = true;
      if (tp.getAttribute('style:text-underline-style') === 'solid') props.underline = true;
      if (tp.getAttribute('style:text-line-through-style') === 'solid') props.strikethrough = true;
    }
    name = style.getAttribute('style:parent-style-name');
  }
  return props;
}

/**
 * Resolve paragraph alignment.
 */
function resolveParaProps(styleName, styleMap) {
  const props = { align: null, lineSpacing: null };
  let name = styleName;
  const visited = new Set();
  while (name && !visited.has(name)) {
    visited.add(name);
    const style = styleMap[name];
    if (!style) break;
    const pp = _g1(style, 'paragraph-properties');
    if (pp) {
      if (!props.align) {
        const a = pp.getAttribute('fo:text-align');
        if (a === 'center') props.align = 'ctr';
        else if (a === 'end') props.align = 'r';
        else if (a === 'justify') props.align = 'just';
        else if (a === 'start' || a === 'left') props.align = 'l';
      }
      if (!props.lineSpacing) props.lineSpacing = pp.getAttribute('fo:line-height');
    }
    name = style.getAttribute('style:parent-style-name');
  }
  return props;
}

/**
 * Resolve drawing-page properties (background).
 */
function resolveDrawingPageProps(styleName, styleMap) {
  const props = { fill: null, fillColor: null };
  let name = styleName;
  const visited = new Set();
  while (name && !visited.has(name)) {
    visited.add(name);
    const style = styleMap[name];
    if (!style) break;
    const dp = _g1(style, 'drawing-page-properties');
    if (dp) {
      if (!props.fill) props.fill = dp.getAttribute('draw:fill');
      if (!props.fillColor) props.fillColor = parseColor(dp.getAttribute('draw:fill-color'));
    }
    name = style.getAttribute('style:parent-style-name');
  }
  return props;
}

// ── ODP → synthetic OOXML conversion ───────────────────────────────────────

/**
 * Convert ODP draw:page content into a synthetic PPTX slide XML string.
 */
function convertPageToSlideXml(page, styleMap, pageIdx) {
  let shapeId = 1; // next shape ID

  // Background
  let bgXml = '';
  const pageStyleName = page.getAttribute('draw:style-name');
  if (pageStyleName) {
    const dpProps = resolveDrawingPageProps(pageStyleName, styleMap);
    if (dpProps.fillColor && dpProps.fill === 'solid') {
      bgXml = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${dpProps.fillColor}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
    }
  }

  let shapesXml = '';

  for (const child of page.children) {
    const ln = child.localName;
    if (ln === 'frame') {
      shapesXml += convertFrame(child, styleMap, ++shapeId);
    } else if (ln === 'rect') {
      shapesXml += convertBasicShape(child, 'rect', styleMap, ++shapeId);
    } else if (ln === 'ellipse') {
      shapesXml += convertBasicShape(child, 'ellipse', styleMap, ++shapeId);
    } else if (ln === 'custom-shape') {
      shapesXml += convertCustomShape(child, styleMap, ++shapeId);
    } else if (ln === 'line') {
      shapesXml += convertLine(child, styleMap, ++shapeId);
    }
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="${NS_P}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">
  <p:cSld>
    ${bgXml}
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      ${shapesXml}
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

/** Convert a draw:frame (text box or image) to p:sp or p:pic. */
function convertFrame(frame, styleMap, id) {
  const x = parseLength(frame.getAttribute('svg:x'));
  const y = parseLength(frame.getAttribute('svg:y'));
  const w = parseLength(frame.getAttribute('svg:width'));
  const h = parseLength(frame.getAttribute('svg:height'));
  const rot = parseRotation(frame.getAttribute('draw:transform'));
  const rotAttr = rot ? ` rot="${rot}"` : '';

  // Check for image
  const image = _g1(frame, 'image');
  if (image) {
    const href = image.getAttribute('xlink:href') || '';
    return `<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="${id}" name="Picture ${id}"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rOdpImg_${id}"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm${rotAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`;
  }

  // Text box
  const textBox = _g1(frame, 'text-box');
  if (!textBox) return '';

  const styleName = frame.getAttribute('draw:style-name');
  const gProps = resolveGraphicProps(styleName, styleMap);

  let fillXml = '<a:noFill/>';
  if (gProps.fill === 'solid' && gProps.fillColor) {
    fillXml = `<a:solidFill><a:srgbClr val="${gProps.fillColor}"/></a:solidFill>`;
  }
  let lnXml = '';
  if (gProps.stroke === 'solid' && gProps.strokeColor) {
    const sw = gProps.strokeWidth ? parseLength(gProps.strokeWidth) : 12700;
    lnXml = `<a:ln w="${sw}"><a:solidFill><a:srgbClr val="${gProps.strokeColor}"/></a:solidFill></a:ln>`;
  }

  const parasXml = convertTextContent(textBox, styleMap);

  return `<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="${id}" name="TextBox ${id}"/>
    <p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm${rotAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    ${fillXml}${lnXml}
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr>
    <a:lstStyle/>
    ${parasXml}
  </p:txBody>
</p:sp>`;
}

/** Convert a draw:rect or draw:ellipse to a p:sp. */
function convertBasicShape(el, shapeType, styleMap, id) {
  const x = parseLength(el.getAttribute('svg:x'));
  const y = parseLength(el.getAttribute('svg:y'));
  const w = parseLength(el.getAttribute('svg:width'));
  const h = parseLength(el.getAttribute('svg:height'));
  const rot = parseRotation(el.getAttribute('draw:transform'));
  const rotAttr = rot ? ` rot="${rot}"` : '';

  const styleName = el.getAttribute('draw:style-name');
  const gProps = resolveGraphicProps(styleName, styleMap);

  let fillXml = '<a:noFill/>';
  if (gProps.fill === 'solid' && gProps.fillColor) {
    fillXml = `<a:solidFill><a:srgbClr val="${gProps.fillColor}"/></a:solidFill>`;
  }
  let lnXml = '';
  if (gProps.stroke === 'solid' && gProps.strokeColor) {
    const sw = gProps.strokeWidth ? parseLength(gProps.strokeWidth) : 12700;
    lnXml = `<a:ln w="${sw}"><a:solidFill><a:srgbClr val="${gProps.strokeColor}"/></a:solidFill></a:ln>`;
  }

  const prst = shapeType === 'ellipse' ? 'ellipse' : 'rect';

  // Check for text inside the shape
  let txBodyXml = '';
  const textContent = convertTextContent(el, styleMap);
  if (textContent) {
    txBodyXml = `<p:txBody>
    <a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/>
    <a:lstStyle/>
    ${textContent}
  </p:txBody>`;
  }

  return `<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="${id}" name="${shapeType} ${id}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm${rotAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
    <a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>
    ${fillXml}${lnXml}
  </p:spPr>
  ${txBodyXml}
</p:sp>`;
}

/** Convert a draw:custom-shape to a p:sp. */
function convertCustomShape(el, styleMap, id) {
  const x = parseLength(el.getAttribute('svg:x'));
  const y = parseLength(el.getAttribute('svg:y'));
  const w = parseLength(el.getAttribute('svg:width'));
  const h = parseLength(el.getAttribute('svg:height'));
  const rot = parseRotation(el.getAttribute('draw:transform'));
  const rotAttr = rot ? ` rot="${rot}"` : '';

  const styleName = el.getAttribute('draw:style-name');
  const gProps = resolveGraphicProps(styleName, styleMap);

  let fillXml = '<a:noFill/>';
  if (gProps.fill === 'solid' && gProps.fillColor) {
    fillXml = `<a:solidFill><a:srgbClr val="${gProps.fillColor}"/></a:solidFill>`;
  }
  let lnXml = '';
  if (gProps.stroke === 'solid' && gProps.strokeColor) {
    const sw = gProps.strokeWidth ? parseLength(gProps.strokeWidth) : 12700;
    lnXml = `<a:ln w="${sw}"><a:solidFill><a:srgbClr val="${gProps.strokeColor}"/></a:solidFill></a:ln>`;
  }

  // Map ODP enhanced-geometry draw:type to OOXML prst
  const geo = _g1(el, 'enhanced-geometry');
  const odpType = geo ? (geo.getAttribute('draw:type') || 'rect') : 'rect';
  const typeMap = {
    'rectangle': 'rect', 'round-rectangle': 'roundRect',
    'isosceles-triangle': 'triangle', 'diamond': 'diamond',
    'pentagon': 'pentagon', 'hexagon': 'hexagon',
    'star5': 'star5', 'star6': 'star6', 'star4': 'star4',
    'heart': 'heart', 'cloud': 'cloud',
    'right-arrow': 'rightArrow', 'left-arrow': 'leftArrow',
    'up-arrow': 'upArrow', 'down-arrow': 'downArrow',
    'cross': 'plus', 'can': 'can', 'ring': 'donut',
    'ellipse': 'ellipse', 'circle': 'ellipse',
  };
  const prst = typeMap[odpType] || 'rect';

  let txBodyXml = '';
  const textContent = convertTextContent(el, styleMap);
  if (textContent) {
    txBodyXml = `<p:txBody>
    <a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/>
    <a:lstStyle/>
    ${textContent}
  </p:txBody>`;
  }

  return `<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="${id}" name="Shape ${id}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm${rotAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
    <a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>
    ${fillXml}${lnXml}
  </p:spPr>
  ${txBodyXml}
</p:sp>`;
}

/** Convert a draw:line to a connector shape. */
function convertLine(el, styleMap, id) {
  const x1 = parseLength(el.getAttribute('svg:x1'));
  const y1 = parseLength(el.getAttribute('svg:y1'));
  const x2 = parseLength(el.getAttribute('svg:x2'));
  const y2 = parseLength(el.getAttribute('svg:y2'));
  const x = Math.min(x1, x2);
  const y = Math.min(y1, y2);
  const w = Math.abs(x2 - x1) || 1;
  const h = Math.abs(y2 - y1) || 1;

  const styleName = el.getAttribute('draw:style-name');
  const gProps = resolveGraphicProps(styleName, styleMap);

  let lnXml = '<a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>';
  if (gProps.stroke === 'solid' && gProps.strokeColor) {
    const sw = gProps.strokeWidth ? parseLength(gProps.strokeWidth) : 12700;
    lnXml = `<a:ln w="${sw}"><a:solidFill><a:srgbClr val="${gProps.strokeColor}"/></a:solidFill></a:ln>`;
  }

  const flipH = x2 < x1 ? ' flipH="1"' : '';
  const flipV = y2 < y1 ? ' flipV="1"' : '';

  return `<p:cxnSp>
  <p:nvCxnSpPr>
    <p:cNvPr id="${id}" name="Line ${id}"/>
    <p:cNvCxnSpPr/>
    <p:nvPr/>
  </p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm${flipH}${flipV}><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
    ${lnXml}
  </p:spPr>
</p:cxnSp>`;
}

/** Convert ODP text:p/text:span elements to OOXML a:p/a:r XML. */
function convertTextContent(container, styleMap) {
  const paragraphs = _children(container, 'p');
  // Also check inside text:list > text:list-item > text:p
  const lists = _gtn(container, 'list');

  if (paragraphs.length === 0 && lists.length === 0) return '';

  let xml = '';

  // Regular paragraphs
  for (const p of paragraphs) {
    xml += convertParagraph(p, styleMap);
  }

  // Lists
  for (const list of lists) {
    const items = _gtn(list, 'list-item');
    for (const item of items) {
      const itemPs = _children(item, 'p');
      for (const p of itemPs) {
        xml += convertParagraph(p, styleMap, true);
      }
    }
  }

  return xml || '<a:p><a:r><a:t></a:t></a:r></a:p>';
}

/** Convert a single text:p to an a:p. */
function convertParagraph(p, styleMap, isBullet = false) {
  const paraStyleName = p.getAttribute('text:style-name');
  const pProps = resolveParaProps(paraStyleName, styleMap);
  const align = pProps.align || 'l';

  let bulletXml = '';
  if (isBullet) {
    bulletXml = '<a:buChar char="\u2022"/>';
  }

  const pPrXml = `<a:pPr algn="${align}"${isBullet ? ' marL="342900" indent="-342900"' : ''}>${bulletXml}</a:pPr>`;

  let runsXml = '';

  // Collect text spans and direct text
  for (const child of p.childNodes) {
    if (child.nodeType === 3) {
      // Direct text node
      const text = child.textContent;
      if (text.trim()) {
        const tProps = resolveTextProps(paraStyleName, styleMap);
        runsXml += buildRunXml(text, tProps);
      }
    } else if (child.nodeType === 1 && child.localName === 'span') {
      const spanStyleName = child.getAttribute('text:style-name');
      const tProps = resolveTextProps(spanStyleName, styleMap);
      // Also inherit from paragraph style
      if (!tProps.fontSize) {
        const pTextProps = resolveTextProps(paraStyleName, styleMap);
        if (pTextProps.fontSize) tProps.fontSize = pTextProps.fontSize;
      }
      if (!tProps.fontFamily) {
        const pTextProps = resolveTextProps(paraStyleName, styleMap);
        if (pTextProps.fontFamily) tProps.fontFamily = pTextProps.fontFamily;
      }
      const text = child.textContent || '';
      runsXml += buildRunXml(text, tProps);
    } else if (child.nodeType === 1 && child.localName === 'line-break') {
      // Line break within a paragraph — split into separate paragraph visually
      // For now, just add a newline run
      runsXml += '<a:br/>';
    }
  }

  if (!runsXml) {
    runsXml = '<a:r><a:rPr lang="en-US" sz="1800" dirty="0"/><a:t> </a:t></a:r>';
  }

  return `<a:p>${pPrXml}${runsXml}</a:p>`;
}

/** Build an a:r (run) XML from text and resolved text properties. */
function buildRunXml(text, tProps) {
  const sz = tProps.fontSize ? parseFontSize(tProps.fontSize) : 1800;
  const color = tProps.color || '000000';
  const fontFamily = tProps.fontFamily || 'Calibri';
  const b = tProps.bold ? ' b="1"' : '';
  const i = tProps.italic ? ' i="1"' : '';
  const u = tProps.underline ? ' u="sng"' : '';
  const strike = tProps.strikethrough ? ' strike="sngStrike"' : '';

  return `<a:r><a:rPr lang="en-US" sz="${sz}"${b}${i}${u}${strike} dirty="0">` +
    `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
    `<a:latin typeface="${escXml(fontFamily)}"/>` +
    `</a:rPr><a:t>${escXml(text)}</a:t></a:r>`;
}

// ── Main export ──────────────────────────────────────────────────────────────

/**
 * Parse an ODP file's ZIP contents and convert to PPTX-like internal files.
 *
 * Mutates the `files` map in-place, adding synthetic OOXML files
 * (presentation.xml, slides, theme, etc.) so PptxRenderer can render them.
 *
 * @param {Record<string, Uint8Array>} files  — ZIP contents from readZip()
 * @returns {{ slideSize: {cx: number, cy: number}, slidePaths: string[] }}
 */
export function convertOdpFiles(files) {
  const dec = new TextDecoder();

  // ── Parse content.xml and styles.xml ───────────────────────────────────
  const contentRaw = files['content.xml'];
  if (!contentRaw) throw new Error('Invalid ODP: missing content.xml');
  const contentDoc = parseXml(dec.decode(contentRaw));

  let stylesDoc = null;
  const stylesRaw = files['styles.xml'];
  if (stylesRaw) stylesDoc = parseXml(dec.decode(stylesRaw));

  const styleMap = buildStyleMap(contentDoc, stylesDoc);

  // ── Determine slide size from styles.xml page layout ──────────────────
  let slideW = 9144000, slideH = 5143500; // defaults (10in × 5.63in)

  if (stylesDoc) {
    const pageLayouts = _gtn(stylesDoc, 'page-layout');
    for (const pl of pageLayouts) {
      const plProps = _g1(pl, 'page-layout-properties');
      if (plProps) {
        const pw = plProps.getAttribute('fo:page-width');
        const ph = plProps.getAttribute('fo:page-height');
        if (pw) slideW = parseLength(pw);
        if (ph) slideH = parseLength(ph);
        break;
      }
    }
  }

  // ── Find draw:page elements ────────────────────────────────────────────
  const presentation = _g1(contentDoc, 'presentation');
  const pages = presentation ? _children(presentation, 'page') : [];

  if (pages.length === 0) {
    throw new Error('Invalid ODP: no slides found');
  }

  // ── Generate synthetic PPTX files ──────────────────────────────────────

  // [Content_Types].xml
  let overrides = '';
  overrides += `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>`;
  overrides += `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`;

  const sldIds = [];
  const slidePaths = [];
  let presRels = '';

  // Theme relationship
  presRels += `<Relationship Id="rId_theme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`;

  for (let i = 0; i < pages.length; i++) {
    const slideNum = i + 1;
    const slidePath = `ppt/slides/slide${slideNum}.xml`;
    const slideXml = convertPageToSlideXml(pages[i], styleMap, i);
    files[slidePath] = enc.encode(slideXml);

    // Build slide relationships (for images)
    const slideRelsXml = buildSlideRels(pages[i], files, i);
    files[`ppt/slides/_rels/slide${slideNum}.xml.rels`] = enc.encode(slideRelsXml);

    sldIds.push(`<p:sldId id="${256 + i}" r:id="rId${slideNum}"/>`);
    presRels += `<Relationship Id="rId${slideNum}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNum}.xml"/>`;
    overrides += `<Override PartName="/ppt/slides/slide${slideNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;
    slidePaths.push(slidePath);
  }

  files['[Content_Types].xml'] = enc.encode(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="gif" ContentType="image/gif"/>
  <Default Extension="svg" ContentType="image/svg+xml"/>
  ${overrides}
</Types>`);

  files['_rels/.rels'] = enc.encode(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);

  files['ppt/presentation.xml'] = enc.encode(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="${NS_P}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">
  <p:sldIdLst>${sldIds.join('')}</p:sldIdLst>
  <p:sldSz cx="${slideW}" cy="${slideH}"/>
  <p:notesSz cx="${slideH}" cy="${slideW}"/>
</p:presentation>`);

  files['ppt/_rels/presentation.xml.rels'] = enc.encode(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${presRels}
</Relationships>`);

  // Theme — extract font info from ODP styles if possible
  let majorFont = 'Calibri Light', minorFont = 'Calibri';
  // Try to detect fonts from ODP font-face-decls
  const fontDecls = _gtn(contentDoc, 'font-face');
  if (fontDecls.length > 0) {
    minorFont = fontDecls[0].getAttribute('style:name') || 'Calibri';
    if (fontDecls.length > 1) {
      majorFont = fontDecls[1].getAttribute('style:name') || minorFont;
    } else {
      majorFont = minorFont;
    }
  }

  files['ppt/theme/theme1.xml'] = enc.encode(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="${NS_A}" name="ODP Theme">
  <a:themeElements>
    <a:clrScheme name="ODP">
      <a:dk1><a:srgbClr val="000000"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="ODP">
      <a:majorFont><a:latin typeface="${escXml(majorFont)}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="${escXml(minorFont)}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="ODP">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`);

  return {
    slideSize: { cx: slideW, cy: slideH },
    slidePaths,
  };
}

/**
 * Build slide relationship XML for ODP image references.
 * Maps ODP image xlink:href paths to synthetic rId entries.
 */
function buildSlideRels(page, files, pageIdx) {
  let rels = '';

  // Find all images in this page
  let shapeId = 1;
  for (const child of page.children) {
    shapeId++;
    const ln = child.localName;
    if (ln === 'frame') {
      const image = _g1(child, 'image');
      if (image) {
        const href = image.getAttribute('xlink:href') || '';
        if (href && !href.startsWith('http')) {
          // ODP stores images as relative paths like "Pictures/image1.png"
          // Copy the image to ppt/media/ path for the PPTX rendering pipeline
          const imgData = files[href];
          if (imgData) {
            const ext = href.split('.').pop().toLowerCase();
            const pptMediaPath = `ppt/media/odp_${pageIdx}_${shapeId}.${ext}`;
            files[pptMediaPath] = imgData;
            rels += `<Relationship Id="rOdpImg_${shapeId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/odp_${pageIdx}_${shapeId}.${ext}"/>`;
          }
        }
      }
    }
  }

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${rels}
</Relationships>`;
}

/**
 * Detect whether a ZIP file is an ODP file.
 * @param {Record<string, Uint8Array>} files
 * @returns {boolean}
 */
export function isOdpFile(files) {
  // ODP files have a 'mimetype' entry with the ODP MIME type
  if (files['mimetype']) {
    const mime = new TextDecoder().decode(files['mimetype']);
    if (mime.includes('opendocument.presentation')) return true;
  }
  // Also check for content.xml (ODP structure)
  if (files['content.xml'] && !files['ppt/presentation.xml']) return true;
  return false;
}
