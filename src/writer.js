/**
 * writer.js — PPTX editor and template engine.
 *
 * Loads an existing PPTX (via a PptxRenderer instance or raw bytes),
 * provides a fluent API to mutate its content, and serializes back to
 * a valid PPTX file that PowerPoint / Keynote / LibreOffice can open.
 *
 * ── Quick start ──────────────────────────────────────────────────────────────
 *
 *   import { PptxWriter } from 'pptx-canvas-renderer';
 *
 *   // From a loaded renderer:
 *   const writer = PptxWriter.fromRenderer(renderer);
 *
 *   // Or from raw bytes:
 *   const writer = await PptxWriter.fromBytes(arrayBuffer);
 *
 *   // Template substitution  ({{tokens}} in shapes / speaker notes)
 *   writer.applyTemplate({ name: 'Acme Corp', year: '2025' });
 *
 *   // Replace text everywhere
 *   writer.replaceText('Old Text', 'New Text');
 *
 *   // Set the text of a specific shape
 *   writer.setShapeText(0, 'Title 1', 'My New Title');
 *
 *   // Swap an image on slide 2 (shape named "Picture 1")
 *   await writer.setShapeImage(1, 'Picture 1', jpegBytes, 'image/jpeg');
 *
 *   // Duplicate slide 0 as a new slide at the end
 *   writer.duplicateSlide(0);
 *
 *   // Remove slide 3
 *   writer.removeSlide(3);
 *
 *   // Reorder slides
 *   writer.reorderSlides([2, 0, 1]);
 *
 *   // Change a theme colour
 *   writer.setThemeColor('accent1', 'FF0000');
 *
 *   // Export
 *   const bytes = await writer.save();        // → Uint8Array
 *   writer.download('edited.pptx');           // trigger browser download
 *
 * ── API reference ─────────────────────────────────────────────────────────────
 *
 *   PptxWriter.create(opts)               — create from scratch (blank slide)
 *   PptxWriter.fromRenderer(renderer)     — clone from loaded PptxRenderer
 *   PptxWriter.fromBytes(buffer)          — parse PPTX bytes fresh
 *
 *   .applyTemplate(data, opts)            — {{token}} substitution
 *   .replaceText(find, replace, opts)     — global find-and-replace
 *   .setShapeText(slideIdx, name, text)   — set text of named shape
 *   .getShapeText(slideIdx, name)         — read text from named shape
 *   .addTextBox(slideIdx, text, style)    — add text box (italic/underline/fill/…)
 *   .addRichText(slideIdx, paragraphs, style)  — mixed-format text runs
 *   .addShape(slideIdx, type, style)      — preset shape with optional text
 *   .addList(slideIdx, items, style)      — bullet/numbered list
 *   .setShapeImage(slideIdx, name, bytes, mime)  — swap shape image
 *   .addImage(slideIdx, bytes, mime, rect)       — add new image shape
 *   .setSlideBackground(slideIdx, color)         — solid background color
 *   .setThemeColor(key, hexRgb)           — change theme colour (no #)
 *   .addSlide(atIdx?)                     — add a blank slide
 *   .duplicateSlide(fromIdx, toIdx?)      — copy slide
 *   .removeSlide(slideIdx)                — delete slide
 *   .reorderSlides(newOrder)              — reorder by index array
 *   .setSlideNotes(slideIdx, text)        — set speaker notes
 *   .getSlidePaths()                      — current slide file paths
 *   .getSlideCount()                      — current slide count
 *   .save()                               → Promise<Uint8Array>  PPTX bytes
 *   .download(filename)                   — save file in browser
 */

import { readZip } from './zip.js';
import { writeZip } from './zip-writer.js';

const dec = new TextDecoder();
const enc = new TextEncoder();
const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
};

// ── XML helpers ───────────────────────────────────────────────────────────────

function parseXml(str) {
  return new DOMParser().parseFromString(str, 'application/xml');
}
function serializeXml(doc) {
  const s = new XMLSerializer().serializeToString(doc);
  // Ensure declaration
  if (s.startsWith('<?xml')) return s;
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s;
}
function xmlBytes(doc) { return enc.encode(serializeXml(doc)); }
function readXml(files, path) {
  const raw = files[path];
  if (!raw) return null;
  return parseXml(dec.decode(raw));
}
function g1(node, name) {
  if (!node) return null;
  const all = node.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) if (all[i].localName === name) return all[i];
  return null;
}
function gtn(node, name) {
  if (!node) return [];
  const r = [];
  const all = node.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) if (all[i].localName === name) r.push(all[i]);
  return r;
}
function attr(el, name, def = null) {
  if (!el) return def;
  const v = el.getAttribute(name);
  return v !== null ? v : def;
}

// ── Relationship helpers ──────────────────────────────────────────────────────

function relsPath(filePath) {
  const parts = filePath.split('/');
  const name  = parts.pop();
  return [...parts, '_rels', name + '.rels'].join('/');
}

function parseRels(files, filePath) {
  const doc = readXml(files, relsPath(filePath));
  if (!doc) return {};
  const map = {};
  for (const rel of Array.from(doc.getElementsByTagName('Relationship'))) {
    const id     = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    const type   = rel.getAttribute('Type') || '';
    let fullPath = target;
    if (!target.startsWith('/') && !target.startsWith('http')) {
      const dir = filePath.split('/').slice(0, -1).join('/');
      fullPath  = dir ? dir + '/' + target.replace(/^\.\.\//, '') : target;
      // Handle ../ traversal
      const parts = fullPath.split('/');
      const resolved = [];
      for (const p of parts) {
        if (p === '..') resolved.pop();
        else resolved.push(p);
      }
      fullPath = resolved.join('/');
    }
    map[id] = { id, target, type, fullPath };
  }
  return map;
}

function buildRelsDoc(rels) {
  const doc = parseXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
  const root = doc.documentElement;
  for (const rel of Object.values(rels)) {
    const el = doc.createElementNS(NS.rel, 'Relationship');
    el.setAttribute('Id', rel.id);
    el.setAttribute('Type', rel.type);
    el.setAttribute('Target', rel.target);
    if (rel.targetMode) el.setAttribute('TargetMode', rel.targetMode);
    root.appendChild(el);
  }
  return doc;
}

function nextRId(rels) {
  const nums = Object.keys(rels)
    .map(id => parseInt(id.replace('rId', ''), 10))
    .filter(n => !isNaN(n));
  return 'rId' + ((nums.length ? Math.max(...nums) : 0) + 1);
}

// ── Shape lookup helpers ──────────────────────────────────────────────────────

function findShapeByName(spTree, name) {
  for (const child of spTree.children) {
    const ln = child.localName;
    if (ln === 'sp' || ln === 'pic' || ln === 'cxnSp') {
      const nvEl  = g1(child, 'nvSpPr') || g1(child, 'nvPicPr') || g1(child, 'nvCxnSpPr');
      const cNvPr = nvEl ? g1(nvEl, 'cNvPr') : null;
      if (cNvPr) {
        const shapeName = cNvPr.getAttribute('name') || '';
        if (shapeName === name) return child;
      }
    } else if (ln === 'grpSp') {
      const found = findShapeByName(child, name);
      if (found) return found;
    }
  }
  return null;
}

function findShapeById(spTree, id) {
  const idStr = String(id);
  for (const child of spTree.children) {
    const ln = child.localName;
    if (ln === 'sp' || ln === 'pic' || ln === 'cxnSp') {
      const nvEl  = g1(child, 'nvSpPr') || g1(child, 'nvPicPr');
      const cNvPr = nvEl ? g1(nvEl, 'cNvPr') : null;
      if (cNvPr && (cNvPr.getAttribute('id') || '') === idStr) return child;
    }
  }
  return null;
}

function getSpTree(slideDoc) {
  const cSld = g1(slideDoc, 'cSld');
  return cSld ? g1(cSld, 'spTree') : null;
}

// ── Text replacement helpers ──────────────────────────────────────────────────

function getAllTextNodes(node) {
  const result = [];
  const walker = node.ownerDocument
    ? node.ownerDocument.createTreeWalker(node, 0x04 /* NodeFilter.SHOW_TEXT */)
    : null;
  if (!walker) return result;
  let n;
  while ((n = walker.nextNode())) result.push(n);
  return result;
}

/** Replace text in a run without disturbing formatting. */
function replaceInDoc(doc, find, replace, caseSensitive = true) {
  // Collect all <a:t> elements and replace within their text content
  for (const t of gtn(doc, 't')) {
    const orig = t.textContent;
    if (!orig) continue;
    const newText = caseSensitive
      ? orig.split(find).join(replace)
      : orig.replace(new RegExp(escapeRegex(find), 'gi'), replace);
    if (newText !== orig) t.textContent = newText;
  }
}

function escapeRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

/** Escape text for use in XML element content / attribute values. */
function escXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ── Content type helpers ──────────────────────────────────────────────────────

const MIME_EXT = {
  'image/jpeg': 'jpeg', 'image/jpg': 'jpeg',
  'image/png': 'png', 'image/gif': 'gif',
  'image/webp': 'webp', 'image/svg+xml': 'svg',
};
const CT_MAP = {
  jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif',
  webp: 'image/webp', svg: 'image/svg+xml',
};

function addContentType(files, ext, partName) {
  const ctPath = '[Content_Types].xml';
  const doc    = readXml(files, ctPath);
  if (!doc) return;
  const root = doc.documentElement;
  // Check if Override already exists
  for (const ov of gtn(doc, 'Override')) {
    if (ov.getAttribute('PartName') === '/' + partName) return;
  }
  const ov = doc.createElementNS(NS.ct, 'Override');
  ov.setAttribute('PartName',    '/' + partName);
  ov.setAttribute('ContentType', CT_MAP[ext] || 'application/octet-stream');
  root.appendChild(ov);
  files[ctPath] = xmlBytes(doc);
}

// ── PptxWriter ────────────────────────────────────────────────────────────────

export class PptxWriter {
  constructor(files) {
    /** @private Mutable copy of all ZIP entries */
    this._files = files;

    // Parse presentation.xml once
    this._presPath = 'ppt/presentation.xml';
    this._presDoc  = readXml(files, this._presPath);
    this._presRels = parseRels(files, this._presPath);

    // Build ordered slide path list
    this._slidePaths = this._buildSlidePaths();
  }

  // ── Factory ─────────────────────────────────────────────────────────────────

  /** Clone from an already-loaded PptxRenderer. O(1) — shares byte arrays. */
  static fromRenderer(renderer) {
    // Deep-copy the files map so mutations don't affect the renderer
    const files = {};
    for (const [k, v] of Object.entries(renderer._files)) {
      files[k] = v instanceof Uint8Array ? v.slice() : v;
    }
    return new PptxWriter(files);
  }

  /** Parse from raw ArrayBuffer or Uint8Array. */
  static async fromBytes(buffer) {
    const files = await readZip(buffer);
    return new PptxWriter(files);
  }

  /**
   * Create a new PPTX from scratch with a blank slide.
   *
   * @param {object} [opts]
   * @param {number} [opts.width=9144000]   slide width in EMU  (default 10in = widescreen)
   * @param {number} [opts.height=5143500]  slide height in EMU (default ~5.63in = 16:9)
   * @param {string} [opts.title='Presentation']
   * @returns {PptxWriter}
   *
   * @example
   *   const writer = PptxWriter.create();
   *   writer.addTextBox(0, 'Hello World', { x: 914400, y: 914400, w: 7000000, h: 900000, fontSize: 4400 });
   *   await writer.download('new.pptx');
   *
   * @example
   *   // 4:3 aspect ratio
   *   const writer = PptxWriter.create({ width: 9144000, height: 6858000 });
   */
  static create(opts = {}) {
    const {
      width  = 9144000,
      height = 5143500,
      title  = 'Presentation',
    } = opts;

    const files = {};

    // ── [Content_Types].xml ──────────────────────────────────────────────
    files['[Content_Types].xml'] = enc.encode(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`);

    // ── _rels/.rels ──────────────────────────────────────────────────────
    files['_rels/.rels'] = enc.encode(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`);

    // ── docProps/core.xml ────────────────────────────────────────────────
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    files['docProps/core.xml'] = enc.encode(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${escXml(title)}</dc:title>
  <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`);

    // ── docProps/app.xml ─────────────────────────────────────────────────
    files['docProps/app.xml'] = enc.encode(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>pptx-browser</Application>
  <Slides>1</Slides>
</Properties>`);

    // ── ppt/presentation.xml ─────────────────────────────────────────────
    files['ppt/presentation.xml'] = enc.encode(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
  <p:sldSz cx="${width}" cy="${height}"/>
  <p:notesSz cx="${height}" cy="${width}"/>
</p:presentation>`);

    files['ppt/_rels/presentation.xml.rels'] = enc.encode(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`);

    // ── ppt/theme/theme1.xml ─────────────────────────────────────────────
    files['ppt/theme/theme1.xml'] = enc.encode(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
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
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`);

    // ── ppt/slideMasters/slideMaster1.xml ────────────────────────────────
    files['ppt/slideMasters/slideMaster1.xml'] = enc.encode(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`);

    files['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = enc.encode(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`);

    // ── ppt/slideLayouts/slideLayout1.xml ─────────────────────────────────
    files['ppt/slideLayouts/slideLayout1.xml'] = enc.encode(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`);

    files['ppt/slideLayouts/_rels/slideLayout1.xml.rels'] = enc.encode(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`);

    // ── ppt/slides/slide1.xml ────────────────────────────────────────────
    files['ppt/slides/slide1.xml'] = enc.encode(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
</p:sld>`);

    files['ppt/slides/_rels/slide1.xml.rels'] = enc.encode(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`);

    return new PptxWriter(files);
  }

  // ── Slide list ──────────────────────────────────────────────────────────────

  _buildSlidePaths() {
    if (!this._presDoc) return [];
    const sldIdLst = g1(this._presDoc, 'sldIdLst');
    if (!sldIdLst) return [];
    const paths = [];
    for (const sldId of sldIdLst.children) {
      if (sldId.localName !== 'sldId') continue;
      const rId = sldId.getAttribute('r:id') || sldId.getAttribute('id');
      const rel = this._presRels[rId];
      if (rel) paths.push(rel.fullPath);
    }
    return paths;
  }

  _savePresDoc() {
    this._files[this._presPath] = xmlBytes(this._presDoc);
  }

  _savePresRels() {
    this._files[relsPath(this._presPath)] = xmlBytes(buildRelsDoc(this._presRels));
  }

  getSlidePaths()  { return [...this._slidePaths]; }
  getSlideCount()  { return this._slidePaths.length; }

  _slideDoc(idx) {
    const path = this._slidePaths[idx];
    if (!path) throw new RangeError(`Slide ${idx} out of range`);
    return readXml(this._files, path);
  }

  _saveSlideDoc(idx, doc) {
    this._files[this._slidePaths[idx]] = xmlBytes(doc);
  }

  // ── Template substitution ────────────────────────────────────────────────────

  /**
   * Replace `{{key}}` placeholders with values from a data object.
   * Applied to every slide, every text shape, and speaker notes.
   *
   * @param {Record<string, string|number>} data
   * @param {object} [opts]
   * @param {string} [opts.open='{{']
   * @param {string} [opts.close='}}']
   * @param {number[]} [opts.slides]  limit to specific slide indices
   */
  applyTemplate(data, opts = {}) {
    const { open = '{{', close = '}}', slides } = opts;
    const indices = slides ?? this._slidePaths.map((_, i) => i);

    for (const idx of indices) {
      const doc = this._slideDoc(idx);

      for (const [key, value] of Object.entries(data)) {
        const token = open + key + close;
        replaceInDoc(doc, token, String(value));
      }

      this._saveSlideDoc(idx, doc);
    }

    // Also apply to speaker notes
    for (const idx of indices) {
      this._applyTemplateToNotes(idx, data, open, close);
    }
  }

  _applyTemplateToNotes(idx, data, open, close) {
    const slideRels = parseRels(this._files, this._slidePaths[idx]);
    const notesRel  = Object.values(slideRels).find(r => r.type?.includes('notesSlide'));
    if (!notesRel) return;
    const notesDoc = readXml(this._files, notesRel.fullPath);
    if (!notesDoc) return;
    for (const [key, value] of Object.entries(data)) {
      replaceInDoc(notesDoc, open + key + close, String(value));
    }
    this._files[notesRel.fullPath] = xmlBytes(notesDoc);
  }

  // ── Global find-and-replace ──────────────────────────────────────────────────

  /**
   * Find and replace text across all (or specified) slides.
   * @param {string} find
   * @param {string} replace
   * @param {object} [opts]
   * @param {boolean} [opts.caseSensitive=true]
   * @param {boolean} [opts.includeNotes=false]
   * @param {number[]} [opts.slides]
   */
  replaceText(find, replace, opts = {}) {
    const { caseSensitive = true, includeNotes = false, slides } = opts;
    const indices = slides ?? this._slidePaths.map((_, i) => i);

    for (const idx of indices) {
      const doc = this._slideDoc(idx);
      replaceInDoc(doc, find, replace, caseSensitive);
      this._saveSlideDoc(idx, doc);

      if (includeNotes) {
        const slideRels = parseRels(this._files, this._slidePaths[idx]);
        const notesRel  = Object.values(slideRels).find(r => r.type?.includes('notesSlide'));
        if (notesRel) {
          const nd = readXml(this._files, notesRel.fullPath);
          if (nd) {
            replaceInDoc(nd, find, replace, caseSensitive);
            this._files[notesRel.fullPath] = xmlBytes(nd);
          }
        }
      }
    }
  }

  // ── Shape text ───────────────────────────────────────────────────────────────

  /**
   * Set the text content of a named shape on a slide.
   * Preserves the formatting of the first run; clears all other runs.
   *
   * @param {number} slideIdx
   * @param {string} shapeName     exact `name` attribute of the shape
   * @param {string} text          new text (use \n for line breaks)
   * @param {object} [opts]
   * @param {boolean} [opts.preserveFormatting=true]
   */
  setShapeText(slideIdx, shapeName, text, opts = {}) {
    const { preserveFormatting = true } = opts;
    const doc    = this._slideDoc(slideIdx);
    const spTree = getSpTree(doc);
    if (!spTree) return this;

    const shape = findShapeByName(spTree, shapeName);
    if (!shape) throw new Error(`Shape "${shapeName}" not found on slide ${slideIdx}`);

    const txBody = g1(shape, 'txBody');
    if (!txBody) return this;

    // Get reference run properties
    const firstRun = g1(txBody, 'r');
    const refRPr   = firstRun ? g1(firstRun, 'rPr') : null;
    const refPPr   = g1(g1(txBody, 'p'), 'pPr');

    // Remove all existing paragraphs
    for (const p of gtn(txBody, 'p')) p.parentNode.removeChild(p);

    const lines = text.split('\n');
    const nsA   = NS.a;

    for (const line of lines) {
      const p = doc.createElementNS(nsA, 'a:p');

      if (refPPr && preserveFormatting) {
        p.appendChild(refPPr.cloneNode(true));
      }

      const r  = doc.createElementNS(nsA, 'a:r');
      if (refRPr && preserveFormatting) {
        r.appendChild(refRPr.cloneNode(true));
      }
      const t = doc.createElementNS(nsA, 'a:t');
      t.textContent = line;
      r.appendChild(t);
      p.appendChild(r);
      txBody.appendChild(p);
    }

    this._saveSlideDoc(slideIdx, doc);
    return this;
  }

  /**
   * Read the plain text of a named shape.
   * @param {number} slideIdx
   * @param {string} shapeName
   * @returns {string}
   */
  getShapeText(slideIdx, shapeName) {
    const doc    = this._slideDoc(slideIdx);
    const spTree = getSpTree(doc);
    if (!spTree) return '';
    const shape  = findShapeByName(spTree, shapeName);
    if (!shape) return '';
    return gtn(shape, 't').map(t => t.textContent).join('');
  }

  // ── Add text box ─────────────────────────────────────────────────────────────

  /**
   * Add a new text box to a slide.
   *
   * @param {number} slideIdx
   * @param {string} text          use \n for line breaks
   * @param {object} style
   * @param {number} style.x           EMU from left edge
   * @param {number} style.y           EMU from top edge
   * @param {number} style.w           EMU width
   * @param {number} style.h           EMU height
   * @param {string} [style.color]     hex colour, no #
   * @param {number} [style.fontSize]  pt * 100  (e.g. 2400 = 24pt)
   * @param {boolean}[style.bold]
   * @param {boolean}[style.italic]
   * @param {boolean}[style.underline]
   * @param {boolean}[style.strikethrough]
   * @param {string} [style.align]     l|ctr|r|just
   * @param {string} [style.vertAlign] t|ctr|b  (vertical alignment)
   * @param {string} [style.fontFamily]
   * @param {string} [style.fill]      shape background, hex no #
   * @param {string} [style.outline]   border colour, hex no #
   * @param {number} [style.outlineWidth] border width EMU (default 12700 = 1pt)
   * @param {number} [style.lineSpacing]  line spacing in hundredths of a percent (e.g. 150000 = 150%)
   * @param {number} [style.rotation]     rotation in 60000ths of a degree (e.g. 5400000 = 90°)
   */
  addTextBox(slideIdx, text, style = {}) {
    const {
      x = 914400, y = 914400, w = 4572000, h = 914400,
      color = '000000', fontSize = 1800, bold = false,
      italic = false, underline = false, strikethrough = false,
      align = 'l', vertAlign, fontFamily = 'Calibri',
      fill, outline, outlineWidth = 12700, lineSpacing, rotation,
    } = style;

    const doc    = this._slideDoc(slideIdx);
    const spTree = getSpTree(doc);
    if (!spTree) return this;

    const maxId = Math.max(0, ...gtn(spTree, 'cNvPr').map(e => parseInt(e.getAttribute('id') || '0', 10)));
    const newId = maxId + 1;
    const name  = `TextBox ${newId}`;
    const nsA = NS.a, nsP = NS.p;

    const fillXml = fill
      ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`
      : `<a:noFill/>`;
    const lnXml = outline
      ? `<a:ln w="${outlineWidth}"><a:solidFill><a:srgbClr val="${outline}"/></a:solidFill></a:ln>`
      : '';
    const rotAttr = rotation ? ` rot="${rotation}"` : '';
    const anchorAttr = vertAlign ? ` anchor="${vertAlign}"` : '';
    const spcAttr = lineSpacing
      ? `<a:lnSpc><a:spcPct val="${lineSpacing}"/></a:lnSpc>`
      : '';

    const lines = text.split('\n');
    const parasXml = lines.map(line =>
      `<a:p><a:pPr algn="${align}">${spcAttr}</a:pPr>` +
      `<a:r><a:rPr lang="en-US" sz="${fontSize}" b="${bold ? 1 : 0}" i="${italic ? 1 : 0}"` +
      `${underline ? ' u="sng"' : ''}${strikethrough ? ' strike="sngStrike"' : ''} dirty="0">` +
      `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
      `<a:latin typeface="${escXml(fontFamily)}"/>` +
      `</a:rPr><a:t>${escXml(line)}</a:t></a:r></a:p>`
    ).join('');

    const xml = `<p:sp xmlns:p="${nsP}" xmlns:a="${nsA}">
  <p:nvSpPr>
    <p:cNvPr id="${newId}" name="${name}"/>
    <p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm${rotAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    ${fillXml}${lnXml}
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" rtlCol="0"${anchorAttr}><a:spAutoFit/></a:bodyPr>
    <a:lstStyle/>
    ${parasXml}
  </p:txBody>
</p:sp>`;

    const frag = parseXml(xml);
    spTree.appendChild(doc.adoptNode(frag.documentElement));
    this._saveSlideDoc(slideIdx, doc);
    return this;
  }

  /**
   * Add a text box with mixed formatting (rich text).
   * Each run can have its own font, size, colour, bold, italic, etc.
   *
   * @param {number} slideIdx
   * @param {Array<Array<{text, color?, fontSize?, bold?, italic?, underline?, strikethrough?, fontFamily?}>>} paragraphs
   *   Array of paragraphs, each paragraph is an array of run objects.
   * @param {object} style
   * @param {number} style.x        EMU from left
   * @param {number} style.y        EMU from top
   * @param {number} style.w        EMU width
   * @param {number} style.h        EMU height
   * @param {string} [style.align]  l|ctr|r|just
   * @param {string} [style.vertAlign] t|ctr|b
   * @param {string} [style.fill]   shape fill hex
   * @param {string} [style.outline] border hex
   * @param {number} [style.outlineWidth]
   * @param {number} [style.lineSpacing]
   * @param {number} [style.rotation]
   *
   * @example
   *   writer.addRichText(0, [
   *     [
   *       { text: 'Bold title', bold: true, fontSize: 3200, color: '1F4E79' },
   *     ],
   *     [
   *       { text: 'Normal text ', fontSize: 1800 },
   *       { text: 'with red highlight', fontSize: 1800, color: 'FF0000', italic: true },
   *     ],
   *   ], { x: 914400, y: 914400, w: 7000000, h: 2000000 });
   */
  addRichText(slideIdx, paragraphs, style = {}) {
    const {
      x = 914400, y = 914400, w = 4572000, h = 914400,
      align = 'l', vertAlign, fill, outline, outlineWidth = 12700,
      lineSpacing, rotation,
    } = style;

    const doc    = this._slideDoc(slideIdx);
    const spTree = getSpTree(doc);
    if (!spTree) return this;

    const maxId = Math.max(0, ...gtn(spTree, 'cNvPr').map(e => parseInt(e.getAttribute('id') || '0', 10)));
    const newId = maxId + 1;
    const nsA = NS.a, nsP = NS.p;

    const fillXml = fill
      ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`
      : `<a:noFill/>`;
    const lnXml = outline
      ? `<a:ln w="${outlineWidth}"><a:solidFill><a:srgbClr val="${outline}"/></a:solidFill></a:ln>`
      : '';
    const rotAttr = rotation ? ` rot="${rotation}"` : '';
    const anchorAttr = vertAlign ? ` anchor="${vertAlign}"` : '';
    const spcXml = lineSpacing ? `<a:lnSpc><a:spcPct val="${lineSpacing}"/></a:lnSpc>` : '';

    let parasXml = '';
    for (const para of paragraphs) {
      parasXml += `<a:p><a:pPr algn="${align}">${spcXml}</a:pPr>`;
      for (const run of para) {
        const sz = run.fontSize ?? 1800;
        const clr = run.color ?? '000000';
        const ff = run.fontFamily ?? 'Calibri';
        parasXml += `<a:r><a:rPr lang="en-US" sz="${sz}" b="${run.bold ? 1 : 0}" i="${run.italic ? 1 : 0}"` +
          `${run.underline ? ' u="sng"' : ''}${run.strikethrough ? ' strike="sngStrike"' : ''} dirty="0">` +
          `<a:solidFill><a:srgbClr val="${clr}"/></a:solidFill>` +
          `<a:latin typeface="${escXml(ff)}"/>` +
          `</a:rPr><a:t>${escXml(run.text)}</a:t></a:r>`;
      }
      parasXml += `</a:p>`;
    }

    const xml = `<p:sp xmlns:p="${nsP}" xmlns:a="${nsA}">
  <p:nvSpPr>
    <p:cNvPr id="${newId}" name="TextBox ${newId}"/>
    <p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm${rotAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    ${fillXml}${lnXml}
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" rtlCol="0"${anchorAttr}><a:spAutoFit/></a:bodyPr>
    <a:lstStyle/>
    ${parasXml}
  </p:txBody>
</p:sp>`;

    const frag = parseXml(xml);
    spTree.appendChild(doc.adoptNode(frag.documentElement));
    this._saveSlideDoc(slideIdx, doc);
    return this;
  }

  /**
   * Add a preset shape (rectangle, ellipse, arrow, etc.) to a slide.
   *
   * @param {number} slideIdx
   * @param {string} shapeType   preset geometry name:
   *   rect, roundRect, ellipse, triangle, diamond, pentagon, hexagon,
   *   star5, star6, rightArrow, leftArrow, upArrow, downArrow,
   *   heart, cloud, line, plus, can, cube, donut, …
   *   (any PowerPoint preset geometry name)
   * @param {object} style
   * @param {number} style.x           EMU from left
   * @param {number} style.y           EMU from top
   * @param {number} style.w           EMU width
   * @param {number} style.h           EMU height
   * @param {string} [style.fill='4472C4']   fill colour hex, no #
   * @param {string} [style.outline]         border colour hex
   * @param {number} [style.outlineWidth=12700]
   * @param {string} [style.text]      optional text inside the shape
   * @param {string} [style.textColor='FFFFFF']
   * @param {number} [style.fontSize=1800]   pt * 100
   * @param {boolean}[style.bold]
   * @param {boolean}[style.italic]
   * @param {string} [style.fontFamily='Calibri']
   * @param {string} [style.align='ctr']
   * @param {string} [style.vertAlign='ctr']  t|ctr|b
   * @param {number} [style.rotation]
   */
  addShape(slideIdx, shapeType, style = {}) {
    const {
      x = 914400, y = 914400, w = 2743200, h = 2743200,
      fill = '4472C4', outline, outlineWidth = 12700,
      text, textColor = 'FFFFFF', fontSize = 1800,
      bold = false, italic = false, fontFamily = 'Calibri',
      align = 'ctr', vertAlign = 'ctr', rotation,
    } = style;

    const doc    = this._slideDoc(slideIdx);
    const spTree = getSpTree(doc);
    if (!spTree) return this;

    const maxId = Math.max(0, ...gtn(spTree, 'cNvPr').map(e => parseInt(e.getAttribute('id') || '0', 10)));
    const newId = maxId + 1;
    const nsA = NS.a, nsP = NS.p;

    const fillXml = fill
      ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`
      : `<a:noFill/>`;
    const lnXml = outline
      ? `<a:ln w="${outlineWidth}"><a:solidFill><a:srgbClr val="${outline}"/></a:solidFill></a:ln>`
      : '';
    const rotAttr = rotation ? ` rot="${rotation}"` : '';

    let txBodyXml = '';
    if (text !== undefined && text !== null) {
      txBodyXml = `<p:txBody>
    <a:bodyPr wrap="square" rtlCol="0" anchor="${vertAlign}"/>
    <a:lstStyle/>
    <a:p><a:pPr algn="${align}"/>
      <a:r><a:rPr lang="en-US" sz="${fontSize}" b="${bold ? 1 : 0}" i="${italic ? 1 : 0}" dirty="0">
        <a:solidFill><a:srgbClr val="${textColor}"/></a:solidFill>
        <a:latin typeface="${escXml(fontFamily)}"/>
      </a:rPr><a:t>${escXml(text)}</a:t></a:r>
    </a:p>
  </p:txBody>`;
    }

    const xml = `<p:sp xmlns:p="${nsP}" xmlns:a="${nsA}">
  <p:nvSpPr>
    <p:cNvPr id="${newId}" name="${shapeType} ${newId}"/>
    <p:cNvSpPr/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm${rotAttr}><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
    <a:prstGeom prst="${escXml(shapeType)}"><a:avLst/></a:prstGeom>
    ${fillXml}${lnXml}
  </p:spPr>
  ${txBodyXml}
</p:sp>`;

    const frag = parseXml(xml);
    spTree.appendChild(doc.adoptNode(frag.documentElement));
    this._saveSlideDoc(slideIdx, doc);
    return this;
  }

  /**
   * Add a bulleted or numbered list to a slide.
   *
   * @param {number} slideIdx
   * @param {string[]} items     list items (strings)
   * @param {object}   style
   * @param {number}   style.x         EMU from left
   * @param {number}   style.y         EMU from top
   * @param {number}   style.w         EMU width
   * @param {number}   style.h         EMU height
   * @param {string}   [style.color='000000']
   * @param {number}   [style.fontSize=1800]   pt * 100
   * @param {boolean}  [style.bold]
   * @param {boolean}  [style.italic]
   * @param {string}   [style.fontFamily='Calibri']
   * @param {string}   [style.bulletChar='•']  set to '' for no bullet, or '1' for numbered
   * @param {string}   [style.bulletColor]     hex, defaults to text color
   * @param {string}   [style.fill]            background fill hex
   * @param {string}   [style.align='l']
   */
  addList(slideIdx, items, style = {}) {
    const {
      x = 914400, y = 914400, w = 7000000, h = 3000000,
      color = '000000', fontSize = 1800, bold = false, italic = false,
      fontFamily = 'Calibri', bulletChar = '\u2022',
      bulletColor, fill, align = 'l',
    } = style;

    const doc    = this._slideDoc(slideIdx);
    const spTree = getSpTree(doc);
    if (!spTree) return this;

    const maxId = Math.max(0, ...gtn(spTree, 'cNvPr').map(e => parseInt(e.getAttribute('id') || '0', 10)));
    const newId = maxId + 1;
    const nsA = NS.a, nsP = NS.p;
    const bClr = bulletColor || color;

    const fillXml = fill
      ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`
      : `<a:noFill/>`;

    const isNumbered = bulletChar === '1';

    let parasXml = '';
    for (let i = 0; i < items.length; i++) {
      let bulletXml;
      if (isNumbered) {
        bulletXml = `<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>`;
      } else if (bulletChar) {
        bulletXml = `<a:buClr><a:srgbClr val="${bClr}"/></a:buClr>` +
          `<a:buSzPct val="100000"/>` +
          `<a:buChar char="${escXml(bulletChar)}"/>`;
      } else {
        bulletXml = '<a:buNone/>';
      }

      parasXml += `<a:p><a:pPr algn="${align}" marL="342900" indent="-342900">${bulletXml}</a:pPr>` +
        `<a:r><a:rPr lang="en-US" sz="${fontSize}" b="${bold ? 1 : 0}" i="${italic ? 1 : 0}" dirty="0">` +
        `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
        `<a:latin typeface="${escXml(fontFamily)}"/>` +
        `</a:rPr><a:t>${escXml(items[i])}</a:t></a:r></a:p>`;
    }

    const xml = `<p:sp xmlns:p="${nsP}" xmlns:a="${nsA}">
  <p:nvSpPr>
    <p:cNvPr id="${newId}" name="TextBox ${newId}"/>
    <p:cNvSpPr txBox="1"><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    ${fillXml}
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr>
    <a:lstStyle/>
    ${parasXml}
  </p:txBody>
</p:sp>`;

    const frag = parseXml(xml);
    spTree.appendChild(doc.adoptNode(frag.documentElement));
    this._saveSlideDoc(slideIdx, doc);
    return this;
  }

  // ── Image replacement ─────────────────────────────────────────────────────────

  /**
   * Replace the image in a named picture shape.
   *
   * @param {number}     slideIdx
   * @param {string}     shapeName
   * @param {Uint8Array} imageBytes
   * @param {string}     [mimeType='image/jpeg']
   */
  async setShapeImage(slideIdx, shapeName, imageBytes, mimeType = 'image/jpeg') {
    const doc    = this._slideDoc(slideIdx);
    const spTree = getSpTree(doc);
    if (!spTree) return this;

    const shape = findShapeByName(spTree, shapeName);
    if (!shape) throw new Error(`Shape "${shapeName}" not found on slide ${slideIdx}`);

    const slideRels = parseRels(this._files, this._slidePaths[slideIdx]);

    // Find existing blip rId
    const blipFill = g1(shape, 'blipFill');
    const blip     = blipFill ? g1(blipFill, 'blip') : null;
    const oldRId   = blip ? (blip.getAttribute('r:embed') || blip.getAttribute('embed')) : null;
    const oldRel   = oldRId ? slideRels[oldRId] : null;

    // Write new media file
    const ext      = MIME_EXT[mimeType] || 'jpeg';
    const mediaIdx = Object.keys(this._files).filter(p => p.startsWith('ppt/media/')).length + 1;
    const mediaPath = `ppt/media/image${mediaIdx}.${ext}`;
    this._files[mediaPath] = imageBytes;

    // Update or create relationship
    let rId;
    if (oldRId && oldRel) {
      // Reuse the old rId, just point to the new file
      rId = oldRId;
      slideRels[rId] = {
        id: rId, type: oldRel.type,
        target: `../media/image${mediaIdx}.${ext}`,
        fullPath: mediaPath,
      };
    } else {
      rId = nextRId(slideRels);
      slideRels[rId] = {
        id: rId,
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        target: `../media/image${mediaIdx}.${ext}`,
        fullPath: mediaPath,
      };
    }

    // Update the blip element
    if (blip) {
      blip.setAttribute('r:embed', rId);
    }

    // Save updated rels and slide doc
    this._files[relsPath(this._slidePaths[slideIdx])] = xmlBytes(buildRelsDoc(slideRels));
    this._saveSlideDoc(slideIdx, doc);

    // Add content type
    addContentType(this._files, ext, mediaPath);
    return this;
  }

  /**
   * Add a new image shape to a slide.
   *
   * @param {number}     slideIdx
   * @param {Uint8Array} imageBytes
   * @param {string}     [mimeType='image/jpeg']
   * @param {object}     rect      { x, y, w, h } in EMU
   */
  async addImage(slideIdx, imageBytes, mimeType = 'image/jpeg', rect = {}) {
    const {
      x = 914400, y = 914400,
      w = 2743200, h = 2057400,
    } = rect;

    const ext       = MIME_EXT[mimeType] || 'jpeg';
    const mediaIdx  = Object.keys(this._files).filter(p => p.startsWith('ppt/media/')).length + 1;
    const mediaPath = `ppt/media/image${mediaIdx}.${ext}`;
    this._files[mediaPath] = imageBytes;

    const slideRels = parseRels(this._files, this._slidePaths[slideIdx]);
    const rId = nextRId(slideRels);
    slideRels[rId] = {
      id: rId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      target: `../media/image${mediaIdx}.${ext}`,
      fullPath: mediaPath,
    };
    this._files[relsPath(this._slidePaths[slideIdx])] = xmlBytes(buildRelsDoc(slideRels));

    const doc    = this._slideDoc(slideIdx);
    const spTree = getSpTree(doc);
    if (!spTree) return this;

    const maxId = Math.max(0, ...gtn(spTree, 'cNvPr').map(e => parseInt(e.getAttribute('id') || '0', 10)));
    const newId = maxId + 1;
    const nsA = NS.a, nsP = NS.p;

    const xml = `<p:pic xmlns:p="${nsP}" xmlns:a="${nsA}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:nvPicPr>
    <p:cNvPr id="${newId}" name="Picture ${newId}"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="${rId}"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`;

    const frag = parseXml(xml);
    spTree.appendChild(doc.adoptNode(frag.documentElement));
    this._saveSlideDoc(slideIdx, doc);
    addContentType(this._files, ext, mediaPath);
    return this;
  }

  // ── Slide background ─────────────────────────────────────────────────────────

  /**
   * Set a solid colour background on a slide.
   * @param {number} slideIdx
   * @param {string} hexRgb   6-digit hex, no '#'
   */
  setSlideBackground(slideIdx, hexRgb) {
    const doc  = this._slideDoc(slideIdx);
    const cSld = g1(doc, 'cSld');
    if (!cSld) return this;

    // Remove existing bg
    const oldBg = g1(cSld, 'bg');
    if (oldBg) cSld.removeChild(oldBg);

    const nsA = NS.a, nsP = NS.p;
    const xml = `<p:bg xmlns:p="${nsP}" xmlns:a="${nsA}">
  <p:bgPr><a:solidFill><a:srgbClr val="${hexRgb}"/></a:solidFill>
  <a:effectLst/></p:bgPr></p:bg>`;
    const bgEl = doc.adoptNode(parseXml(xml).documentElement);
    // Insert bg as first child of cSld
    cSld.insertBefore(bgEl, cSld.firstChild);
    this._saveSlideDoc(slideIdx, doc);
    return this;
  }

  // ── Theme colours ─────────────────────────────────────────────────────────────

  /**
   * Override a theme colour.
   * Key: dk1|lt1|dk2|lt2|accent1…accent6|hlink|folHlink
   * Value: 6-digit hex RGB, no '#'
   *
   * @param {string} key
   * @param {string} hexRgb
   */
  setThemeColor(key, hexRgb) {
    // Find theme file via presentation rels
    const presRels = this._presRels;
    let themePath  = Object.values(presRels).find(r => r.type?.includes('theme'))?.fullPath;

    if (!themePath) {
      // Try via slide master
      const masterRel = Object.values(presRels).find(r => r.type?.includes('slideMaster'));
      if (masterRel) {
        const mr = parseRels(this._files, masterRel.fullPath);
        themePath = Object.values(mr).find(r => r.type?.includes('theme'))?.fullPath;
      }
    }
    if (!themePath) return this;

    const doc = readXml(this._files, themePath);
    if (!doc) return this;

    // Map theme key to element path: e.g. accent1 → a:accent1 > a:srgbClr
    const fmtScheme = g1(doc, 'fmtScheme');
    const clrScheme = g1(doc, 'clrScheme');
    if (!clrScheme) return this;

    // Find the element with matching local name
    for (const child of clrScheme.children) {
      if (child.localName === key) {
        // Replace or set inner colour element
        const srgb = g1(child, 'srgbClr');
        if (srgb) {
          srgb.setAttribute('val', hexRgb);
        } else {
          while (child.firstChild) child.removeChild(child.firstChild);
          const nsA = NS.a;
          const el  = doc.createElementNS(nsA, 'a:srgbClr');
          el.setAttribute('val', hexRgb);
          child.appendChild(el);
        }
        break;
      }
    }

    this._files[themePath] = xmlBytes(doc);
    return this;
  }

  // ── Slide operations ──────────────────────────────────────────────────────────

  /**
   * Duplicate a slide.
   * @param {number} fromIdx       source slide index
   * @param {number} [toIdx]       insert position (default: end)
   */
  duplicateSlide(fromIdx, toIdx) {
    const insertAt = toIdx ?? this._slidePaths.length;
    const srcPath  = this._slidePaths[fromIdx];
    if (!srcPath) throw new RangeError(`Slide ${fromIdx} out of range`);

    // Find next available slide number
    const nums = Object.keys(this._files)
      .map(p => p.match(/ppt\/slides\/slide(\d+)\.xml/))
      .filter(Boolean).map(m => parseInt(m[1], 10));
    const nextNum = (nums.length ? Math.max(...nums) : 0) + 1;

    const newSlidePath = `ppt/slides/slide${nextNum}.xml`;
    const newRelsPath  = relsPath(newSlidePath);

    // Copy slide XML
    this._files[newSlidePath] = this._files[srcPath].slice();

    // Copy slide rels (images etc. are shared)
    const srcRelsPath = relsPath(srcPath);
    if (this._files[srcRelsPath]) {
      this._files[newRelsPath] = this._files[srcRelsPath].slice();
    }

    // Add relationship in presentation.xml.rels
    const newRId = nextRId(this._presRels);
    const target = `slides/slide${nextNum}.xml`;
    this._presRels[newRId] = {
      id: newRId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
      target,
      fullPath: newSlidePath,
    };
    this._savePresRels();

    // Add sldId to presentation.xml sldIdLst
    const sldIdLst = g1(this._presDoc, 'sldIdLst');
    if (sldIdLst) {
      const ids = gtn(sldIdLst, 'sldId').map(el => parseInt(el.getAttribute('id') || '0', 10));
      const nextId = (ids.length ? Math.max(...ids) : 255) + 1;
      const nsP  = NS.p;
      const sldIdEl = this._presDoc.createElementNS(nsP, 'p:sldId');
      sldIdEl.setAttribute('id', String(nextId));
      sldIdEl.setAttributeNS(NS.r, 'r:id', newRId);

      // Insert at correct position
      const children = Array.from(sldIdLst.children);
      if (insertAt >= children.length) {
        sldIdLst.appendChild(sldIdEl);
      } else {
        sldIdLst.insertBefore(sldIdEl, children[insertAt]);
      }
    }
    this._savePresDoc();

    // Rebuild slide path list
    this._slidePaths = this._buildSlidePaths();

    // Add content type for new slide
    const ctPath = '[Content_Types].xml';
    const ctDoc  = readXml(this._files, ctPath);
    if (ctDoc) {
      const root = ctDoc.documentElement;
      const ov   = ctDoc.createElementNS(NS.ct, 'Override');
      ov.setAttribute('PartName',    '/' + newSlidePath);
      ov.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
      root.appendChild(ov);
      this._files[ctPath] = xmlBytes(ctDoc);
    }

    return this;
  }

  /**
   * Remove a slide.
   * @param {number} slideIdx
   */
  removeSlide(slideIdx) {
    if (this._slidePaths.length <= 1) throw new Error('Cannot remove the last slide');
    const path = this._slidePaths[slideIdx];
    if (!path) throw new RangeError(`Slide ${slideIdx} out of range`);

    // Remove from sldIdLst
    const sldIdLst = g1(this._presDoc, 'sldIdLst');
    if (sldIdLst) {
      for (const sldId of Array.from(sldIdLst.children)) {
        const rId = sldId.getAttribute('r:id') || sldId.getAttribute('id');
        const rel = this._presRels[rId];
        if (rel && rel.fullPath === path) {
          sldIdLst.removeChild(sldId);
          delete this._presRels[rId];
          break;
        }
      }
    }
    this._savePresDoc();
    this._savePresRels();
    this._slidePaths = this._buildSlidePaths();
    return this;
  }

  /**
   * Add a new blank slide.
   * @param {number} [atIdx]  insert position (default: end)
   * @returns {PptxWriter}
   */
  addSlide(atIdx) {
    const insertAt = atIdx ?? this._slidePaths.length;

    // Find next available slide number
    const nums = Object.keys(this._files)
      .map(p => p.match(/ppt\/slides\/slide(\d+)\.xml/))
      .filter(Boolean).map(m => parseInt(m[1], 10));
    const nextNum = (nums.length ? Math.max(...nums) : 0) + 1;

    const newSlidePath = `ppt/slides/slide${nextNum}.xml`;
    const nsP = NS.p, nsA = NS.a;

    this._files[newSlidePath] = enc.encode(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="${nsP}" xmlns:a="${nsA}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
    </p:spTree>
  </p:cSld>
</p:sld>`);

    // Point the new slide at a layout (use the first layout available)
    const layoutTarget = Object.keys(this._files).find(p => p.match(/ppt\/slideLayouts\/slideLayout\d+\.xml/));
    const layoutRelTarget = layoutTarget ? '../slideLayouts/' + layoutTarget.split('/').pop() : '../slideLayouts/slideLayout1.xml';
    this._files[relsPath(newSlidePath)] = enc.encode(
`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="${layoutRelTarget}"/>
</Relationships>`);

    // Add relationship in presentation.xml.rels
    const newRId = nextRId(this._presRels);
    this._presRels[newRId] = {
      id: newRId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
      target: `slides/slide${nextNum}.xml`,
      fullPath: newSlidePath,
    };
    this._savePresRels();

    // Add sldId to presentation.xml
    const sldIdLst = g1(this._presDoc, 'sldIdLst');
    if (sldIdLst) {
      const ids = gtn(sldIdLst, 'sldId').map(el => parseInt(el.getAttribute('id') || '0', 10));
      const nextId = (ids.length ? Math.max(...ids) : 255) + 1;
      const sldIdEl = this._presDoc.createElementNS(NS.p, 'p:sldId');
      sldIdEl.setAttribute('id', String(nextId));
      sldIdEl.setAttributeNS(NS.r, 'r:id', newRId);

      const children = Array.from(sldIdLst.children);
      if (insertAt >= children.length) {
        sldIdLst.appendChild(sldIdEl);
      } else {
        sldIdLst.insertBefore(sldIdEl, children[insertAt]);
      }
    }
    this._savePresDoc();
    this._slidePaths = this._buildSlidePaths();

    // Add content type
    const ctPath = '[Content_Types].xml';
    const ctDoc  = readXml(this._files, ctPath);
    if (ctDoc) {
      const root = ctDoc.documentElement;
      const ov   = ctDoc.createElementNS(NS.ct, 'Override');
      ov.setAttribute('PartName',    '/' + newSlidePath);
      ov.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
      root.appendChild(ov);
      this._files[ctPath] = xmlBytes(ctDoc);
    }

    return this;
  }

  /**
   * Reorder slides.
   * @param {number[]} newOrder  e.g. [2, 0, 1] to put slide 2 first
   */
  reorderSlides(newOrder) {
    if (newOrder.length !== this._slidePaths.length) {
      throw new Error('newOrder must have the same length as the current slide count');
    }

    const sldIdLst = g1(this._presDoc, 'sldIdLst');
    if (!sldIdLst) return this;

    const children = Array.from(sldIdLst.children).filter(el => el.localName === 'sldId');
    // Detach all
    for (const c of children) sldIdLst.removeChild(c);
    // Re-attach in new order
    for (const idx of newOrder) {
      if (children[idx]) sldIdLst.appendChild(children[idx]);
    }
    this._savePresDoc();
    this._slidePaths = this._buildSlidePaths();
    return this;
  }

  // ── Speaker notes ─────────────────────────────────────────────────────────────

  /**
   * Set the speaker notes for a slide. Creates the notes slide if absent.
   * @param {number} slideIdx
   * @param {string} text
   */
  setSlideNotes(slideIdx, text) {
    const slidePath = this._slidePaths[slideIdx];
    const slideRels = parseRels(this._files, slidePath);
    const notesRel  = Object.values(slideRels).find(r => r.type?.includes('notesSlide'));

    if (notesRel) {
      const nd = readXml(this._files, notesRel.fullPath);
      if (nd) {
        for (const sp of gtn(nd, 'sp')) {
          const nvPr = g1(g1(sp, 'nvSpPr'), 'nvPr');
          const ph   = nvPr ? g1(nvPr, 'ph') : null;
          if (ph && attr(ph, 'type') !== 'sldNum') {
            for (const t of gtn(sp, 't')) t.textContent = '';
            const firstT = g1(sp, 't');
            if (firstT) firstT.textContent = text;
            break;
          }
        }
        this._files[notesRel.fullPath] = xmlBytes(nd);
      }
    } else {
      // Create a minimal notes slide
      this._createNotesSlide(slideIdx, slidePath, slideRels, text);
    }
    return this;
  }

  _createNotesSlide(slideIdx, slidePath, slideRels, text) {
    const num = Object.keys(this._files).filter(p => p.startsWith('ppt/notesSlides/')).length + 1;
    const nsP = NS.p, nsA = NS.a;
    const notesPath = `ppt/notesSlides/notesSlide${num}.xml`;

    const notesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:p="${nsP}" xmlns:a="${nsA}">
  <p:cSld><p:spTree>
    <p:sp>
      <p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/>
        <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
        <p:nvPr><p:ph type="body" idx="1"/></p:nvPr>
      </p:nvSpPr>
      <p:spPr/>
      <p:txBody><a:bodyPr/><a:lstStyle/>
        <a:p><a:r><a:t>${text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}</a:t></a:r></a:p>
      </p:txBody>
    </p:sp>
  </p:spTree></p:cSld>
</p:notes>`;

    this._files[notesPath] = enc.encode(notesXml);

    const newRId = nextRId(slideRels);
    slideRels[newRId] = {
      id: newRId,
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
      target: `../notesSlides/notesSlide${num}.xml`,
      fullPath: notesPath,
    };
    this._files[relsPath(slidePath)] = xmlBytes(buildRelsDoc(slideRels));
  }

  // ── Serialisation ─────────────────────────────────────────────────────────────

  /**
   * Serialize the edited PPTX to bytes.
   * @returns {Promise<Uint8Array>}
   */
  async save() {
    return writeZip(this._files);
  }

  /**
   * Download as a PPTX file in the browser.
   * @param {string} [filename='edited.pptx']
   */
  async download(filename = 'edited.pptx') {
    const bytes = await this.save();
    const blob  = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
    const url   = URL.createObjectURL(blob);
    const a     = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  // ── Utility ───────────────────────────────────────────────────────────────────

  /**
   * List all shape names on a slide.
   * @param {number} slideIdx
   * @returns {Array<{id, name, type}>}
   */
  listShapes(slideIdx) {
    const doc    = this._slideDoc(slideIdx);
    const spTree = getSpTree(doc);
    if (!spTree) return [];
    const shapes = [];
    for (const child of spTree.children) {
      const ln = child.localName;
      if (!['sp','pic','cxnSp','graphicFrame'].includes(ln)) continue;
      const nvEl  = g1(child, 'nvSpPr') || g1(child, 'nvPicPr') || g1(child, 'nvGraphicFramePr') || g1(child, 'nvCxnSpPr');
      const cNvPr = nvEl ? g1(nvEl, 'cNvPr') : null;
      shapes.push({
        id:   cNvPr?.getAttribute('id') || '',
        name: cNvPr?.getAttribute('name') || '',
        type: ln,
      });
    }
    return shapes;
  }
}
