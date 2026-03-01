/**
 * odp.js — ODP (OpenDocument Presentation) writer.
 *
 * Creates ODP files from scratch or converts loaded PPTX presentations
 * to ODP format. Zero dependencies — uses the same native ZIP writer.
 *
 * ODP is the ISO-standardised presentation format used by LibreOffice Impress,
 * Apache OpenOffice, Google Docs, and others.
 *
 * ── Quick start ──────────────────────────────────────────────────────────────
 *
 *   import { OdpWriter } from 'pptx-browser/odp';
 *
 *   // Create from scratch:
 *   const odp = OdpWriter.create();
 *   odp.addTextBox(0, 'Hello World', { x: 2, y: 2, w: 20, h: 3, fontSize: 36 });
 *   await odp.download('hello.odp');
 *
 *   // Convert from a loaded PPTX:
 *   import { PptxRenderer } from 'pptx-browser';
 *   const renderer = new PptxRenderer();
 *   await renderer.load(pptxFile);
 *   const odp = OdpWriter.fromRenderer(renderer);
 *   await odp.download('converted.odp');
 *
 * ── API reference ────────────────────────────────────────────────────────────
 *
 *   OdpWriter.create(opts)              — new blank ODP
 *   OdpWriter.fromRenderer(renderer)    — convert loaded PPTX to ODP
 *
 *   .addSlide()                         — add a blank slide
 *   .removeSlide(idx)                   — remove a slide
 *   .addTextBox(slideIdx, text, style)  — add text box (cm units)
 *   .addImage(slideIdx, bytes, mime, rect) — add image (cm units)
 *   .setSlideBackground(slideIdx, hex)  — solid background
 *   .getSlideCount()                    — number of slides
 *   .save()                             → Promise<Uint8Array>  ODP bytes
 *   .download(filename)                 — browser download
 */

import { writeZip } from './zip-writer.js';

const enc = new TextEncoder();

// ── ODP XML namespaces ──────────────────────────────────────────────────────

const NSMAP = {
  office:   'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  style:    'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  text:     'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  table:    'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  draw:     'urn:oasis:names:tc:opendocument:xmlns:drawing:1.0',
  fo:       'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  xlink:    'http://www.w3.org/1999/xlink',
  dc:       'http://purl.org/dc/elements/1.1/',
  meta:     'urn:oasis:names:tc:opendocument:xmlns:meta:1.0',
  svg:      'urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0',
  presentation: 'urn:oasis:names:tc:opendocument:xmlns:presentation:1.0',
  smil:     'urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0',
  anim:     'urn:oasis:names:tc:opendocument:xmlns:animation:1.0',
  manifest: 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
};

// ── XML helpers ─────────────────────────────────────────────────────────────

function escXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Build the xmlns declarations string used in content.xml and styles.xml. */
function nsDecls() {
  return Object.entries(NSMAP)
    .filter(([k]) => k !== 'manifest')
    .map(([k, v]) => `xmlns:${k}="${v}"`)
    .join(' ');
}

/** Convert EMU to centimeters. 1 inch = 914400 EMU = 2.54 cm. */
function emuToCm(emu) {
  return (emu / 914400 * 2.54);
}

/** Format a number as an ODP length string (cm). */
function cm(val) {
  return val.toFixed(3) + 'cm';
}

// MIME type → file extension
const MIME_EXT = {
  'image/jpeg': 'jpg', 'image/jpg': 'jpg',
  'image/png': 'png', 'image/gif': 'gif',
  'image/webp': 'webp', 'image/svg+xml': 'svg',
};

// ── Slide data structure ────────────────────────────────────────────────────

function createSlideData() {
  return {
    shapes: [],       // { type, xml } — rendered into <draw:page>
    background: null, // hex string or null
    name: '',
  };
}

// ── OdpWriter ───────────────────────────────────────────────────────────────

export class OdpWriter {
  constructor() {
    /** @type {Array<{shapes: Array, background: string|null, name: string}>} */
    this._slides = [];
    /** @type {Record<string, Uint8Array>} media files: path → bytes */
    this._media = {};
    this._mediaCounter = 0;
    /** @type {Array<{name: string, css: string}>} style definitions */
    this._styles = [];
    this._styleCounter = 0;
    /** Slide width and height in cm */
    this._width  = 25.4;   // 10 inches
    this._height = 14.288;  // ~5.63 inches (16:9)
    this._title = 'Presentation';
  }

  // ── Factory ───────────────────────────────────────────────────────────────

  /**
   * Create a new ODP from scratch with one blank slide.
   *
   * @param {object} [opts]
   * @param {number} [opts.width=25.4]    slide width in cm  (default 10in = 25.4cm)
   * @param {number} [opts.height=14.288] slide height in cm (default ~5.63in)
   * @param {string} [opts.title='Presentation']
   * @returns {OdpWriter}
   *
   * @example
   *   const odp = OdpWriter.create();
   *   odp.addTextBox(0, 'Hello!', { x: 5, y: 5, w: 15, h: 3, fontSize: 48 });
   *   await odp.download('new.odp');
   *
   * @example
   *   // 4:3 aspect ratio (25.4cm × 19.05cm)
   *   const odp = OdpWriter.create({ width: 25.4, height: 19.05 });
   */
  static create(opts = {}) {
    const w = new OdpWriter();
    if (opts.width  !== undefined) w._width  = opts.width;
    if (opts.height !== undefined) w._height = opts.height;
    if (opts.title  !== undefined) w._title  = opts.title;
    w._slides.push(createSlideData());
    return w;
  }

  /**
   * Convert a loaded PptxRenderer to an ODP file.
   * Extracts text content, images, and backgrounds from the PPTX slides.
   *
   * @param {import('./index.js').default} renderer  loaded PptxRenderer
   * @returns {OdpWriter}
   */
  static fromRenderer(renderer) {
    const w = new OdpWriter();
    w._width  = emuToCm(renderer.slideSize.cx);
    w._height = emuToCm(renderer.slideSize.cy);

    const dec = new TextDecoder();

    for (let i = 0; i < renderer.slideCount; i++) {
      const slide = createSlideData();
      slide.name = `Slide ${i + 1}`;

      const slidePath = renderer.slidePaths[i];
      const raw = renderer._files[slidePath];
      if (!raw) { w._slides.push(slide); continue; }

      const slideXml = dec.decode(raw);
      const doc = new DOMParser().parseFromString(slideXml, 'application/xml');

      // Extract background colour
      const bg = _extractBgColor(doc);
      if (bg) slide.background = bg;

      // Extract shapes from spTree
      const cSld = _g1(doc, 'cSld');
      const spTree = cSld ? _g1(cSld, 'spTree') : null;
      if (spTree) {
        _convertShapes(w, slide, spTree, renderer, slidePath, i);
      }

      w._slides.push(slide);
    }

    return w;
  }

  // ── Slide operations ──────────────────────────────────────────────────────

  getSlideCount() { return this._slides.length; }

  /**
   * Add a new blank slide.
   * @param {number} [atIdx]  insert position (default: end)
   * @returns {OdpWriter}
   */
  addSlide(atIdx) {
    const slide = createSlideData();
    if (atIdx !== undefined) {
      this._slides.splice(atIdx, 0, slide);
    } else {
      this._slides.push(slide);
    }
    return this;
  }

  /**
   * Remove a slide.
   * @param {number} slideIdx
   * @returns {OdpWriter}
   */
  removeSlide(slideIdx) {
    if (this._slides.length <= 1) throw new Error('Cannot remove the last slide');
    this._slides.splice(slideIdx, 1);
    return this;
  }

  // ── Shape creation ────────────────────────────────────────────────────────

  /**
   * Add a text box to a slide.
   *
   * @param {number} slideIdx
   * @param {string} text       use \n for line breaks
   * @param {object} [style]
   * @param {number} [style.x=2]         cm from left
   * @param {number} [style.y=2]         cm from top
   * @param {number} [style.w=20]        cm width
   * @param {number} [style.h=3]         cm height
   * @param {string} [style.color='000000']   hex, no #
   * @param {number} [style.fontSize=18]      pt
   * @param {boolean}[style.bold=false]
   * @param {string} [style.align='start']    start|center|end
   * @param {string} [style.fontFamily='Calibri']
   * @returns {OdpWriter}
   */
  addTextBox(slideIdx, text, style = {}) {
    const {
      x = 2, y = 2, w = 20, h = 3,
      color = '000000', fontSize = 18,
      bold = false, align = 'start',
      fontFamily = 'Calibri',
    } = style;

    const slide = this._slides[slideIdx];
    if (!slide) throw new RangeError(`Slide ${slideIdx} out of range`);

    // Create a graphic style for the text box frame
    const frameStyleName = this._addStyle(
      'graphic',
      `<style:graphic-properties svg:stroke-color="#000000" draw:stroke="none" draw:fill="none" draw:auto-grow-height="true"/>`,
    );

    // Create a paragraph style
    const paraStyleName = this._addStyle(
      'paragraph',
      `<style:paragraph-properties fo:text-align="${align}"/>`,
    );

    // Create a text style
    const textStyleName = this._addStyle(
      'text',
      `<style:text-properties fo:font-size="${fontSize}pt" fo:color="#${color}"` +
      (bold ? ' fo:font-weight="bold"' : '') +
      ` style:font-name="${escXml(fontFamily)}"/>`,
    );

    const lines = text.split('\n');
    const parasXml = lines.map(line =>
      `<text:p text:style-name="${paraStyleName}">` +
      `<text:span text:style-name="${textStyleName}">${escXml(line)}</text:span>` +
      `</text:p>`
    ).join('');

    slide.shapes.push({
      type: 'frame',
      xml: `<draw:frame draw:style-name="${frameStyleName}" ` +
           `svg:x="${cm(x)}" svg:y="${cm(y)}" svg:width="${cm(w)}" svg:height="${cm(h)}" ` +
           `presentation:class="subtitle">` +
           `<draw:text-box>${parasXml}</draw:text-box>` +
           `</draw:frame>`,
    });

    return this;
  }

  /**
   * Add an image to a slide.
   *
   * @param {number}     slideIdx
   * @param {Uint8Array} imageBytes
   * @param {string}     [mimeType='image/png']
   * @param {object}     [rect]        { x, y, w, h } in cm
   * @returns {OdpWriter}
   */
  addImage(slideIdx, imageBytes, mimeType = 'image/png', rect = {}) {
    const { x = 2, y = 2, w = 10, h = 7.5 } = rect;

    const slide = this._slides[slideIdx];
    if (!slide) throw new RangeError(`Slide ${slideIdx} out of range`);

    const ext = MIME_EXT[mimeType] || 'png';
    const mediaPath = `Pictures/image${++this._mediaCounter}.${ext}`;
    this._media[mediaPath] = imageBytes;

    slide.shapes.push({
      type: 'image',
      xml: `<draw:frame svg:x="${cm(x)}" svg:y="${cm(y)}" svg:width="${cm(w)}" svg:height="${cm(h)}">` +
           `<draw:image xlink:href="${mediaPath}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
           `</draw:frame>`,
    });

    return this;
  }

  /**
   * Set a solid background colour on a slide.
   * @param {number} slideIdx
   * @param {string} hexRgb  6-digit hex, no '#'
   * @returns {OdpWriter}
   */
  setSlideBackground(slideIdx, hexRgb) {
    const slide = this._slides[slideIdx];
    if (!slide) throw new RangeError(`Slide ${slideIdx} out of range`);
    slide.background = hexRgb;
    return this;
  }

  // ── Serialisation ─────────────────────────────────────────────────────────

  /**
   * Serialize the ODP to bytes.
   * @returns {Promise<Uint8Array>}
   */
  async save() {
    const files = {};

    // ── mimetype (must be first, uncompressed) ────────────────────────────
    files['mimetype'] = enc.encode('application/vnd.oasis.opendocument.presentation');

    // ── META-INF/manifest.xml ─────────────────────────────────────────────
    let manifestEntries =
      `<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.presentation"/>\n` +
      `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>\n` +
      `<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>\n` +
      `<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>\n`;

    for (const [path, bytes] of Object.entries(this._media)) {
      const ext = path.split('.').pop();
      const mime = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', webp: 'image/webp', svg: 'image/svg+xml' }[ext] || 'application/octet-stream';
      manifestEntries += `<manifest:file-entry manifest:full-path="${escXml(path)}" manifest:media-type="${mime}"/>\n`;
    }

    files['META-INF/manifest.xml'] = enc.encode(
`<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="${NSMAP.manifest}" manifest:version="1.2">
${manifestEntries}</manifest:manifest>`);

    // ── meta.xml ──────────────────────────────────────────────────────────
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    files['meta.xml'] = enc.encode(
`<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta ${nsDecls()}>
  <office:meta>
    <dc:title>${escXml(this._title)}</dc:title>
    <meta:creation-date>${now}</meta:creation-date>
    <meta:generator>pptx-browser</meta:generator>
  </office:meta>
</office:document-meta>`);

    // ── styles.xml ────────────────────────────────────────────────────────
    files['styles.xml'] = enc.encode(this._buildStylesXml());

    // ── content.xml ───────────────────────────────────────────────────────
    files['content.xml'] = enc.encode(this._buildContentXml());

    // ── Media files ───────────────────────────────────────────────────────
    for (const [path, bytes] of Object.entries(this._media)) {
      files[path] = bytes;
    }

    return writeZip(files);
  }

  /**
   * Download as an ODP file in the browser.
   * @param {string} [filename='presentation.odp']
   */
  async download(filename = 'presentation.odp') {
    const bytes = await this.save();
    const blob = new Blob([bytes], { type: 'application/vnd.oasis.opendocument.presentation' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  // ── Internal helpers ──────────────────────────────────────────────────────

  _addStyle(family, propertiesXml) {
    const name = `s${++this._styleCounter}`;
    this._styles.push({ name, family, propertiesXml });
    return name;
  }

  _buildStylesXml() {
    // Drawing-page styles for backgrounds
    let dpStyles = '';
    for (let i = 0; i < this._slides.length; i++) {
      const slide = this._slides[i];
      const bgProp = slide.background
        ? `<style:drawing-page-properties draw:fill="solid" draw:fill-color="#${slide.background}"/>`
        : `<style:drawing-page-properties draw:fill="solid" draw:fill-color="#ffffff"/>`;
      dpStyles += `<style:style style:name="dp${i}" style:family="drawing-page">${bgProp}</style:style>\n`;
    }

    return `<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles ${nsDecls()}>
  <office:styles>
    <style:style style:name="standard" style:family="graphic">
      <style:graphic-properties draw:stroke="none" draw:fill="none"/>
    </style:style>
  </office:styles>
  <office:automatic-styles>
    <style:page-layout style:name="PM0">
      <style:page-layout-properties fo:page-width="${cm(this._width)}" fo:page-height="${cm(this._height)}" fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm"/>
    </style:page-layout>
${dpStyles}  </office:automatic-styles>
  <office:master-styles>
    <style:master-page style:name="Default" style:page-layout-name="PM0" draw:style-name="dp0"/>
  </office:master-styles>
</office:document-styles>`;
  }

  _buildContentXml() {
    // Automatic styles (shape styles)
    let autoStyles = '';
    for (const s of this._styles) {
      autoStyles += `<style:style style:name="${s.name}" style:family="${s.family}">${s.propertiesXml}</style:style>\n`;
    }

    // Drawing page styles inside content.xml automatic-styles
    for (let i = 0; i < this._slides.length; i++) {
      const slide = this._slides[i];
      const bgProp = slide.background
        ? `<style:drawing-page-properties draw:fill="solid" draw:fill-color="#${slide.background}"/>`
        : `<style:drawing-page-properties draw:fill="solid" draw:fill-color="#ffffff"/>`;
      autoStyles += `<style:style style:name="cdp${i}" style:family="drawing-page">${bgProp}</style:style>\n`;
    }

    // Slide pages
    let pages = '';
    for (let i = 0; i < this._slides.length; i++) {
      const slide = this._slides[i];
      const name = slide.name || `Slide ${i + 1}`;
      pages += `<draw:page draw:name="${escXml(name)}" draw:style-name="cdp${i}" draw:master-page-name="Default" presentation:presentation-page-layout-name="">\n`;
      for (const shape of slide.shapes) {
        pages += '  ' + shape.xml + '\n';
      }
      pages += `</draw:page>\n`;
    }

    return `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content ${nsDecls()}>
  <office:automatic-styles>
${autoStyles}  </office:automatic-styles>
  <office:body>
    <office:presentation>
${pages}    </office:presentation>
  </office:body>
</office:document-content>`;
  }
}

// ── PPTX → ODP conversion helpers ──────────────────────────────────────────

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

/** Extract background solid colour from a slide doc. */
function _extractBgColor(doc) {
  const bg = _g1(doc, 'bg');
  if (!bg) return null;
  const srgb = _g1(bg, 'srgbClr');
  if (srgb) return srgb.getAttribute('val');
  return null;
}

/** Convert PPTX spTree shapes to ODP draw elements. */
function _convertShapes(writer, slide, spTree, renderer, slidePath, slideIdx) {
  const dec = new TextDecoder();

  for (const child of spTree.children) {
    const ln = child.localName;

    if (ln === 'sp') {
      _convertTextShape(writer, slide, child);
    } else if (ln === 'pic') {
      _convertPicture(writer, slide, child, renderer, slidePath);
    } else if (ln === 'grpSp') {
      // Recurse into group shapes
      _convertShapes(writer, slide, child, renderer, slidePath, slideIdx);
    }
  }
}

/** Convert a PPTX shape (sp) to an ODP text frame. */
function _convertTextShape(writer, slide, sp) {
  // Get position and size
  const xfrm = _g1(sp, 'xfrm');
  if (!xfrm) return;

  const off = _g1(xfrm, 'off');
  const ext = _g1(xfrm, 'ext');
  if (!off || !ext) return;

  const x = emuToCm(parseInt(off.getAttribute('x') || '0', 10));
  const y = emuToCm(parseInt(off.getAttribute('y') || '0', 10));
  const w = emuToCm(parseInt(ext.getAttribute('cx') || '0', 10));
  const h = emuToCm(parseInt(ext.getAttribute('cy') || '0', 10));

  if (w === 0 && h === 0) return;

  // Extract text content
  const txBody = _g1(sp, 'txBody');
  if (!txBody) return;

  const paragraphs = _gtn(txBody, 'p');
  if (paragraphs.length === 0) return;

  // Check if there's any actual text
  const allT = _gtn(txBody, 't');
  const fullText = allT.map(t => t.textContent).join('');
  if (!fullText.trim()) return;

  // Extract text formatting from first run
  let fontSize = 18;
  let fontFamily = 'Calibri';
  let color = '000000';
  let bold = false;
  let align = 'start';

  const firstRun = _g1(txBody, 'r');
  if (firstRun) {
    const rPr = _g1(firstRun, 'rPr');
    if (rPr) {
      const sz = rPr.getAttribute('sz');
      if (sz) fontSize = parseInt(sz, 10) / 100;
      const b = rPr.getAttribute('b');
      if (b === '1' || b === 'true') bold = true;

      const srgb = _g1(rPr, 'srgbClr');
      if (srgb) color = srgb.getAttribute('val') || '000000';

      const latin = _g1(rPr, 'latin');
      if (latin) fontFamily = latin.getAttribute('typeface') || 'Calibri';
    }
  }

  // Check paragraph alignment
  const pPr = _g1(paragraphs[0], 'pPr');
  if (pPr) {
    const algn = pPr.getAttribute('algn');
    if (algn === 'ctr') align = 'center';
    else if (algn === 'r') align = 'end';
    else if (algn === 'just') align = 'justify';
  }

  // Build paragraph XML for each paragraph
  const frameStyleName = writer._addStyle(
    'graphic',
    `<style:graphic-properties draw:stroke="none" draw:fill="none" draw:auto-grow-height="true"/>`,
  );
  const paraStyleName = writer._addStyle(
    'paragraph',
    `<style:paragraph-properties fo:text-align="${align}"/>`,
  );
  const textStyleName = writer._addStyle(
    'text',
    `<style:text-properties fo:font-size="${fontSize}pt" fo:color="#${color}"` +
    (bold ? ' fo:font-weight="bold"' : '') +
    ` style:font-name="${escXml(fontFamily)}"/>`,
  );

  let parasXml = '';
  for (const p of paragraphs) {
    const runs = _gtn(p, 'r');
    let paraText = '';
    for (const r of runs) {
      const tEl = _g1(r, 't');
      if (tEl && tEl.textContent) {
        paraText += escXml(tEl.textContent);
      }
    }
    parasXml += `<text:p text:style-name="${paraStyleName}">` +
      `<text:span text:style-name="${textStyleName}">${paraText}</text:span>` +
      `</text:p>`;
  }

  slide.shapes.push({
    type: 'frame',
    xml: `<draw:frame draw:style-name="${frameStyleName}" ` +
         `svg:x="${cm(x)}" svg:y="${cm(y)}" svg:width="${cm(w)}" svg:height="${cm(h)}">` +
         `<draw:text-box>${parasXml}</draw:text-box>` +
         `</draw:frame>`,
  });
}

/** Convert a PPTX picture (pic) to an ODP image frame. */
function _convertPicture(writer, slide, pic, renderer, slidePath) {
  const xfrm = _g1(pic, 'xfrm');
  if (!xfrm) return;

  const off = _g1(xfrm, 'off');
  const ext = _g1(xfrm, 'ext');
  if (!off || !ext) return;

  const x = emuToCm(parseInt(off.getAttribute('x') || '0', 10));
  const y = emuToCm(parseInt(off.getAttribute('y') || '0', 10));
  const w = emuToCm(parseInt(ext.getAttribute('cx') || '0', 10));
  const h = emuToCm(parseInt(ext.getAttribute('cy') || '0', 10));

  // Find image relationship
  const blipFill = _g1(pic, 'blipFill');
  const blip = blipFill ? _g1(blipFill, 'blip') : null;
  if (!blip) return;

  const rEmbed = blip.getAttribute('r:embed') || blip.getAttribute('embed');
  if (!rEmbed) return;

  // Resolve relationship to find the image path
  const slideDir = slidePath.split('/').slice(0, -1).join('/');
  const relsPath = slideDir + '/_rels/' + slidePath.split('/').pop() + '.rels';
  const relsRaw = renderer._files[relsPath];
  if (!relsRaw) return;

  const relsDoc = new DOMParser().parseFromString(new TextDecoder().decode(relsRaw), 'application/xml');
  let imagePath = null;
  for (const rel of relsDoc.getElementsByTagName('Relationship')) {
    if (rel.getAttribute('Id') === rEmbed) {
      let target = rel.getAttribute('Target');
      if (target && !target.startsWith('/') && !target.startsWith('http')) {
        // Resolve relative path
        const parts = (slideDir + '/' + target).split('/');
        const resolved = [];
        for (const p of parts) {
          if (p === '..') resolved.pop();
          else resolved.push(p);
        }
        imagePath = resolved.join('/');
      } else {
        imagePath = target;
      }
      break;
    }
  }

  if (!imagePath || !renderer._files[imagePath]) return;

  // Determine MIME type from extension
  const imgExt = imagePath.split('.').pop().toLowerCase();
  const mimeMap = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', webp: 'image/webp', svg: 'image/svg+xml' };
  const mime = mimeMap[imgExt] || 'image/png';

  // Store the image in media
  const odpExt = MIME_EXT[mime] || 'png';
  const mediaPath = `Pictures/image${++writer._mediaCounter}.${odpExt}`;
  writer._media[mediaPath] = renderer._files[imagePath];

  slide.shapes.push({
    type: 'image',
    xml: `<draw:frame svg:x="${cm(x)}" svg:y="${cm(y)}" svg:width="${cm(w)}" svg:height="${cm(h)}">` +
         `<draw:image xlink:href="${mediaPath}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
         `</draw:frame>`,
  });
}
