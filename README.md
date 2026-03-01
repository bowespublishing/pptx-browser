# pptx-browser

Render PowerPoint (`.pptx`) slides directly onto an HTML `<canvas>` element — **zero dependencies**, no server required.

Uses only native browser/Node APIs: `DecompressionStream` for ZIP decompression, `DOMParser` for XML, and the Canvas 2D API for rendering.

## Browser support

Chrome 80+ · Firefox 113+ · Safari 16.4+ · Node 18+ (with [`node-canvas`](https://github.com/Automattic/node-canvas))

---

## Installation

```bash
npm install pptx-browser
```

No `JSZip`, no `pptxjs`, no other dependencies.

---

## Quick start

```js
import { PptxRenderer } from 'pptx-browser';

const renderer = new PptxRenderer();

// Load from a <input type="file"> element
const [file] = fileInput.files;
await renderer.load(file, (progress, msg) => console.log(progress, msg));

console.log(renderer.slideCount); // number of slides

// Render slide 0 onto a canvas
const canvas = document.getElementById('my-canvas');
await renderer.renderSlide(0, canvas, 1280); // 1280px wide

// Clean up blob URLs when done
renderer.destroy();
```

---

## API

### `new PptxRenderer()`

Creates a new renderer instance. Each instance is independent and can load one PPTX at a time.

---

### `renderer.load(source, onProgress?)`

Load a PPTX file.

| Parameter    | Type                                         | Description                         |
|--------------|----------------------------------------------|-------------------------------------|
| `source`     | `File \| Blob \| ArrayBuffer \| Uint8Array`  | The PPTX data                       |
| `onProgress` | `(progress: number, message: string) => void`| Optional. Called with 0–1 progress  |

Returns: `Promise<void>`

---

### `renderer.renderSlide(index, canvas, width?)`

Render a single slide.

| Parameter | Type                 | Default | Description              |
|-----------|----------------------|---------|--------------------------|
| `index`   | `number`             | —       | 0-based slide index      |
| `canvas`  | `HTMLCanvasElement`  | —       | Target canvas element    |
| `width`   | `number`             | `1280`  | Output width in pixels   |

Canvas height is set automatically to maintain the correct aspect ratio.

Returns: `Promise<void>`

---

### `renderer.renderAllSlides(width?)`

Render all slides and return an array of canvas elements. Useful for thumbnail strips.

```js
const canvases = await renderer.renderAllSlides(320);
canvases.forEach((c, i) => document.body.appendChild(c));
```

Returns: `Promise<HTMLCanvasElement[]>`

---

### `renderer.slideCount`

`number` — total number of slides. Available after `load()` resolves.

---

### `renderer.destroy()`

Releases all `blob:` URLs created during rendering. Call this when you're done with the renderer or loading a new file.

---

## React example

```jsx
import { useRef, useState } from 'react';
import { PptxRenderer } from 'pptx-browser';

export function PptxViewer() {
  const canvasRef = useRef(null);
  const [renderer] = useState(() => new PptxRenderer());
  const [slideCount, setSlideCount] = useState(0);
  const [current, setCurrent] = useState(0);

  async function onFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    renderer.destroy();
    await renderer.load(file);
    setSlideCount(renderer.slideCount);
    await renderer.renderSlide(0, canvasRef.current, 1280);
    setCurrent(0);
  }

  async function goTo(i) {
    await renderer.renderSlide(i, canvasRef.current, 1280);
    setCurrent(i);
  }

  return (
    <div>
      <input type="file" accept=".pptx" onChange={onFile} />
      <p>Slide {current + 1} / {slideCount}</p>
      <canvas ref={canvasRef} style={{ maxWidth: '100%' }} />
      <div>
        {current > 0 && <button onClick={() => goTo(current - 1)}>←</button>}
        {current < slideCount - 1 && <button onClick={() => goTo(current + 1)}>→</button>}
      </div>
    </div>
  );
}
```

---

## What's rendered

| Feature                    | Status  |
|----------------------------|---------|
| Theme colours (all 12)     | ✅      |
| Colour transforms (lumMod, lumOff, tint, shade, sat, hue, alpha…) | ✅ |
| Gradient fills (linear + radial) | ✅ |
| Image fills + cropping     | ✅      |
| Pattern fills              | ✅ (simplified) |
| 50+ preset shapes          | ✅      |
| Custom geometry (custGeom) | ✅      |
| Drop shadows + glow        | ✅      |
| Text rendering (wrapping, alignment, superscript/subscript) | ✅ |
| Font mapping MS→Google Fonts | ✅ (80+ fonts) |
| normAutoFit text scaling   | ✅      |
| Tables                     | ✅      |
| Images in shapes/backgrounds | ✅    |
| Slide master & layout inheritance | ✅ |
| Placeholder position inheritance | ✅ |
| Rotation & flip            | ✅      |
| Group shapes               | ✅      |
| Charts/SmartArt            | Placeholder box |

---

## How ZIP decompression works (no JSZip)

PPTX files are ZIP archives. Instead of bundling JSZip (~100KB), this library uses the browser-native `DecompressionStream('deflate-raw')` API, available since:

- Chrome 80 (Apr 2020)
- Firefox 113 (May 2023)
- Safari 16.4 (Mar 2023)
- Node.js 18 (Apr 2022)

The custom ZIP parser is ~80 lines of code vs JSZip's 3,000+.

---

## Font handling

Microsoft Office fonts (Calibri, Cambria, Franklin Gothic, etc.) aren't available in browsers. This library automatically maps them to the nearest metric-compatible Google Font and loads it via the Google Fonts API before rendering. Key mappings:

| Office font        | Web substitute          | Notes                              |
|--------------------|-------------------------|------------------------------------|
| Calibri            | **Carlito**             | Metric-compatible (identical widths) |
| Calibri Light      | Carlito Light           | Metric-compatible                  |
| Cambria            | **Caladea**             | Metric-compatible                  |
| Aptos (new default)| Inter                   |                                    |
| Franklin Gothic    | Libre Franklin          |                                    |
| Gill Sans          | Quattrocento Sans       |                                    |
| Segoe UI           | Inter                   |                                    |

Fonts are loaded once and cached for the session.

---

## License

MIT
