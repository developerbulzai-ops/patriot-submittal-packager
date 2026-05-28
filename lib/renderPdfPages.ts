import path from "path";
import { pathToFileURL } from "url";
import type { RenderParameters } from "pdfjs-dist/types/src/display/api";

export interface PageImage {
  buffer: Buffer;
  width: number;
  height: number;
}

type CanvasLike = {
  getContext: (t: string) => unknown;
  width: number;
  height: number;
  toBuffer: (t: string) => Buffer;
};

class NodeCanvasFactory {
  private _create: (w: number, h: number) => CanvasLike;

  constructor(createCanvas: (w: number, h: number) => CanvasLike) {
    this._create = createCanvas;
  }

  create(width: number, height: number) {
    const canvas = this._create(width, height);
    return { canvas, context: canvas.getContext("2d") };
  }

  reset(pair: { canvas: CanvasLike }, w: number, h: number) {
    pair.canvas.width = w;
    pair.canvas.height = h;
  }

  destroy(pair: { canvas: CanvasLike }) {
    pair.canvas.width = 0;
    pair.canvas.height = 0;
  }
}

export async function renderPdfPages(pdfBuffer: Buffer): Promise<PageImage[]> {
  // webpackIgnore: pdfjs-dist is ESM-only and must NOT be bundled by webpack —
  // it has to load natively via Node.js import() at runtime. Bundling it causes
  // module-initialization failures on Lambda that crash the entire route.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const pdfjsLib: any = await import(/* webpackIgnore: true */ "pdfjs-dist/legacy/build/pdf.mjs");

  // canvas is in webpack externals — webpack emits require("canvas") which Node.js
  // resolves to the native binary at runtime (errors here are caught by the caller)
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  const { createCanvas } = require("canvas") as { createCanvas: (w: number, h: number) => CanvasLike };

  // Worker must reference the actual file on disk via file:// URL
  pdfjsLib.GlobalWorkerOptions.workerSrc = pathToFileURL(
    path.join(process.cwd(), "node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs")
  ).href;

  const canvasFactory = new NodeCanvasFactory(createCanvas);

  const pdfDoc = await pdfjsLib
    .getDocument({
      data: new Uint8Array(pdfBuffer),
      cMapUrl: path.join(process.cwd(), "node_modules/pdfjs-dist/cmaps/"),
      cMapPacked: true,
      standardFontDataUrl: path.join(
        process.cwd(),
        "node_modules/pdfjs-dist/standard_fonts/"
      ),
      canvasFactory,
    } as unknown)
    .promise;

  const pages: PageImage[] = [];

  // Skip page 1 (supplier cover), render pages 2-N
  for (let pageNum = 2; pageNum <= pdfDoc.numPages; pageNum++) {
    const page = await pdfDoc.getPage(pageNum);
    const viewport = page.getViewport({ scale: 1.5 }); // ~144 DPI
    const w = Math.ceil(viewport.width);
    const h = Math.ceil(viewport.height);

    const { canvas, context } = canvasFactory.create(w, h);

    await page.render({
      canvasContext: context,
      viewport,
    } as unknown as RenderParameters).promise;

    pages.push({ buffer: canvas.toBuffer("image/png"), width: w, height: h });
    canvasFactory.destroy({ canvas });
  }

  return pages;
}
