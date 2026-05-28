import path from "path";
import { pathToFileURL } from "url";
import * as pdfjsLib from "pdfjs-dist";
import type { RenderParameters } from "pdfjs-dist/types/src/display/api";

// In Node.js, pdfjs-dist needs a file:// URL for its worker thread
pdfjsLib.GlobalWorkerOptions.workerSrc = pathToFileURL(
  path.join(process.cwd(), "node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs")
).href;

export interface PageImage {
  buffer: Buffer;
  width: number;
  height: number;
}

class NodeCanvasFactory {
  private _createCanvas: (w: number, h: number) => unknown;

  constructor(createCanvas: (w: number, h: number) => unknown) {
    this._createCanvas = createCanvas;
  }

  create(width: number, height: number) {
    const canvas = this._createCanvas(width, height) as {
      getContext: (t: string) => unknown;
      width: number;
      height: number;
      toBuffer: (t: string) => Buffer;
    };
    const context = canvas.getContext("2d");
    return { canvas, context };
  }

  reset(pair: ReturnType<NodeCanvasFactory["create"]>, w: number, h: number) {
    pair.canvas.width = w;
    pair.canvas.height = h;
  }

  destroy(pair: ReturnType<NodeCanvasFactory["create"]>) {
    pair.canvas.width = 0;
    pair.canvas.height = 0;
  }
}

export async function renderPdfPages(pdfBuffer: Buffer): Promise<PageImage[]> {
  // canvas is a native module externalized from webpack — require at runtime
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  const { createCanvas } = require("canvas") as {
    createCanvas: (w: number, h: number) => {
      getContext: (t: string) => unknown;
      width: number;
      height: number;
      toBuffer: (t: string) => Buffer;
    };
  };

  const canvasFactory = new NodeCanvasFactory(createCanvas);

  const pdfDoc = await pdfjsLib
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    .getDocument({
      data: new Uint8Array(pdfBuffer),
      cMapUrl: path.join(process.cwd(), "node_modules/pdfjs-dist/cmaps/"),
      cMapPacked: true,
      standardFontDataUrl: path.join(
        process.cwd(),
        "node_modules/pdfjs-dist/standard_fonts/"
      ),
      canvasFactory: canvasFactory,
    } as any)
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
    canvasFactory.destroy({ canvas, context });
  }

  return pages;
}
