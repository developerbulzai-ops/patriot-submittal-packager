export interface PageImage {
  buffer: Buffer;
  width: number;
  height: number;
}

export async function renderPdfPages(pdfBuffer: Buffer): Promise<PageImage[]> {
  // mupdf is WASM-based — no native binaries, works on all Node.js environments.
  // webpackIgnore prevents webpack from bundling it (WASM wrappers break when bundled).
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const mupdf: any = await import(/* webpackIgnore: true */ "mupdf");

  const doc = mupdf.Document.openDocument(
    new Uint8Array(pdfBuffer),
    "application/pdf"
  );

  const numPages: number = doc.countPages();
  const pages: PageImage[] = [];

  // i=0 is supplier cover page — skip it, start at i=1
  for (let i = 1; i < numPages; i++) {
    const page = doc.loadPage(i);
    const scale = 1.5; // ~108 DPI (72 DPI base × 1.5)
    const pixmap = page.toPixmap(
      [scale, 0, 0, scale, 0, 0],
      mupdf.ColorSpace.DeviceRGB,
      false
    );

    pages.push({
      buffer: Buffer.from(pixmap.asPNG() as Uint8Array),
      width: pixmap.getWidth() as number,
      height: pixmap.getHeight() as number,
    });

    pixmap.destroy();
    page.destroy();
  }

  return pages;
}
