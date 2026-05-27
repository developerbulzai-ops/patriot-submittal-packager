import { PDFDocument } from "pdf-lib";

export async function mergePdfs(
  patriotCoverBytes: Uint8Array,
  supplierPdfBytes: Buffer
): Promise<Uint8Array> {
  const mergedDoc = await PDFDocument.create();

  // Insert Patriot cover as page 1
  const coverDoc = await PDFDocument.load(patriotCoverBytes);
  const [coverPage] = await mergedDoc.copyPages(coverDoc, [0]);
  mergedDoc.addPage(coverPage);

  // Append supplier data pages, skipping their cover (page 1 = index 0)
  const supplierDoc = await PDFDocument.load(supplierPdfBytes);
  const count = supplierDoc.getPageCount();
  const indices = Array.from({ length: count - 1 }, (_, i) => i + 1); // skip index 0
  const supplierPages = await mergedDoc.copyPages(supplierDoc, indices);
  for (const p of supplierPages) {
    mergedDoc.addPage(p);
  }

  return mergedDoc.save();
}
