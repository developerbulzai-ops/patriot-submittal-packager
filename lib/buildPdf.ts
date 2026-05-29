import { PDFDocument } from "pdf-lib";
import { buildCoverPage } from "./buildCoverPage";
import type { SubmittalData } from "@/types/submittal";

export async function buildPdf(
  data: SubmittalData,
  logoBuffer: Buffer | undefined,
  supplierPdfBuffer: Buffer,
): Promise<Buffer> {
  const logoBytes = logoBuffer ? new Uint8Array(logoBuffer) : undefined;
  const titleBytes = await buildCoverPage(data, logoBytes);

  const finalDoc = await PDFDocument.create();

  const titleDoc = await PDFDocument.load(titleBytes);
  const [titlePage] = await finalDoc.copyPages(titleDoc, [0]);
  finalDoc.addPage(titlePage);

  // Blank page 2
  finalDoc.addPage([612, 792]);

  // Supplier data pages — skip their cover (index 0)
  const supplierDoc = await PDFDocument.load(supplierPdfBuffer);
  const count = supplierDoc.getPageCount();
  if (count > 1) {
    const indices = Array.from({ length: count - 1 }, (_, i) => i + 1);
    const pages = await finalDoc.copyPages(supplierDoc, indices);
    for (const p of pages) finalDoc.addPage(p);
  }

  return Buffer.from(await finalDoc.save());
}
