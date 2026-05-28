import {
  PDFDocument,
  StandardFonts,
  rgb,
  PDFPage,
  PDFFont,
  PDFImage,
} from "pdf-lib";
import type { SubmittalData } from "@/types/submittal";

// Letter size in points (72pt per inch)
const PAGE_W = 612;
const PAGE_H = 792;
const MARGIN = 50;
const CONTENT_W = PAGE_W - MARGIN * 2; // 512

const BLACK = rgb(0, 0, 0);
const GRAY_FILL = rgb(0.9, 0.9, 0.9);

// ─── helpers ──────────────────────────────────────────────────────────────────

function drawHLine(page: PDFPage, x: number, y: number, width: number) {
  page.drawLine({
    start: { x, y },
    end: { x: x + width, y },
    thickness: 0.5,
    color: BLACK,
  });
}

function drawVLine(page: PDFPage, x: number, y: number, height: number) {
  page.drawLine({
    start: { x, y },
    end: { x, y: y - height },
    thickness: 0.5,
    color: BLACK,
  });
}

/** Draw a single-row cell rectangle with optional fill. */
function drawCell(
  page: PDFPage,
  x: number,
  y: number,
  w: number,
  h: number,
  fill?: ReturnType<typeof rgb>
) {
  page.drawRectangle({
    x,
    y: y - h,
    width: w,
    height: h,
    color: fill,
    borderColor: BLACK,
    borderWidth: 0.5,
  });
}

/** Draw left-aligned text inside a cell with standard padding. */
function cellText(
  page: PDFPage,
  text: string,
  x: number,
  y: number,
  h: number,
  font: PDFFont,
  size: number,
  rightAlign = false,
  colWidth = 0
) {
  if (!text) return;
  const tw = font.widthOfTextAtSize(text, size);
  const tx = rightAlign ? x + colWidth - tw - 4 : x + 4;
  page.drawText(text, {
    x: tx,
    y: y - h + 5,
    size,
    font,
    color: BLACK,
  });
}

// ─── main export ──────────────────────────────────────────────────────────────

export async function buildCoverPage(data: SubmittalData, logoBytes?: Uint8Array): Promise<Uint8Array> {
  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([PAGE_W, PAGE_H]);

  const helv = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const helvBold = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

  // ── Logo ────────────────────────────────────────────────────────────────────
  let logoImage: PDFImage | null = null;
  let logoW = 0;
  let logoH = 0;
  if (logoBytes) {
    try {
      logoImage = await pdfDoc.embedPng(logoBytes);
      const dims = logoImage.scale(1);
      const scale = 200 / dims.width;
      logoW = 200;
      logoH = dims.height * scale;
    } catch {
      // Falls back to text placeholder below
    }
  }

  // ── Header section (y: 742 → 650) ──────────────────────────────────────────
  const headerTopY = PAGE_H - MARGIN; // 742

  // Left: company address block
  const addressLines = [
    "Patriot Pipeline Inc.",
    "PO Box 1487",
    "Wildomar, Ca 92595",
    "Phone: 951-679-8364",
    "Fax: 951-304-0684",
  ];
  let addrY = headerTopY - 8;
  for (const line of addressLines) {
    page.drawText(line, { x: MARGIN, y: addrY, size: 10, font: helv, color: BLACK });
    addrY -= 14;
  }

  // Center: logo image or text fallback
  if (logoImage) {
    const logoX = (PAGE_W - logoW) / 2;
    const logoY = headerTopY - logoH - 8;
    page.drawImage(logoImage, { x: logoX, y: logoY, width: logoW, height: logoH });
  } else {
    // Text placeholder — two lines styled to suggest the logo
    const line1 = "PATRIOT";
    const line2 = "PIPELINE";
    const s1 = 18;
    const s2 = 14;
    page.drawText(line1, {
      x: (PAGE_W - helvBold.widthOfTextAtSize(line1, s1)) / 2,
      y: headerTopY - 30,
      size: s1,
      font: helvBold,
      color: rgb(0.13, 0.27, 0.53),
    });
    page.drawText(line2, {
      x: (PAGE_W - helvBold.widthOfTextAtSize(line2, s2)) / 2,
      y: headerTopY - 50,
      size: s2,
      font: helvBold,
      color: rgb(0.13, 0.27, 0.53),
    });
  }

  // Right: "Submittals" heading
  const submittalsStr = "Submittals";
  const submittalsSize = 22;
  const submittalsW = helvBold.widthOfTextAtSize(submittalsStr, submittalsSize);
  page.drawText(submittalsStr, {
    x: PAGE_W - MARGIN - submittalsW,
    y: headerTopY - 10,
    size: submittalsSize,
    font: helvBold,
    color: BLACK,
  });

  // Right: Job No and Date (right-aligned, bold, below heading)
  const rightX = PAGE_W - MARGIN;
  const jobStr = data.jobNo ? `Job No: ${data.jobNo}` : "Job No: ___";
  const dateStr = data.date ? `Date: ${data.date}` : "Date: ___";
  const metaSize = 11;

  const jobW = helvBold.widthOfTextAtSize(jobStr, metaSize);
  page.drawText(jobStr, {
    x: rightX - jobW,
    y: headerTopY - 55,
    size: metaSize,
    font: helvBold,
    color: BLACK,
  });

  const dateW = helvBold.widthOfTextAtSize(dateStr, metaSize);
  page.drawText(dateStr, {
    x: rightX - dateW,
    y: headerTopY - 71,
    size: metaSize,
    font: helvBold,
    color: BLACK,
  });

  // ── To / Subject table ──────────────────────────────────────────────────────
  // Start ~120pt below page top (y = 622), give it enough rows for 5 data lines
  const toLines = [
    data.recipient.company,
    data.recipient.attention ? `Attn: ${data.recipient.attention}` : "",
    data.recipient.address1,
    data.recipient.city,
  ].filter(Boolean);

  const subjectLines = [
    data.subject.projectName,
    data.subject.location,
  ].filter(Boolean);

  const bodyRows = Math.max(toLines.length, subjectLines.length, 4);
  const hdrH = 18;
  const rowH = 16;
  const col1W = Math.floor(CONTENT_W / 2); // 256
  const col2W = CONTENT_W - col1W;         // 256

  const toSubjectTop = headerTopY - 118; // ~624

  // Header row
  drawCell(page, MARGIN, toSubjectTop, col1W, hdrH, GRAY_FILL);
  drawCell(page, MARGIN + col1W, toSubjectTop, col2W, hdrH, GRAY_FILL);
  cellText(page, "To:", MARGIN, toSubjectTop, hdrH, helvBold, 10);
  cellText(page, "Subject:", MARGIN + col1W, toSubjectTop, hdrH, helvBold, 10);

  // Body rows (one merged tall rectangle per column, individual row lines)
  const bodyTopY = toSubjectTop - hdrH;
  drawCell(page, MARGIN, bodyTopY, col1W, bodyRows * rowH);
  drawCell(page, MARGIN + col1W, bodyTopY, col2W, bodyRows * rowH);

  // Draw text lines in each column
  const textSize = 10;
  for (let i = 0; i < toLines.length; i++) {
    page.drawText(toLines[i], {
      x: MARGIN + 4,
      y: bodyTopY - (i + 1) * rowH + 5,
      size: textSize,
      font: helv,
      color: BLACK,
    });
  }
  for (let i = 0; i < subjectLines.length; i++) {
    page.drawText(subjectLines[i], {
      x: MARGIN + col1W + 4,
      y: bodyTopY - (i + 1) * rowH + 5,
      size: textSize,
      font: helv,
      color: BLACK,
    });
  }

  // ── TOC table ───────────────────────────────────────────────────────────────
  const tocTop = bodyTopY - bodyRows * rowH - 18; // gap below To/Subject table

  const pageNumColW = 72;
  const itemColW = CONTENT_W - pageNumColW;

  // TOC header row
  drawCell(page, MARGIN, tocTop, itemColW, hdrH, GRAY_FILL);
  drawCell(page, MARGIN + itemColW, tocTop, pageNumColW, hdrH, GRAY_FILL);
  cellText(page, "Item", MARGIN, tocTop, hdrH, helvBold, 10);
  cellText(page, "Page Number", MARGIN + itemColW, tocTop, hdrH, helvBold, 10, true, pageNumColW);

  let tocY = tocTop - hdrH;

  // Category sections
  for (const category of data.categories) {
    // Category header row (full width, bold, centered)
    drawCell(page, MARGIN, tocY, CONTENT_W, hdrH);
    const catW = helvBold.widthOfTextAtSize(category.name, 10);
    page.drawText(category.name, {
      x: MARGIN + (CONTENT_W - catW) / 2,
      y: tocY - hdrH + 5,
      size: 10,
      font: helvBold,
      color: BLACK,
    });
    tocY -= hdrH;

  // Line item rows (inside category loop — closed below)
  for (const item of category.lineItems) {
    const pageStr =
      item.startPage === item.endPage
        ? `${item.startPage}`
        : `${item.startPage}-${item.endPage}`;

    drawCell(page, MARGIN, tocY, itemColW, rowH);
    drawCell(page, MARGIN + itemColW, tocY, pageNumColW, rowH);
    cellText(page, item.description, MARGIN, tocY, rowH, helv, 10);
    cellText(page, pageStr, MARGIN + itemColW, tocY, rowH, helv, 10, true, pageNumColW);

    tocY -= rowH;
  }
  } // end category loop

  // Bottom border line for last row if table didn't draw one automatically
  drawHLine(page, MARGIN, tocY, CONTENT_W);

  return pdfDoc.save();
}
