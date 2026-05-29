import ExcelJS from "exceljs";
import type { SubmittalData } from "@/types/submittal";
import type { PageImage } from "./renderPdfPages";

// ── Layout ────────────────────────────────────────────────────────────────────
const COL_A = 1, COL_B = 2, COL_C = 3, COL_D = 4;
// [26, 22, 35, 13] = 96 chars → ~720 px (Letter – 1" margins)
// A+B = 48, C+D = 48 — equal halves so To/Subject boxes are balanced
const COL_WIDTHS = [26, 22, 35, 13];

// Data-page image constants
const IMG_DISPLAY_W = 720; // px — fills the 720 px printable width exactly
const PX_TO_PT      = 0.75; // 96-DPI screen → 72 pt/in
const IMG_ROW_H     = 90;  // pt per image-row bucket
const BLANK_ROWS    = 8;   // 8 × 90 pt = 720 pt ≈ 10" blank-page section

// ── Styles ────────────────────────────────────────────────────────────────────
const GRAY: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE0E0E0" } };
const THIN: Partial<ExcelJS.Border> = { style: "thin" };
const ALL_BORDERS: Partial<ExcelJS.Borders> = { top: THIN, bottom: THIN, left: THIN, right: THIN };
const FONT = "Calibri";

// ── Helpers ───────────────────────────────────────────────────────────────────

function sc(
  ws: ExcelJS.Worksheet, row: number, col: number,
  opts: {
    value?: ExcelJS.CellValue;
    bold?: boolean; size?: number;
    hAlign?: ExcelJS.Alignment["horizontal"];
    vAlign?: ExcelJS.Alignment["vertical"];
    wrap?: boolean; indent?: number;
    fill?: ExcelJS.Fill; border?: Partial<ExcelJS.Borders>;
  }
) {
  const c = ws.getRow(row).getCell(col);
  if (opts.value !== undefined) c.value = opts.value;
  c.font   = { name: FONT, size: opts.size ?? 10, bold: opts.bold ?? false };
  c.alignment = {
    horizontal: opts.hAlign ?? "left",
    vertical:   opts.vAlign ?? "middle",
    wrapText:   opts.wrap   ?? false,
    indent:     opts.indent ?? 0,
  };
  if (opts.fill)   c.fill   = opts.fill;
  if (opts.border) c.border = opts.border;
}

function merge(ws: ExcelJS.Worksheet, r1: number, c1: number, r2: number, c2: number) {
  ws.mergeCells(r1, c1, r2, c2);
}

// ── Main export ───────────────────────────────────────────────────────────────

export async function buildExcel(
  data: SubmittalData,
  logoBuffer?: Buffer,
  pageImages?: PageImage[],
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Patriot Pipeline Inc.";
  wb.created = new Date();

  const ws = wb.addWorksheet("Submittal", {
    pageSetup: {
      paperSize:         1 as number,
      orientation:       "portrait",
      fitToPage:         true,
      fitToWidth:        1,
      fitToHeight:       0,
      horizontalCentered: true,
      verticalCentered:  false,
    },
  });
  ws.pageSetup.margins = { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0, footer: 0 };
  COL_WIDTHS.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  let row = 1;

  // ══════════════════════════════════════════════════════════════════════════
  // SECTION 1 — Cover sheet
  // ══════════════════════════════════════════════════════════════════════════

  // Address block — col A, rows 1–5
  const ADDRESS = [
    "Patriot Pipeline Inc.",
    "PO Box 1487",
    "Wildomar, Ca 92595",
    "Phone: 951-679-8364",
    "Fax: 951-304-0684",
  ];
  ADDRESS.forEach((line, i) => {
    ws.getRow(row + i).height = 15;
    sc(ws, row + i, COL_A, { value: line, size: 9, vAlign: "middle" });
  });

  // "Submittals" — row 1, cols C+D
  merge(ws, row, COL_C, row, COL_D);
  sc(ws, row, COL_C, { value: "Submittals", size: 22, bold: true, hAlign: "right", vAlign: "middle" });
  ws.getRow(row).height = 22;

  // Job No — row 3, cols C+D
  merge(ws, row + 2, COL_C, row + 2, COL_D);
  sc(ws, row + 2, COL_C, {
    value: data.jobNo ? `Job No: ${data.jobNo}` : "Job No:",
    size: 10, bold: true, hAlign: "right", vAlign: "middle",
  });

  // Date — row 4, cols C+D
  merge(ws, row + 3, COL_C, row + 3, COL_D);
  sc(ws, row + 3, COL_C, {
    value: data.date ? `Date: ${data.date}` : "Date:",
    size: 10, bold: true, hAlign: "right", vAlign: "middle",
  });

  // Logo — perfectly page-centered, rows 1–5
  // patriot_logo_sq.png: 225×225 px (1:1). Display at 110×110 px.
  // Page total = 96 chars. Center = char 48 from left.
  // Logo (110 px = 14.67 chars) left edge at char 48 − 7.33 = 40.67.
  // Col B starts at char 26 (width 22) → offset 14.67/22 = 0.667 → col: 1.667
  if (logoBuffer) {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const imgId = wb.addImage({ buffer: logoBuffer as any, extension: "png" });
      ws.addImage(imgId, {
        tl: { col: COL_B - 1 + 0.667, row: 0.0 },
        ext: { width: 110, height: 110 },
      });
    } catch { /* skip logo on embed failure */ }
  }

  row += ADDRESS.length; // → 6

  // Spacer
  row++; ws.getRow(row).height = 6; row++;

  // To / Subject header (row 8)
  merge(ws, row, COL_A, row, COL_B);
  sc(ws, row, COL_A, { value: "To:", bold: true, indent: 1, fill: GRAY });
  merge(ws, row, COL_C, row, COL_D);
  sc(ws, row, COL_C, { value: "Subject:", bold: true, indent: 1, fill: GRAY });
  ws.getRow(row).height = 18;
  row++;

  // To / Subject body (row 9)
  const toLines = [
    data.recipient.company,
    data.recipient.attention ? `Attn: ${data.recipient.attention}` : "",
    data.recipient.address1,
    data.recipient.city,
  ].filter(Boolean);
  const subjectLines = [data.subject.projectName, data.subject.location].filter(Boolean);
  const bodyH = Math.max(toLines.length, subjectLines.length, 4) * 15 + 8;

  merge(ws, row, COL_A, row, COL_B);
  sc(ws, row, COL_A, { value: toLines.join("\n"), wrap: true, vAlign: "top", indent: 1 });
  ws.getRow(row).height = bodyH;
  merge(ws, row, COL_C, row, COL_D);
  sc(ws, row, COL_C, { value: subjectLines.join("\n"), wrap: true, vAlign: "top", indent: 1 });
  row++;

  // Spacer
  row++; ws.getRow(row).height = 6; row++;

  // TOC header (row 12)
  merge(ws, row, COL_A, row, COL_C);
  sc(ws, row, COL_A, { value: "Item", bold: true, indent: 1, fill: GRAY });
  sc(ws, row, COL_D, { value: "Page Number", bold: true, hAlign: "right", indent: 1, fill: GRAY });
  ws.getRow(row).height = 18;
  row++;

  // TOC rows (row 13+)
  for (const cat of data.categories) {
    merge(ws, row, COL_A, row, COL_D);
    sc(ws, row, COL_A, { value: cat.name, bold: true, hAlign: "center" });
    ws.getRow(row).height = 16;
    row++;
    for (const item of cat.lineItems) {
      const pg = item.startPage === item.endPage ? `${item.startPage}` : `${item.startPage}-${item.endPage}`;
      merge(ws, row, COL_A, row, COL_C);
      sc(ws, row, COL_A, { value: item.description, indent: 2 });
      sc(ws, row, COL_D, { value: pg, hAlign: "right", indent: 1 });
      ws.getRow(row).height = 15;
      row++;
    }
  }

  // Page break after cover sheet
  ws.getRow(row - 1).addPageBreak();

  // ══════════════════════════════════════════════════════════════════════════
  // SECTION 2 — Blank page
  // ══════════════════════════════════════════════════════════════════════════
  for (let i = 0; i < BLANK_ROWS; i++) {
    ws.getRow(row).height = IMG_ROW_H;
    row++;
  }
  ws.getRow(row - 1).addPageBreak();

  // ══════════════════════════════════════════════════════════════════════════
  // SECTION 3 — Supplier data pages (images flowing continuously)
  // ══════════════════════════════════════════════════════════════════════════
  if (pageImages && pageImages.length > 0) {
    for (const img of pageImages) {
      const imgStartRow = row;

      // Scale to printable width, maintain aspect ratio
      const displayH = Math.round(IMG_DISPLAY_W * img.height / img.width);

      // Distribute image height across fixed-height rows
      const totalPt = Math.round(displayH * PX_TO_PT);
      const numRows = Math.max(1, Math.ceil(totalPt / IMG_ROW_H));
      const rowH    = Math.ceil(totalPt / numRows);

      for (let i = 0; i < numRows; i++) {
        ws.getRow(row).height = rowH;
        row++;
      }

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const imgId = wb.addImage({ buffer: img.buffer as any, extension: "png" });
      ws.addImage(imgId, {
        tl: { col: 0, row: imgStartRow - 1 },
        ext: { width: IMG_DISPLAY_W, height: displayH },
      });

      ws.getRow(row - 1).addPageBreak();
    }
  }

  ws.pageSetup.printArea = `A1:D${row - 1}`;

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}
