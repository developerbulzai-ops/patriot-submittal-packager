import ExcelJS from "exceljs";
import type { SubmittalData } from "@/types/submittal";
import type { PageImage } from "./renderPdfPages";

// ── Column layout ─────────────────────────────────────────────────────────────
// 6 columns matching the reference screenshot (A–F active, G–H default empty):
//   A: 2  — narrow left gutter
//   B: 20 — address block / item descriptions
//   C: 26 — centre-left content
//   D: 20 — centre-right content
//   E: 15 — right content
//   F: 13 — page numbers / right-side labels
// Total = 96 chars ≈ 720 px (Letter – 1" margins)
// Left half  A+B+C = 48 chars  ← To / Item / description side
// Right half D+E+F = 48 chars  ← Subject / Submittals / page-number side
const COL_A = 1, COL_B = 2, COL_C = 3, COL_D = 4, COL_E = 5, COL_F = 6;
const COL_WIDTHS = [2, 20, 26, 20, 15, 13];

// Page-center logo placement:
//   char 48 = centre of 96-char page width
//   logo 110 px = 14.67 chars  →  left edge at char 40.67
//   col C starts at char 48−26=22; offset into C = (40.67−22)/26 = 0.718
//   → tl.col = COL_C − 1 + 0.718 = 2.718
const LOGO_COL = COL_C - 1 + 0.718; // 2.718

// Data-page image constants
const IMG_DISPLAY_W = 720;
const PX_TO_PT      = 0.75;
const IMG_ROW_H     = 90;
const BLANK_ROWS    = 8;

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
    fill?: ExcelJS.Fill;
    border?: Partial<ExcelJS.Borders>;
  }
) {
  const c = ws.getRow(row).getCell(col);
  if (opts.value !== undefined) c.value = opts.value;
  c.font = { name: FONT, size: opts.size ?? 10, bold: opts.bold ?? false };
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

function spacer(ws: ExcelJS.Worksheet, row: number, h = 8) {
  ws.getRow(row).height = h;
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
      paperSize:          1 as number,
      orientation:        "portrait",
      fitToPage:          true,
      fitToWidth:         1,
      fitToHeight:        0,
      horizontalCentered: true,
      verticalCentered:   false,
    },
  });
  ws.pageSetup.margins = { left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0, footer: 0 };
  COL_WIDTHS.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  let row = 1;

  // ══════════════════════════════════════════════════════════════════════════
  // SECTION 1 — Cover sheet
  // ══════════════════════════════════════════════════════════════════════════

  // ── Address block — col B, rows 1–5 ──────────────────────────────────────
  // Col A is a narrow left gutter; all address text starts in col B.
  const ADDRESS = [
    "Patriot Pipeline Inc.",
    "PO Box 1487",
    "Wildomar, Ca 92595",
    "Phone: 951-679-8364",
    "Fax: 951-304-0684",
  ];
  ADDRESS.forEach((line, i) => {
    ws.getRow(row + i).height = 15;
    sc(ws, row + i, COL_B, { value: line, size: 9, vAlign: "middle" });
  });

  // ── "Submittals" — right half (D+E+F), row 1 ─────────────────────────────
  merge(ws, row, COL_D, row, COL_F);
  sc(ws, row, COL_D, {
    value: "Submittals",
    size: 22, bold: true, hAlign: "right", vAlign: "middle",
  });
  ws.getRow(row).height = 22;

  // ── Job No — right half, row 3 ────────────────────────────────────────────
  merge(ws, row + 2, COL_D, row + 2, COL_F);
  sc(ws, row + 2, COL_D, {
    value: data.jobNo ? `Job No: ${data.jobNo}` : "Job No:",
    size: 10, bold: true, hAlign: "right", vAlign: "middle",
  });

  // ── Date — right half, row 4 ──────────────────────────────────────────────
  merge(ws, row + 3, COL_D, row + 3, COL_F);
  sc(ws, row + 3, COL_D, {
    value: data.date ? `Date: ${data.date}` : "Date:",
    size: 10, bold: true, hAlign: "right", vAlign: "middle",
  });

  // ── Logo — page-centered (tl.col = 2.718), rows 1–5 ─────────────────────
  // patriot_logo_sq.png: 225×225 → display at 110×110 px
  if (logoBuffer) {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const imgId = wb.addImage({ buffer: logoBuffer as any, extension: "png" });
      ws.addImage(imgId, {
        tl: { col: LOGO_COL, row: 0.0 },
        ext: { width: 110, height: 110 },
      });
    } catch { /* skip logo on embed failure */ }
  }

  row += ADDRESS.length; // row → 6

  // ── Spacer ────────────────────────────────────────────────────────────────
  spacer(ws, row); row++;

  // ── To / Subject header ───────────────────────────────────────────────────
  merge(ws, row, COL_A, row, COL_C);
  sc(ws, row, COL_A, { value: "To:", bold: true, indent: 1, fill: GRAY, border: ALL_BORDERS });
  merge(ws, row, COL_D, row, COL_F);
  sc(ws, row, COL_D, { value: "Subject:", bold: true, indent: 1, fill: GRAY, border: ALL_BORDERS });
  ws.getRow(row).height = 18;
  row++;

  // ── To / Subject body — one row per line (matching reference layout) ──────
  const toLines = [
    data.recipient.company,
    data.recipient.attention ? `Attn: ${data.recipient.attention}` : "",
    data.recipient.address1,
    data.recipient.city,
  ].filter(Boolean);

  // Split multi-line location strings the AI may produce
  const subjectLines = [
    data.subject.projectName,
    ...data.subject.location.split("\n").filter(Boolean),
  ].filter(Boolean);

  const bodyRows = Math.max(toLines.length, subjectLines.length, 3) + 1; // +1 blank buffer

  for (let i = 0; i < bodyRows; i++) {
    merge(ws, row, COL_A, row, COL_C);
    sc(ws, row, COL_A, {
      ...(toLines[i] ? { value: toLines[i] } : {}),
      indent: 1, vAlign: "middle", border: ALL_BORDERS,
    });

    merge(ws, row, COL_D, row, COL_F);
    sc(ws, row, COL_D, {
      ...(subjectLines[i] ? { value: subjectLines[i] } : {}),
      indent: 1, vAlign: "middle", border: ALL_BORDERS,
    });

    ws.getRow(row).height = 16;
    row++;
  }

  // ── Spacer ────────────────────────────────────────────────────────────────
  spacer(ws, row); row++;

  // ── TOC header ────────────────────────────────────────────────────────────
  merge(ws, row, COL_A, row, COL_E);
  sc(ws, row, COL_A, { value: "Item", bold: true, indent: 1, fill: GRAY, border: ALL_BORDERS });
  sc(ws, row, COL_F, { value: "Page Number", bold: true, hAlign: "right", indent: 1, fill: GRAY, border: ALL_BORDERS });
  ws.getRow(row).height = 18;
  row++;

  // ── Spacer after Item header ──────────────────────────────────────────────
  spacer(ws, row, 6); row++;

  // ── TOC rows ──────────────────────────────────────────────────────────────
  for (const cat of data.categories) {
    // Category header — full width, bold, centered
    merge(ws, row, COL_A, row, COL_F);
    sc(ws, row, COL_A, { value: cat.name, bold: true, hAlign: "center" });
    ws.getRow(row).height = 16;
    row++;

    // Spacer after category header
    spacer(ws, row, 6); row++;

    // Line items
    for (const item of cat.lineItems) {
      const pg = item.startPage === item.endPage
        ? `${item.startPage}`
        : `${item.startPage}-${item.endPage}`;
      merge(ws, row, COL_A, row, COL_E);
      sc(ws, row, COL_A, { value: item.description, indent: 2 });
      sc(ws, row, COL_F, { value: pg, bold: true, hAlign: "right" });
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
      const displayH = Math.round(IMG_DISPLAY_W * img.height / img.width);
      const totalPt  = Math.round(displayH * PX_TO_PT);
      const numRows  = Math.max(1, Math.ceil(totalPt / IMG_ROW_H));
      const rowH     = Math.ceil(totalPt / numRows);

      for (let i = 0; i < numRows; i++) { ws.getRow(row).height = rowH; row++; }

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const imgId = wb.addImage({ buffer: img.buffer as any, extension: "png" });
      ws.addImage(imgId, {
        tl: { col: 0, row: imgStartRow - 1 },
        ext: { width: IMG_DISPLAY_W, height: displayH },
      });
      ws.getRow(row - 1).addPageBreak();
    }
  }

  ws.pageSetup.printArea = `A1:F${row - 1}`;

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}
