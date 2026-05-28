import ExcelJS from "exceljs";
import type { SubmittalData } from "@/types/submittal";
import type { PageImage } from "./renderPdfPages";

// ── Layout constants ──────────────────────────────────────────────────────────
// 4-column layout: A=left address  B=logo area  C=right header/content  D=page nums
// Widths scaled so total ≈ 720px (Letter printable width: 8.5" − 1" margins = 7.5" × 96 dpi)
const COL_A = 1, COL_B = 2, COL_C = 3, COL_D = 4;
const COL_WIDTHS = [26, 26, 31, 13]; // 96 chars × 7.5 px/char ≈ 720 px

// Display width for supplier page images — matches printable Letter width so
// fitToWidth=1 maps one image to exactly one printed page (7.5" × 96 dpi = 720 px).
const IMG_DISPLAY_W = 720;
// px-to-points conversion (96 dpi screen → 72 pt/in: 72/96 = 0.75)
const PX_TO_PT = 0.75;
// Row height used for image rows and blank page rows (points)
const IMG_ROW_H = 90;

const GRAY = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFE0E0E0" } };
const BORDER_THIN: Partial<ExcelJS.Border> = { style: "thin" };
const ALL_BORDERS = { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN };
const FONT = "Calibri";

// ── Helpers ───────────────────────────────────────────────────────────────────

function cell(ws: ExcelJS.Worksheet, row: number, col: number) {
  return ws.getRow(row).getCell(col);
}

function merge(ws: ExcelJS.Worksheet, r1: number, c1: number, r2: number, c2: number) {
  ws.mergeCells(r1, c1, r2, c2);
}

function styleCell(
  c: ExcelJS.Cell,
  opts: {
    value?: ExcelJS.CellValue;
    bold?: boolean;
    size?: number;
    italic?: boolean;
    hAlign?: ExcelJS.Alignment["horizontal"];
    vAlign?: ExcelJS.Alignment["vertical"];
    wrap?: boolean;
    indent?: number;
    fill?: ExcelJS.Fill;
    border?: Partial<ExcelJS.Borders>;
  }
) {
  if (opts.value !== undefined) c.value = opts.value;
  c.font = { name: FONT, size: opts.size ?? 10, bold: opts.bold ?? false, italic: opts.italic ?? false };
  c.alignment = {
    horizontal: opts.hAlign ?? "left",
    vertical: opts.vAlign ?? "middle",
    wrapText: opts.wrap ?? false,
    indent: opts.indent ?? 0,
  };
  if (opts.fill) c.fill = opts.fill;
  if (opts.border) c.border = opts.border;
}

// ── Main export ───────────────────────────────────────────────────────────────

export async function buildExcel(
  data: SubmittalData,
  logoBuffer?: Buffer,
  pageImages?: PageImage[]
): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Patriot Pipeline Inc.";
  wb.created = new Date();

  // Single continuous sheet — one long document, page breaks define print pages
  const ws = wb.addWorksheet("Submittal", {
    pageSetup: {
      paperSize: 1 as number, // US Letter
      orientation: "portrait",
      fitToPage: true,
      fitToWidth: 1,  // scale so all columns fit one page wide
      fitToHeight: 0, // allow as many pages tall as needed
      horizontalCentered: true,
      verticalCentered: false,
    },
  });

  ws.pageSetup.margins = {
    left: 0.5, right: 0.5,
    top: 0.5, bottom: 0.5,
    header: 0, footer: 0,
  };

  COL_WIDTHS.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // addPageBreak(row) helper — inserts a manual row page break after that row
  const addBreak = (r: number) => ws.getRow(r).addPageBreak();

  let row = 1;

  // ═══════════════════════════════════════════════════════════════════════════
  // Section 1 — Patriot Title Sheet
  // ═══════════════════════════════════════════════════════════════════════════

  const ADDRESS_LINES = [
    "Patriot Pipeline Inc.",
    "PO Box 1487",
    "Wildomar, Ca 92595",
    "Phone: 951-679-8364",
    "Fax: 951-304-0684",
  ];

  ADDRESS_LINES.forEach((line, i) => {
    const r = ws.getRow(row + i);
    r.height = 15;
    styleCell(r.getCell(COL_A), { value: line, size: 9, vAlign: "middle" });
  });

  // "Submittals" heading — cols C+D, row 1
  merge(ws, row, COL_C, row, COL_D);
  styleCell(cell(ws, row, COL_C), {
    value: "Submittals",
    size: 22,
    bold: true,
    hAlign: "right",
    vAlign: "middle",
  });
  ws.getRow(row).height = 22;

  // Job No — cols C+D, row 3
  merge(ws, row + 2, COL_C, row + 2, COL_D);
  styleCell(cell(ws, row + 2, COL_C), {
    value: data.jobNo ? `Job No: ${data.jobNo}` : "Job No:",
    size: 10,
    bold: true,
    hAlign: "right",
    vAlign: "middle",
  });

  // Date — cols C+D, row 4
  merge(ws, row + 3, COL_C, row + 3, COL_D);
  styleCell(cell(ws, row + 3, COL_C), {
    value: data.date ? `Date: ${data.date}` : "Date:",
    size: 10,
    bold: true,
    hAlign: "right",
    vAlign: "middle",
  });

  // Logo — proportionally scaled to the wider column layout
  if (logoBuffer) {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const imgId = wb.addImage({ buffer: logoBuffer as any, extension: "png" });
      ws.addImage(imgId, {
        tl: { col: COL_B - 1 + 0.05, row: 0.1 },
        ext: { width: 225, height: 87 }, // scaled from original 190×73 for wider cols
      });
    } catch { /* skip logo if embed fails */ }
  }

  row += ADDRESS_LINES.length; // row 6

  // Spacer
  row++;
  ws.getRow(row).height = 6;
  row++;

  // ── To / Subject table ──────────────────────────────────────────────────────
  merge(ws, row, COL_A, row, COL_B);
  styleCell(cell(ws, row, COL_A), { value: "To:", bold: true, indent: 1, fill: GRAY, border: ALL_BORDERS });
  merge(ws, row, COL_C, row, COL_D);
  styleCell(cell(ws, row, COL_C), { value: "Subject:", bold: true, indent: 1, fill: GRAY, border: ALL_BORDERS });
  ws.getRow(row).height = 18;
  row++;

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

  const bodyH = Math.max(toLines.length, subjectLines.length, 4) * 15 + 8;

  merge(ws, row, COL_A, row, COL_B);
  styleCell(cell(ws, row, COL_A), {
    value: toLines.join("\n"),
    wrap: true, vAlign: "top", indent: 1, border: ALL_BORDERS,
  });
  ws.getRow(row).height = bodyH;

  merge(ws, row, COL_C, row, COL_D);
  styleCell(cell(ws, row, COL_C), {
    value: subjectLines.join("\n"),
    wrap: true, vAlign: "top", indent: 1, border: ALL_BORDERS,
  });

  row++;

  // Spacer
  row++;
  ws.getRow(row).height = 6;
  row++;

  // ── TOC table ───────────────────────────────────────────────────────────────
  merge(ws, row, COL_A, row, COL_C);
  styleCell(cell(ws, row, COL_A), { value: "Item", bold: true, indent: 1, fill: GRAY, border: ALL_BORDERS });
  styleCell(cell(ws, row, COL_D), {
    value: "Page Number",
    bold: true, hAlign: "right", indent: 1, fill: GRAY, border: ALL_BORDERS,
  });
  ws.getRow(row).height = 18;
  row++;

  for (const category of data.categories) {
    merge(ws, row, COL_A, row, COL_D);
    styleCell(cell(ws, row, COL_A), {
      value: category.name,
      bold: true, hAlign: "center", border: ALL_BORDERS,
    });
    ws.getRow(row).height = 16;
    row++;

    for (const item of category.lineItems) {
      const pageStr = item.startPage === item.endPage
        ? `${item.startPage}`
        : `${item.startPage}-${item.endPage}`;

      merge(ws, row, COL_A, row, COL_C);
      styleCell(cell(ws, row, COL_A), { value: item.description, indent: 2, border: ALL_BORDERS });
      styleCell(cell(ws, row, COL_D), { value: pageStr, hAlign: "right", indent: 1, border: ALL_BORDERS });
      ws.getRow(row).height = 15;
      row++;
    }
  }

  // Page break after title sheet
  addBreak(row - 1);

  // ═══════════════════════════════════════════════════════════════════════════
  // Section 2 — Blank Page
  // ═══════════════════════════════════════════════════════════════════════════
  // 8 rows × 90 pt = 720 pt = 10" of content height (Letter with 0.5" margins)
  const BLANK_ROWS = 8;
  for (let i = 0; i < BLANK_ROWS; i++) {
    ws.getRow(row).height = IMG_ROW_H;
    row++;
  }
  addBreak(row - 1);

  // ═══════════════════════════════════════════════════════════════════════════
  // Section 3 — Supplier Data Sheet Pages (continuous, one per print page)
  // ═══════════════════════════════════════════════════════════════════════════
  if (pageImages && pageImages.length > 0) {
    for (const img of pageImages) {
      const imageStartRow = row; // 1-indexed

      // Scale image to printable Letter width; maintain source aspect ratio
      const displayH = Math.round(IMG_DISPLAY_W * img.height / img.width);

      // Rows must collectively span the image height so nothing overlaps below
      const totalPt = Math.round(displayH * PX_TO_PT);
      const numRows = Math.max(1, Math.ceil(totalPt / IMG_ROW_H));
      const rowH = Math.ceil(totalPt / numRows);

      for (let i = 0; i < numRows; i++) {
        ws.getRow(row).height = rowH;
        row++;
      }

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const imgId = wb.addImage({ buffer: img.buffer as any, extension: "png" });
      ws.addImage(imgId, {
        tl: { col: 0, row: imageStartRow - 1 }, // row index is 0-based in ExcelJS
        ext: { width: IMG_DISPLAY_W, height: displayH },
      });

      addBreak(row - 1);
    }
  }

  ws.pageSetup.printArea = `A1:D${row - 1}`;

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}
