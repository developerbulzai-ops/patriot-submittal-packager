import ExcelJS from "exceljs";
import type { SubmittalData } from "@/types/submittal";

// ── layout constants ──────────────────────────────────────────────────────────
// 4-column layout:  A=left address  B=logo area  C=right header/content  D=page nums
const COL_A = 1, COL_B = 2, COL_C = 3, COL_D = 4;
const COL_WIDTHS = [22, 22, 26, 11]; // characters

const GRAY = { type: "pattern" as const, pattern: "solid" as const, fgColor: { argb: "FFE0E0E0" } };
const BORDER_THIN: Partial<ExcelJS.Border> = { style: "thin" };
const ALL_BORDERS = { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN };
const FONT = "Calibri";

// ── helpers ───────────────────────────────────────────────────────────────────

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

// ── main export ───────────────────────────────────────────────────────────────

export async function buildExcel(data: SubmittalData, logoBuffer?: Buffer): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Patriot Pipeline Inc.";
  wb.created = new Date();

  // ════════════════════════════════════════════════════════════════════════════
  // Sheet 1 — Title Sheet
  // ════════════════════════════════════════════════════════════════════════════
  const ws = wb.addWorksheet("Title Sheet", {
    pageSetup: {
      paperSize: 1 as number,           // US Letter
      orientation: "portrait",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      horizontalCentered: true,
      verticalCentered: false,
    },
  });

  ws.pageSetup.margins = {
    left: 0.5, right: 0.5,
    top: 0.6, bottom: 0.6,
    header: 0.3, footer: 0.3,
  };

  COL_WIDTHS.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  let row = 1;

  // ── Header block (rows 1-6) ─────────────────────────────────────────────────
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

  // Logo image centred in cols B-C, rows 1-5
  if (logoBuffer) {
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const imgId = wb.addImage({ buffer: logoBuffer as any, extension: "png" });
      ws.addImage(imgId, {
        tl: { col: COL_B - 1 + 0.05, row: 0.1 },  // col index is 0-based
        ext: { width: 190, height: 73 },            // px at 96 DPI
      });
    } catch { /* skip logo if embed fails */ }
  }

  row += ADDRESS_LINES.length; // row 6

  // ── Spacer ──────────────────────────────────────────────────────────────────
  row++;
  ws.getRow(row).height = 6;
  row++;

  // ── To / Subject table ──────────────────────────────────────────────────────
  // Header row
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

  // Body cells (merged tall blocks)
  merge(ws, row, COL_A, row, COL_B);
  const toBody = cell(ws, row, COL_A);
  styleCell(toBody, {
    value: toLines.join("\n"),
    wrap: true,
    vAlign: "top",
    indent: 1,
    border: ALL_BORDERS,
  });
  ws.getRow(row).height = bodyH;

  merge(ws, row, COL_C, row, COL_D);
  const subjectBody = cell(ws, row, COL_C);
  styleCell(subjectBody, {
    value: subjectLines.join("\n"),
    wrap: true,
    vAlign: "top",
    indent: 1,
    border: ALL_BORDERS,
  });

  row++;

  // ── Spacer ──────────────────────────────────────────────────────────────────
  row++;
  ws.getRow(row).height = 6;
  row++;

  // ── TOC table ───────────────────────────────────────────────────────────────
  // Header row
  merge(ws, row, COL_A, row, COL_C);
  styleCell(cell(ws, row, COL_A), { value: "Item", bold: true, indent: 1, fill: GRAY, border: ALL_BORDERS });
  styleCell(cell(ws, row, COL_D), {
    value: "Page Number",
    bold: true,
    hAlign: "right",
    indent: 1,
    fill: GRAY,
    border: ALL_BORDERS,
  });
  ws.getRow(row).height = 18;
  row++;

  for (const category of data.categories) {
    // Category header — full width, bold, centred
    merge(ws, row, COL_A, row, COL_D);
    styleCell(cell(ws, row, COL_A), {
      value: category.name,
      bold: true,
      hAlign: "center",
      indent: 0,
      border: ALL_BORDERS,
    });
    ws.getRow(row).height = 16;
    row++;

    for (const item of category.lineItems) {
      const pageStr = item.startPage === item.endPage
        ? `${item.startPage}`
        : `${item.startPage}-${item.endPage}`;

      merge(ws, row, COL_A, row, COL_C);
      styleCell(cell(ws, row, COL_A), { value: item.description, indent: 2, border: ALL_BORDERS });
      styleCell(cell(ws, row, COL_D), {
        value: pageStr,
        hAlign: "right",
        indent: 1,
        border: ALL_BORDERS,
      });
      ws.getRow(row).height = 15;
      row++;
    }
  }

  ws.pageSetup.printArea = `A1:D${row - 1}`;

  // ════════════════════════════════════════════════════════════════════════════
  // Sheet 2 — Blank (separator before supplier pages)
  // ════════════════════════════════════════════════════════════════════════════
  wb.addWorksheet("Blank", {
    pageSetup: { paperSize: 1 as number, orientation: "portrait" },
  });

  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}
