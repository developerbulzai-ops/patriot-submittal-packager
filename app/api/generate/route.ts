import { NextRequest, NextResponse } from "next/server";
import { buildPdf } from "@/lib/buildPdf";
import fs from "fs";
import path from "path";
import type { SubmittalData } from "@/types/submittal";

export const runtime = "nodejs";

function slugify(str: string): string {
  return str
    .replace(/[^a-zA-Z0-9\s-]/g, "")
    .trim()
    .replace(/\s+/g, "_")
    .substring(0, 40);
}

export async function POST(req: NextRequest) {
  try {
    let submittalData: SubmittalData;
    let pdfBuffer: Buffer | undefined;

    const contentType = req.headers.get("content-type") ?? "";
    if (contentType.includes("multipart/form-data")) {
      const fd = await req.formData();
      const dataStr = fd.get("data");
      if (typeof dataStr !== "string") throw new Error("Missing form field: data");
      submittalData = JSON.parse(dataStr) as SubmittalData;
      const pdfFile = fd.get("pdf");
      if (pdfFile instanceof Blob) {
        pdfBuffer = Buffer.from(await pdfFile.arrayBuffer());
      }
    } else {
      submittalData = (await req.json()) as SubmittalData;
    }

    if (!pdfBuffer) throw new Error("Supplier PDF is required");

    let logoBuffer: Buffer | undefined;
    try {
      const logoPath = path.join(process.cwd(), "public", "assets", "patriot_logo.png");
      logoBuffer = fs.readFileSync(logoPath);
    } catch { /* text fallback */ }

    const pdfOut = await buildPdf(submittalData, logoBuffer, pdfBuffer);

    const jobPart = slugify(submittalData.jobNo || "Job");
    const projPart = slugify(submittalData.subject.projectName || "Submittal");
    const filename = `${jobPart}_${projPart}_Submittal.pdf`;

    return new NextResponse(pdfOut as unknown as BodyInit, {
      status: 200,
      headers: {
        "Content-Type": "application/pdf",
        "Content-Disposition": `attachment; filename="${filename}"`,
      },
    });
  } catch (err) {
    console.error("Generate error:", err);
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Generation failed" },
      { status: 500 }
    );
  }
}
