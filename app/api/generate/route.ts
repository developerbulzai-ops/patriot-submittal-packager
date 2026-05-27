import { NextRequest, NextResponse } from "next/server";
import { buildCoverPage } from "@/lib/buildCoverPage";
import { mergePdfs } from "@/lib/mergePdfs";
import fs from "fs";
import path from "path";
import type { SubmittalData } from "@/types/submittal";

export const runtime = "nodejs";
export const maxDuration = 60;

function slugify(str: string): string {
  return str
    .replace(/[^a-zA-Z0-9\s-]/g, "")
    .trim()
    .replace(/\s+/g, "_")
    .substring(0, 40);
}

export async function POST(req: NextRequest) {
  if (!process.env.ANTHROPIC_API_KEY) {
    return NextResponse.json(
      { error: "ANTHROPIC_API_KEY is not configured. Add it to your Vercel environment variables." },
      { status: 500 }
    );
  }
  try {
    const formData = await req.formData();
    const file = formData.get("pdf");
    const dataStr = formData.get("data");

    if (!file || typeof file === "string") {
      return NextResponse.json({ error: "No PDF file provided" }, { status: 400 });
    }
    if (!dataStr || typeof dataStr !== "string") {
      return NextResponse.json({ error: "No submittal data provided" }, { status: 400 });
    }

    const submittalData: SubmittalData = JSON.parse(dataStr);
    const supplierBuffer = Buffer.from(await (file as File).arrayBuffer());

    // Load logo from public assets
    let logoBytes: Uint8Array | undefined;
    try {
      const logoPath = path.join(process.cwd(), "public", "assets", "patriot_logo.png");
      logoBytes = fs.readFileSync(logoPath);
    } catch { /* use text fallback */ }

    const coverBytes = await buildCoverPage(submittalData, logoBytes);
    const mergedBytes = await mergePdfs(coverBytes, supplierBuffer);

    const jobPart = slugify(submittalData.jobNo || "Job");
    const projPart = slugify(submittalData.subject.projectName || "Submittal");
    const filename = `${jobPart}_${projPart}_Submittal.pdf`;

    return new NextResponse(Buffer.from(mergedBytes), {
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
