import { NextRequest, NextResponse } from "next/server";
import { buildCoverPage } from "@/lib/buildCoverPage";
import { mergePdfs } from "@/lib/mergePdfs";
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

    // Build the Patriot-branded cover page
    const coverBytes = await buildCoverPage(submittalData);

    // Merge cover + all supplier pages
    const mergedBytes = await mergePdfs(coverBytes, supplierBuffer);

    // Build a clean filename
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
