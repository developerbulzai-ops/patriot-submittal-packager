import { NextRequest, NextResponse } from "next/server";
import { extractSubmittal } from "@/lib/extractSubmittal";

export const runtime = "nodejs";
export const maxDuration = 60;

export async function POST(req: NextRequest) {
  if (!process.env.ANTHROPIC_API_KEY) {
    return NextResponse.json(
      { error: "ANTHROPIC_API_KEY is not configured. Add it to your Vercel environment variables." },
      { status: 500 }
    );
  }
  try {
    const { blobUrl } = (await req.json()) as { blobUrl: string };
    if (!blobUrl) {
      return NextResponse.json({ error: "No blobUrl provided" }, { status: 400 });
    }

    const pdfRes = await fetch(blobUrl);
    if (!pdfRes.ok) throw new Error(`Failed to fetch PDF from storage: ${pdfRes.status}`);
    const buffer = Buffer.from(await pdfRes.arrayBuffer());

    const data = await extractSubmittal(buffer);
    return NextResponse.json(data);
  } catch (err) {
    console.error("Extract error:", err);
    return NextResponse.json(
      { error: err instanceof Error ? err.message : "Extraction failed" },
      { status: 500 }
    );
  }
}
