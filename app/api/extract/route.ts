import { NextRequest, NextResponse } from "next/server";
import { extractSubmittal } from "@/lib/extractSubmittal";

export const runtime = "edge";
export const maxDuration = 60;

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

    if (!file || typeof file === "string") {
      return NextResponse.json({ error: "No PDF file provided" }, { status: 400 });
    }

    const buffer = Buffer.from(await (file as File).arrayBuffer());
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
