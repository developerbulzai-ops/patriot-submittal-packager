import Anthropic from "@anthropic-ai/sdk";
import type { SubmittalData } from "@/types/submittal";

const EXTRACTION_PROMPT = `You are analyzing a supplier submittal PDF for a construction project.

Extract the following information and return ONLY a valid JSON object (no markdown, no explanation):

1. From the cover page (page 1 of the PDF):
   - Project/submittal title (becomes category header in our TOC, e.g. "On-Site Rough Grade Storm Drain Pipe and Fittings")
   - Project name and location (e.g. "Walmart Wildomar Phase 1 Utilities", location address)
   - Any recipient info, job number, or date if visible

2. For pages 2 through the end: identify groups of pages by manufacturer/product:
   - Each group = one TOC line item
   - Format description as "ManufacturerName - ProductDescription"
   - Group related products from the same manufacturer together
   - If two manufacturers supply the same product type, list them under a single item led by the primary/first manufacturer
   - Record exact start and end page numbers within THIS PDF (1-indexed, page 1 = supplier cover)
   - Page 1 (supplier cover) is NOT a line item

Return this exact JSON structure:
{
  "jobNo": null,
  "date": null,
  "recipient": {
    "company": null,
    "attention": null,
    "address1": null,
    "city": null
  },
  "subject": {
    "projectName": null,
    "location": null
  },
  "category": null,
  "lineItems": [
    {
      "description": "Manufacturer Name - Product Description",
      "startPage": 2,
      "endPage": 4
    }
  ]
}

Rules:
- category: a concise description of the overall product type (e.g. "On-Site Rough Grade Storm Drain Pipe and Fittings")
- lineItems must cover all pages from page 2 to the last page with no gaps and no overlaps
- Page numbers are in the SUPPLIER PDF (not the final output)
- Combine manufacturers making the same product into one line item spanning all their pages`;

function extractJSON(text: string): unknown {
  // Strip markdown code fences if present
  const cleaned = text
    .replace(/^```(?:json)?\s*/m, "")
    .replace(/\s*```\s*$/m, "")
    .trim();
  return JSON.parse(cleaned);
}

export async function extractSubmittal(
  pdfBuffer: Buffer
): Promise<Partial<SubmittalData> & { lineItems: Array<{ description: string; startPage: number; endPage: number }> }> {
  const anthropic = new Anthropic({
    apiKey: process.env.ANTHROPIC_API_KEY,
  });

  const base64Pdf = pdfBuffer.toString("base64");

  const content: Anthropic.MessageParam["content"] = [
    {
      type: "document",
      source: {
        type: "base64",
        media_type: "application/pdf",
        data: base64Pdf,
      },
    } as Anthropic.DocumentBlockParam,
    {
      type: "text",
      text: EXTRACTION_PROMPT,
    },
  ];

  const message = await anthropic.messages.create({
    model: "claude-sonnet-4-6",
    max_tokens: 4096,
    messages: [{ role: "user", content }],
  });

  const responseText =
    message.content[0].type === "text" ? message.content[0].text : "{}";

  const raw = extractJSON(responseText) as {
    jobNo: string | null;
    date: string | null;
    recipient: { company: string | null; attention: string | null; address1: string | null; city: string | null };
    subject: { projectName: string | null; location: string | null };
    category: string | null;
    lineItems: Array<{ description: string; startPage: number; endPage: number }>;
  };

  // Page number mapping: supplier cover (page 1) is stripped, Patriot cover becomes page 1.
  // Supplier page 2 → output page 2, so offset = 0. No adjustment needed.
  const lineItems = (raw.lineItems || []).map((item) => ({
    description: item.description,
    startPage: item.startPage,
    endPage: item.endPage,
  }));

  return {
    jobNo: raw.jobNo || "",
    date: raw.date || "",
    recipient: {
      company: raw.recipient?.company || "",
      attention: raw.recipient?.attention || "",
      address1: raw.recipient?.address1 || "",
      city: raw.recipient?.city || "",
    },
    subject: {
      projectName: raw.subject?.projectName || "",
      location: raw.subject?.location || "",
    },
    category: raw.category || "",
    lineItems,
  };
}
