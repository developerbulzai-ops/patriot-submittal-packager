import Anthropic from "@anthropic-ai/sdk";
import type { SubmittalData, CategoryGroup } from "@/types/submittal";

const EXTRACTION_PROMPT = `You are analyzing a supplier submittal PDF for a construction project.

CRITICAL — HIGHLIGHTED TEXT: Suppliers highlight (yellow or colored background) the specific sizes,
gauges, and specs that apply to this particular job. When naming line items, use ONLY the highlighted
specifications, not the full range printed on the data sheet. For example: if a pipe data sheet lists
sizes 4"–48" but only 18" is highlighted, the line item title must say "18\" SDR-35 PVC Pipe" — not
the full range. If no highlighting is visible, use the most prominent or first-listed specification.

CATEGORIES: Supplier cover sheets and section headers indicate the submittal category
(e.g. "On-Site Rough Grade Storm Drain Pipe and Fittings", "Off-Site Storm Drain", "Fire",
"Water", "Sewer"). Detect ALL categories present in the package and group line items under them.
Common categories: On-Site, Off-Site, Fire, Water, Storm Drain, Sewer — but use whatever the
document actually says.

Extract the following and return ONLY a valid JSON object (no markdown, no explanation):

1. From the cover page (page 1): recipient info, job number, date, project name, location.
2. For pages 2 through the end: group by manufacturer/product into categories.
   - Each group = one TOC line item
   - Format description as "ManufacturerName - HighlightedSpec ProductType"
   - Record exact start and end page numbers in THIS PDF (1-indexed, page 1 = supplier cover)
   - Page 1 (supplier cover) is NOT a line item
   - If two manufacturers supply the same product, list them under one item

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
  "categories": [
    {
      "name": "On-Site Rough Grade Storm Drain Pipe and Fittings",
      "lineItems": [
        {
          "description": "JM Eagle - 18\" SDR-35 PVC Pipe",
          "startPage": 2,
          "endPage": 6
        }
      ]
    }
  ]
}

Rules:
- categories: detect ALL utility categories in the package; each gets its own object
- lineItems within each category must cover all pages in that section with no gaps or overlaps
- Page numbers are in the SUPPLIER PDF (not the final output)
- Use highlighted/circled specs for item titles; never list the full range from the data sheet`;

function extractJSON(text: string): unknown {
  const cleaned = text
    .replace(/^```(?:json)?\s*/m, "")
    .replace(/\s*```\s*$/m, "")
    .trim();
  return JSON.parse(cleaned);
}

export async function extractSubmittal(
  pdfBuffer: Buffer
): Promise<Partial<SubmittalData> & { categories: CategoryGroup[] }> {
  const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
  const base64Pdf = pdfBuffer.toString("base64");

  const content: Anthropic.MessageParam["content"] = [
    {
      type: "document",
      source: { type: "base64", media_type: "application/pdf", data: base64Pdf },
    } as Anthropic.DocumentBlockParam,
    { type: "text", text: EXTRACTION_PROMPT },
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
    categories: Array<{
      name: string;
      lineItems: Array<{ description: string; startPage: number; endPage: number }>;
    }>;
  };

  // Page number mapping: supplier cover (page 1) stripped, Patriot title = page 1,
  // blank sheet = page 2, so supplier page K → output page K+1.
  const categories: CategoryGroup[] = (raw.categories || []).map((cat) => ({
    name: cat.name || "",
    lineItems: (cat.lineItems || []).map((item) => ({
      description: item.description,
      startPage: item.startPage + 1,
      endPage: item.endPage + 1,
    })),
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
    categories,
  };
}
