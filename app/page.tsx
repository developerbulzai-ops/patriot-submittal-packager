"use client";

import { useState, useRef, useCallback } from "react";
import type { SubmittalData, LineItem } from "@/types/submittal";

type Step = "upload" | "review" | "generating" | "done";

const emptyData = (): SubmittalData => ({
  jobNo: "",
  date: "",
  recipient: { company: "", attention: "", address1: "", city: "" },
  subject: { projectName: "", location: "" },
  category: "",
  lineItems: [],
});

// ── small field component ─────────────────────────────────────────────────────
function Field({
  label,
  value,
  onChange,
  placeholder,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  placeholder?: string;
}) {
  return (
    <div className="flex flex-col gap-1">
      <label className="text-xs font-semibold uppercase tracking-widest text-slate-400">
        {label}
      </label>
      <input
        type="text"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className="rounded bg-slate-700 border border-slate-600 px-3 py-2 text-sm text-slate-100
                   placeholder:text-slate-500 focus:outline-none focus:ring-2 focus:ring-blue-500"
      />
    </div>
  );
}

// ── main page ─────────────────────────────────────────────────────────────────
export default function Home() {
  const [step, setStep] = useState<Step>("upload");
  const [dragging, setDragging] = useState(false);
  const [file, setFile] = useState<File | null>(null);
  const [extracting, setExtracting] = useState(false);
  const [error, setError] = useState("");
  const [form, setForm] = useState<SubmittalData>(emptyData());
  const [downloadUrl, setDownloadUrl] = useState("");
  const [downloadName, setDownloadName] = useState("submittal.pdf");
  const fileInputRef = useRef<HTMLInputElement>(null);

  // ── drag and drop ─────────────────────────────────────────────────────────
  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setDragging(false);
    const dropped = e.dataTransfer.files[0];
    if (dropped?.type === "application/pdf") {
      setFile(dropped);
      setError("");
    } else {
      setError("Please upload a PDF file.");
    }
  }, []);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const picked = e.target.files?.[0];
    if (picked) {
      setFile(picked);
      setError("");
    }
  };

  // ── extract ───────────────────────────────────────────────────────────────
  const handleExtract = async () => {
    if (!file) return;
    if (file.size > 4 * 1024 * 1024) {
      setError("File is too large for the current host (max 4 MB). Compress the PDF and try again.");
      return;
    }
    setExtracting(true);
    setError("");
    try {
      const fd = new FormData();
      fd.append("pdf", file);
      const res = await fetch("/api/extract", { method: "POST", body: fd });
      if (!res.ok) {
        const text = await res.text();
        let msg = "Extraction failed";
        try { msg = (JSON.parse(text) as { error?: string }).error || msg; } catch { msg = text.slice(0, 300) || msg; }
        throw new Error(msg);
      }
      const data = await res.json();
      setForm({ ...emptyData(), ...data });
      setStep("review");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Extraction failed");
    } finally {
      setExtracting(false);
    }
  };

  // ── generate ──────────────────────────────────────────────────────────────
  const handleGenerate = async () => {
    if (!file) return;
    setStep("generating");
    setError("");
    try {
      const fd = new FormData();
      fd.append("pdf", file);
      fd.append("data", JSON.stringify(form));
      const res = await fetch("/api/generate", { method: "POST", body: fd });
      if (!res.ok) {
        const text = await res.text();
        let msg = "Generation failed";
        try { msg = (JSON.parse(text) as { error?: string }).error || msg; } catch { msg = text.slice(0, 300) || msg; }
        throw new Error(msg);
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const disposition = res.headers.get("Content-Disposition") || "";
      const match = disposition.match(/filename="([^"]+)"/);
      setDownloadName(match?.[1] ?? "Patriot_Submittal.pdf");
      setDownloadUrl(url);
      setStep("done");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Generation failed");
      setStep("review");
    }
  };

  // ── line item helpers ─────────────────────────────────────────────────────
  const updateLineItem = (idx: number, field: keyof LineItem, val: string | number) => {
    setForm((f) => {
      const items = [...f.lineItems];
      items[idx] = { ...items[idx], [field]: val };
      return { ...f, lineItems: items };
    });
  };

  const addLineItem = () => {
    setForm((f) => ({
      ...f,
      lineItems: [...f.lineItems, { description: "", startPage: 0, endPage: 0 }],
    }));
  };

  const removeLineItem = (idx: number) => {
    setForm((f) => ({ ...f, lineItems: f.lineItems.filter((_, i) => i !== idx) }));
  };

  // ── render ────────────────────────────────────────────────────────────────
  return (
    <main className="min-h-screen bg-slate-900 text-slate-100">
      {/* Header */}
      <header className="border-b border-slate-800 px-8 py-5 flex items-center gap-4">
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img
          src="/assets/patriot_logo.png"
          alt="Patriot Pipeline"
          className="h-10 w-auto"
          onError={(e) => {
            (e.currentTarget as HTMLImageElement).style.display = "none";
          }}
        />
        <div>
          <h1 className="text-lg font-bold tracking-tight">Patriot Pipeline Inc.</h1>
          <p className="text-xs text-slate-400 tracking-widest uppercase">
            Submittal Packager
          </p>
        </div>
      </header>

      <div className="max-w-3xl mx-auto px-6 py-10">
        {/* ── STEP: upload ─────────────────────────────────────────────── */}
        {(step === "upload") && (
          <div className="flex flex-col gap-8">
            <div>
              <h2 className="text-2xl font-bold mb-1">New Submittal</h2>
              <p className="text-slate-400 text-sm">
                Upload a supplier submittal PDF to generate a Patriot-branded package.
              </p>
            </div>

            {/* Drop zone */}
            <div
              className={`relative flex flex-col items-center justify-center gap-3 rounded-xl border-2 border-dashed
                          p-14 transition-colors cursor-pointer
                          ${dragging ? "border-blue-400 bg-blue-500/10" : "border-slate-600 hover:border-slate-400 hover:bg-slate-800/60"}
                          ${file ? "border-green-500 bg-green-500/10" : ""}`}
              onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={handleDrop}
              onClick={() => fileInputRef.current?.click()}
            >
              <input
                ref={fileInputRef}
                type="file"
                accept="application/pdf"
                className="sr-only"
                onChange={handleFileChange}
              />
              {file ? (
                <>
                  <svg className="w-10 h-10 text-green-400" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                  </svg>
                  <p className="font-semibold text-green-400">{file.name}</p>
                  <p className="text-xs text-slate-400">{(file.size / 1024 / 1024).toFixed(2)} MB · Click to change</p>
                </>
              ) : (
                <>
                  <svg className="w-10 h-10 text-slate-500" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.5}>
                    <path strokeLinecap="round" strokeLinejoin="round"
                      d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m6.75 12l-3-3m0 0l-3 3m3-3v6m-1.5-15H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
                  </svg>
                  <p className="text-slate-300 font-medium">Drop supplier PDF here</p>
                  <p className="text-xs text-slate-500">or click to browse</p>
                </>
              )}
            </div>

            {error && <p className="text-red-400 text-sm">{error}</p>}

            <button
              onClick={handleExtract}
              disabled={!file || extracting}
              className="w-full rounded-lg bg-blue-600 hover:bg-blue-500 disabled:bg-slate-700
                         disabled:text-slate-500 py-3 font-semibold text-sm tracking-wide
                         transition-colors disabled:cursor-not-allowed"
            >
              {extracting ? (
                <span className="flex items-center justify-center gap-2">
                  <svg className="animate-spin w-4 h-4" fill="none" viewBox="0 0 24 24">
                    <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                    <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z" />
                  </svg>
                  Extracting with AI…
                </span>
              ) : (
                "Extract & Review →"
              )}
            </button>
          </div>
        )}

        {/* ── STEP: review ─────────────────────────────────────────────── */}
        {step === "review" && (
          <div className="flex flex-col gap-8">
            <div className="flex items-center justify-between">
              <div>
                <h2 className="text-2xl font-bold mb-1">Review & Edit</h2>
                <p className="text-slate-400 text-sm">Confirm or fill in the extracted details.</p>
              </div>
              <button
                onClick={() => { setStep("upload"); setError(""); }}
                className="text-xs text-slate-400 hover:text-slate-200 underline"
              >
                ← Upload different file
              </button>
            </div>

            {/* Job info */}
            <section className="rounded-xl bg-slate-800 border border-slate-700 p-5 flex flex-col gap-4">
              <h3 className="text-xs font-bold uppercase tracking-widest text-slate-400">Job Info</h3>
              <div className="grid grid-cols-2 gap-4">
                <Field label="Job No" value={form.jobNo} onChange={(v) => setForm((f) => ({ ...f, jobNo: v }))} placeholder="e.g. 1919" />
                <Field label="Date" value={form.date} onChange={(v) => setForm((f) => ({ ...f, date: v }))} placeholder="e.g. 3/20/2026" />
              </div>
            </section>

            {/* Recipient */}
            <section className="rounded-xl bg-slate-800 border border-slate-700 p-5 flex flex-col gap-4">
              <h3 className="text-xs font-bold uppercase tracking-widest text-slate-400">To</h3>
              <div className="grid grid-cols-2 gap-4">
                <Field label="Company" value={form.recipient.company} onChange={(v) => setForm((f) => ({ ...f, recipient: { ...f.recipient, company: v } }))} placeholder="General contractor name" />
                <Field label="Attention" value={form.recipient.attention} onChange={(v) => setForm((f) => ({ ...f, recipient: { ...f.recipient, attention: v } }))} placeholder="Contact name" />
                <Field label="Address" value={form.recipient.address1} onChange={(v) => setForm((f) => ({ ...f, recipient: { ...f.recipient, address1: v } }))} placeholder="Street address" />
                <Field label="City / State / Zip" value={form.recipient.city} onChange={(v) => setForm((f) => ({ ...f, recipient: { ...f.recipient, city: v } }))} placeholder="City, CA 12345" />
              </div>
            </section>

            {/* Subject */}
            <section className="rounded-xl bg-slate-800 border border-slate-700 p-5 flex flex-col gap-4">
              <h3 className="text-xs font-bold uppercase tracking-widest text-slate-400">Subject</h3>
              <Field label="Project Name" value={form.subject.projectName} onChange={(v) => setForm((f) => ({ ...f, subject: { ...f.subject, projectName: v } }))} placeholder="Project name" />
              <Field label="Location" value={form.subject.location} onChange={(v) => setForm((f) => ({ ...f, subject: { ...f.subject, location: v } }))} placeholder="Address / cross-streets, City, CA 12345" />
            </section>

            {/* Category + TOC */}
            <section className="rounded-xl bg-slate-800 border border-slate-700 p-5 flex flex-col gap-4">
              <h3 className="text-xs font-bold uppercase tracking-widest text-slate-400">Table of Contents</h3>
              <Field
                label="Category Header"
                value={form.category}
                onChange={(v) => setForm((f) => ({ ...f, category: v }))}
                placeholder="e.g. On-Site Rough Grade Storm Drain Pipe and Fittings"
              />

              <div className="flex flex-col gap-2 mt-1">
                <div className="grid grid-cols-[1fr_64px_64px_24px] gap-2 text-xs font-semibold uppercase tracking-wider text-slate-400 px-1">
                  <span>Description</span>
                  <span className="text-center">Start</span>
                  <span className="text-center">End</span>
                  <span />
                </div>
                {form.lineItems.map((item, i) => (
                  <div key={i} className="grid grid-cols-[1fr_64px_64px_24px] gap-2 items-center">
                    <input
                      type="text"
                      value={item.description}
                      onChange={(e) => updateLineItem(i, "description", e.target.value)}
                      className="rounded bg-slate-700 border border-slate-600 px-3 py-2 text-sm text-slate-100
                                 placeholder:text-slate-500 focus:outline-none focus:ring-2 focus:ring-blue-500"
                      placeholder="Manufacturer - Product"
                    />
                    <input
                      type="number"
                      value={item.startPage}
                      onChange={(e) => updateLineItem(i, "startPage", parseInt(e.target.value) || 0)}
                      className="rounded bg-slate-700 border border-slate-600 px-2 py-2 text-sm text-slate-100 text-center
                                 focus:outline-none focus:ring-2 focus:ring-blue-500"
                    />
                    <input
                      type="number"
                      value={item.endPage}
                      onChange={(e) => updateLineItem(i, "endPage", parseInt(e.target.value) || 0)}
                      className="rounded bg-slate-700 border border-slate-600 px-2 py-2 text-sm text-slate-100 text-center
                                 focus:outline-none focus:ring-2 focus:ring-blue-500"
                    />
                    <button
                      onClick={() => removeLineItem(i)}
                      className="text-slate-500 hover:text-red-400 transition-colors text-lg leading-none"
                      title="Remove row"
                    >
                      ×
                    </button>
                  </div>
                ))}
                <button
                  onClick={addLineItem}
                  className="mt-1 text-sm text-blue-400 hover:text-blue-300 text-left transition-colors"
                >
                  + Add row
                </button>
              </div>
            </section>

            {error && <p className="text-red-400 text-sm">{error}</p>}

            <button
              onClick={handleGenerate}
              className="w-full rounded-lg bg-blue-600 hover:bg-blue-500 py-3 font-semibold
                         text-sm tracking-wide transition-colors"
            >
              Generate Submittal →
            </button>
          </div>
        )}

        {/* ── STEP: generating ─────────────────────────────────────────── */}
        {step === "generating" && (
          <div className="flex flex-col items-center justify-center gap-6 py-24">
            <svg className="animate-spin w-12 h-12 text-blue-500" fill="none" viewBox="0 0 24 24">
              <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
              <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z" />
            </svg>
            <p className="text-slate-300 font-semibold">Building your submittal…</p>
            <p className="text-slate-500 text-sm">Generating cover page and merging PDFs</p>
          </div>
        )}

        {/* ── STEP: done ───────────────────────────────────────────────── */}
        {step === "done" && (
          <div className="flex flex-col items-center gap-8 py-16">
            <div className="flex flex-col items-center gap-3">
              <div className="w-16 h-16 rounded-full bg-green-500/20 flex items-center justify-center">
                <svg className="w-8 h-8 text-green-400" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                  <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
                </svg>
              </div>
              <h2 className="text-2xl font-bold">Submittal Ready</h2>
              <p className="text-slate-400 text-sm">{downloadName}</p>
            </div>

            <a
              href={downloadUrl}
              download={downloadName}
              className="rounded-lg bg-green-600 hover:bg-green-500 px-8 py-3 font-semibold
                         text-sm tracking-wide transition-colors flex items-center gap-2"
            >
              <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
              </svg>
              Download PDF
            </a>

            <button
              onClick={() => {
                setStep("upload");
                setFile(null);
                setForm(emptyData());
                setDownloadUrl("");
                setError("");
              }}
              className="text-sm text-slate-400 hover:text-slate-200 underline"
            >
              Start a new submittal
            </button>
          </div>
        )}
      </div>
    </main>
  );
}
