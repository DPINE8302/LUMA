import type { Material, MaterialType } from "./types";

export function fileTypeFromName(name: string): MaterialType {
  const lower = name.toLowerCase();
  if (lower.endsWith(".pdf")) return "pdf";
  if (/\.(png|jpg|jpeg|webp|gif)$/.test(lower)) return "image";
  if (/\.(docx|doc)$/.test(lower)) return "doc";
  if (/\.(pptx|ppt)$/.test(lower)) return "slide";
  if (/\.(xlsx|xls|csv)$/.test(lower)) return "sheet";
  if (lower.startsWith("http")) return "link";
  return "text";
}

export async function extractFileText(file: File): Promise<string> {
  const type = fileTypeFromName(file.name);
  if (type === "pdf") {
    const pdfjs = await import("pdfjs-dist");
    pdfjs.GlobalWorkerOptions.workerSrc = new URL("pdfjs-dist/build/pdf.worker.mjs", import.meta.url).toString();
    const document = await pdfjs.getDocument({ data: await file.arrayBuffer() }).promise;
    const pages = await Promise.all(
      Array.from({ length: document.numPages }, async (_, index) => {
        const page = await document.getPage(index + 1);
        const text = await page.getTextContent();
        return text.items.map((item) => ("str" in item ? item.str : "")).join(" ");
      }),
    );
    return pages.join("\n\n");
  }

  if (type === "doc") {
    const mammoth = await import("mammoth");
    const result = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
    return result.value;
  }

  if (type === "sheet") {
    if (file.name.toLowerCase().endsWith(".csv")) {
      return file.text();
    }
    const readXlsxFile = (await import("read-excel-file/browser")).default;
    const sheets = await readXlsxFile(file);
    return sheets
      .map((sheet) => [`[${sheet.sheet}]`, ...sheet.data.map((row) => row.map((cell) => String(cell ?? "")).join(", "))].join("\n"))
      .join("\n\n");
  }

  if (type === "image") {
    const { createWorker } = await import("tesseract.js");
    const worker = await createWorker("eng");
    const result = await worker.recognize(file);
    await worker.terminate();
    return result.data.text;
  }

  return file.text();
}

export async function materialFromFile(file: File, subjectId: string): Promise<Material> {
  const now = new Date().toISOString();
  let content = "";
  try {
    content = await extractFileText(file);
  } catch (error) {
    content = `LUMA could not fully extract this file yet. ${(error as Error).message}`;
  }

  return {
    id: crypto.randomUUID(),
    title: file.name.replace(/\.[^.]+$/, ""),
    subjectId,
    type: fileTypeFromName(file.name),
    sourceName: file.name,
    folder: "Uploads",
    tags: ["uploaded"],
    content,
    createdAt: now,
    updatedAt: now,
    sharedWith: [],
  };
}
