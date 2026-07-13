
import React, { useMemo, useRef, useState } from "react";
import { UploadCloud, FileText, FileSpreadsheet, X, CheckCircle2, AlertTriangle, Download, ShieldCheck, Settings2, Loader2, ScanText } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Progress } from "@/components/ui/progress";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";
import { createWorker } from "tesseract.js";
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`;

const MAX_FILE_SIZE = 25 * 1024 * 1024;

function formatBytes(bytes) {
  if (!bytes) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / 1024 ** index).toFixed(index ? 1 : 0)} ${units[index]}`;
}

function safeSheetName(name, used) {
  const base = name.replace(/[\\/?*:[\]]/g, " ").trim().slice(0, 27) || "PDF";
  let candidate = base;
  let n = 2;
  while (used.has(candidate)) candidate = `${base.slice(0, 24)} ${n++}`;
  used.add(candidate);
  return candidate;
}

function groupItemsIntoLines(items) {
  const normalized = items
    .filter((item) => item.str?.trim())
    .map((item) => ({ text: item.str.trim(), x: Number(item.transform?.[4] || 0), y: Number(item.transform?.[5] || 0), width: Number(item.width || 0) }))
    .sort((a, b) => Math.abs(b.y - a.y) > 3 ? b.y - a.y : a.x - b.x);

  const lines = [];
  for (const item of normalized) {
    let line = lines.find((row) => Math.abs(row.y - item.y) <= 3);
    if (!line) { line = { y: item.y, items: [] }; lines.push(line); }
    line.items.push(item);
  }

  return lines.sort((a, b) => b.y - a.y).map((line) => {
    const cells = [];
    const sorted = line.items.sort((a, b) => a.x - b.x);
    for (const item of sorted) {
      const previous = cells[cells.length - 1];
      const gap = previous ? item.x - previous.end : Infinity;
      if (previous && gap < 14) {
        previous.text += ` ${item.text}`;
        previous.end = Math.max(previous.end, item.x + item.width);
      } else cells.push({ text: item.text, x: item.x, end: item.x + item.width });
    }
    return { text: sorted.map((item) => item.text).join(" "), cells };
  });
}

function ocrTextIntoLines(text) {
  return text.split(/\r?
/).map((value) => value.trim()).filter(Boolean).map((value) => ({
    text: value,
    cells: value.split(/\s{2,}|\t/).map((cell) => ({ text: cell.trim() })).filter((cell) => cell.text),
  }));
}



function rowsFromLines(lines, pageNumber, fileName, source) {
  return lines.map((line, lineIndex) => {
    const row = { Arquivo: fileName, Pagina: pageNumber, Linha: lineIndex + 1, Origem: source };
    line.cells.forEach((cell, index) => { row[`Campo_${index + 1}`] = cell.text; });
    row.Texto_integral = line.text;
    return row;
  });
}

async function renderPageForOcr(page) {
  const viewport = page.getViewport({ scale: 2.2 });
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d", { willReadFrequently: true });
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: context, viewport }).promise;
  return canvas;
}


async function extractPdf(file, options, onProgress, onOcrStatus) {
  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const rows = [];
  const rawRows = [];
  let emptyPages = 0;
  let ocrPages = 0;
  let worker = null;
  try {
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      const page = await pdf.getPage(pageNumber);
      const content = await page.getTextContent();
      let lines = groupItemsIntoLines(content.items);
      let source = "Texto do PDF";


      if (!lines.length && options.useOcr) {
        onOcrStatus(`OCR na pagina ${pageNumber} de ${pdf.numPages}`);
        if (!worker) worker = await createWorker(options.ocrLanguage);
        const canvas = await renderPageForOcr(page);
        const result = await worker.recognize(canvas);
        lines = ocrTextIntoLines(result.data.text || "");
        source = "OCR";
        if (lines.length) ocrPages += 1;
      }


      if (!lines.length) emptyPages += 1;
      rows.push(...rowsFromLines(lines, pageNumber, file.name, source));
      rawRows.push({
        Arquivo: file.name,
        Pagina: pageNumber,
        Origem: source,
        Texto_extraido: lines.map((line) => line.text).join("
"),
      });
      onProgress(Math.round((pageNumber / pdf.numPages) * 100));
    }
  } finally {
    if (worker) await worker.terminate();
  }

  return { rows, rawRows, pageCount: pdf.numPages, emptyPages, ocrPages };
}

function autosize(worksheet, rows) {
  if (!rows.length) return;
  const keys = [...new Set(rows.flatMap((row) => Object.keys(row)))];
  worksheet["!cols"] = keys.map((key) => ({ wch: Math.min(60, Math.max(key.length + 2, ...rows.slice(0, 500).map((row) => String(row[key] ?? "").length + 2))) }));
  worksheet["!autofilter"] = worksheet["!ref"] ? { ref: worksheet["!ref"] } : undefined;
}

export default function App() {
  const inputRef = useRef(null);
  const [files, setFiles] = useState([]);
  const [dragging, setDragging] = useState(false);
  const [processing, setProcessing] = useState(false);
  const [message, setMessage] = useState("");
  const [strictMode, setStrictMode] = useState(true);
  const [useOcr, setUseOcr] = useState(true);
  const [ocrLanguage, setOcrLanguage] = useState("por");
  const totalSize = useMemo(() => files.reduce((sum, item) => sum + item.file.size, 0), [files]);

  function addFiles(fileList) {
    const incoming = Array.from(fileList || []);
    const accepted = [];
    const rejected = [];
    incoming.forEach((file) => {
      const isPdf = file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf");
      if (!isPdf) rejected.push(`${file.name}: formato invalido`);
      else if (file.size > MAX_FILE_SIZE) rejected.push(`${file.name}: excede 25 MB`);
      else if (!files.some((item) => item.file.name === file.name && item.file.size === file.size)) accepted.push({ file, status: "Pronto", progress: 0, error: "" });
    });
    setFiles((current) => [...current, ...accepted]);
    setMessage(rejected.length ? rejected.join(" | ") : "");
  }

  function updateFile(index, patch) { setFiles((current) => current.map((item, i) => i === index ? { ...item, ...patch } : item)); }
  function removeFile(index) { if (!processing) setFiles((current) => current.filter((_, i) => i !== index)); }
  async function convert() {
    if (!files.length || processing) return;
    setProcessing(true);
    setMessage("");
    const workbook = XLSX.utils.book_new();
    const allRows = [], allRawRows = [], auditRows = [];
    const usedNames = new Set();

    for (let index = 0; index < files.length; index += 1) {
      const item = files[index];
      updateFile(index, { status: "Processando", progress: 2, error: "" });
      try {
        const result = await extractPdf(
          item.file,
          { useOcr, ocrLanguage },
          (progress) => updateFile(index, { progress }),
          (status) => updateFile(index, { status })
        );
        allRows.push(...result.rows);
        allRawRows.push(...result.rawRows);
        const sheetRows = result.rows.length ? result.rows : [{ Arquivo: item.file.name, Aviso: "Nenhum texto foi localizado, inclusive por OCR." }];
        const worksheet = XLSX.utils.json_to_sheet(sheetRows);
        autosize(worksheet, sheetRows);
        XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName(item.file.name.replace(/\.pdf$/i, ""), usedNames));
        auditRows.push({
          Arquivo: item.file.name, Tamanho: formatBytes(item.file.size), Paginas: result.pageCount,
          Linhas_extraidas: result.rows.length, Paginas_processadas_com_OCR: result.ocrPages,
          Paginas_sem_texto: result.emptyPages, Resultado: result.emptyPages ? "Revisar paginas sem texto" : "Concluido"
        });
        updateFile(index, { status: result.emptyPages ? "Revisar" : "Concluido", progress: 100 });
      } catch (error) {
        auditRows.push({ Arquivo: item.file.name, Tamanho: formatBytes(item.file.size), Resultado: "Erro", Detalhes: error.message });
        updateFile(index, { status: "Erro", progress: 100, error: error.message });
      }
    }

    if (strictMode) {
      const consolidated = allRows.length ? allRows : [{ Aviso: "Nenhum dado foi extraido." }];
      const sheet = XLSX.utils.json_to_sheet(consolidated); autosize(sheet, consolidated);
      XLSX.utils.book_append_sheet(workbook, sheet, safeSheetName("Dados consolidados", usedNames));
    }
    const raw = allRawRows.length ? allRawRows : [{ Aviso: "Sem texto integral disponivel." }];
    const rawSheet = XLSX.utils.json_to_sheet(raw); autosize(rawSheet, raw);
    XLSX.utils.book_append_sheet(workbook, rawSheet, safeSheetName("Texto integral", usedNames));
    const auditSheet = XLSX.utils.json_to_sheet(auditRows); autosize(auditSheet, auditRows);
    XLSX.utils.book_append_sheet(workbook, auditSheet, safeSheetName("Auditoria", usedNames));
    XLSX.writeFile(workbook, `conversao_pdf_excel_${new Date().toISOString().slice(0, 10)}.xlsx`, { compression: true });
    setProcessing(false);
    setMessage("Conversao finalizada. O Excel inclui dados extraidos, OCR, texto integral e auditoria.");
  }

  return (
    <main className="min-h-screen bg-slate-50 text-slate-900">
      <header className="border-b border-slate-200 bg-white">
        <div className="mx-auto flex max-w-6xl items-center justify-between px-6 py-5">
          <div className="flex items-center gap-3">
            <div className="rounded-2xl bg-emerald-600 p-2.5 text-white shadow-sm"><FileSpreadsheet className="h-6 w-6" /></div>
            <div><h1 className="text-xl font-bold tracking-tight">PDF para Excel com OCR</h1><p className="text-sm text-slate-500">Extracao organizada, rastreavel e pronta para conferencia</p></div>
          </div>
          <Badge variant="outline" className="gap-1.5 rounded-full border-emerald-200 bg-emerald-50 px-3 py-1 text-emerald-700"><ShieldCheck className="h-3.5 w-3.5" /> Processamento local</Badge>
        </div>
      </header>

      <section className="mx-auto grid max-w-6xl gap-6 px-6 py-8 lg:grid-cols-[1fr_340px]">
        <div className="space-y-6">
          <Card className="rounded-2xl border-slate-200 shadow-sm"><CardContent className="p-6">
            <button type="button" onClick={() => inputRef.current?.click()} onDragOver={(e) => { e.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={(e) => { e.preventDefault(); setDragging(false); addFiles(e.dataTransfer.files); }} className={`flex min-h-64 w-full flex-col items-center justify-center rounded-2xl border-2 border-dashed p-8 text-center transition ${dragging ? "border-emerald-500 bg-emerald-50" : "border-slate-300 bg-slate-50 hover:border-emerald-400 hover:bg-emerald-50/50"}`}>
              <div className="mb-4 rounded-2xl bg-white p-4 text-emerald-600 shadow-sm"><UploadCloud className="h-9 w-9" /></div>
              <h2 className="text-lg font-semibold">Arraste os PDFs para esta area</h2>
              <p className="mt-2 max-w-md text-sm leading-6 text-slate-500">Aceita PDFs com texto e documentos digitalizados. Limite de 25 MB por arquivo.</p>
              <span className="mt-5 rounded-xl bg-slate-900 px-4 py-2.5 text-sm font-semibold text-white">Selecionar arquivos</span>
            </button>
            <input ref={inputRef} type="file" accept="application/pdf,.pdf" multiple className="hidden" onChange={(e) => { addFiles(e.target.files); e.target.value = ""; }} />
          </CardContent></Card>

          {files.length > 0 && <Card className="rounded-2xl border-slate-200 shadow-sm">
            <CardHeader className="flex flex-row items-center justify-between pb-3"><CardTitle className="text-base">Arquivos adicionados</CardTitle><span className="text-sm text-slate-500">{files.length} arquivo(s) · {formatBytes(totalSize)}</span></CardHeader>
            <CardContent className="space-y-3">{files.map((item, index) => <div key={`${item.file.name}-${item.file.size}`} className="rounded-xl border border-slate-200 bg-white p-4">
              <div className="flex items-start gap-3"><div className="rounded-xl bg-red-50 p-2 text-red-600"><FileText className="h-5 w-5" /></div><div className="min-w-0 flex-1">
                <div className="flex items-center justify-between gap-3"><p className="truncate text-sm font-semibold">{item.file.name}</p><button disabled={processing} onClick={() => removeFile(index)} className="rounded-lg p-1 text-slate-400 hover:bg-slate-100 disabled:opacity-40"><X className="h-4 w-4" /></button></div>
                <div className="mt-1 flex items-center justify-between text-xs text-slate-500"><span>{formatBytes(item.file.size)}</span><span className={item.status === "Erro" ? "text-red-600" : item.status === "Revisar" ? "text-amber-600" : item.status === "Concluido" ? "text-emerald-600" : ""}>{item.status}</span></div>
                {(processing || item.progress > 0) && <Progress value={item.progress} className="mt-3 h-1.5" />}{item.error && <p className="mt-2 text-xs text-red-600">{item.error}</p>}
              </div></div>
            </div>)}</CardContent>
          </Card>}

          {message && <div className={`flex items-start gap-3 rounded-2xl border p-4 text-sm ${message.startsWith("Conversao") ? "border-emerald-200 bg-emerald-50 text-emerald-800" : "border-amber-200 bg-amber-50 text-amber-800"}`}>{message.startsWith("Conversao") ? <CheckCircle2 className="h-5 w-5 shrink-0" /> : <AlertTriangle className="h-5 w-5 shrink-0" />}<span>{message}</span></div>}
        </div>

        <aside className="space-y-6">
          <Card className="rounded-2xl border-slate-200 shadow-sm"><CardHeader><CardTitle className="flex items-center gap-2 text-base"><Settings2 className="h-4 w-4" /> Configuracao</CardTitle></CardHeader><CardContent className="space-y-5">
            <label className="flex cursor-pointer items-start gap-3"><input type="checkbox" checked={useOcr} onChange={(e) => setUseOcr(e.target.checked)} className="mt-1 h-4 w-4 accent-emerald-600" /><span><span className="flex items-center gap-2 text-sm font-semibold"><ScanText className="h-4 w-4" /> Ativar OCR</span><span className="mt-1 block text-xs leading-5 text-slate-500">Aplica OCR automaticamente apenas nas paginas sem texto pesquisavel.</span></span></label>
            {useOcr && <label className="block"><span className="mb-2 block text-sm font-semibold">Idioma do documento</span><select value={ocrLanguage} onChange={(e) => setOcrLanguage(e.target.value)} className="h-11 w-full rounded-xl border border-slate-300 bg-white px-3 text-sm"><option value="por">Portugues</option><option value="eng">Ingles</option><option value="spa">Espanhol</option><option value="por+eng">Portugues + Ingles</option><option value="por+spa">Portugues + Espanhol</option></select></label>}
            <label className="flex cursor-pointer items-start gap-3"><input type="checkbox" checked={strictMode} onChange={(e) => setStrictMode(e.target.checked)} className="mt-1 h-4 w-4 accent-emerald-600" /><span><span className="block text-sm font-semibold">Planilha consolidada</span><span className="mt-1 block text-xs leading-5 text-slate-500">Reune dados e preserva arquivo, pagina, linha e origem da extracao.</span></span></label>
            <div className="rounded-xl bg-slate-50 p-4 text-xs leading-5 text-slate-600">O OCR roda no navegador. Na primeira utilizacao, os arquivos do idioma escolhido precisam ser carregados.</div>
            <Button onClick={convert} disabled={!files.length || processing} className="h-12 w-full rounded-xl bg-emerald-600 font-semibold hover:bg-emerald-700">{processing ? <><Loader2 className="mr-2 h-4 w-4 animate-spin" /> Convertendo...</> : <><Download className="mr-2 h-4 w-4" /> Converter e baixar</>}</Button>
          </CardContent></Card>

          <Card className="rounded-2xl border-slate-200 bg-slate-900 text-white shadow-sm"><CardContent className="p-5"><h3 className="font-semibold">Preservacao e conferencia</h3><ul className="mt-4 space-y-3 text-sm text-slate-300"><li className="flex gap-2"><CheckCircle2 className="mt-0.5 h-4 w-4 text-emerald-400" /> OCR somente quando necessario</li><li className="flex gap-2"><CheckCircle2 className="mt-0.5 h-4 w-4 text-emerald-400" /> Origem marcada como PDF ou OCR</li><li className="flex gap-2"><CheckCircle2 className="mt-0.5 h-4 w-4 text-emerald-400" /> Texto integral e auditoria no Excel</li></ul></CardContent></Card>
          <p className="px-1 text-xs leading-5 text-slate-500"><strong>Importante:</strong> OCR reduz perdas em documentos digitalizados, mas campos manuscritos, imagens ruins e tabelas complexas ainda devem ser conferidos.</p>
        </aside>
      </section>
    </main>
  );
}
