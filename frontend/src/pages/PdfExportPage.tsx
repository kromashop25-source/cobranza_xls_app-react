import { useEffect, useState } from "react";
import type { FormEvent } from "react";
import FileField from "../components/FileField";
import ProgressBar from "../components/ProgressBar";
import { StatusLine } from "../components/StatusLine";
import { downloadBlob, xhrPostWithProgress } from "../api/client";

type BlockItem = {
  id: string;
  name: string;
  sheet: string;
  include: boolean;
};

const DEFAULT_CONSOLIDATED_ORDER = [
  "SALDOS COBRANZA",
  "MANUEL CARRASCO",
  "PITER HUAYTA",
  "LEONEL MEZA",
  "CA\u00d1ETE",
  "BEATRIZ ROJAS",
  "LURIN",
  "MANCHAY",
  "CIUDAD",
  "UNICACHI",
  "NORTE - ROSA",
  "CAQUETA (ROSA)",
  "SURCO (OSCAR)",
  "SURQUILLO (OSCAR)",
  "SAN LUIS (OSCAR)",
  "RAUL ARROYO",
];

const ORDER_ALIASES: Record<string, string[]> = {
  "SALDOS COBRANZA": ["SALDOS COBRANZA"],
  "MANUEL CARRASCO": ["MANUEL CARRASCO"],
  "PITER HUAYTA": ["PITER HUAYTA"],
  "LEONEL MEZA": ["LEONEL MEZA"],
  "CA\u00d1ETE": ["CA\u00d1ETE", "CA\u00d1ETE - (MANUEL)", "CA\u00d1ETE (MANUEL)"],
  "BEATRIZ ROJAS": ["BEATRIZ ROJAS"],
  "LURIN": ["LURIN", "LURIN - (ROSA)"],
  "MANCHAY": ["MANCHAY", "MANCHAY - (ROSA)"],
  "CIUDAD": ["CIUDAD", "CIUDAD - (ROSA)"],
  "UNICACHI": ["UNICACHI", "UNICACHI SUR - (ROSA)", "UNICACHI SUR (ROSA)"],
  "NORTE - ROSA": ["NORTE - ROSA", "NORTE ROSA", "NORTE-ROSA"],
  "CAQUETA (ROSA)": ["CAQUETA (ROSA)", "CAQUETA - (ROSA)", "CAQUETA ROSA"],
  "SURCO (OSCAR)": ["SURCO (OSCAR)", "SURCO - (OSCAR)", "SURCO OSCAR"],
  "SURQUILLO (OSCAR)": [
    "SURQUILLO (OSCAR)",
    "SURQ/SURCO - (OSCAR)",
    "SURQ/SURCO (OSCAR)",
  ],
  "SAN LUIS (OSCAR)": ["SAN LUIS (OSCAR)", "SAN LUIS - (OSCAR)", "SAN LUIS OSCAR"],
  "RAUL ARROYO": ["RAUL ARROYO"],
};

function buildZipName(file: File | null) {
  if (!file?.name) return "PDFS_COBRANZA.zip";
  const stem = file.name.replace(/\.[^.]+$/i, "");
  return `PDFS_${stem}.zip`;
}

function stripLeadingCode(name: string) {
  return name.replace(/^\s*\d+\s+/, "").trim();
}

function normalizeOrderName(name: string) {
  return stripLeadingCode(name)
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function applyDefaultOrder(blocks: BlockItem[]) {
  const orderMap = new Map<string, number>();
  DEFAULT_CONSOLIDATED_ORDER.forEach((name, index) => {
    const aliases = ORDER_ALIASES[name] ?? [name];
    aliases.forEach((alias) => {
      orderMap.set(normalizeOrderName(alias), index);
    });
  });

  return blocks
    .map((block, index) => ({ ...block, _idx: index }))
    .sort((a, b) => {
      const aKey = normalizeOrderName(a.name);
      const bKey = normalizeOrderName(b.name);
      const aOrder = orderMap.has(aKey)
        ? orderMap.get(aKey)!
        : Number.MAX_SAFE_INTEGER;
      const bOrder = orderMap.has(bKey)
        ? orderMap.get(bKey)!
        : Number.MAX_SAFE_INTEGER;
      if (aOrder !== bOrder) return aOrder - bOrder;
      return a._idx - b._idx;
    })
    .map(({ _idx, ...rest }) => rest);
}

function parseDateFromFilename(name: string) {
  const match = name.match(/(\d{2})[-_](\d{2})[-_](\d{2,4})/);
  if (!match) return "";
  const day = match[1];
  const month = match[2];
  const rawYear = match[3];
  const year = rawYear.length === 2 ? `20${rawYear}` : rawYear;
  return `${year}-${month}-${day}`;
}

export default function PdfExportPage() {
  const [excel, setExcel] = useState<File | null>(null);
  const [hojaBase, setHojaBase] = useState("");
  const [pdfDate, setPdfDate] = useState("");
  const [blocks, setBlocks] = useState<BlockItem[]>([]);
  const [analysisStatus, setAnalysisStatus] = useState("");
  const [analysisOk, setAnalysisOk] = useState(true);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [status, setStatus] = useState("");
  const [statusOk, setStatusOk] = useState(true);
  const [progress, setProgress] = useState(0);
  const [isSubmitting, setIsSubmitting] = useState(false);

  const analyzeFile = async (file: File | null, hoja: string) => {
    if (!file) return;
    setIsAnalyzing(true);
    setAnalysisOk(true);
    setAnalysisStatus("Analizando XLS...");

    try {
      const fd = new FormData();
      fd.append("excel", file);
      if (hoja.trim()) fd.append("hoja_base", hoja.trim());

      const resp = await fetch("/pdf/preview-upload", {
        method: "POST",
        body: fd,
      });

      if (!resp.ok) {
        const data = await resp.json().catch(() => null);
        throw new Error(data?.detail ?? "No se pudo analizar el XLS.");
      }

      const data = (await resp.json()) as {
        blocks: { id: string; name: string; sheet: string }[];
        count: number;
      };

      const items = data.blocks.map((block) => ({
        id: block.id,
        name: block.name,
        sheet: block.sheet,
        include: true,
      }));

      const ordered = applyDefaultOrder(items);
      setBlocks(ordered);
      setAnalysisStatus(`Detectados ${data.count} vendedores.`);
      setAnalysisOk(true);
    } catch (error) {
      setAnalysisOk(false);
      setAnalysisStatus(
        error instanceof Error ? error.message : "No se pudo analizar el XLS."
      );
      setBlocks([]);
    } finally {
      setIsAnalyzing(false);
    }
  };

  useEffect(() => {
    if (!excel) {
      setBlocks([]);
      return;
    }
    const inferredDate = parseDateFromFilename(excel.name);
    if (inferredDate) {
      setPdfDate(inferredDate);
    }
    void analyzeFile(excel, hojaBase);
  }, [excel]);

  const handleMove = (index: number, direction: -1 | 1) => {
    setBlocks((prev) => {
      const next = [...prev];
      const target = index + direction;
      if (target < 0 || target >= next.length) return prev;
      [next[index], next[target]] = [next[target], next[index]];
      return next;
    });
  };

  const handleToggleInclude = (id: string) => {
    setBlocks((prev) =>
      prev.map((item) =>
        item.id === id ? { ...item, include: !item.include } : item
      )
    );
  };

  const handleApplyDefaultOrder = () => {
    setBlocks((prev) => applyDefaultOrder(prev));
  };

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!excel) {
      setStatusOk(false);
      setStatus("Adjunta el XLS previamente generado en /merge.");
      return;
    }
    if (!pdfDate) {
      setStatusOk(false);
      setStatus("Selecciona la fecha usada al generar el Excel.");
      return;
    }

    setIsSubmitting(true);
    setStatus("Generando PDFs...");
    setStatusOk(true);
    setProgress(1);

    try {
      const fd = new FormData();
      fd.append("excel", excel);
      if (hojaBase.trim()) fd.append("hoja_base", hojaBase.trim());
      if (pdfDate) fd.append("pdf_date", pdfDate);
      if (blocks.length) {
        fd.append("orden", JSON.stringify(blocks.map((block) => block.id)));
        fd.append(
          "excluir",
          JSON.stringify(blocks.filter((block) => !block.include).map((block) => block.id))
        );
      }

      const blob = await xhrPostWithProgress(
        "/pdf/export-upload",
        fd,
        (pct) => setProgress(Math.min(70, Math.round(pct))),
        (pct) => setProgress(Math.round(pct))
      );
      const zipName = buildZipName(excel);
      downloadBlob(blob, zipName);
      setStatus(`ZIP descargado (${zipName}).`);
      setStatusOk(true);
    } catch (error) {
      setStatusOk(false);
      setStatus(
        error instanceof Error ? error.message : "No se pudo generar el ZIP."
      );
    } finally {
      setProgress(0);
      setIsSubmitting(false);
    }
  };

  return (
    <div className="container py-4">
      <h1 className="h4 mb-2 text-body-emphasis">Exportar PDFs desde XLS formateado</h1>
      <p className="text-muted mb-4">
        Usa el archivo generado en "Copiar a Maestro". El backend devolvera un
        ZIP con los PDFs detectados.
      </p>

      <div className="d-flex justify-content-center">
        <div className="card shadow-sm app-panel w-100">
          <div className="card-body">
            <form onSubmit={handleSubmit}>
              <FileField
                label="XLS formateado"
                accept=".xls"
                required
                onChange={setExcel}
                hint="Es el resultado del paso /merge."
              />

              <div className="mb-3">
                <label htmlFor="hojaBase" className="form-label">
                  Hoja base (opcional)
                </label>
                <input
                  id="hojaBase"
                  type="text"
                  className="form-control"
                  placeholder='Ej: "OFICINA (VES)"'
                  value={hojaBase}
                  onChange={(e) => setHojaBase(e.target.value)}
                />
                <div className="form-text">
                  Solo si deseas limitar a una hoja especifica.
                </div>
              </div>

              <div className="mb-3">
                <label htmlFor="pdfDate" className="form-label">
                  Fecha para nombres de PDF
                </label>
                <input
                  id="pdfDate"
                  type="date"
                  className="form-control"
                  value={pdfDate}
                  required
                  onChange={(e) => setPdfDate(e.target.value)}
                />
                <div className="form-text">
                  Usa la misma fecha seleccionada al generar el Excel.
                </div>
              </div>

              <div className="d-flex flex-wrap gap-2 mb-2">
                <button
                  className="btn btn-outline-secondary"
                  type="button"
                  disabled={!excel || isAnalyzing}
                  onClick={() => analyzeFile(excel, hojaBase)}
                >
                  {isAnalyzing ? "Analizando..." : "Analizar XLS"}
                </button>
                <button
                  className="btn btn-outline-secondary"
                  type="button"
                  disabled={!blocks.length}
                  onClick={handleApplyDefaultOrder}
                >
                  Orden sugerido
                </button>
              </div>

              {analysisStatus && (
                <div className="mb-3">
                  <StatusLine text={analysisStatus} ok={analysisOk} />
                </div>
              )}

              {blocks.length > 0 && (
                <div className="mb-3">
                  <div className="d-flex justify-content-between align-items-center mb-2">
                    <div className="fw-semibold">Orden para consolidado</div>
                    <div className="text-muted small">
                      {blocks.filter((b) => !b.include).length} excluidos
                    </div>
                  </div>

                  <div className="list-group">
                    {blocks.map((block, index) => (
                      <div
                        key={block.id}
                        className="list-group-item d-flex align-items-center gap-2"
                      >
                        <span className="badge bg-primary">{index + 1}</span>
                        <div className="flex-grow-1">
                          <div className="fw-semibold">{block.name}</div>
                          <div className="small text-muted">{block.sheet}</div>
                        </div>
                        <div className="form-check form-switch">
                          <input
                            className="form-check-input"
                            type="checkbox"
                            checked={block.include}
                            onChange={() => handleToggleInclude(block.id)}
                          />
                          <label className="form-check-label small">
                            Incluir
                          </label>
                        </div>
                        <div className="btn-group" role="group">
                          <button
                            type="button"
                            className="btn btn-sm btn-outline-secondary"
                            onClick={() => handleMove(index, -1)}
                            disabled={index === 0}
                          >
                            Subir
                          </button>
                          <button
                            type="button"
                            className="btn btn-sm btn-outline-secondary"
                            onClick={() => handleMove(index, 1)}
                            disabled={index === blocks.length - 1}
                          >
                            Bajar
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              <button
                className="btn btn-outline-primary"
                type="submit"
                disabled={isSubmitting}
              >
                {isSubmitting ? "Generando..." : "Generar PDFs"}
              </button>
            </form>

            <div className="mt-3">
              <StatusLine text={status} ok={statusOk} />
              <ProgressBar value={progress} />
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
