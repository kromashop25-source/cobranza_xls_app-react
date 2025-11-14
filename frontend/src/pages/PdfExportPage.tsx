import { useState } from "react";
import type { FormEvent } from "react";
import FileField from "../components/FileField";
import ProgressBar from "../components/ProgressBar";
import { StatusLine } from "../components/StatusLine";
import { downloadBlob, xhrPostWithProgress } from "../api/client";

function buildZipName(file: File | null) {
  if (!file?.name) return "PDFS_COBRANZA.zip";
  const stem = file.name.replace(/\.[^.]$/i, "");
  return `PDFS_${stem}.zip`;
}

export default function PdfExportPage() {
  const [excel, setExcel] = useState<File | null>(null);
  const [hojaBase, setHojaBase] = useState("");
  const [status, setStatus] = useState("");
  const [statusOk, setStatusOk] = useState(true);
  const [progress, setProgress] = useState(0);
  const [isSubmitting, setIsSubmitting] = useState(false);

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!excel) {
      setStatusOk(false);
      setStatus("Adjunta el XLS previamente generado en /merge.");
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

      {/* Wrapper flex propio: centrado consistente con ventana maximizada */}
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
