import { useEffect, useMemo, useState } from "react";
import type { FormEvent } from "react";
import FileField from "../components/FileField";
import ProgressBar from "../components/ProgressBar";
import { StatusLine } from "../components/StatusLine";
import { downloadBlob, xhrPostWithProgress } from "../api/client";

type MasterInfo = {
  exists: boolean;
  name?: string | null;
  debug_path?: string | null;
};

function buildDownloadName(source: File | null, hdrDate: string) {
  if (source?.name) {
    return source.name;
  }
  if (hdrDate) {
    const [y, m, d] = hdrDate.split("-");
    if (y && m && d) {
      return `COBRANZA_${d}-${m}-${y.slice(-2)}.xls`;
    }
  }
  return "COBRANZA.xls";
}

export default function MergePage() {
  const [source, setSource] = useState<File | null>(null);
  const [master, setMaster] = useState<File | null>(null);
  const [hdrDate, setHdrDate] = useState("");
  const [useDefault, setUseDefault] = useState(true);
  const [status, setStatus] = useState("");
  const [statusOk, setStatusOk] = useState(true);
  const [progress, setProgress] = useState(0);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [masterInfo, setMasterInfo] = useState<MasterInfo | null>(null);

  useEffect(() => {
    let active = true;
    fetch("/master/default-info")
      .then((resp) => resp.json())
      .then((data: MasterInfo) => {
        if (active) setMasterInfo(data);
      })
      .catch(() => {
        if (active) setMasterInfo({ exists: false });
      });
    return () => {
      active = false;
    };
  }, []);

  const defaultMasterHint = useMemo(() => {
    if (!masterInfo) return "Verificando maestro por defecto...";
    if (!masterInfo.exists) {
      const extra = masterInfo.debug_path
        ? ` (Ruta: ${masterInfo.debug_path})`
        : "";
      return `No se encontro el maestro alojado.${extra}`;
    }
    return `Usar ${masterInfo.name ?? "maestro por defecto"}`;
  }, [masterInfo]);

  const handleToggleDefault = (checked: boolean) => {
    setUseDefault(checked);
    if (checked) setMaster(null);
  };

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    if (!source) {
      setStatusOk(false);
      setStatus("Carga el archivo de tecnicos (.XLS).");
      return;
    }
    if (!useDefault && !master) {
      setStatusOk(false);
      setStatus('Carga el maestro o activa "Usar maestro por defecto".');
      return;
    }

    setIsSubmitting(true);
    setStatus("Procesando, por favor espera...");
    setStatusOk(true);
    setProgress(1);

    try {
      const fd = new FormData();
      fd.append("source", source);
      if (!useDefault && master) fd.append("master", master);
      if (hdrDate) fd.append("hdr_date", hdrDate);
      fd.append("use_default_master", useDefault ? "1" : "0");

      const blob = await xhrPostWithProgress(
        "/merge",
        fd,
        (pct) => setProgress(Math.min(70, Math.round(pct))),
        (pct) => setProgress(Math.round(pct))
      );

      const filename = buildDownloadName(source, hdrDate);
      downloadBlob(blob, filename);
      setStatus(`Listo. Se descargo ${filename}.`);
      setStatusOk(true);
    } catch (error) {
      setStatusOk(false);
      setStatus(
        error instanceof Error ? error.message : "Error desconocido al procesar."
      );
    } finally {
      setProgress(0);
      setIsSubmitting(false);
    }
  };

  return (
    <div className="container py-4">
      <h1 className="h4 mb-2 text-body-emphasis">Cobranza XLS - Copiar a Maestro</h1>
      <p className="text-muted mb-4">
        Requiere Windows con Microsoft Excel instalado. Extension .XLS antigua.
      </p>

      {/* Wrapper flex propio: evita depender de .row que Adminator modifica */}
      <div className="d-flex justify-content-center">
        <div className="card shadow-sm app-panel w-100">
          <div className="card-body">
              <form onSubmit={handleSubmit}>

            <FileField
              label="Origen (.XLS tecnicos)"
              accept=".xls"
              required
              onChange={setSource}
              hint="Archivo exportado por el equipo de cobranzas."
            />

            <div className="form-check form-switch mb-2">
              <input
                id="useDefaultMaster"
                className="form-check-input"
                type="checkbox"
                checked={useDefault}
                onChange={(e) => handleToggleDefault(e.target.checked)}
              />
              <label className="form-check-label" htmlFor="useDefaultMaster">
                Usar maestro por defecto
              </label>
              <div className="form-text">{defaultMasterHint}</div>
            </div>

            {!useDefault && (
              <FileField
                label="Maestro (.XLS)"
                accept=".xls"
                required
                onChange={setMaster}
                hint="Archivo maestro actualizado."
              />
            )}

            <div className="mb-3">
              <label htmlFor="hdrDate" className="form-label">
                Fecha (opcional)
              </label>
              <input
                id="hdrDate"
                type="date"
                className="form-control"
                value={hdrDate}
                onChange={(e) => setHdrDate(e.target.value)}
              />
              <div className="form-text">
                Se usara para actualizar encabezados y hojas SUR/NORTE.
              </div>
            </div>

            <button
                  className="btn btn-primary"
                  type="submit"
                  disabled={isSubmitting}
                >
                  {isSubmitting ? "Procesando..." : "Copiar a Maestro"}
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
