import { useState } from "react";
import MergePage from "./pages/MergePage";
import PdfExportPage from "./pages/PdfExportPage";
import ThemeToggle from "./components/ThemeToggle";

type View = "merge" | "pdf";

export default function App() {
  const [view, setView] = useState<View>("merge");

  return (
    <div className="app-shell">
      <header className="border-bottom sticky-top bg-body">
        <div className="container py-3 d-flex flex-column flex-md-row align-items-center gap-2">
          <div className="me-auto">
            <span className="fw-semibold">Cobranza XLS App</span>
            <span className="text-muted ms-2 small">React UI</span>
          </div>
          <div className="btn-group">
            <button
              className={`btn btn-sm ${
                view === "merge" ? "btn-primary" : "btn-outline-primary"
              }`}
              type="button"
              onClick={() => setView("merge")}
            >
              Copiar a Maestro
            </button>
            <button
              className={`btn btn-sm ${
                view === "pdf" ? "btn-primary" : "btn-outline-primary"
              }`}
              type="button"
              onClick={() => setView("pdf")}
            >
              Exportar PDFs
            </button>
          </div>
          <div className="ms-2">
            <ThemeToggle />
          </div>
        </div>
      </header>
      <main>{view === "merge" ? <MergePage /> : <PdfExportPage />}</main>
    </div>
  );
}
