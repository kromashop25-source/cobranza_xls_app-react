import { useCallback, useState } from "react";
import { motion } from "framer-motion";
import { FileSpreadsheet, FileText } from "lucide-react";

import GlobalFileDrop from "./components/GlobalFileDrop";
import MergePage from "./pages/MergePage";
import PdfExportPage from "./pages/PdfExportPage";
import ThemeToggle from "./components/ThemeToggle";
import { Button } from "./components/ui/button";

type View = "merge" | "pdf";
type DropHandler = (file: File) => void;

const tabItems: Array<{ id: View; label: string; icon: typeof FileSpreadsheet }> = [
  { id: "merge", label: "Copiar a Maestro", icon: FileSpreadsheet },
  { id: "pdf", label: "Exportar PDFs", icon: FileText },
];

export default function App() {
  const [view, setView] = useState<View>("merge");
  const [mergeDropHandler, setMergeDropHandler] = useState<DropHandler | null>(null);
  const [pdfDropHandler, setPdfDropHandler] = useState<DropHandler | null>(null);

  const registerMergeDropHandler = useCallback((handler: DropHandler | null) => {
    setMergeDropHandler(() => handler);
  }, []);

  const registerPdfDropHandler = useCallback((handler: DropHandler | null) => {
    setPdfDropHandler(() => handler);
  }, []);

  const activeDropHandler = view === "merge" ? mergeDropHandler : pdfDropHandler;

  const handleGlobalFileDrop = useCallback(
    (file: File) => {
      activeDropHandler?.(file);
    },
    [activeDropHandler]
  );

  const dropTitle =
    view === "merge"
      ? "Suelta tu archivo XLS para Copiar a Maestro"
      : "Suelta tu archivo XLS para Exportar PDFs";
  const dropDescription =
    view === "merge"
      ? "Puedes soltar el archivo en cualquier parte de la ventana."
      : "El archivo se asignara automaticamente al campo XLS formateado.";

  return (
    <div className="min-h-screen bg-[radial-gradient(circle_at_top,_hsl(var(--primary)/0.12)_0%,_hsl(var(--background))_38%)]">
      <GlobalFileDrop
        enabled={Boolean(activeDropHandler)}
        title={dropTitle}
        description={dropDescription}
        onFileDrop={handleGlobalFileDrop}
      />

      <header className="sticky top-0 z-30 border-b border-border/80 bg-background/85 backdrop-blur">
        <div className="mx-auto flex w-full max-w-6xl flex-col gap-3 px-4 py-4 md:flex-row md:items-center md:justify-between">
          <div>
            <p className="text-lg font-semibold tracking-tight text-foreground">Cobranza XLS App</p>
            <p className="text-sm text-muted-foreground">UI moderna React para flujo XLS y exportacion PDF</p>
          </div>

          <div className="flex items-center gap-2">
            <div className="grid grid-cols-2 gap-2 rounded-lg border border-border bg-card/70 p-1">
              {tabItems.map(({ id, label, icon: Icon }) => (
                <Button
                  key={id}
                  size="sm"
                  variant={view === id ? "default" : "ghost"}
                  className="gap-2"
                  onClick={() => setView(id)}
                  type="button"
                >
                  <Icon className="h-4 w-4" />
                  {label}
                </Button>
              ))}
            </div>
            <ThemeToggle />
          </div>
        </div>
      </header>

      <motion.main
        key={view}
        initial={{ opacity: 0, y: 10 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.25, ease: "easeOut" }}
        className="mx-auto w-full max-w-6xl px-4 py-6 md:py-8"
      >
        {view === "merge" ? (
          <MergePage registerDropHandler={registerMergeDropHandler} />
        ) : (
          <PdfExportPage registerDropHandler={registerPdfDropHandler} />
        )}
      </motion.main>
    </div>
  );
}
