import { useEffect, useRef, useState } from "react";
import { AnimatePresence, motion } from "framer-motion";
import { CheckCircle2, FileSpreadsheet, Sparkles } from "lucide-react";

type GlobalFileDropProps = {
  enabled: boolean;
  title: string;
  description: string;
  onFileDrop: (file: File) => void;
};

function hasFiles(event: DragEvent) {
  return Array.from(event.dataTransfer?.types ?? []).includes("Files");
}

export default function GlobalFileDrop({
  enabled,
  title,
  description,
  onFileDrop,
}: GlobalFileDropProps) {
  const [isDragging, setIsDragging] = useState(false);
  const [showDroppedState, setShowDroppedState] = useState(false);
  const dragDepth = useRef(0);
  const dropHandlerRef = useRef(onFileDrop);

  useEffect(() => {
    dropHandlerRef.current = onFileDrop;
  }, [onFileDrop]);

  useEffect(() => {
    if (!enabled) {
      setIsDragging(false);
      setShowDroppedState(false);
      dragDepth.current = 0;
      return;
    }

    const onDragEnter = (event: DragEvent) => {
      if (!hasFiles(event)) return;
      event.preventDefault();
      dragDepth.current += 1;
      setIsDragging(true);
    };

    const onDragOver = (event: DragEvent) => {
      if (!hasFiles(event)) return;
      event.preventDefault();
      if (event.dataTransfer) {
        event.dataTransfer.dropEffect = "copy";
      }
    };

    const onDragLeave = (event: DragEvent) => {
      if (!hasFiles(event)) return;
      event.preventDefault();
      dragDepth.current = Math.max(0, dragDepth.current - 1);
      if (dragDepth.current === 0) {
        setIsDragging(false);
      }
    };

    const onDrop = (event: DragEvent) => {
      if (!hasFiles(event)) return;
      event.preventDefault();
      dragDepth.current = 0;
      setIsDragging(false);

      const file = event.dataTransfer?.files?.[0];
      if (!file) return;

      dropHandlerRef.current(file);
      setShowDroppedState(true);
      window.setTimeout(() => setShowDroppedState(false), 520);
    };

    window.addEventListener("dragenter", onDragEnter);
    window.addEventListener("dragover", onDragOver);
    window.addEventListener("dragleave", onDragLeave);
    window.addEventListener("drop", onDrop);

    return () => {
      window.removeEventListener("dragenter", onDragEnter);
      window.removeEventListener("dragover", onDragOver);
      window.removeEventListener("dragleave", onDragLeave);
      window.removeEventListener("drop", onDrop);
    };
  }, [enabled]);

  const visible = enabled && (isDragging || showDroppedState);

  return (
    <AnimatePresence>
      {visible ? (
        <motion.div
          className="pointer-events-none fixed inset-0 z-[90] flex items-center justify-center bg-[hsl(var(--background)/0.72)] backdrop-blur-sm"
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          transition={{ duration: 0.18 }}
        >
          <motion.div
            initial={{ opacity: 0, y: 16, scale: 0.98 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: 8, scale: 0.98 }}
            className="relative w-[min(92vw,38rem)] overflow-hidden rounded-2xl border border-primary/45 bg-card/95 p-8 shadow-[0_20px_65px_-20px_hsl(var(--primary)/0.6)]"
          >
            <motion.div
              className="absolute -right-10 -top-12 h-36 w-36 rounded-full bg-primary/20 blur-2xl"
              animate={{ scale: [1, 1.12, 1] }}
              transition={{ duration: 1.2, repeat: Number.POSITIVE_INFINITY, ease: "easeInOut" }}
            />

            <motion.div
              className="absolute -left-8 -top-6 rounded-xl border border-primary/40 bg-card/80 p-2"
              animate={{ y: [0, -6, 0], rotate: [0, -4, 0] }}
              transition={{ duration: 1.4, repeat: Number.POSITIVE_INFINITY, ease: "easeInOut" }}
            >
              <FileSpreadsheet className="h-5 w-5 text-primary" />
            </motion.div>

            <motion.div
              className="absolute -right-6 top-10 rounded-xl border border-border bg-background/80 p-2"
              animate={{ y: [0, 7, 0], rotate: [0, 3, 0] }}
              transition={{ duration: 1.5, repeat: Number.POSITIVE_INFINITY, ease: "easeInOut", delay: 0.1 }}
            >
              <Sparkles className="h-4 w-4 text-primary" />
            </motion.div>

            <div className="relative flex flex-col items-center gap-3 rounded-xl border border-dashed border-primary/50 bg-background/60 p-8 text-center">
              <motion.div
                animate={isDragging ? { scale: [1, 1.08, 1] } : { scale: 1 }}
                transition={{ duration: 0.65, repeat: isDragging ? Number.POSITIVE_INFINITY : 0 }}
                className="rounded-2xl border border-primary/45 bg-primary/15 p-3"
              >
                {showDroppedState ? (
                  <CheckCircle2 className="h-8 w-8 text-emerald-500" />
                ) : (
                  <FileSpreadsheet className="h-8 w-8 text-primary" />
                )}
              </motion.div>
              <p className="text-lg font-semibold text-foreground">
                {showDroppedState ? "Archivo agregado" : title}
              </p>
              <p className="max-w-md text-sm text-muted-foreground">
                {showDroppedState ? "Se aplico el archivo al formulario actual." : description}
              </p>
            </div>
          </motion.div>
        </motion.div>
      ) : null}
    </AnimatePresence>
  );
}
