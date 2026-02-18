import { useId, useRef } from "react";
import { Upload } from "lucide-react";

import { cn } from "../lib/utils";
import { Label } from "./ui/label";

type Props = {
  label: string;
  accept?: string;
  required?: boolean;
  disabled?: boolean;
  name?: string;
  hint?: string;
  error?: string;
  value?: File | null;
  onChange: (file: File | null) => void;
};

export default function FileField({
  label,
  accept,
  required,
  disabled,
  name,
  hint,
  error,
  value,
  onChange,
}: Props) {
  const inputId = useId();
  const inputRef = useRef<HTMLInputElement>(null);

  const openPicker = () => {
    if (disabled) return;
    if (!inputRef.current) return;
    inputRef.current.value = "";
    inputRef.current.click();
  };

  return (
    <div className="space-y-2">
      <Label htmlFor={inputId}>{label}</Label>
      <div
        className={cn(
          "flex h-10 items-center gap-3 rounded-md border border-input bg-background px-3 text-sm text-foreground",
          disabled ? "cursor-not-allowed opacity-60" : "cursor-pointer",
          error ? "border-destructive/70" : ""
        )}
        onClick={openPicker}
      >
        <button
          type="button"
          disabled={disabled}
          onClick={(event) => {
            event.preventDefault();
            openPicker();
          }}
          className="inline-flex shrink-0 items-center gap-2 rounded-md bg-primary/12 px-3 py-1.5 font-medium text-primary transition-colors hover:bg-primary/18"
        >
          <Upload className="h-4 w-4" />
          <span>Seleccionar archivo</span>
        </button>

        <span className="truncate text-foreground/90">{value?.name ?? "Ningún archivo seleccionado"}</span>

        <input
          ref={inputRef}
          id={inputId}
          name={name}
          type="file"
          accept={accept}
          aria-required={required}
          disabled={disabled}
          className="sr-only"
          onChange={(e) => onChange(e.target.files?.[0] ?? null)}
        />
      </div>
      {value?.name ? <p className="text-xs text-muted-foreground">Seleccionado: {value.name}</p> : null}
      {hint ? <p className="text-xs text-muted-foreground">{hint}</p> : null}
      {error ? <p className="text-xs text-destructive">{error}</p> : null}
    </div>
  );
}
