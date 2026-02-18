import { AlertCircle, CheckCircle2 } from "lucide-react";

import { cn } from "../lib/utils";

export function StatusLine({
  text,
  ok = true,
}: {
  text: string;
  ok?: boolean;
}) {
  if (!text) return null;

  return (
    <div
      className={cn(
        "flex items-center gap-2 text-sm",
        ok ? "text-emerald-600 dark:text-emerald-400" : "text-destructive"
      )}
      aria-live="polite"
      role="status"
    >
      {ok ? <CheckCircle2 className="h-4 w-4" /> : <AlertCircle className="h-4 w-4" />}
      <span>{text}</span>
    </div>
  );
}

