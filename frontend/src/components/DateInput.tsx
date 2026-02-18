import { useRef } from "react";
import { CalendarDays } from "lucide-react";

import { Input, type InputProps } from "./ui/input";
import { cn } from "../lib/utils";

type DateInputProps = Omit<InputProps, "type">;

export default function DateInput({ className, ...props }: DateInputProps) {
  const inputRef = useRef<HTMLInputElement>(null);

  const openDatePicker = () => {
    const element = inputRef.current;
    if (!element || element.disabled) return;

    if ("showPicker" in element && typeof element.showPicker === "function") {
      element.showPicker();
      return;
    }

    element.focus();
    element.click();
  };

  return (
    <div className="relative">
      <Input ref={inputRef} type="date" className={cn("pr-11", className)} {...props} />
      <button
        type="button"
        className="absolute right-2 top-1/2 inline-flex h-7 w-7 -translate-y-1/2 items-center justify-center rounded-md text-muted-foreground transition-colors hover:bg-muted hover:text-foreground"
        onMouseDown={(event) => event.preventDefault()}
        onClick={openDatePicker}
        aria-label="Abrir calendario"
      >
        <CalendarDays className="h-4 w-4" />
      </button>
    </div>
  );
}
