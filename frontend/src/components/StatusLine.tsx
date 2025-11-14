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
      className={ok ? "text-success" : "text-danger"}
      aria-live="polite"
      role="status"
    >
      {text}
    </div>
  );
}
