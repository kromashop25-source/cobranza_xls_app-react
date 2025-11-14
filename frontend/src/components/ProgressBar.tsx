export default function ProgressBar({ value }: { value: number }) {
  if (value <= 0 || value >= 100) return null;
  return (
    <div
      className="progress my-2"
      role="progressbar"
      aria-valuemin={0}
      aria-valuemax={100}
      aria-valuenow={value}
    >
      <div className="progress-bar" style={{ width: `${value}%` }} />
    </div>
  );
}
