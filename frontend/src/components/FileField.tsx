import { useId } from "react";

type Props = {
  label: string;
  accept?: string;
  required?: boolean;
  disabled?: boolean;
  name?: string;
  hint?: string;
  onChange: (file: File | null) => void;
};

export default function FileField({
  label,
  accept,
  required,
  disabled,
  name,
  hint,
  onChange,
}: Props) {
  const inputId = useId();
  return (
    <div className="mb-3">
      <label className="form-label" htmlFor={inputId}>
        {label}
      </label>
      <input
        id={inputId}
        name={name}
        type="file"
        className="form-control"
        accept={accept}
        required={required}
        disabled={disabled}
        onChange={(e) => onChange(e.target.files?.[0] ?? null)}
      />
      {hint ? <div className="form-text">{hint}</div> : null}
    </div>
  );
}
