const API_BASE =
  (import.meta.env.VITE_API_BASE as string | undefined)?.replace(/\/$/, "") ||
  window.location.origin;

export type UploadProgressCb = (pct: number) => void;
export type CancelUploadFn = () => void;

export type CancelableUpload = {
  promise: Promise<Blob>;
  cancel: CancelUploadFn;
};

const NETWORK_ERROR = "Error de red";

function readBlobAsText(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(NETWORK_ERROR);
    reader.onload = () => resolve(String(reader.result ?? ""));
    reader.readAsText(blob);
  });
}

async function extractErrorMessage(blob: Blob, status: number): Promise<string> {
  try {
    const text = await readBlobAsText(blob);
    if (!text) return `HTTP ${status}`;
    try {
      const parsed = JSON.parse(text) as { detail?: unknown };
      if (parsed?.detail) return String(parsed.detail);
    } catch {
      // ignore JSON parse errors
    }
    return text;
  } catch {
    return `HTTP ${status}`;
  }
}

export async function fetchJson<T>(url: string, init?: RequestInit): Promise<T> {
  const response = await fetch(`${API_BASE}${url}`, init);
  if (!response.ok) {
    const text = await response.text().catch(() => "");
    if (text) {
      try {
        const parsed = JSON.parse(text) as { detail?: unknown };
        throw new Error(parsed?.detail ? String(parsed.detail) : text);
      } catch {
        throw new Error(text);
      }
    }
    throw new Error(`HTTP ${response.status}`);
  }
  return (await response.json()) as T;
}

export function xhrPostWithProgressCancelable(
  url: string,
  formData: FormData,
  onUpload?: UploadProgressCb,
  onDownload?: UploadProgressCb
): CancelableUpload {
  const xhr = new XMLHttpRequest();
  let settled = false;

  const promise = new Promise<Blob>((resolve, reject) => {
    const resolveOnce = (value: Blob) => {
      if (settled) return;
      settled = true;
      resolve(value);
    };

    const rejectOnce = (error: Error) => {
      if (settled) return;
      settled = true;
      reject(error);
    };

    xhr.open("POST", `${API_BASE}${url}`);
    xhr.responseType = "blob";

    xhr.upload.onprogress = (event) => {
      if (event.lengthComputable && onUpload) {
        onUpload((event.loaded / event.total) * 70);
      }
    };

    xhr.onprogress = (event) => {
      if (event.lengthComputable && onDownload) {
        onDownload(70 + (event.loaded / event.total) * 30);
      }
    };

    xhr.onabort = () => rejectOnce(new Error("Proceso cancelado por el usuario."));
    xhr.onerror = () => rejectOnce(new Error(NETWORK_ERROR));

    xhr.onload = async () => {
      if (xhr.status >= 200 && xhr.status < 300) {
        resolveOnce(xhr.response);
        return;
      }
      const blob = xhr.response instanceof Blob ? xhr.response : new Blob();
      rejectOnce(new Error(await extractErrorMessage(blob, xhr.status)));
    };

    xhr.send(formData);
  });

  const cancel = () => {
    if (settled) return;
    xhr.abort();
  };

  return { promise, cancel };
}

export async function xhrPostWithProgress(
  url: string,
  formData: FormData,
  onUpload?: UploadProgressCb,
  onDownload?: UploadProgressCb
): Promise<Blob> {
  const { promise } = xhrPostWithProgressCancelable(url, formData, onUpload, onDownload);
  return promise;
}

export function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

