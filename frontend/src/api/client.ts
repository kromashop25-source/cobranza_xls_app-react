const API_BASE = (import.meta.env.VITE_API_BASE as string | undefined)?.replace(
  /\/$/,
  ""
) || window.location.origin;

export type UploadProgressCb = (pct: number) => void;

const NETWORK_ERROR = "Error de red";

function readBlobAsText(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(NETWORK_ERROR);
    reader.onload = () => resolve(String(reader.result ?? ""));
    reader.readAsText(blob);
  });
}

async function extractErrorMessage(
  blob: Blob,
  status: number
): Promise<string> {
  try {
    const text = await readBlobAsText(blob);
    if (!text) return `HTTP ${status}`;
    try {
      const parsed = JSON.parse(text);
      if (parsed?.detail) return String(parsed.detail);
    } catch {
      /* ignore json parse */
    }
    return text;
  } catch {
    return `HTTP ${status}`;
  }
}

export async function xhrPostWithProgress(
  url: string,
  formData: FormData,
  onUpload?: UploadProgressCb,
  onDownload?: UploadProgressCb
): Promise<Blob> {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("POST", `${API_BASE}${url}`);
    xhr.responseType = "blob";

    xhr.upload.onprogress = (e) => {
      if (e.lengthComputable && onUpload) {
        onUpload((e.loaded / e.total) * 70);
      }
    };

    xhr.onprogress = (e) => {
      if (e.lengthComputable && onDownload) {
        onDownload(70 + (e.loaded / e.total) * 30);
      }
    };

    xhr.onerror = () => reject(new Error(NETWORK_ERROR));

    xhr.onload = async () => {
      if (xhr.status >= 200 && xhr.status < 300) {
        resolve(xhr.response);
        return;
      }
      const blob = xhr.response instanceof Blob ? xhr.response : new Blob();
      reject(new Error(await extractErrorMessage(blob, xhr.status)));
    };

    xhr.send(formData);
  });
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
