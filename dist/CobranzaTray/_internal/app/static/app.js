// /static/app.js
// @ts-nocheck
(() => {
  const $ = (s) => document.querySelector(s);

  // ---- Elementos de entrada / UI ----
  const src = $("#src");
  const mst = $("#mst");
  const hdrDate = $("#hdrDate");
  const btn = $("#go");

  // Toggle "usar maestro por defecto" + etiqueta con el nombre del default
  const useDefaultMaster = $("#useDefaultMaster");
  const defaultMasterName = $("#defaultMasterName");

  // Panel PDFs
  const actions = document.getElementById("pdf-actions");
  const hojaBaseEl = document.getElementById("hojaBase");
  const btnGenPDFs = document.getElementById("btnGenPDFs");
  const btnZipPDFs = document.getElementById("btnZipPDFs");
  const pdfMsg = document.getElementById("pdfMsg");

  // XLS generado en memoria (para PDFs)
  let lastBlob = null;
  let lastName = null;

  // ---- Estado / Progreso ----
  const statusEl = () => document.getElementById("status");
  // ---- Progreso (UI robusto) ----
const progress = {
  el: document.getElementById("progress"),
  bar: document.querySelector("#progress .progress__bar"),
  ensure() {
    // 1) Usa el #progress del HTML si existe
    if (!this.el) this.el = document.getElementById("progress");
    // Si no existe, créalo y ponlo debajo del status
    if (!this.el) {
      this.el = document.createElement("div");
      this.el.id = "progress";
      this.el.className = "progress is-hidden";
      (statusEl()?.parentNode || document.body).insertBefore(
        this.el,
        statusEl()?.nextSibling || null
      );
    }
    // 2) Usa la barra que ya existe o créala
    if (!this.bar) {
      this.bar =
        this.el.querySelector(".progress__bar") || document.createElement("div");
      if (!this.bar.parentNode) {
        this.bar.className = "progress__bar";
        this.el.appendChild(this.bar);
      }
    }
  },
  show() {
    this.ensure();
    // Asegura visibilidad aunque el CSS requiera .show
    this.el.style.display = "";
    this.el.classList.remove("is-hidden", "progress--error");
    this.el.classList.add("show");
    this.bar.style.width = "0%";
    requestAnimationFrame(() => (this.bar.style.width = "5%"));
  },
  set(p) {
    if (!this.bar) return;
    const v = Math.max(0, Math.min(100, Math.floor(p)));
    this.bar.style.width = v + "%";
  },
  hide(ok = true) {
    if (!this.el) return;
    if (!ok) this.el.classList.add("progress--error");
    this.set(ok ? 100 : 96);
    // Si tu CSS anima con .show, la quitamos antes de ocultar
    this.el.classList.remove("show");
    setTimeout(() => this.el.classList.add("is-hidden"), ok ? 500 : 900);
  },
  error() { this.hide(false); },
};

  function setStatus(msg, ok = true) {
    const el = statusEl();
    if (!el) return;
    el.textContent = msg;
    el.className = ok ? "status ok" : "status err";
  }

  // ---- Sincronizar toggle maestro por defecto ----
  function syncToggle() {
    const useDefault = !!(useDefaultMaster && useDefaultMaster.checked);
    if (mst) {
      mst.disabled = useDefault;
      if (useDefault) mst.value = "";
    }
  }
  useDefaultMaster && useDefaultMaster.addEventListener("change", syncToggle);
  syncToggle();

  // Mostrar nombre del maestro alojado
  if (defaultMasterName) {
    fetch("/master/default-info")
      .then((r) => r.json())
      .then((info) => {
        defaultMasterName.textContent =
          info && info.exists && info.name ? `(${info.name})` : "(no encontrado)";
      })
      .catch(() => {});
  }

  // ---- POST genérico con progreso ----
  function xhrPost(url, formData, onUpload, onDownload, responseType = "blob") {
    return new Promise((resolve, reject) => {
      const xhr = new XMLHttpRequest();
      xhr.open("POST", url);
      xhr.responseType = responseType;

      xhr.upload.onprogress = (e) => {
        if (e.lengthComputable && onUpload) onUpload(e.loaded, e.total);
      };
      xhr.onprogress = (e) => {
        if (e.lengthComputable && onDownload) onDownload(e.loaded, e.total);
      };

      xhr.onerror = () => reject(new Error("Error de red"));
      xhr.onload = () => {
        if (xhr.status >= 200 && xhr.status < 300) {
          resolve({ body: xhr.response, headers: xhr.getAllResponseHeaders() });
        } else {
          try {
            const reader = new FileReader();
            reader.onload = () =>
              reject(new Error(String(reader.result || `HTTP ${xhr.status}`)));
            reader.onerror = () => reject(new Error(`HTTP ${xhr.status}`));
            reader.readAsText(xhr.response);
          } catch {
            reject(new Error(`HTTP ${xhr.status}`));
          }
        }
      };

      xhr.send(formData);
    });
  }

  function filenameFromHeaders(headers, fallback) {
    if (!headers) return fallback;
    try {
      const m =
        /content-disposition:.*?filename\*=UTF-8''([^;\r\n]+)|content-disposition:.*?filename="?([^";\r\n]+)"?/i.exec(
          headers
        );
      if (m) return decodeURIComponent(m[1] || m[2]);
    } catch {}
    return fallback;
  }

  // ---- Copiar a Maestro ----
  btn &&
    btn.addEventListener("click", async () => {
      try {
        const useDefault = !!(useDefaultMaster && useDefaultMaster.checked);

        // Oculta panel PDFs al iniciar
        if (actions) {
          actions.style.display = "none";
          if (pdfMsg) pdfMsg.textContent = "";
          if (btnZipPDFs) btnZipPDFs.style.display = "none";
        }

        // Validaciones
        if (!src?.files?.[0]) {
          setStatus("Carga el archivo de técnicos (.XLS).", false);
          return;
        }
        if (!useDefault && !mst?.files?.[0]) {
          setStatus('Carga el maestro (.XLS) o activa "Usar maestro por defecto".', false);
          return;
        }
        if (!hdrDate?.value) {
          setStatus("Selecciona la fecha de encabezado.", false);
          return;
        }

        setStatus("Procesando… por favor espera.");
        progress.show();

        const fd = new FormData();
        fd.append("source", src.files[0]);
        if (!useDefault) fd.append("master", mst.files[0]);
        fd.append("hdr_date", hdrDate.value);
        fd.append("use_default_master", useDefault ? "1" : "0");

        // Progreso: 5–70 subida, 70–99 descarga
        const { body: blob, headers } = await xhrPost(
          "/merge",
          fd,
          (loaded, total) => progress.set(5 + (loaded / total) * 65),
          (loaded, total) => progress.set(70 + (loaded / total) * 29),
          "blob"
        );

        const [y, mm, dd] = (hdrDate.value || "").split("-");
        let downloadName =
          y && mm && dd ? `COBRANZA ${dd}-${mm}-${y.slice(2)}.xls` : "COBRANZA.xls";
        downloadName = filenameFromHeaders(headers, downloadName);

        // Descargar XLS
        {
          const url = URL.createObjectURL(blob);
          const a = document.createElement("a");
          a.href = url;
          a.download = downloadName;
          document.body.appendChild(a);
          a.click();
          a.remove();
          URL.revokeObjectURL(url);
        }

        // Guardar para PDFs y mostrar panel
        lastBlob = blob;
        lastName = downloadName;
        if (actions) {
          actions.style.display = "flex";
          if (pdfMsg) pdfMsg.textContent = "Listo para generar PDFs.";
          if (btnZipPDFs) btnZipPDFs.style.display = "none";
        }

        progress.hide(true);
        setStatus(`Listo. Se descargó ${downloadName}`);
      } catch (err) {
        progress.error();
        setStatus(`Error: ${err && err.message ? err.message : err}`, false);
      }
    });

  // ---- Generar PDFs ----
  btnGenPDFs &&
    btnGenPDFs.addEventListener("click", async () => {
      if (!lastBlob) {
        if (pdfMsg) pdfMsg.textContent = 'Primero genera el XLS con "Copiar a Maestro".';
        return;
      }
      try {
        if (pdfMsg) pdfMsg.textContent = "Generando PDFs...";
        setStatus("Generando PDFs…", true);
        progress.show();

        const excelFile = new File([lastBlob], lastName || "COBRANZA.xls", {
          type: "application/vnd.ms-excel",
        });
        const fd = new FormData();
        fd.append("excel", excelFile);
        const hojaBase = hojaBaseEl && hojaBaseEl.value ? hojaBaseEl.value : "";
        if (hojaBase) fd.append("hoja_base", hojaBase);

        // 0–30 subida, 95–100 descarga (la fase COM no emite progreso real)
        const { body: zipBlob } = await xhrPost(
          "/pdf/export-upload",
          fd,
          (loaded, total) => progress.set((loaded / total) * 30),
          (loaded, total) => progress.set(95 + (loaded / total) * 5),
          "blob"
        );

        const url = URL.createObjectURL(zipBlob);
        if (btnZipPDFs) {
          btnZipPDFs.href = url;
          btnZipPDFs.download = `PDFS_${(lastName || "COBRANZA").replace(/\.xls$/i, "")}.zip`;
          btnZipPDFs.style.display = "inline-block";
          btnZipPDFs.click(); // auto descarga
          setTimeout(() => URL.revokeObjectURL(url), 20000);
        }

        if (pdfMsg) pdfMsg.textContent = "PDFs generados correctamente.";
        progress.hide(true);
        setStatus("PDFs generados.", true);
      } catch (e) {
        if (pdfMsg) pdfMsg.textContent = `Error: ${e.message || e}`;
        progress.error();
        setStatus(`Error al generar PDFs: ${e && e.message ? e.message : e}`, false);
      }
    });
})();
