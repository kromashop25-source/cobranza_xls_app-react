import { useEffect, useMemo, useRef, useState } from "react";
import { Controller, useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { useMutation, useQuery } from "@tanstack/react-query";
import { motion } from "framer-motion";
import { z } from "zod";

import DateInput from "../components/DateInput";
import FileField from "../components/FileField";
import ProgressBar from "../components/ProgressBar";
import { StatusLine } from "../components/StatusLine";
import { downloadBlob, fetchJson, xhrPostWithProgressCancelable } from "../api/client";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "../components/ui/card";
import { Button } from "../components/ui/button";
import { Label } from "../components/ui/label";
import { Switch } from "../components/ui/switch";

type MasterInfo = {
  exists: boolean;
  name?: string | null;
  debug_path?: string | null;
};

type DropHandler = (file: File) => void;

type MergePageProps = {
  registerDropHandler?: (handler: DropHandler | null) => void;
};

const mergeFormSchema = z
  .object({
    source: z.union([z.instanceof(File), z.null()]),
    master: z.union([z.instanceof(File), z.null()]),
    hdrDate: z.string().optional(),
    useDefaultMaster: z.boolean(),
  })
  .superRefine((values, ctx) => {
    if (!(values.source instanceof File)) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["source"],
        message: "Carga el archivo de tecnicos (.XLS).",
      });
    }

    if (!values.useDefaultMaster && !(values.master instanceof File)) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: ["master"],
        message: 'Carga el maestro o activa "Usar maestro por defecto".',
      });
    }
  });

type MergeFormValues = z.infer<typeof mergeFormSchema>;

function buildDownloadName(source: File | null, hdrDate: string | undefined) {
  if (source?.name) {
    return source.name;
  }
  if (hdrDate) {
    const [year, month, day] = hdrDate.split("-");
    if (year && month && day) {
      return `COBRANZA_${day}-${month}-${year.slice(-2)}.xls`;
    }
  }
  return "COBRANZA.xls";
}

function isXlsFile(file: File) {
  return file.name.toLowerCase().endsWith(".xls");
}

export default function MergePage({ registerDropHandler }: MergePageProps) {
  const [status, setStatus] = useState("");
  const [statusOk, setStatusOk] = useState(true);
  const [progress, setProgress] = useState(0);
  const activeMergeCancelRef = useRef<(() => void) | null>(null);

  const {
    control,
    handleSubmit,
    setValue,
    watch,
    formState: { isSubmitting },
  } = useForm<MergeFormValues>({
    resolver: zodResolver(mergeFormSchema),
    defaultValues: {
      source: null,
      master: null,
      hdrDate: "",
      useDefaultMaster: true,
    },
  });

  const useDefaultMaster = watch("useDefaultMaster");
  const selectedSource = watch("source");
  const selectedMaster = watch("master");

  useEffect(() => {
    if (!registerDropHandler) return;

    registerDropHandler((file) => {
      if (!isXlsFile(file)) {
        setStatusOk(false);
        setStatus("Solo se permiten archivos .XLS en esta pantalla.");
        return;
      }

      setValue("source", file, { shouldDirty: true, shouldValidate: true });
      setStatusOk(true);
      setStatus(`Archivo cargado por arrastre: ${file.name}`);
    });

    return () => registerDropHandler(null);
  }, [registerDropHandler, setValue]);

  const masterInfoQuery = useQuery({
    queryKey: ["master-default-info"],
    queryFn: () => fetchJson<MasterInfo>("/master/default-info"),
    staleTime: 30_000,
  });

  const submitMergeMutation = useMutation({
    mutationFn: async (values: MergeFormValues) => {
      const formData = new FormData();
      if (values.source instanceof File) formData.append("source", values.source);
      if (!values.useDefaultMaster && values.master instanceof File) {
        formData.append("master", values.master);
      }
      if (values.hdrDate) formData.append("hdr_date", values.hdrDate);
      formData.append("use_default_master", values.useDefaultMaster ? "1" : "0");

      const task = xhrPostWithProgressCancelable(
        "/merge",
        formData,
        (pct) => setProgress(Math.min(70, Math.round(pct))),
        (pct) => setProgress(Math.round(pct))
      );

      activeMergeCancelRef.current = task.cancel;
      try {
        return await task.promise;
      } finally {
        activeMergeCancelRef.current = null;
      }
    },
  });

  const defaultMasterHint = useMemo(() => {
    if (masterInfoQuery.isPending) return "Verificando maestro por defecto...";
    const info = masterInfoQuery.data;
    if (!info?.exists) {
      const extra = info?.debug_path ? ` (Ruta: ${info.debug_path})` : "";
      return `No se encontro el maestro alojado.${extra}`;
    }
    return `Usar ${info.name ?? "maestro por defecto"}`;
  }, [masterInfoQuery.data, masterInfoQuery.isPending]);

  const handleCancel = () => {
    if (!activeMergeCancelRef.current) return;
    activeMergeCancelRef.current();
    activeMergeCancelRef.current = null;
    setProgress(0);
    setStatusOk(false);
    setStatus("Proceso cancelado por el usuario.");
  };

  const onSubmit = handleSubmit(async (values) => {
    setStatus("Procesando, por favor espera...");
    setStatusOk(true);
    setProgress(1);

    try {
      const blob = await submitMergeMutation.mutateAsync(values);
      const filename = buildDownloadName(values.source, values.hdrDate);
      downloadBlob(blob, filename);
      setStatus(`Listo. Se descargo ${filename}.`);
      setStatusOk(true);
    } catch (error) {
      setStatusOk(false);
      setStatus(error instanceof Error ? error.message : "Error desconocido al procesar.");
    } finally {
      setProgress(0);
    }
  });

  const isBusy = isSubmitting || submitMergeMutation.isPending;

  return (
    <section className="mx-auto w-full max-w-4xl">
      <motion.div initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.24 }}>
        <Card className="border-border/70 bg-card/95">
          <CardHeader>
            <CardTitle>Cobranza XLS - Copiar a Maestro</CardTitle>
            <CardDescription>
              Requiere Windows con Microsoft Excel instalado. Se conserva el mismo flujo de procesamiento XLS.
            </CardDescription>
          </CardHeader>
          <CardContent>
            <form className="space-y-5" onSubmit={onSubmit} noValidate>
              <Controller
                name="source"
                control={control}
                render={({ field, fieldState }) => (
                  <FileField
                    label="Origen (.XLS tecnicos)"
                    accept=".xls"
                    required
                    value={field.value}
                    onChange={field.onChange}
                    error={fieldState.error?.message}
                    disabled={isBusy}
                    hint="Archivo exportado por el equipo de cobranzas."
                  />
                )}
              />

              <div className="rounded-lg border border-border p-4">
                <div className="flex items-center justify-between gap-3">
                  <div className="space-y-1">
                    <Label htmlFor="use-default-master">Usar maestro por defecto</Label>
                    <p className="text-xs text-muted-foreground">{defaultMasterHint}</p>
                  </div>
                  <Controller
                    name="useDefaultMaster"
                    control={control}
                    render={({ field }) => (
                      <Switch
                        id="use-default-master"
                        checked={field.value}
                        onCheckedChange={(checked) => {
                          field.onChange(checked);
                          if (checked) {
                            setValue("master", null, { shouldValidate: true });
                          }
                        }}
                        disabled={isBusy}
                      />
                    )}
                  />
                </div>
              </div>

              {!useDefaultMaster ? (
                <Controller
                  name="master"
                  control={control}
                  render={({ field, fieldState }) => (
                    <FileField
                      label="Maestro (.XLS)"
                      accept=".xls"
                      required
                      value={field.value}
                      onChange={field.onChange}
                      error={fieldState.error?.message}
                      disabled={isBusy}
                      hint="Archivo maestro actualizado."
                    />
                  )}
                />
              ) : null}

              <div className="space-y-2">
                <Label htmlFor="hdrDate">Fecha (opcional)</Label>
                <Controller
                  name="hdrDate"
                  control={control}
                  render={({ field }) => (
                    <DateInput
                      id="hdrDate"
                      value={field.value ?? ""}
                      onChange={field.onChange}
                      disabled={isBusy}
                    />
                  )}
                />
                <p className="text-xs text-muted-foreground">
                  Se usara para actualizar encabezados y hojas SUR/NORTE.
                </p>
              </div>

              <div className="mt-3 border-t border-border/60 pt-4">
                <div className="flex flex-wrap items-center gap-2 sm:gap-3">
                  <Button type="submit" disabled={isBusy}>
                    {isBusy ? "Procesando..." : "Copiar a Maestro"}
                  </Button>
                  {submitMergeMutation.isPending ? (
                    <Button type="button" variant="destructive" onClick={handleCancel}>
                      Cancelar
                    </Button>
                  ) : null}
                </div>

                <div className="mt-2 flex flex-wrap gap-x-4 gap-y-1">
                  <p className="text-xs text-muted-foreground">
                    Archivo origen: {selectedSource?.name ?? "sin seleccionar"}
                  </p>
                  {!useDefaultMaster ? (
                    <p className="text-xs text-muted-foreground">
                      Maestro: {selectedMaster?.name ?? "sin seleccionar"}
                    </p>
                  ) : null}
                </div>
              </div>
            </form>

            <div className="mt-5 space-y-2">
              <StatusLine text={status} ok={statusOk} />
              <ProgressBar value={progress} />
            </div>
          </CardContent>
        </Card>
      </motion.div>
    </section>
  );
}
