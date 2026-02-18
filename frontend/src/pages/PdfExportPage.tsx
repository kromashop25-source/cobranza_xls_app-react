import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Controller, useForm } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { useMutation } from "@tanstack/react-query";
import {
  createColumnHelper,
  flexRender,
  getCoreRowModel,
  useReactTable,
} from "@tanstack/react-table";
import { motion } from "framer-motion";
import { RefreshCw } from "lucide-react";
import { z } from "zod";

import DateInput from "../components/DateInput";
import FileField from "../components/FileField";
import ProgressBar from "../components/ProgressBar";
import { StatusLine } from "../components/StatusLine";
import { downloadBlob, xhrPostWithProgressCancelable } from "../api/client";
import { Badge } from "../components/ui/badge";
import { Button } from "../components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "../components/ui/card";
import { Input } from "../components/ui/input";
import { Label } from "../components/ui/label";
import { Switch } from "../components/ui/switch";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "../components/ui/table";

type BlockItem = {
  id: string;
  name: string;
  sheet: string;
  include: boolean;
};

type PreviewResponse = {
  blocks: { id: string; name: string; sheet: string }[];
  count: number;
};

type DropHandler = (file: File) => void;

type PdfExportPageProps = {
  registerDropHandler?: (handler: DropHandler | null) => void;
};

const pdfFormSchema = z.object({
  excel: z.union([z.instanceof(File), z.null()]).refine((value): value is File => value instanceof File, {
    message: "Adjunta el XLS previamente generado en /merge.",
  }),
  hojaBase: z.string().optional(),
  pdfDate: z.string().min(1, "Selecciona la fecha usada al generar el Excel."),
});

type PdfFormInput = z.input<typeof pdfFormSchema>;
type PdfFormValues = z.output<typeof pdfFormSchema>;

const DEFAULT_CONSOLIDATED_ORDER = [
  "SALDOS COBRANZA",
  "MANUEL CARRASCO",
  "PITER HUAYTA",
  "LEONEL MEZA",
  "CANETE",
  "BEATRIZ ROJAS",
  "LURIN",
  "MANCHAY",
  "CIUDAD",
  "UNICACHI",
  "NORTE - ROSA",
  "CAQUETA (ROSA)",
  "SURCO (OSCAR)",
  "SURQUILLO (OSCAR)",
  "SAN LUIS (OSCAR)",
  "RAUL ARROYO",
] as const;

const ORDER_ALIASES: Record<string, string[]> = {
  "SALDOS COBRANZA": ["SALDOS COBRANZA"],
  "MANUEL CARRASCO": ["MANUEL CARRASCO"],
  "PITER HUAYTA": ["PITER HUAYTA"],
  "LEONEL MEZA": ["LEONEL MEZA"],
  CANETE: ["CANETE", "CANETE - (MANUEL)", "CANETE (MANUEL)", "CAÑETE", "CAÑETE - (MANUEL)", "CAÑETE (MANUEL)"],
  "BEATRIZ ROJAS": ["BEATRIZ ROJAS"],
  LURIN: ["LURIN", "LURIN - (ROSA)"],
  MANCHAY: ["MANCHAY", "MANCHAY - (ROSA)"],
  CIUDAD: ["CIUDAD", "CIUDAD - (ROSA)"],
  UNICACHI: ["UNICACHI", "UNICACHI SUR - (ROSA)", "UNICACHI SUR (ROSA)"],
  "NORTE - ROSA": ["NORTE - ROSA", "NORTE ROSA", "NORTE-ROSA"],
  "CAQUETA (ROSA)": ["CAQUETA (ROSA)", "CAQUETA - (ROSA)", "CAQUETA ROSA"],
  "SURCO (OSCAR)": ["SURCO (OSCAR)", "SURCO - (OSCAR)", "SURCO OSCAR"],
  "SURQUILLO (OSCAR)": ["SURQUILLO (OSCAR)", "SURQ/SURCO - (OSCAR)", "SURQ/SURCO (OSCAR)"],
  "SAN LUIS (OSCAR)": ["SAN LUIS (OSCAR)", "SAN LUIS - (OSCAR)", "SAN LUIS OSCAR"],
  "RAUL ARROYO": ["RAUL ARROYO"],
};

function buildZipName(file: File | null) {
  if (!file?.name) return "PDFS_COBRANZA.zip";
  const stem = file.name.replace(/\.[^.]+$/i, "");
  return `PDFS_${stem}.zip`;
}

function stripLeadingCode(name: string) {
  return name.replace(/^\s*\d+\s+/, "").trim();
}

function normalizeOrderName(name: string) {
  return stripLeadingCode(name)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function applyDefaultOrder(blocks: BlockItem[]) {
  const orderMap = new Map<string, number>();
  DEFAULT_CONSOLIDATED_ORDER.forEach((name, index) => {
    const aliases = ORDER_ALIASES[name] ?? [name];
    aliases.forEach((alias) => {
      orderMap.set(normalizeOrderName(alias), index);
    });
  });

  return blocks
    .map((block, index) => ({ ...block, _idx: index }))
    .sort((a, b) => {
      const aKey = normalizeOrderName(a.name);
      const bKey = normalizeOrderName(b.name);
      const aOrder = orderMap.has(aKey) ? orderMap.get(aKey)! : Number.MAX_SAFE_INTEGER;
      const bOrder = orderMap.has(bKey) ? orderMap.get(bKey)! : Number.MAX_SAFE_INTEGER;
      if (aOrder !== bOrder) return aOrder - bOrder;
      return a._idx - b._idx;
    })
    .map((item) => {
      const { _idx: droppedIndex, ...rest } = item;
      void droppedIndex;
      return rest;
    });
}

function parseDateFromFilename(name: string) {
  const match = name.match(/(\d{2})[-_](\d{2})[-_](\d{2,4})/);
  if (!match) return "";
  const day = match[1];
  const month = match[2];
  const rawYear = match[3];
  const year = rawYear.length === 2 ? `20${rawYear}` : rawYear;
  return `${year}-${month}-${day}`;
}

function isXlsFile(file: File) {
  return file.name.toLowerCase().endsWith(".xls");
}

async function parseJsonOrThrow(response: Response) {
  const data = (await response.json().catch(() => null)) as { detail?: string } | null;
  if (!response.ok) {
    throw new Error(data?.detail ?? "No se pudo procesar la solicitud.");
  }
  return data;
}

export default function PdfExportPage({ registerDropHandler }: PdfExportPageProps) {
  const [blocks, setBlocks] = useState<BlockItem[]>([]);
  const [analysisStatus, setAnalysisStatus] = useState("");
  const [analysisOk, setAnalysisOk] = useState(true);
  const [status, setStatus] = useState("");
  const [statusOk, setStatusOk] = useState(true);
  const [progress, setProgress] = useState(0);
  const activeExportCancelRef = useRef<(() => void) | null>(null);

  const {
    control,
    handleSubmit,
    setValue,
    watch,
    formState: { isSubmitting },
  } = useForm<PdfFormInput, unknown, PdfFormValues>({
    resolver: zodResolver(pdfFormSchema),
    defaultValues: {
      excel: null,
      hojaBase: "",
      pdfDate: "",
    },
  });

  const excel = watch("excel");
  const hojaBase = watch("hojaBase");

  useEffect(() => {
    if (!registerDropHandler) return;

    registerDropHandler((file) => {
      if (!isXlsFile(file)) {
        setStatusOk(false);
        setStatus("Solo se permiten archivos .XLS en esta pantalla.");
        return;
      }

      setValue("excel", file, { shouldDirty: true, shouldValidate: true });
      setStatusOk(true);
      setStatus(`Archivo cargado por arrastre: ${file.name}`);
    });

    return () => registerDropHandler(null);
  }, [registerDropHandler, setValue]);

  const { mutateAsync: analyzePreview, isPending: isAnalyzing } = useMutation({
    mutationFn: async ({ file, baseSheet }: { file: File; baseSheet: string }) => {
      const formData = new FormData();
      formData.append("excel", file);
      if (baseSheet.trim()) formData.append("hoja_base", baseSheet.trim());

      const response = await fetch("/pdf/preview-upload", {
        method: "POST",
        body: formData,
      });
      const data = (await parseJsonOrThrow(response)) as PreviewResponse;
      return data;
    },
  });

  const analyzePreviewRef = useRef(analyzePreview);

  useEffect(() => {
    analyzePreviewRef.current = analyzePreview;
  }, [analyzePreview]);

  const submitPdfMutation = useMutation({
    mutationFn: async (values: PdfFormValues) => {
      if (!(values.excel instanceof File)) {
        throw new Error("Adjunta el XLS previamente generado en /merge.");
      }
      const formData = new FormData();
      formData.append("excel", values.excel);
      if (values.hojaBase?.trim()) formData.append("hoja_base", values.hojaBase.trim());
      if (values.pdfDate) formData.append("pdf_date", values.pdfDate);
      if (blocks.length) {
        formData.append("orden", JSON.stringify(blocks.map((block) => block.id)));
        formData.append(
          "excluir",
          JSON.stringify(blocks.filter((block) => !block.include).map((block) => block.id))
        );
      }

      const task = xhrPostWithProgressCancelable(
        "/pdf/export-upload",
        formData,
        (pct) => setProgress(Math.min(70, Math.round(pct))),
        (pct) => setProgress(Math.round(pct))
      );

      activeExportCancelRef.current = task.cancel;
      try {
        return await task.promise;
      } finally {
        activeExportCancelRef.current = null;
      }
    },
  });

  const runAnalysis = useCallback(
    async (file: File | null, baseSheet: string) => {
      if (!file) return;

      setAnalysisStatus("Analizando XLS...");
      setAnalysisOk(true);
      try {
        const data = await analyzePreviewRef.current({ file, baseSheet });
        const ordered = applyDefaultOrder(
          data.blocks.map((block) => ({
            id: block.id,
            name: block.name,
            sheet: block.sheet,
            include: true,
          }))
        );
        setBlocks(ordered);
        setAnalysisStatus(`Detectados ${data.count} vendedores.`);
        setAnalysisOk(true);
      } catch (error) {
        setAnalysisStatus(error instanceof Error ? error.message : "No se pudo analizar el XLS.");
        setAnalysisOk(false);
        setBlocks([]);
      }
    },
    []
  );

  useEffect(() => {
    if (!excel) {
      setBlocks([]);
      setAnalysisStatus("");
      return;
    }

    const inferredDate = parseDateFromFilename(excel.name);
    if (inferredDate) {
      setValue("pdfDate", inferredDate, { shouldValidate: true });
    }

    void runAnalysis(excel, hojaBase ?? "");
  }, [excel, hojaBase, runAnalysis, setValue]);

  const handleMove = useCallback((index: number, direction: -1 | 1) => {
    setBlocks((prev) => {
      const next = [...prev];
      const target = index + direction;
      if (target < 0 || target >= next.length) return prev;
      [next[index], next[target]] = [next[target], next[index]];
      return next;
    });
  }, []);

  const handleToggleInclude = useCallback((id: string) => {
    setBlocks((prev) => prev.map((item) => (item.id === id ? { ...item, include: !item.include } : item)));
  }, []);

  const columns = useMemo(() => {
    const helper = createColumnHelper<BlockItem>();
    return [
      helper.display({
        id: "order",
        header: "#",
        cell: (ctx) => <Badge variant="secondary">{ctx.row.index + 1}</Badge>,
      }),
      helper.accessor("name", {
        header: "Vendedor",
        cell: (ctx) => (
          <div>
            <p className="font-medium text-foreground">{ctx.row.original.name}</p>
            <p className="text-xs text-muted-foreground">{ctx.row.original.sheet}</p>
          </div>
        ),
      }),
      helper.display({
        id: "include",
        header: "Incluir",
        cell: (ctx) => (
          <Switch
            checked={ctx.row.original.include}
            onCheckedChange={() => handleToggleInclude(ctx.row.original.id)}
            aria-label={`Incluir ${ctx.row.original.name}`}
          />
        ),
      }),
      helper.display({
        id: "actions",
        header: "Orden",
        cell: (ctx) => (
          <div className="flex gap-2">
            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={() => handleMove(ctx.row.index, -1)}
              disabled={ctx.row.index === 0}
            >
              Subir
            </Button>
            <Button
              type="button"
              variant="outline"
              size="sm"
              onClick={() => handleMove(ctx.row.index, 1)}
              disabled={ctx.row.index === blocks.length - 1}
            >
              Bajar
            </Button>
          </div>
        ),
      }),
    ];
  }, [blocks.length, handleMove, handleToggleInclude]);

  const table = useReactTable({
    data: blocks,
    columns,
    getCoreRowModel: getCoreRowModel(),
  });

  const handleCancel = () => {
    if (!activeExportCancelRef.current) return;
    activeExportCancelRef.current();
    activeExportCancelRef.current = null;
    setProgress(0);
    setStatusOk(false);
    setStatus("Proceso cancelado por el usuario.");
  };

  const onSubmit = handleSubmit(async (values) => {
    setStatus("Generando PDFs...");
    setStatusOk(true);
    setProgress(1);

    try {
      const blob = await submitPdfMutation.mutateAsync(values);
      const zipName = buildZipName(values.excel);
      downloadBlob(blob, zipName);
      setStatus(`ZIP descargado (${zipName}).`);
      setStatusOk(true);
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "No se pudo generar el ZIP.");
      setStatusOk(false);
    } finally {
      setProgress(0);
    }
  });

  const isSubmitPending = isSubmitting || submitPdfMutation.isPending;

  return (
    <section className="mx-auto w-full max-w-5xl">
      <motion.div initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.24 }}>
        <Card className="border-border/70 bg-card/95">
          <CardHeader>
            <CardTitle>Exportar PDFs desde XLS formateado</CardTitle>
            <CardDescription>
              Usa el archivo generado en "Copiar a Maestro". El backend mantiene el mismo motor Excel/COM.
            </CardDescription>
          </CardHeader>
          <CardContent>
            <form className="space-y-5" onSubmit={onSubmit} noValidate>
              <Controller
                name="excel"
                control={control}
                render={({ field, fieldState }) => (
                  <FileField
                    label="XLS formateado"
                    accept=".xls"
                    required
                    value={field.value}
                    onChange={field.onChange}
                    error={fieldState.error?.message}
                    disabled={isSubmitPending}
                    hint="Es el resultado del paso /merge."
                  />
                )}
              />

              <div className="grid gap-4 md:grid-cols-2">
                <div className="space-y-2">
                  <Label htmlFor="hojaBase">Hoja base (opcional)</Label>
                  <Controller
                    name="hojaBase"
                    control={control}
                    render={({ field }) => (
                      <Input
                        id="hojaBase"
                        type="text"
                        placeholder='Ej: "OFICINA (VES)"'
                        value={field.value ?? ""}
                        onChange={field.onChange}
                        disabled={isSubmitPending}
                      />
                    )}
                  />
                  <p className="text-xs text-muted-foreground">
                    Solo si deseas limitar el analisis a una hoja especifica.
                  </p>
                </div>

                <div className="space-y-2">
                  <Label htmlFor="pdfDate">Fecha para nombres de PDF</Label>
                  <Controller
                    name="pdfDate"
                    control={control}
                    render={({ field, fieldState }) => (
                      <>
                        <DateInput
                          id="pdfDate"
                          value={field.value}
                          onChange={field.onChange}
                          disabled={isSubmitPending}
                        />
                        {fieldState.error?.message ? (
                          <p className="text-xs text-destructive">{fieldState.error.message}</p>
                        ) : null}
                      </>
                    )}
                  />
                  <p className="text-xs text-muted-foreground">
                    Usa la misma fecha seleccionada al generar el Excel.
                  </p>
                </div>
              </div>

              <div className="mt-3 flex flex-wrap items-center gap-2 pl-0.5">
                <Button
                  type="button"
                  variant="outline"
                  className="gap-2"
                  disabled={!excel || isAnalyzing || isSubmitPending}
                  onClick={() => void runAnalysis(excel, hojaBase ?? "")}
                >
                  <RefreshCw className={`h-4 w-4 ${isAnalyzing ? "animate-spin" : ""}`} />
                  {isAnalyzing ? "Analizando..." : "Analizar XLS"}
                </Button>
                <Button
                  type="button"
                  variant="outline"
                  disabled={!blocks.length || isSubmitPending}
                  onClick={() => setBlocks((prev) => applyDefaultOrder(prev))}
                >
                  Orden sugerido
                </Button>
              </div>

              {analysisStatus ? <StatusLine text={analysisStatus} ok={analysisOk} /> : null}

              {blocks.length > 0 ? (
                <div className="space-y-3 rounded-lg border border-border p-4">
                  <div className="flex items-center justify-between">
                    <p className="font-medium text-foreground">Orden para consolidado</p>
                    <p className="text-xs text-muted-foreground">
                      {blocks.filter((block) => !block.include).length} excluidos
                    </p>
                  </div>

                  <Table>
                    <TableHeader>
                      {table.getHeaderGroups().map((headerGroup) => (
                        <TableRow key={headerGroup.id}>
                          {headerGroup.headers.map((header) => (
                            <TableHead key={header.id}>
                              {header.isPlaceholder
                                ? null
                                : flexRender(header.column.columnDef.header, header.getContext())}
                            </TableHead>
                          ))}
                        </TableRow>
                      ))}
                    </TableHeader>
                    <TableBody>
                      {table.getRowModel().rows.map((row) => (
                        <TableRow key={row.id}>
                          {row.getVisibleCells().map((cell) => (
                            <TableCell key={cell.id}>
                              {flexRender(cell.column.columnDef.cell, cell.getContext())}
                            </TableCell>
                          ))}
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </div>
              ) : null}

              <div className="mt-3 border-t border-border/60 pt-4">
                <div className="flex flex-wrap items-center gap-2 sm:gap-3">
                  <Button type="submit" disabled={isSubmitPending}>
                    {isSubmitPending ? "Generando..." : "Generar PDFs"}
                  </Button>
                  {submitPdfMutation.isPending ? (
                    <Button type="button" variant="destructive" onClick={handleCancel}>
                      Cancelar
                    </Button>
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
