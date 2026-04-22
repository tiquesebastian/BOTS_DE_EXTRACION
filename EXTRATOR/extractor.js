import fs from "node:fs/promises";
import path from "node:path";
import { createWorker } from "tesseract.js";
import sharp from "sharp";
import pdfParse from "pdf-parse";

const PROJECT_ROOT = path.resolve("..");

const CONFIG = {
  rootPath: process.env.EXTRACTOR_ROOT_PATH || "Z:/IA 10/NUEVO/PARTE3/890701715",
  targetRadicadosPath: process.env.EXTRACTOR_TARGET_FILE || path.join(PROJECT_ROOT, "faltan estos radicados"),
  outputTxtPath: process.env.EXTRACTOR_OUTPUT_TXT || path.join(PROJECT_ROOT, "890701715_fechas_ingreso_js.txt"),
  outputDetalleCsvPath:
    process.env.EXTRACTOR_OUTPUT_DETAIL || path.join(PROJECT_ROOT, "890701715_fechas_ingreso_js_detalle.csv"),
  maxOcrPages: Number(process.env.EXTRACTOR_MAX_OCR_PAGES || 7),
  maxPdfTextPages: Number(process.env.EXTRACTOR_MAX_PDF_TEXT_PAGES || 7),
  maxFilesPerRadicado: Number(process.env.EXTRACTOR_MAX_FILES || 10),
  fileKeyword: normalizeToken(process.env.EXTRACTOR_FILE_KEYWORD || "tapa factura hoja de atencion otros"),
  ingresoKeyword: normalizeToken(process.env.EXTRACTOR_INGRESO_KEYWORD || "ingreso"),
  language: process.env.EXTRACTOR_OCR_LANG || "spa",
  extensions: new Set([".pdf", ".tif", ".tiff", ".png", ".jpg", ".jpeg", ".bmp", ".gif"])
};

const DATE_PATTERN = /(\d{1,2})\s*[\/\-.]\s*(\d{1,2})\s*[\/\-.]\s*(\d{2,4})/;
const ATENCION_PATTERN = /atencion\s*:?\s*(\d{8})\d*/i;

function normalizeToken(input) {
  return (input || "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[_-]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeRadicado(input) {
  const match = String(input || "").match(/\d{6,15}/);
  return match ? match[0] : String(input || "").trim();
}

function normalizeFecha(raw) {
  const match = String(raw || "").match(DATE_PATTERN);
  if (!match) return "";

  const dd = Number(match[1]);
  const mm = Number(match[2]);
  let yy = Number(match[3]);

  if (match[3].length === 2) yy += yy <= 50 ? 2000 : 1900;

  const dt = new Date(yy, mm - 1, dd);
  if (Number.isNaN(dt.getTime())) return "";
  if (dt.getFullYear() !== yy || dt.getMonth() !== mm - 1 || dt.getDate() !== dd) return "";

  return `${dt.getDate()}/${String(dt.getMonth() + 1).padStart(2, "0")}/${String(dt.getFullYear()).slice(-2)}`;
}

function buildIngresoPatterns(keyword) {
  const escaped = keyword.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return [
    new RegExp(`${escaped}.{0,220}?fecha.{0,60}?(\\d{1,2}\\s*[\\/\\-.]\\s*\\d{1,2}\\s*[\\/\\-.]\\s*\\d{2,4})`, "i"),
    new RegExp(`${escaped}.{0,220}?(\\d{1,2}\\s*[\\/\\-.]\\s*\\d{1,2}\\s*[\\/\\-.]\\s*\\d{2,4})`, "i")
  ];
}

const ingresoPatterns = buildIngresoPatterns(CONFIG.ingresoKeyword || "ingreso");
const filePriorityTokens = [...new Set((CONFIG.fileKeyword || "").split(" ").filter(Boolean))];

function extractFechaFromText(text) {
  const normalized = normalizeToken(String(text || "").replace(/\x00/g, " "));
  if (!normalized) return { fecha: "", cruda: "" };

  for (const pattern of ingresoPatterns) {
    const match = normalized.match(pattern);
    if (match) {
      const cruda = match[1].replace(/\s+/g, "");
      const fecha = normalizeFecha(cruda);
      if (fecha) return { fecha, cruda };
    }
  }

  const atencion = normalized.match(ATENCION_PATTERN);
  if (atencion) {
    const value = atencion[1];
    const yyyy = Number(value.slice(0, 4));
    const mm = Number(value.slice(4, 6));
    const dd = Number(value.slice(6, 8));
    const dt = new Date(yyyy, mm - 1, dd);
    if (!Number.isNaN(dt.getTime())) {
      return {
        fecha: `${dt.getDate()}/${String(mm).padStart(2, "0")}/${String(yyyy).slice(-2)}`,
        cruda: value
      };
    }
  }

  const idx = normalized.indexOf(CONFIG.ingresoKeyword || "ingreso");
  if (idx >= 0) {
    const segment = normalized.slice(idx, idx + 240);
    const match = segment.match(DATE_PATTERN);
    if (match) {
      const cruda = match[0].replace(/\s+/g, "");
      const fecha = normalizeFecha(cruda);
      if (fecha) return { fecha, cruda };
    }
  }

  return { fecha: "", cruda: "" };
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function computeConfidenceScore({ ocrConfidence = 0, width = 0, height = 0, source = "ocr", matchedKeyword = false }) {
  if (source === "pdf-texto") return 10;

  const pixels = width * height;
  let resolutionScore = 1;
  if (pixels >= 4_000_000) resolutionScore = 10;
  else if (pixels >= 2_500_000) resolutionScore = 9;
  else if (pixels >= 1_500_000) resolutionScore = 8;
  else if (pixels >= 1_000_000) resolutionScore = 7;
  else if (pixels >= 700_000) resolutionScore = 6;
  else if (pixels >= 450_000) resolutionScore = 5;
  else if (pixels >= 250_000) resolutionScore = 4;
  else if (pixels >= 120_000) resolutionScore = 3;
  else resolutionScore = 2;

  const tesseractScore = clamp(Math.round((ocrConfidence || 0) / 10), 1, 10);
  const keywordBonus = matchedKeyword ? 1 : 0;
  return clamp(Math.round((resolutionScore * 0.45) + (tesseractScore * 0.55) + keywordBonus), 1, 10);
}

async function loadTargetRadicados() {
  try {
    const content = await fs.readFile(CONFIG.targetRadicadosPath, "utf-8");
    const out = [];
    for (const line of content.split(/\r?\n/)) {
      const value = normalizeRadicado(line.trim());
      if (/^\d{6,15}$/.test(value)) out.push(value);
    }
    return [...new Set(out)];
  } catch {
    const children = await fs.readdir(CONFIG.rootPath, { withFileTypes: true });
    return children
      .filter((item) => item.isDirectory())
      .map((item) => normalizeRadicado(item.name))
      .filter((value) => /^\d{6,15}$/.test(value))
      .sort();
  }
}

async function listFilesRecursive(dir) {
  const out = [];
  const stack = [dir];

  while (stack.length) {
    const current = stack.pop();
    const items = await fs.readdir(current, { withFileTypes: true });
    for (const item of items) {
      const full = path.join(current, item.name);
      if (item.isDirectory()) {
        stack.push(full);
        continue;
      }
      if (!item.isFile()) continue;
      if (CONFIG.extensions.has(path.extname(item.name).toLowerCase())) out.push(full);
    }
  }

  return out;
}

function sortByPriority(files) {
  const rank = (file) => {
    const normalizedName = normalizeToken(path.basename(file));
    for (let i = 0; i < filePriorityTokens.length; i += 1) {
      if (normalizedName.includes(filePriorityTokens[i])) return i;
    }
    return filePriorityTokens.length;
  };

  return [...files].sort((a, b) => {
    const ra = rank(a);
    const rb = rank(b);
    if (ra !== rb) return ra - rb;
    return a.localeCompare(b);
  });
}

async function getImageMetrics(filePath, page = 0) {
  try {
    const metadata = await sharp(filePath, { page }).metadata();
    return {
      width: metadata.width || 0,
      height: metadata.height || 0,
      pages: metadata.pages || 1
    };
  } catch {
    return { width: 0, height: 0, pages: 1 };
  }
}

async function recognizeBuffer(worker, imageBuffer) {
  const result = await worker.recognize(imageBuffer);
  return {
    text: result.data?.text || "",
    confidence: result.data?.confidence || 0
  };
}

async function extractPdf(filePath, worker) {
  try {
    const pdfBuffer = await fs.readFile(filePath);
    const data = await pdfParse(pdfBuffer, { max: CONFIG.maxPdfTextPages });
    const byText = extractFechaFromText(data.text || "");
    if (byText.fecha) {
      return {
        fecha: byText.fecha,
        cruda: byText.cruda,
        pagina: 0,
        metodo: "pdf-texto",
        motivo: "Fecha detectada por texto del PDF",
        score: 10,
        ocrConfidence: 100
      };
    }
  } catch {
    // Pasa a OCR.
  }

  try {
    const metrics = await getImageMetrics(filePath, 0);
    const totalPages = Math.max(1, metrics.pages || 1);
    const pages = Math.min(totalPages, CONFIG.maxOcrPages);

    for (let page = 0; page < pages; page += 1) {
      const currentMetrics = await getImageMetrics(filePath, page);
      const imageBuffer = await sharp(filePath, { density: 220, page }).png().toBuffer();
      const result = await recognizeBuffer(worker, imageBuffer);
      const parsed = extractFechaFromText(result.text);
      if (parsed.fecha) {
        const score = computeConfidenceScore({
          ocrConfidence: result.confidence,
          width: currentMetrics.width,
          height: currentMetrics.height,
          source: "pdf-ocr",
          matchedKeyword: result.text.toLowerCase().includes(CONFIG.ingresoKeyword)
        });
        return {
          fecha: parsed.fecha,
          cruda: parsed.cruda,
          pagina: page + 1,
          metodo: "pdf-ocr",
          motivo: "Fecha detectada por OCR de PDF",
          score,
          ocrConfidence: result.confidence
        };
      }
    }
  } catch {
    return {
      fecha: "",
      cruda: "",
      pagina: 0,
      metodo: "pdf-error",
      motivo: "No se pudo procesar el PDF",
      score: 1,
      ocrConfidence: 0
    };
  }

  return {
    fecha: "",
    cruda: "",
    pagina: 0,
    metodo: "pdf-sin-fecha",
    motivo: "No se detecto fecha en PDF",
    score: 1,
    ocrConfidence: 0
  };
}

async function extractTiff(filePath, worker) {
  try {
    const metadata = await sharp(filePath).metadata();
    const totalPages = Math.max(1, metadata.pages || 1);
    const pages = Math.min(totalPages, CONFIG.maxOcrPages);

    for (let page = 0; page < pages; page += 1) {
      const currentMetrics = await getImageMetrics(filePath, page);
      const imageBuffer = await sharp(filePath, { page }).png().toBuffer();
      const result = await recognizeBuffer(worker, imageBuffer);
      const parsed = extractFechaFromText(result.text);
      if (parsed.fecha) {
        const score = computeConfidenceScore({
          ocrConfidence: result.confidence,
          width: currentMetrics.width,
          height: currentMetrics.height,
          source: "tiff-ocr",
          matchedKeyword: result.text.toLowerCase().includes(CONFIG.ingresoKeyword)
        });
        return {
          fecha: parsed.fecha,
          cruda: parsed.cruda,
          pagina: page + 1,
          metodo: "tiff-ocr",
          motivo: "Fecha detectada por OCR de multi-TIF",
          score,
          ocrConfidence: result.confidence
        };
      }
    }
  } catch {
    return {
      fecha: "",
      cruda: "",
      pagina: 0,
      metodo: "tiff-error",
      motivo: "No se pudo procesar el multi-TIF",
      score: 1,
      ocrConfidence: 0
    };
  }

  return {
    fecha: "",
    cruda: "",
    pagina: 0,
    metodo: "tiff-sin-fecha",
    motivo: "No se detecto fecha en multi-TIF",
    score: 1,
    ocrConfidence: 0
  };
}

async function extractImage(filePath, worker) {
  try {
    const metrics = await getImageMetrics(filePath, 0);
    const imageBuffer = await sharp(filePath).png().toBuffer();
    const result = await recognizeBuffer(worker, imageBuffer);
    const parsed = extractFechaFromText(result.text);
    if (parsed.fecha) {
      const score = computeConfidenceScore({
        ocrConfidence: result.confidence,
        width: metrics.width,
        height: metrics.height,
        source: "imagen-ocr",
        matchedKeyword: result.text.toLowerCase().includes(CONFIG.ingresoKeyword)
      });
      return {
        fecha: parsed.fecha,
        cruda: parsed.cruda,
        pagina: 1,
        metodo: "imagen-ocr",
        motivo: "Fecha detectada por OCR de imagen",
        score,
        ocrConfidence: result.confidence
      };
    }
  } catch {
    return {
      fecha: "",
      cruda: "",
      pagina: 0,
      metodo: "imagen-error",
      motivo: "No se pudo procesar la imagen",
      score: 1,
      ocrConfidence: 0
    };
  }

  return {
    fecha: "",
    cruda: "",
    pagina: 0,
    metodo: "imagen-sin-fecha",
    motivo: "No se detecto fecha en imagen",
    score: 1,
    ocrConfidence: 0
  };
}

async function extractFromFile(filePath, worker) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".pdf") return extractPdf(filePath, worker);
  if (ext === ".tif" || ext === ".tiff") return extractTiff(filePath, worker);
  return extractImage(filePath, worker);
}

function csvEscape(value) {
  const raw = String(value ?? "");
  if (raw.includes(",") || raw.includes("\n") || raw.includes("\r") || raw.includes('"')) {
    return `"${raw.replace(/"/g, '""')}"`;
  }
  return raw;
}

async function appendLine(filePath, line) {
  for (let i = 0; i < 8; i += 1) {
    try {
      await fs.appendFile(filePath, `${line}\n`, "utf-8");
      return;
    } catch {
      await new Promise((resolve) => setTimeout(resolve, 200));
    }
  }
}

function formatProgress(processed, total, found, notFound, averageScore) {
  return `[${processed}/${total}] encontradas=${found} sin_fecha=${notFound} score_prom=${averageScore.toFixed(1)}`;
}

async function run() {
  const rootExists = await fs.access(CONFIG.rootPath).then(() => true).catch(() => false);
  if (!rootExists) {
    throw new Error(`No existe la ruta raiz: ${CONFIG.rootPath}`);
  }

  const radicados = await loadTargetRadicados();
  if (!radicados.length) {
    throw new Error("No se encontraron radicados para procesar");
  }

  await fs.writeFile(CONFIG.outputTxtPath, "", "utf-8");
  await fs.writeFile(
    CONFIG.outputDetalleCsvPath,
    "radicado,ingreso,score,estado,metodo,archivo_origen,pagina,fecha_cruda,ocr_confidence,motivo\n",
    "utf-8"
  );

  console.log(`Ruta: ${CONFIG.rootPath}`);
  console.log(`Archivo salida TXT: ${CONFIG.outputTxtPath}`);
  console.log(`Palabra clave archivo: ${CONFIG.fileKeyword}`);
  console.log(`Palabra clave ingreso: ${CONFIG.ingresoKeyword}`);

  const worker = await createWorker(CONFIG.language);
  const results = [];
  let found = 0;
  let notFound = 0;
  let totalScore = 0;

  try {
    for (let index = 0; index < radicados.length; index += 1) {
      const radicado = radicados[index];
      const carpeta = path.join(CONFIG.rootPath, radicado);

      let result = {
        radicado,
        ingreso: "",
        score: 1,
        estado: "sin_carpeta",
        metodo: "",
        archivo_origen: "",
        pagina: 0,
        fecha_cruda: "",
        ocrConfidence: 0,
        motivo: "No existe carpeta del radicado"
      };

      const exists = await fs.access(carpeta).then(() => true).catch(() => false);
      if (exists) {
        const files = sortByPriority(await listFilesRecursive(carpeta)).slice(0, CONFIG.maxFilesPerRadicado);
        if (!files.length) {
          result = { ...result, estado: "sin_archivos", motivo: "No hay archivos soportados" };
        } else {
          const reasons = [];
          for (const filePath of files) {
            const extracted = await extractFromFile(filePath, worker);
            if (extracted.fecha) {
              result = {
                radicado,
                ingreso: extracted.fecha,
                score: extracted.score,
                estado: "ok",
                metodo: extracted.metodo,
                archivo_origen: filePath,
                pagina: extracted.pagina,
                fecha_cruda: extracted.cruda,
                ocrConfidence: Math.round(extracted.ocrConfidence || 0),
                motivo: extracted.motivo
              };
              break;
            }
            reasons.push(`${path.basename(filePath)}: ${extracted.metodo}`);
          }

          if (!result.ingreso) {
            result = {
              radicado,
              ingreso: "",
              score: 1,
              estado: "sin_fecha",
              metodo: "",
              archivo_origen: files[0],
              pagina: 0,
              fecha_cruda: "",
              ocrConfidence: 0,
              motivo: reasons.slice(0, 4).join(" | ")
            };
          }
        }
      }

      results.push(result);
      if (result.ingreso) {
        found += 1;
        totalScore += result.score;
      } else {
        notFound += 1;
      }

      await appendLine(CONFIG.outputTxtPath, `${result.radicado},${result.ingreso}`);
      await appendLine(
        CONFIG.outputDetalleCsvPath,
        [
          result.radicado,
          result.ingreso,
          result.score,
          result.estado,
          result.metodo,
          result.archivo_origen,
          result.pagina,
          result.fecha_cruda,
          result.ocrConfidence,
          result.motivo
        ].map(csvEscape).join(",")
      );

      const averageScore = found ? totalScore / found : 0;
      const resumen = result.ingreso
        ? `${result.radicado} -> ${result.ingreso} | score=${result.score}/10 | ${result.metodo}`
        : `${result.radicado} -> SIN_FECHA`;
      console.log(`${formatProgress(index + 1, radicados.length, found, notFound, averageScore)} | ${resumen}`);
    }
  } finally {
    await worker.terminate();
  }

  const averageScore = found ? totalScore / found : 0;
  console.log("Proceso finalizado");
  console.log(`Total radicados: ${results.length}`);
  console.log(`Con fecha: ${found}`);
  console.log(`Sin fecha: ${notFound}`);
  console.log(`Score promedio: ${averageScore.toFixed(1)}/10`);
  console.log(`TXT: ${CONFIG.outputTxtPath}`);
  console.log(`Detalle: ${CONFIG.outputDetalleCsvPath}`);
}

run().catch((error) => {
  console.error("Fallo extractor JS:", error.message || error);
  process.exitCode = 1;
});
