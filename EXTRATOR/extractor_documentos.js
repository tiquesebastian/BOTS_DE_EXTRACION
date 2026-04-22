import fs from "node:fs/promises";
import path from "node:path";
import { createWorker } from "tesseract.js";
import sharp from "sharp";
import pdfParse from "pdf-parse";

const PROJECT_ROOT = path.resolve("..");

const CONFIG = {
  rootPath: process.env.DOC_ROOT_PATH || "Z:/IA 10/NUEVO/PARTE3/890701715",
  targetRadicadosPath: process.env.DOC_TARGET_FILE || path.join(PROJECT_ROOT, "numeros_pendientes"),
  outputCsvPath: process.env.DOC_OUTPUT_CSV || path.join(PROJECT_ROOT, "890701715_documentos.csv"),
  outputDetailCsvPath:
    process.env.DOC_OUTPUT_DETAIL || path.join(PROJECT_ROOT, "890701715_documentos_detalle.csv"),
  language: process.env.DOC_OCR_LANG || "spa",
  fastScanPages: Number(process.env.DOC_FAST_SCAN_PAGES || 10),
  maxFilesPerRadicado: Number(process.env.DOC_MAX_FILES || 12),
  fileKeyword: normalizeToken(process.env.DOC_FILE_KEYWORD || "imprimir liquidacion tapa hoja de atencion factura"),
  extensions: new Set([".pdf", ".tif", ".tiff", ".png", ".jpg", ".jpeg", ".bmp", ".gif"]),
  blockedDocuments: new Set(["890701715"])
};

const IDENT_USER_PATTERN = /(?:nombre\s+usuario|usuario|estado\s+afiliacion\s+usuario).{0,260}?identif[a-z]{2,18}\s*[:\-]?\s*(\d{6,15})/i;
const IDENT_PATTERN = /(?:tipo\s+)?identif[a-z]{2,18}\s*[:\-]?\s*(\d{6,15})/i;
const TYPE_DOC_PATTERN = /\b(?:cc|ti|ce|rc|pa|pep|ppt|dni)\s*[-:]?\s*(\d{6,15})\b/i;
const GENERIC_PATTERN = /\b(\d{7,15})\b/g;

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
  const m = String(input || "").match(/\d{6,15}/);
  return m ? m[0] : String(input || "").trim();
}

function normalizeDoc(input) {
  const m = String(input || "").match(/\d{6,15}/);
  return m ? m[0] : "";
}

function compactDigits(text) {
  return String(text || "").replace(/(?<=\d)\s+(?=\d)/g, "");
}

function isLikelyDateNumber(value) {
  if (!/^\d{8}$/.test(value)) return false;

  const y = Number(value.slice(0, 4));
  const m = Number(value.slice(4, 6));
  const d = Number(value.slice(6, 8));
  if (y >= 1900 && y <= 2099 && m >= 1 && m <= 12 && d >= 1 && d <= 31) return true;

  const d2 = Number(value.slice(0, 2));
  const m2 = Number(value.slice(2, 4));
  const y2 = Number(value.slice(4, 8));
  return d2 >= 1 && d2 <= 31 && m2 >= 1 && m2 <= 12 && y2 >= 1900 && y2 <= 2099;
}

function isInstitutionalContext(text, number) {
  const idx = text.indexOf(number);
  if (idx < 0) return false;
  const window = text.slice(Math.max(0, idx - 90), Math.min(text.length, idx + number.length + 90));
  return ["nit", "hospital", "ips", "prestador", "razon social", "empresa"].some((w) => window.includes(w));
}

function isValidDocument(doc, radicado, text) {
  if (!doc) return false;
  if (doc === radicado) return false;
  if (CONFIG.blockedDocuments.has(doc)) return false;
  if (isLikelyDateNumber(doc)) return false;
  if (isInstitutionalContext(text, doc)) return false;
  return true;
}

function extractDocumentFromText(text, radicado) {
  const clean = compactDigits(normalizeToken(String(text || "").replace(/\x00/g, " ")));
  if (!clean) return { documento: "", score: 1, metodo: "vacio", motivo: "Texto vacio" };

  const m1 = clean.match(IDENT_USER_PATTERN);
  if (m1) {
    const doc = normalizeDoc(m1[1]);
    if (isValidDocument(doc, radicado, clean)) {
      return { documento: doc, score: 10, metodo: "usuario-identificacion", motivo: "Usuario + identificacion" };
    }
  }

  const m2 = clean.match(IDENT_PATTERN);
  if (m2) {
    const doc = normalizeDoc(m2[1]);
    if (isValidDocument(doc, radicado, clean)) {
      const plus = clean.includes("imprimir liquidacion") ? 1 : 0;
      return { documento: doc, score: 9 + plus, metodo: "identificacion", motivo: "Etiqueta identificacion" };
    }
  }

  const m3 = clean.match(TYPE_DOC_PATTERN);
  if (m3) {
    const doc = normalizeDoc(m3[1]);
    if (isValidDocument(doc, radicado, clean)) {
      return { documento: doc, score: 8, metodo: "tipo-doc", motivo: "Tipo de documento (CC/TI/CE/etc)" };
    }
  }

  let generic;
  while ((generic = GENERIC_PATTERN.exec(clean)) !== null) {
    const doc = normalizeDoc(generic[1]);
    if (isValidDocument(doc, radicado, clean)) {
      return { documento: doc, score: 5, metodo: "generico", motivo: "Numero generico filtrado" };
    }
  }

  return { documento: "", score: 1, metodo: "sin_match", motivo: "No se detecto documento" };
}

async function loadTargetRadicados() {
  const content = await fs.readFile(CONFIG.targetRadicadosPath, "utf-8");
  const out = [];
  for (const line of content.split(/\r?\n/)) {
    const value = normalizeRadicado(line.trim());
    if (/^\d{6,15}$/.test(value)) out.push(value);
  }
  return [...new Set(out)];
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
  const tokens = [...new Set(CONFIG.fileKeyword.split(" ").filter(Boolean))];
  const rank = (file) => {
    const name = normalizeToken(path.basename(file));
    for (let i = 0; i < tokens.length; i += 1) {
      if (name.includes(tokens[i])) return i;
    }
    return tokens.length;
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
    return { width: metadata.width || 0, height: metadata.height || 0, pages: metadata.pages || 1 };
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

function scoreFromConfidence(base, confidence) {
  if (confidence >= 92) return Math.min(10, base + 1);
  if (confidence <= 40) return Math.max(1, base - 1);
  return base;
}

async function extractFromPdf(filePath, radicado, worker) {
  try {
    const pdfBuffer = await fs.readFile(filePath);
    const data = await pdfParse(pdfBuffer);
    const parsed = extractDocumentFromText(data.text || "");
    if (parsed.documento) {
      return {
        documento: parsed.documento,
        score: 10,
        metodo: `pdf-texto-${parsed.metodo}`,
        motivo: parsed.motivo,
        pagina: 0,
        ocrConfidence: 100
      };
    }
  } catch {
    // pasa a OCR
  }

  const metrics = await getImageMetrics(filePath, 0);
  const totalPages = Math.max(1, metrics.pages || 1);
  const fastPages = Math.min(totalPages, CONFIG.fastScanPages);

  const scanRange = async (start, end) => {
    for (let page = start; page < end; page += 1) {
      const imageBuffer = await sharp(filePath, { density: 220, page }).png().toBuffer();
      const result = await recognizeBuffer(worker, imageBuffer);
      const parsed = extractDocumentFromText(result.text, radicado);
      if (parsed.documento) {
        return {
          documento: parsed.documento,
          score: scoreFromConfidence(parsed.score, result.confidence),
          metodo: `pdf-ocr-${parsed.metodo}`,
          motivo: parsed.motivo,
          pagina: page + 1,
          ocrConfidence: Math.round(result.confidence)
        };
      }
    }
    return null;
  };

  const fastResult = await scanRange(0, fastPages);
  if (fastResult) return fastResult;

  if (fastPages < totalPages) {
    const deepResult = await scanRange(fastPages, totalPages);
    if (deepResult) return deepResult;
  }

  return {
    documento: "",
    score: 1,
    metodo: "pdf-sin-documento",
    motivo: "No se detecto documento en PDF",
    pagina: 0,
    ocrConfidence: 0
  };
}

async function extractFromTiff(filePath, radicado, worker) {
  try {
    const metadata = await sharp(filePath).metadata();
    const totalPages = Math.max(1, metadata.pages || 1);
    const fastPages = Math.min(totalPages, CONFIG.fastScanPages);

    const scanRange = async (start, end) => {
      for (let page = start; page < end; page += 1) {
        const imageBuffer = await sharp(filePath, { page }).png().toBuffer();
        const result = await recognizeBuffer(worker, imageBuffer);
        const parsed = extractDocumentFromText(result.text, radicado);
        if (parsed.documento) {
          return {
            documento: parsed.documento,
            score: scoreFromConfidence(parsed.score, result.confidence),
            metodo: `tiff-ocr-${parsed.metodo}`,
            motivo: parsed.motivo,
            pagina: page + 1,
            ocrConfidence: Math.round(result.confidence)
          };
        }
      }
      return null;
    };

    const fastResult = await scanRange(0, fastPages);
    if (fastResult) return fastResult;

    if (fastPages < totalPages) {
      const deepResult = await scanRange(fastPages, totalPages);
      if (deepResult) return deepResult;
    }

    return {
      documento: "",
      score: 1,
      metodo: "tiff-sin-documento",
      motivo: "No se detecto documento en multi-TIF",
      pagina: 0,
      ocrConfidence: 0
    };
  } catch {
    return {
      documento: "",
      score: 1,
      metodo: "tiff-error",
      motivo: "No se pudo procesar multi-TIF",
      pagina: 0,
      ocrConfidence: 0
    };
  }
}

async function extractFromImage(filePath, radicado, worker) {
  try {
    const imageBuffer = await sharp(filePath).png().toBuffer();
    const result = await recognizeBuffer(worker, imageBuffer);
    const parsed = extractDocumentFromText(result.text, radicado);
    if (parsed.documento) {
      return {
        documento: parsed.documento,
        score: scoreFromConfidence(parsed.score, result.confidence),
        metodo: `imagen-ocr-${parsed.metodo}`,
        motivo: parsed.motivo,
        pagina: 1,
        ocrConfidence: Math.round(result.confidence)
      };
    }
    return {
      documento: "",
      score: 1,
      metodo: "imagen-sin-documento",
      motivo: "No se detecto documento en imagen",
      pagina: 0,
      ocrConfidence: 0
    };
  } catch {
    return {
      documento: "",
      score: 1,
      metodo: "imagen-error",
      motivo: "No se pudo procesar imagen",
      pagina: 0,
      ocrConfidence: 0
    };
  }
}

async function extractFromFile(filePath, radicado, worker) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".pdf") return extractFromPdf(filePath, radicado, worker);
  if (ext === ".tif" || ext === ".tiff") return extractFromTiff(filePath, radicado, worker);
  return extractFromImage(filePath, radicado, worker);
}

function csvEscape(value) {
  const raw = String(value ?? "");
  if (raw.includes(",") || raw.includes("\n") || raw.includes("\r") || raw.includes('"')) {
    return `"${raw.replace(/"/g, '""')}"`;
  }
  return raw;
}

async function appendLine(filePath, line) {
  await fs.appendFile(filePath, `${line}\n`, "utf-8");
}

async function run() {
  const radicados = await loadTargetRadicados();
  if (!radicados.length) throw new Error("No se encontraron radicados objetivo");

  await fs.writeFile(CONFIG.outputCsvPath, "radicado,numero_documento,score_confianza\n", "utf-8");
  await fs.writeFile(
    CONFIG.outputDetailCsvPath,
    "radicado,numero_documento,score_confianza,estado,metodo,archivo_origen,pagina,ocr_confidence,motivo\n",
    "utf-8"
  );

  console.log(`Ruta: ${CONFIG.rootPath}`);
  console.log(`Objetivo: ${radicados.length} radicados`);
  console.log(`Salida CSV: ${CONFIG.outputCsvPath}`);

  const worker = await createWorker(CONFIG.language);

  let found = 0;
  let noDoc = 0;

  try {
    for (let i = 0; i < radicados.length; i += 1) {
      const radicado = radicados[i];
      const carpeta = path.join(CONFIG.rootPath, radicado);

      let row = {
        radicado,
        numero_documento: "SIN_DATO",
        score_confianza: 1,
        estado: "sin_carpeta",
        metodo: "",
        archivo_origen: "",
        pagina: 0,
        ocr_confidence: 0,
        motivo: "No existe carpeta del radicado"
      };

      const exists = await fs.access(carpeta).then(() => true).catch(() => false);
      if (exists) {
        const files = sortByPriority(await listFilesRecursive(carpeta));
        const toScan = CONFIG.maxFilesPerRadicado > 0 ? files.slice(0, CONFIG.maxFilesPerRadicado) : files;

        if (!toScan.length) {
          row = { ...row, estado: "sin_archivos", motivo: "No hay archivos soportados" };
        } else {
          const reasons = [];
          for (const filePath of toScan) {
            const ext = await extractFromFile(filePath, radicado, worker);
            if (ext.documento) {
              row = {
                radicado,
                numero_documento: ext.documento,
                score_confianza: ext.score,
                estado: "ok",
                metodo: ext.metodo,
                archivo_origen: filePath,
                pagina: ext.pagina,
                ocr_confidence: ext.ocrConfidence,
                motivo: ext.motivo
              };
              break;
            }
            reasons.push(`${path.basename(filePath)}:${ext.metodo}`);
          }

          if (row.estado !== "ok") {
            row = {
              radicado,
              numero_documento: "SIN_DATO",
              score_confianza: 1,
              estado: "sin_documento",
              metodo: "",
              archivo_origen: toScan[0],
              pagina: 0,
              ocr_confidence: 0,
              motivo: reasons.slice(0, 4).join(" | ") || "No se detecto documento"
            };
          }
        }
      }

      if (row.estado === "ok") found += 1;
      else noDoc += 1;

      await appendLine(CONFIG.outputCsvPath, `${row.radicado},${row.numero_documento},${row.score_confianza}`);
      await appendLine(
        CONFIG.outputDetailCsvPath,
        [
          row.radicado,
          row.numero_documento,
          row.score_confianza,
          row.estado,
          row.metodo,
          row.archivo_origen,
          row.pagina,
          row.ocr_confidence,
          row.motivo
        ]
          .map(csvEscape)
          .join(",")
      );

      const resumen = row.estado === "ok"
        ? `${row.radicado} -> ${row.numero_documento} | score=${row.score_confianza} | ${row.metodo}`
        : `${row.radicado} -> SIN_DATO`;
      console.log(`[${i + 1}/${radicados.length}] con_doc=${found} sin_doc=${noDoc} | ${resumen}`);
    }
  } finally {
    await worker.terminate();
  }

  console.log("Proceso finalizado");
  console.log(`Con documento: ${found}`);
  console.log(`Sin documento: ${noDoc}`);
  console.log(`CSV: ${CONFIG.outputCsvPath}`);
  console.log(`Detalle: ${CONFIG.outputDetailCsvPath}`);
}

run().catch((error) => {
  console.error("Fallo extractor documentos JS:", error.message || error);
  process.exitCode = 1;
});
