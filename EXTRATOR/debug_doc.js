import { createWorker } from "tesseract.js";
import sharp from "sharp";

const filePath = "Z:/IA 10/NUEVO/PARTE3/890701715/243612940/Tapa__243612940_477124081_0.tif";

const worker = await createWorker("spa");
try {
  const imageBuffer = await sharp(filePath, { page: 0 }).png().toBuffer();
  const result = await worker.recognize(imageBuffer);
  const text = (result.data?.text || "").replace(/\s+/g, " ").trim();
  const nums = [...text.matchAll(/\d{7,15}/g)].map((m) => m[0]).slice(0, 12);
  console.log("chars", text.length);
  console.log("nums", nums);
  console.log(text.slice(0, 900));
} finally {
  await worker.terminate();
}
