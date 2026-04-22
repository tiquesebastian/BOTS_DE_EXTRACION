from pdf2image import convert_from_path
from PIL import ImageOps, ImageFilter
import numpy as np
from rapidocr_onnxruntime import RapidOCR
import re

pdf = r"Z:\IA 10\NUEVO\PARTE3\890701715\267217743\632396162_Tapa_-_267217743_999111543.pdf"
engine = RapidOCR()

pages = convert_from_path(pdf, dpi=220, first_page=1, last_page=2, poppler_path=r"C:\poppler-25.11.0\Library\bin")
for i, pg in enumerate(pages, start=1):
    g = ImageOps.grayscale(pg)
    g = ImageOps.autocontrast(g)
    g = g.filter(ImageFilter.SHARPEN)
    arr = np.array(g)
    res, _ = engine(arr)
    text = " ".join([r[1] for r in (res or [])])
    low = re.sub(r"\s+", " ", text).lower()
    print("\nPAGE", i, "chars", len(low))
    print("has cedula", "cedula" in low, "has identific", "ident" in low, "has usuario", "usuario" in low)
    print("nums", re.findall(r"\d{6,11}", low)[:30])
    print(low[:1200])
