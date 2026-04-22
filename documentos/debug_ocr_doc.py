from PIL import Image, ImageOps, ImageFilter
import numpy as np
import re
from rapidocr_onnxruntime import RapidOCR

path = r"Z:\IA 10\NUEVO\PARTE3\890701715\243612940\Tapa__243612940_477124081_0.tif"
eng = RapidOCR()

with Image.open(path) as img:
    total = int(getattr(img, "n_frames", 1))
    print("frames", total)
    for i in range(min(total, 10)):
        img.seek(i)
        im = img.convert("RGB")
        g = ImageOps.grayscale(im)
        g = ImageOps.autocontrast(g)
        g = g.filter(ImageFilter.SHARPEN)
        arr = np.array(g)
        res, _ = eng(arr)
        lines = [r[1] for r in (res or [])]
        text = " ".join(lines)
        low = re.sub(r"\s+", " ", text).lower()
        nums = re.findall(r"\d{7,15}", low)
        has = any(k in low for k in ["ident", "usuario", "paciente", "cc", "ti"])
        print(f"page={i+1} chars={len(low)} nums={nums[:10]} keys={has}")
        if has:
            print(low[:800])
            print("---")
