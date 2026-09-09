from img2table.document import Image
from img2table.ocr import TesseractOCR

# --- Hardcoded input and output paths ---
IMAGE_PATH = "Callum.png"
OUTPUT_EXCEL = "Callum.xlsx"

# --- Initialize OCR engine (Tesseract) ---
ocr = TesseractOCR(lang="eng")

# --- Load the image ---
doc = Image(src=IMAGE_PATH)

# --- Extract tables ---
# as_excel=True tells img2table to generate an Excel file
doc.to_xlsx(
    ocr=ocr,
    output_path=OUTPUT_EXCEL
)

print(f"Extraction complete. Saved to {OUTPUT_EXCEL}")
