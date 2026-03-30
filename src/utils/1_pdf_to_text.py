# 01_pdf_to_text.py
# =================
# Step 1 of the Fondazione del Monte extraction pipeline.
 
# For each project folder, converts all PDFs into a single clean .txt file
# ready to be uploaded to ChatGPT.
 
# HOW TO USE
# ----------
# 1. Install dependencies:
#         pip install pypdf pdfplumber Pillow pytesseract pypdfium2
 
# 2. Install Tesseract:
#         Download and run the installer from:
#         https://github.com/UB-Mannheim/tesseract/wiki
#         Use the default install location.
 
# 3. Organise your files like this:
 
#         projects/
#             casa_adolescente/
#                 Modulo-rendicontazione-parte-1.pdf
#                 Modulo-rendicontazione-parte-2.pdf
#                 Template_comunicazione.pdf
#             altro_progetto/
#                 ...
 
#     (Excel/XLS files in the same folders are ignored here —
#      they are handled by script 02)
 
# 4. Set INPUT_DIR and OUTPUT_DIR below if needed (defaults should work
#    if you run the script from the same folder as the projects/ directory).
 
# 5. Run:
#         python 01_pdf_to_text.py
 
# 6. Output: one .txt file per project in output_texts/
#            e.g.  output_texts/casa_adolescente.txt

 
import os
import sys
from pathlib import Path
 
import pypdfium2 as pdfium
import pdfplumber
from pypdf import PdfReader
from PIL import Image
import pytesseract
 
# ── Configuration ─────────────────────────────────────────────────────────────
 
INPUT_DIR  = "src/utils/projects"       # folder containing one sub-folder per project
OUTPUT_DIR = "src/utils/output_texts"   # where the .txt files will be written
 
# Tesseract path — update this if your Tesseract is installed elsewhere.
# The path below is the default for the UB Mannheim Windows installer.
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
 
# Tesseract language.
# "eng" works well for Italian text too.
# If you installed the Italian language pack during Tesseract setup,
# you can change this to "ita+eng" for slightly better accent handling.
OCR_LANG = "ita+eng"
 
# Resolution for rasterising scanned pages before OCR.
# 300 DPI is recommended — good quality without being too slow.
OCR_DPI = 300
 
# A page with fewer than this many characters from direct extraction
# is treated as scanned and sent to OCR instead.
TEXT_MIN_CHARS = 50
 
# ── Setup ─────────────────────────────────────────────────────────────────────
 
# Point pytesseract at the Tesseract executable
if os.path.exists(TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
else:
    print(f"  ⚠  Tesseract not found at '{TESSERACT_PATH}'.")
    print("     If Tesseract is on your PATH, this may still work.")
    print("     Otherwise update TESSERACT_PATH at the top of the script.\n")
 
# ── Helpers ───────────────────────────────────────────────────────────────────
 
def is_text_extractable(pdf_path: Path) -> bool:
    
    # Return True if the first page of the PDF has extractable text.
    # Returns False for scanned PDFs (pages are images with no text layer).
    
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            if not pdf.pages:
                return False
            text = pdf.pages[0].extract_text() or ""
            return len(text.strip()) >= TEXT_MIN_CHARS
    except Exception:
        return False
 
 
def extract_text_pdfplumber(pdf_path: Path) -> str:
    
    # Extract text from a text-based PDF using pdfplumber.
    # Also extracts tables and formats them as plain text grids.
    # Falls back to pypdf if pdfplumber fails.
    
    pages_text = []
 
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                text = page.extract_text() or ""
 
                # Extract tables and append as readable grids
                tables = page.extract_tables() or []
                table_blocks = []
                for table in tables:
                    rows = []
                    for row in table:
                        cells = [str(c or "").strip() for c in row]
                        rows.append(" | ".join(cells))
                    table_blocks.append("\n".join(rows))
 
                page_content = text
                if table_blocks:
                    page_content += "\n\n[TABLE]\n" + "\n\n[TABLE]\n".join(table_blocks)
 
                pages_text.append(f"--- Page {i} ---\n{page_content.strip()}")
 
    except Exception as e:
        print(f"    ⚠  pdfplumber failed ({e}), trying pypdf ...")
        try:
            reader = PdfReader(str(pdf_path))
            for i, page in enumerate(reader.pages, 1):
                pages_text.append(f"--- Page {i} ---\n{page.extract_text() or ''}")
        except Exception as e2:
            print(f"    ✗  pypdf also failed: {e2}")
 
    return "\n\n".join(pages_text)
 
 
def rasterise_pdf_pages(pdf_path: Path, dpi: int) -> list:
    
    # Use pypdfium2 to render each page of a PDF to a PIL Image.
    # Returns a list of PIL Images, one per page.
    
    images = []
    try:
        doc = pdfium.PdfDocument(str(pdf_path))
        scale = dpi / 72  # pypdfium2 uses 72 DPI as its base unit
        for page in doc:
            bitmap = page.render(scale=scale, rotation=0)
            pil_image = bitmap.to_pil()
            images.append(pil_image)
        doc.close()
    except Exception as e:
        print(f"    ✗  Failed to rasterise {pdf_path.name}: {e}")
    return images
 
 
def ocr_pdf(pdf_path: Path) -> str:
    
    # Rasterise each page of a scanned PDF and run Tesseract OCR on it.
    # Returns the full extracted text.
    
    pages_text = []
 
    print(f"    → Rasterising pages at {OCR_DPI} DPI ...")
    images = rasterise_pdf_pages(pdf_path, OCR_DPI)
 
    if not images:
        return ""
 
    for i, img in enumerate(images, 1):
        try:
            # Grayscale conversion helps Tesseract accuracy
            if img.mode != "L":
                img = img.convert("L")
            text = pytesseract.image_to_string(img, lang=OCR_LANG)
            pages_text.append(f"--- Page {i} ---\n{text.strip()}")
            print(f"    ✓  OCR page {i}/{len(images)}")
        except Exception as e:
            print(f"    ⚠  OCR failed on page {i}: {e}")
 
    return "\n\n".join(pages_text)
 
 
def process_pdf(pdf_path: Path) -> str:
    
    # Auto-detect whether a PDF is text-based or scanned, then extract
    # text using the appropriate method.
    
    print(f"  Processing: {pdf_path.name}")
 
    if is_text_extractable(pdf_path):
        print(f"    → Text-based PDF, extracting directly ...")
        return extract_text_pdfplumber(pdf_path)
    else:
        print(f"    → Scanned PDF, using OCR ...")
        return ocr_pdf(pdf_path)
 
 
def build_project_text(project_dir: Path) -> str:
    
    # Process all PDFs in a project folder and combine them into one
    # structured text block.
    
    sections = []
    sections.append(f"PROJECT: {project_dir.name}\n{'=' * 60}\n")
 
    pdf_files = sorted(
        [f for f in project_dir.iterdir() if f.suffix.lower() == ".pdf"]
    )
 
    if not pdf_files:
        print(f"  ⚠  No PDF files found in '{project_dir.name}'")
        return ""
 
    for pdf_path in pdf_files:
        text = process_pdf(pdf_path)
        if text.strip():
            sections.append(
                f"\n{'─' * 60}\n"
                f"SOURCE FILE: {pdf_path.name}\n"
                f"{'─' * 60}\n"
                f"{text}\n"
            )
        else:
            print(f"    ⚠  No text extracted from {pdf_path.name}")
 
    return "\n".join(sections)
 
 
# ── Main ──────────────────────────────────────────────────────────────────────
 
def main():
    input_dir  = Path(INPUT_DIR)
    output_dir = Path(OUTPUT_DIR)
 
    if not input_dir.exists():
        print(f"\nERROR: Input directory '{INPUT_DIR}' not found.")
        print("Please create it and place one sub-folder per project inside it.")
        print("Example structure:")
        print("  projects/")
        print("      casa_adolescente/")
        print("          file1.pdf")
        print("          file2.pdf")
        sys.exit(1)
 
    output_dir.mkdir(parents=True, exist_ok=True)
 
    project_dirs = sorted([d for d in input_dir.iterdir() if d.is_dir()])
 
    if not project_dirs:
        print(f"No project sub-folders found in '{INPUT_DIR}'.")
        sys.exit(1)
 
    print(f"\nFound {len(project_dirs)} project(s) in '{INPUT_DIR}'\n")
 
    for project_dir in project_dirs:
        print(f"\n{'━' * 60}")
        print(f"PROJECT: {project_dir.name}")
        print(f"{'━' * 60}")
 
        combined_text = build_project_text(project_dir)
 
        if combined_text.strip():
            out_path = output_dir / f"{project_dir.name}.txt"
            out_path.write_text(combined_text, encoding="utf-8")
            size_kb = out_path.stat().st_size / 1024
            print(f"\n  ✓  Saved → {out_path}  ({size_kb:.1f} KB)")
        else:
            print(f"\n  ✗  No text extracted for '{project_dir.name}', skipping.")
 
    print(f"\n{'━' * 60}")
    print(f"Done. Text files saved to: {output_dir}/")
    print("Next step: run 02_excel_to_text.py to process the Excel files.")
 
 
if __name__ == "__main__":
    main()