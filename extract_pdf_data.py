import pdfplumber
from pdf2image import convert_from_path
import pytesseract

def extract_pdf_full_text(pdf_path, txt_path):
    raw_texts = []
    need_ocr = []
    # Step 1: Try to extract RAW text, record which pages need OCR
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            print(f"Extracting text from page {i+1}...")
            text = page.extract_text() or ""
            if text.strip():
                raw_texts.append(f"\n--- PAGE {i+1} RAW TEXT ---\n{text.strip()}")
            else:
                raw_texts.append(None)  # Mark that we need OCR for this page
                need_ocr.append(i)
    
    # Step 2: OCR only those pages with no RAW text
    print("Running OCR where RAW text is missing...")
    images = convert_from_path(pdf_path, dpi=300)
    for idx in need_ocr:
        ocr_text = pytesseract.image_to_string(images[idx])
        raw_texts[idx] = f"\n--- PAGE {idx+1} OCR TEXT ---\n{ocr_text.strip()}"
    
    # Step 3: Save to file (skip any leftover Nones, but there shouldn't be any)
    result = [txt for txt in raw_texts if txt]
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(result))
    print(f"✅ Saved deduped full text to {txt_path}")

if __name__ == "__main__":
    extract_pdf_full_text("test1.pdf", "pdf_all_text_full.txt")