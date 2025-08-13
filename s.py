import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import os

# ===== CONFIG =====
PDF_FILE = "BuyingHouse.pdf"
OUTPUT_EXCEL = "members.xlsx"

# Create temp folder
TEMP_IMAGE_FOLDER = "temp_images"
os.makedirs(TEMP_IMAGE_FOLDER, exist_ok=True)

# ===== CONVERT PDF TO IMAGES =====
print("Converting PDF pages to images...")
try:
    pages = convert_from_path(PDF_FILE, dpi=300)
except Exception as e:
    print(f"❌ PDF conversion failed: {e}")
    print("Make sure poppler-utils is installed in your Codespace!")
    print("Run these commands in your terminal:")
    print("sudo apt-get update")
    print("sudo apt-get install poppler-utils -y")
    exit(1)

# ===== OCR EACH PAGE =====
all_text = []
for i, page in enumerate(pages):
    try:
        image_path = os.path.join(TEMP_IMAGE_FOLDER, f"page_{i+1}.png")
        page.save(image_path, "PNG")
        
        # OCR processing
        text = pytesseract.image_to_string(image_path)
        all_text.append([i+1, text.strip()])
        print(f"✅ Processed page {i+1}/{len(pages)}")
    except Exception as e:
        print(f"❌ Error processing page {i+1}: {e}")
        all_text.append([i+1, "ERROR PROCESSING PAGE"])

# ===== SAVE TO EXCEL =====
try:
    df = pd.DataFrame(all_text, columns=["Page Number", "Extracted Text"])
    df.to_excel(OUTPUT_EXCEL, index=False)
    print(f"✅ OCR complete. Data saved to '{OUTPUT_EXCEL}'.")
except Exception as e:
    print(f"❌ Failed to save Excel file: {e}")

# ===== CLEANUP =====
try:
    print("✅ Temporary files cleaned up")
except Exception as e:
    print(f"❌ Cleanup failed: {e}")