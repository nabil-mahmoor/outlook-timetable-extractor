import os
import fitz
from config import KEYWORDS
fitz.TOOLS.mupdf_display_errors(False)

def find_timetable_page(pdf_path) -> int | None:
    pdf = fitz.open(pdf_path)

    for page_index in range(len(pdf)):
        page = pdf.load_page(page_index)
        text = page.get_text()

        if all(keyword.lower() in text.lower() for keyword in KEYWORDS):
            print(f"Found timetable on page {page_index + 1}")
            pdf.close()
            return page_index
            
    pdf.close()
    
def save_page_as_image(pdf_path, page_index, date_str, output_folder="output"):
    pdf = fitz.open(pdf_path)
    page = pdf.load_page(page_index)

    pix = page.get_pixmap(dpi=200)

    os.makedirs(output_folder, exist_ok=True)

    output_path = os.path.join(output_folder, f"timetable_{date_str}.png")
    pix.save(output_path)
    pdf.close()

    print(f"Timetable image saved to {output_path}")
    return output_path