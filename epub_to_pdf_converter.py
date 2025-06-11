import os
from ebooklib import epub, ITEM_DOCUMENT
from bs4 import BeautifulSoup
from reportlab.lib.pagesizes import A4, LETTER, landscape, portrait
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

# Mapping numeric options
PAGE_SIZES = {
    1: ('A4', A4),
    2: ('Letter', LETTER)
}

ORIENTATIONS = {
    1: 'Portrait',
    2: 'Landscape'
}




def get_page_size(size_choice, orientation_choice):
    _, base_size = PAGE_SIZES.get(size_choice, ('A4', A4))
    return landscape(base_size) if orientation_choice == 2 else portrait(base_size)

def epub_to_text(epub_path):
    book = epub.read_epub(epub_path)
    full_text = ''
    for item in book.get_items():
        if item.get_type() == ITEM_DOCUMENT:
            soup = BeautifulSoup(item.get_body_content(), 'html.parser')
            full_text += soup.get_text() + '\n\n'
    return full_text





def is_heading(paragraph):
    # Simple heuristics: uppercase or starts with keywords
    if paragraph.isupper():
        return True
    for kw in ['chapter', 'section', 'part', 'chapter ', 'section ', 'part ']:
        if paragraph.lower().startswith(kw):
            return True
    return False

def text_to_pdf(text, output_pdf_path, page_size):
    c = canvas.Canvas(output_pdf_path, pagesize=page_size)
    width, height = page_size
    x_margin = 50
    y_margin = 50
    usable_width = width - 2 * x_margin

    # Register fonts if you want (using built-in Helvetica here)
    normal_font = 'Helvetica'
    bold_font = 'Helvetica-Bold'

    normal_font_size = 12
    heading_font_size = 16
    line_height = normal_font_size * 1.3  # 1.3 line spacing

    y = height - y_margin
    indent = 15  # indentation for paragraphs

    for paragraph in text.split('\n'):
        para = paragraph.strip()
        if not para:
            # Empty line: add gap and continue
            y -= line_height / 2
            continue

        if is_heading(para):
            # Heading style
            font = bold_font
            font_size = heading_font_size
            gap_after = line_height  # bigger gap after heading
            indent_x = x_margin
        else:
            # Normal paragraph style
            font = normal_font
            font_size = normal_font_size
            gap_after = line_height / 2
            indent_x = x_margin + indent

        # Wrap lines within usable width
        lines = simpleSplit(para, font, font_size, usable_width - (indent_x - x_margin))

        for line in lines:
            if y < y_margin:
                c.showPage()
                y = height - y_margin
            c.setFont(font, font_size)
            c.drawString(indent_x, y, line)
            y -= line_height

        y -= gap_after

    c.save()


# def text_to_pdf(text, output_pdf_path, page_size):
#     c = canvas.Canvas(output_pdf_path, pagesize=page_size)
#     width, height = page_size
#     x_margin = 50
#     y_margin = 50
#     line_height = 14
#     usable_width = width - 2 * x_margin
#     x = x_margin
#     y = height - y_margin

#     # Split the full text into lines that fit within the page
#     for paragraph in text.split('\n'):
#         lines = simpleSplit(paragraph, 'Helvetica', 12, usable_width)
#         for line in lines:
#             if y < y_margin:
#                 c.showPage()
#                 y = height - y_margin
#             c.drawString(x, y, line)
#             y -= line_height

#     c.save()


def convert_epubs(folder_path, orientation_choice, size_choice):
    page_size = get_page_size(size_choice, orientation_choice)
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.epub'):
            epub_path = os.path.join(folder_path, filename)
            base_name = os.path.splitext(filename)[0]
            output_pdf_path = os.path.join(folder_path, f"{base_name}.pdf")

            print(f"🔄 Converting: {filename} → {base_name}.pdf")
            text = epub_to_text(epub_path)
            text_to_pdf(text, output_pdf_path, page_size)
    print("\n✅ All EPUB files converted to PDF!")

def main():
    folder_path = input("📁 Enter the full path to the folder containing EPUB files: ").strip()

    if not os.path.isdir(folder_path):
        print("❌ Invalid folder path.")
        return

    print("\n📐 Select Page Orientation:")
    print("1: Portrait")
    print("2: Landscape")
    orientation_choice = int(input("Enter option (1 or 2): ").strip())

    print("\n📄 Select Page Size:")
    print("1: A4")
    print("2: Letter")
    size_choice = int(input("Enter option (1 or 2): ").strip())

    convert_epubs(folder_path, orientation_choice, size_choice)

if __name__ == "__main__":
    main()
