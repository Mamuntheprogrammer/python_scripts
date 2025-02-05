import fitz  # PyMuPDF

def crop_and_merge_pdf(input_pdf, output_pdf, selections):
    doc = fitz.open(input_pdf)
    new_doc = fitz.open()

    # Create a new page with the same size as the original
    ref_page = doc[0]  # Reference page
    page_width, page_height = ref_page.rect.width, ref_page.rect.height
    new_page = new_doc.new_page(width=page_width, height=page_height)
    print("Page Size:", page_width, "x", page_height)

    # Tracking placement of previous selection
    x_offset = 0  # Initial X position
    y_offset = 0  # Tracks vertical placement of sections

    for idx, sel in enumerate(selections):
        page_num, rect = sel['page'], sel['rect']
        # page_num, rect = 0, fitz.Rect(0, 0, 1056, 75)
        if idx==1:
            rect = fitz.Rect(0+20, 76+5, 102+20, 814+5)

        page = doc[page_num]

        # Adjust selection 3 and 4 positions
        if idx == 2:  # Selection 3 should be next to Selection 2
            rect = fitz.Rect(102+20, 76+5, 438+20, 814+5)

        elif idx == 3:  # Selection 4 should be next to Selection 3
            rect = fitz.Rect(15+20, 68, 1056+20, 809.5)

        # # Adjust vertical placement after the first section
        # if idx > 0 and idx < 4:
        #     y_offset += 5 # Add vertical gap
        #     rect = fitz.Rect(rect.x0, rect.y0 + y_offset, rect.x1, rect.y1 + y_offset)

        # Extract and paste content with full quality
        new_page.show_pdf_page(rect, doc, page_num, clip=sel['rect'])  # Preserves vector content & images

    new_doc.save(output_pdf)
    # print(f"High-quality merged PDF saved as: {output_pdf}")

# Define crop areas (page index, cropping rectangle)
selections = [
    {"page": 0, "rect": fitz.Rect(0, 0, 1056, 75)},    # Selection 1 (Placed correctly)
    {"page": 0, "rect": fitz.Rect(0, 76, 102, 814)},   # Selection 2 (Placed correctly)
    {"page": 0, "rect": fitz.Rect(720, 76, 1056, 814)},  # Selection 3 (Will be repositioned)
    {"page": 1, "rect": fitz.Rect(790, 76, 1056, 814)},  # Selection 4 (Will be repositioned)
]

crop_and_merge_pdf("STP.pdf", "STO-ONCOLOGY.pdf", selections)