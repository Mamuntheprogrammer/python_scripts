import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def crop_and_merge_pdf(input_pdf, output_pdf, selections):
    doc = fitz.open(input_pdf)
    new_doc = fitz.open()

    # Create a new page with the same size as the original
    ref_page = doc[0]  # Reference page
    page_width, page_height = ref_page.rect.width, ref_page.rect.height
    new_page = new_doc.new_page(width=page_width, height=page_height)
    # print("Page Size:", page_width, "x", page_height)

    for idx, sel in enumerate(selections):
        page_num, rect = sel['page'], sel['rect']
        if idx == 1:
            rect = fitz.Rect(0+20, 76+5, 102+20, 814+5)
        
        page = doc[page_num]

        if idx == 2:  # Selection 3 should be next to Selection 2
            rect = fitz.Rect(102+20, 76+5, 438+20, 814+5)
        elif idx == 3:  # Selection 4 should be next to Selection 3
            rect = fitz.Rect(15+20, 68, 1056+20, 809.5)
        
        new_page.show_pdf_page(rect, doc, page_num, clip=sel['rect'])  # Preserves vector content & images
    
    new_doc.save(output_pdf)
    # print(f"High-quality merged PDF saved as: {output_pdf}")
    messagebox.showinfo("Success", "PDF processing completed successfully!")

def select_input_pdf():
    input_pdf = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if input_pdf:
        file_label.config(text=os.path.basename(input_pdf))
        output_pdf = os.path.join(os.path.dirname(input_pdf), "STP-Oncology.pdf")
        selections = [
            {"page": 0, "rect": fitz.Rect(0, 0, 1056, 75)},
            {"page": 0, "rect": fitz.Rect(0, 76, 102, 814)},
            {"page": 0, "rect": fitz.Rect(720, 76, 1056, 814)},
            {"page": 1, "rect": fitz.Rect(790, 76, 1056, 814)},
        ]
        crop_and_merge_pdf(input_pdf, output_pdf, selections)

# Create Tkinter UI
root = tk.Tk()
root.title("PDF Crop and Merge")
root.geometry("300x200")

tk.Label(root, text="Select Input PDF").pack(pady=10)
tk.Button(root, text="Browse", command=select_input_pdf).pack()

file_label = tk.Label(root, text="", fg="blue")
file_label.pack(pady=5)

root.mainloop()