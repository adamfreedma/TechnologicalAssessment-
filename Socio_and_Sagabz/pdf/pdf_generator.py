import shutil
import fitz  # PyMuPDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject
from reportlab.lib import colors
import sys
import os
import comtypes.client
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# Step 1: Find the position of the text
def find_text_coordinates(pdf_path, search_text):
    doc = fitz.open(pdf_path)
    page_numbers = {}
    for page_num in range(len(doc)):
        page = doc[page_num]
        text_instances = page.search_for(search_text)
        if not page_num in page_numbers:
            page_numbers[page_num] = []

        if text_instances:
            for text_instance in text_instances:
                page_numbers[page_num].append(text_instance)

    return page_numbers


# Step 2: Create overlay with white rectangle + textbox
def create_overlay_with_box(overlay_path, rect, page_size, index):
    c = canvas.Canvas(overlay_path, pagesize=page_size)

    # Convert PyMuPDF coords to ReportLab coords

    rect_base_width = 270
    total_width = rect.width + rect_base_width

    width = rect.width
    height = rect.height

    x = rect.x0 - total_width + 30
    y = page_size[1] - rect.y1 + 50 # Flip y-axis

    # c.setFillColorRGB(1, 0, 0)
    # c.circle(x, y, 3, fill=1)

    # Cover original text
    c.setFillColorRGB(1, 1, 1)
    c.setLineWidth(0)
    c.rect(x, y, width + rect_base_width, height + 2, fill=0, stroke=0)

    # Add fillable text box with underline
    a = c.acroForm.textfield(
        name=f'dynamic_box{index}',
        x=x, y=y, width=width + rect_base_width, height=height,
        forceBorder=False, fillColor=colors.white, borderWidth=0,
        fontSize=12,
    )
    # Add another text box a row below
    b = c.acroForm.textfield(
        name=f'dynamic_box_below{index}',
        x=x, y=y - height - 2, width=width + rect_base_width, height=height,
        forceBorder=False, fillColor=colors.white, borderWidth=0,
        fontSize=12,
    )

    c.save()


# Step 3: Merge overlay onto the correct page
def merge_overlay(base_pdf_path, overlay_path, output_pdf_path, target_page):
    base_pdf = PdfReader(base_pdf_path)
    overlay_pdf = PdfReader(overlay_path)
    writer = PdfWriter()

    for i, page in enumerate(base_pdf.pages):
        if i == target_page:
            page.merge_page(overlay_pdf.pages[0])
        writer.add_page(page)

    if "/AcroForm" in overlay_pdf.trailer["/Root"]:
        writer._root_object.update({
            NameObject("/AcroForm"): overlay_pdf.trailer["/Root"]["/AcroForm"]
        })

    with open(output_pdf_path, "wb") as f_out:
        writer.write(f_out)


def replace_blanks(input_file, output_file):
    search_text = "{box}"  # Example placeholder text
    base_pdf_file = input_file
    pdf_file = base_pdf_file

    page_numbers = find_text_coordinates(pdf_file, search_text)
    counter = 0
    
    for page_num in page_numbers:
        for rect in page_numbers[page_num]:
            overlay_name = f"overlay.pdf"
            if rect:
                create_overlay_with_box(overlay_name, rect, letter, index=counter)
                merge_overlay(pdf_file, overlay_name, output_file, page_num)
            else:
                print("Text not found.")

            counter += 1
            if pdf_file == base_pdf_file:
                pdf_file = output_file

import win32com.client
from docx2pdf import convert


def docx_to_pdf(docx_path, pdf_path):
    # Start Microsoft Word application
    word = win32com.client.Dispatch("Word.Application")
    
    # Make Word invisible (optional)
    word.Visible = False
    
    # Open the .docx file
    doc = word.Documents.Open(docx_path)
    
    # Save the document as a PDF
    doc.SaveAs(pdf_path, FileFormat=17)  # 17 corresponds to PDF format in Word
    
    # Close the document and quit Word
    doc.Close()
    word.Quit()


    

from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject, NumberObject

def fix_textfield_alignment(input_pdf_path, output_pdf_path):
    reader = PdfReader(input_pdf_path)
    writer = PdfWriter()

    # Copy AcroForm from source to writer
    if "/AcroForm" in reader.trailer["/Root"]:
        writer._root_object.update({
            NameObject("/AcroForm"): reader.trailer["/Root"]["/AcroForm"]
        })

    # Iterate pages and annotations
    for page in reader.pages:
        annots = page.get("/Annots", [])
        for annot_ref in annots:
            annot = annot_ref.get_object()
            # Identify text fields
            if annot.get("/Subtype") == "/Widget" and annot.get("/FT") == "/Tx":
                name = annot.get("/T")
                if name and name.startswith("dynamic_box"):
                    # Set Q = 2 for right alignment
                    annot.update({ NameObject("/Q"): NumberObject(2) })
        writer.add_page(page)

    # Write out the fixed PDF
    with open(output_pdf_path, "wb") as out_f:
        writer.write(out_f)


if __name__ == "__main__":
    for file in os.listdir("input/"):
        if file.startswith("~$"):
            os.remove(os.path.join("input/", file))
    convert("input/")
    
    for file in os.listdir("input/"):
        if file.endswith(".pdf"):
            input_docx = os.path.join("input/", file)
            shutil.move(input_docx, os.path.join("output/", file))
    
    for file in os.listdir("output/"):
        if file.endswith(".pdf"):
            input_pdf = os.path.join("output/", file)
            output_pdf = "temp.pdf"
            print(file)
            
            replace_blanks(input_pdf, output_pdf)
            fix_textfield_alignment(output_pdf, output_pdf)
            
            os.remove(input_pdf)
            os.remove("overlay.pdf")
            os.rename(output_pdf, input_pdf)
            print(f"Processed {input_pdf} and saved as {input_pdf}.")