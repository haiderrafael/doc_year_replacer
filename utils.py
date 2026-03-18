import os
from docx import Document
import fitz  # PyMuPDF

def replace_year_in_docx(input_filepath, output_filepath, old_year, new_year):
    """
    Replaces all occurrences of old_year with new_year in a docx document.
    """
    doc = Document(input_filepath)

    # Replace in paragraphs
    for paragraph in doc.paragraphs:
        if old_year in paragraph.text:
            for run in paragraph.runs:
                if old_year in run.text:
                    run.text = run.text.replace(old_year, new_year)
            
            # Fallback if text spans multiple runs
            if old_year in paragraph.text:
                paragraph.text = paragraph.text.replace(old_year, new_year)

    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if old_year in cell.text:
                    for paragraph in cell.paragraphs:
                        if old_year in paragraph.text:
                            for run in paragraph.runs:
                                if old_year in run.text:
                                    run.text = run.text.replace(old_year, new_year)
                            if old_year in paragraph.text:
                                paragraph.text = paragraph.text.replace(old_year, new_year)

    doc.save(output_filepath)

def replace_year_in_pdf(input_filepath, output_filepath, old_year, new_year):
    """
    Replaces occurrences of old_year with new_year in a PDF file using PyMuPDF.
    This works by redacting the text and adding the new text, but visual fidelity
    might be slightly altered depending on the font.
    """
    doc = fitz.open(input_filepath)

    for page in doc:
        # Search for the old_year text
        text_instances = page.search_for(old_year)
        
        for inst in text_instances:
            # We must clear the existing text using redaction.
            # Using add_redact_annot carefully on the rectangle.
            page.add_redact_annot(inst, fill=(1, 1, 1)) # Try filling with white to obscure. Or use standard redaction.
            
        page.apply_redactions()

        # Insert new text where the old instances were
        for inst in text_instances:
            rect = fitz.Rect(inst)
            
            # Estimate font size based on original text height
            fsize = (rect.y1 - rect.y0) * 0.85
            if fsize < 8:
                fsize = 11
                
            # Use insert_text which ignores bounding box clipping rules, 
            # placing the text exactly at the bottom-left coordinate of the original text
            p = fitz.Point(rect.x0, rect.y1 - (rect.height * 0.2))
            page.insert_text(p, new_year, fontsize=fsize, fontname="helv", color=(0, 0, 0))
            
    doc.save(output_filepath)
