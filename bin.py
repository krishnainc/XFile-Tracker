import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import tempfile
from docx import Document
from docx2pdf import convert
import pdfplumber
import difflib

original_file = None
modified_file = None
temp_dir = tempfile.mkdtemp()

def upload_original():
    global original_file
    original_file = filedialog.askopenfilename(
        title="Select Original Word Document",
        filetypes=[("Word Files", "*.docx")]
    )
    if original_file:
        messagebox.showinfo("File Selected", f"Original document:\n{original_file}")

def upload_modified():
    global modified_file
    modified_file = filedialog.askopenfilename(
        title="Select Modified Word Document",
        filetypes=[("Word Files", "*.docx")]
    )
    if modified_file:
        messagebox.showinfo("File Selected", f"Modified document:\n{modified_file}")

def convert_to_pdf(docx_path):
    pdf_path = os.path.join(temp_dir, os.path.basename(docx_path).replace(".docx", ".pdf"))
    convert(docx_path, pdf_path)
    return pdf_path

def extract_docx_paragraphs(docx_path):
    """Extract all non-empty paragraphs with their nearest heading."""
    doc = Document(docx_path)
    paragraphs = []
    current_heading = "No Heading"
    
    for para in doc.paragraphs:
        if para.style.name.startswith('Heading'):
            current_heading = para.text.strip() or current_heading
        text = para.text.strip()
        if text:
            paragraphs.append((text, current_heading))
    return paragraphs

def extract_pdf_text_per_page(pdf_path):
    """Extract plain text per page for page number mapping."""
    page_texts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            page_texts.append(text.strip())
    return page_texts

def find_page_number(paragraph_text, pdf_pages):
    """Try to locate which page a paragraph appears on."""
    for idx, page in enumerate(pdf_pages, start=1):
        if paragraph_text in page:
            return idx
    return "Unknown"

def compare_documents():
    if not original_file or not modified_file:
        messagebox.showwarning("Missing Files", "Please upload both documents first.")
        return

    # Step 1: Convert to PDF to get page mapping
    original_pdf = convert_to_pdf(original_file)
    modified_pdf = convert_to_pdf(modified_file)
    modified_pdf_pages = extract_pdf_text_per_page(modified_pdf)

    # Step 2: Extract paragraphs from Word docs
    original_paras = [p[0] for p in extract_docx_paragraphs(original_file)]
    modified_paras = extract_docx_paragraphs(modified_file)

    # Step 3: Detect added/removed paragraphs using set difference
    original_texts = set(original_paras)
    modified_texts = set([p[0] for p in modified_paras])

    added = [p for p in modified_paras if p[0] not in original_texts]
    removed = [p for p in extract_docx_paragraphs(original_file) if p[0] not in modified_texts]

    # Step 4: Build difference report
    differences = []

    for text, heading in added:
        page_num = find_page_number(text, modified_pdf_pages)
        differences.append(f"Page {page_num} | Heading: {heading}\n  + Added: {text}\n")

    for text, heading in removed:
        # Map removed paragraph to original page if needed
        page_num = find_page_number(text, modified_pdf_pages)
        differences.append(f"Page {page_num} | Heading: {heading}\n  - Removed: {text}\n")

    # Step 5: Display results
    if not differences:
        messagebox.showinfo("Comparison Result", "No changes detected!")
        return

    result_window = tk.Toplevel(root)
    result_window.title("Comparison Result")
    txt = scrolledtext.ScrolledText(result_window, wrap="word", width=120, height=40)
    txt.pack(padx=10, pady=10, fill="both", expand=True)
    txt.insert("1.0", "\n".join(differences))
    txt.config(state="disabled")

# ------------------ Tkinter GUI ------------------
root = tk.Tk()
root.title("Word Document Comparison Tool")

btn1 = tk.Button(root, text="Upload Original Document", command=upload_original, width=30)
btn1.pack(pady=10)

btn2 = tk.Button(root, text="Upload Modified Document", command=upload_modified, width=30)
btn2.pack(pady=10)

btn3 = tk.Button(root, text="Compare Documents", command=compare_documents, width=30)
btn3.pack(pady=20)

root.mainloop()
