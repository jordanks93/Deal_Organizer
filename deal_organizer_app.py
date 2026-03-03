import os
import shutil
import pdfplumber
from docx import Document
from pypdf import PdfWriter, PdfReader
import tkinter as tk
from tkinter import filedialog, messagebox

# ===== Folder Classification Map =====
FOLDER_STRUCTURE = {
    "01_Credit_Writeup": ["writeup", "credit memo"],
    "02_Application": ["application", "app"],
    "03_Invoice": ["invoice"],
    "04_Sales_Order": ["sales order"],
    "05_PayNet": ["paynet", "master score"],
    "06_Personal_Credit": ["experian", "transunion", "equifax", "fico"],
    "07_Financials": ["balance sheet", "income statement", "financial statement"],
    "08_Tax_Returns": ["1120", "1065", "schedule c", "tax return"],
    "09_Personal_Financial_Statement": ["personal financial statement", "pfs"],
    "10_Bank_Statements": ["bank statement"],
    "11_Insurance": ["insurance"],
    "12_Misc": []
}

selected_folder = ""

# ===== Helper Functions =====

def get_pdf_text(path):
    try:
        with pdfplumber.open(path) as pdf:
            text = ""
            for page in pdf.pages[:2]:
                text += page.extract_text() or ""
        return text.lower()
    except:
        return ""

def get_docx_text(path):
    try:
        doc = Document(path)
        return "\n".join([p.text for p in doc.paragraphs]).lower()
    except:
        return ""

def classify_file(file_name, full_path):
    lower_name = file_name.lower()

    content_text = ""
    if file_name.endswith(".pdf"):
        content_text = get_pdf_text(full_path)
    elif file_name.endswith(".docx"):
        content_text = get_docx_text(full_path)

    combined = lower_name + " " + content_text

    for folder, keywords in FOLDER_STRUCTURE.items():
        for keyword in keywords:
            if keyword in combined:
                return folder

    return "12_Misc"

def organize_files(folder):
    for subfolder in FOLDER_STRUCTURE:
        os.makedirs(os.path.join(folder, subfolder), exist_ok=True)

    for file in os.listdir(folder):
        full_path = os.path.join(folder, file)

        if os.path.isfile(full_path):
            category = classify_file(file, full_path)
            destination = os.path.join(folder, category, file)

            if not os.path.exists(destination):
                shutil.move(full_path, destination)

def combine_pdfs(folder):
    writer = PdfWriter()

    for subfolder in FOLDER_STRUCTURE:
        subfolder_path = os.path.join(folder, subfolder)

        if os.path.exists(subfolder_path):
            files = sorted(os.listdir(subfolder_path))

            for file in files:
                if file.lower().endswith(".pdf"):
                    pdf_path = os.path.join(subfolder_path, file)
                    reader = PdfReader(pdf_path)
                    for page in reader.pages:
                        writer.add_page(page)

    output_file = os.path.join(folder, os.path.basename(folder) + "_PRINT_PACKAGE.pdf")

    with open(output_file, "wb") as f:
        writer.write(f)

# ===== GUI Functions =====

def select_folder():
    global selected_folder
    selected_folder = filedialog.askdirectory()
    folder_label.config(text=selected_folder)

def process_deal():
    if not selected_folder:
        messagebox.showerror("Error", "Please select a deal folder first.")
        return

    try:
        organize_files(selected_folder)
        combine_pdfs(selected_folder)
        messagebox.showinfo("Success", "Deal organized and print package created!")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# ===== GUI Layout =====

root = tk.Tk()
root.title("Credit Deal Organizer")

root.geometry("500x250")

select_btn = tk.Button(root, text="Select Deal Folder", command=select_folder)
select_btn.pack(pady=10)

folder_label = tk.Label(root, text="No folder selected", wraplength=450)
folder_label.pack()

process_btn = tk.Button(root, text="Organize + Create Print Package", command=process_deal)
process_btn.pack(pady=20)

root.mainloop()