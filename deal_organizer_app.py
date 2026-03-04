import os
import shutil
import pdfplumber
from pypdf import PdfWriter, PdfReader
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from io import BytesIO
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox

# ==============================
# CONFIG
# ==============================

FOLDER_STRUCTURE = {
    "01_Credit_Writeup": ["credit submission for customer"],
    "02_Application": ["application"],
    "03_Invoice": ["invoice"],
    "04_Spec_Sheet": ["specs", "spec sheet", "specifications"],
    "05_PayNet": ["paynet", "Paynet"],
    "06_Personal_Credit": ["experian", "transunion", "equifax", "fico"],
    "07_Financials": ["balance sheet", "income statement", "Profit and Loss", "cash flow", "financial statement", "financials"],
    "08_Tax_Returns": ["1120", "1065", "schedule c", "tax return"],
    "09_Personal_Financial_Statement": ["personal financial statement", "pfs"],
    "10_Bank_Statements": ["bank statement"],
    "11_Insurance": ["insurance"],
    "12_Misc": []
}

selected_folder = ""

# ==============================
# CONVERSION FUNCTIONS
# ==============================

def convert_docx_to_pdf(input_path, output_path):
    subprocess.run([
        "powershell",
        "-Command",
        f"""$word = New-Object -ComObject Word.Application;
        $word.Visible = $false;
        $doc = $word.Documents.Open('{input_path}');
        $doc.SaveAs('{output_path}', 17);
        $doc.Close();
        $word.Quit();"""
    ])

def convert_xlsx_to_pdf(input_path, output_path):
    subprocess.run([
        "powershell",
        "-Command",
        f"""$excel = New-Object -ComObject Excel.Application;
        $excel.Visible = $false;
        $workbook = $excel.Workbooks.Open('{input_path}');
        $workbook.ExportAsFixedFormat(0, '{output_path}');
        $workbook.Close($false);
        $excel.Quit();"""
    ])

def convert_txt_to_pdf(input_path, output_path):
    from reportlab.platypus import SimpleDocTemplate, Paragraph
    from reportlab.lib.styles import getSampleStyleSheet

    doc = SimpleDocTemplate(output_path, pagesize=letter)
    styles = getSampleStyleSheet()
    elements = []

    with open(input_path, "r", encoding="utf-8") as f:
        for line in f:
            elements.append(Paragraph(line.strip(), styles["Normal"]))

    doc.build(elements)

def convert_image_to_pdf(input_path, output_path):
    image = Image.open(input_path)
    if image.mode == "RGBA":
        image = image.convert("RGB")
    image.save(output_path, "PDF")

# ==============================
# PRE-CONVERT ALL FILES
# ==============================

def preconvert_all_files(folder):
    for file in os.listdir(folder):
        full_path = os.path.join(folder, file)

        if not os.path.isfile(full_path):
            continue

        name, ext = os.path.splitext(file)
        ext = ext.lower()

        output_pdf = os.path.join(folder, name + ".pdf")

        if ext == ".docx":
            convert_docx_to_pdf(full_path, output_pdf)

        elif ext == ".txt":
            convert_txt_to_pdf(full_path, output_pdf)

        elif ext in [".jpg", ".jpeg", ".png"]:
            convert_image_to_pdf(full_path, output_pdf)

# ==============================
# CLASSIFICATION
# ==============================

def get_pdf_text(path):
    try:
        with pdfplumber.open(path) as pdf:
            text = ""
            for page in pdf.pages[:2]:
                text += page.extract_text() or ""
        return text.lower()
    except:
        return ""

def classify_file(file_name, full_path):
    combined = file_name.lower() + " " + get_pdf_text(full_path)

    for folder, keywords in FOLDER_STRUCTURE.items():
        for keyword in keywords:
            if keyword in combined:
                return folder

    return "12_Misc"

def organize_pdfs(folder):
    for subfolder in FOLDER_STRUCTURE:
        os.makedirs(os.path.join(folder, subfolder), exist_ok=True)

    for file in os.listdir(folder):
        full_path = os.path.join(folder, file)

        if os.path.isfile(full_path) and file.lower().endswith(".pdf"):
            category = classify_file(file, full_path)
            destination = os.path.join(folder, category, file)

            if not os.path.exists(destination):
                shutil.move(full_path, destination)

# ==============================
# DIVIDER + COMBINE
# ==============================

def create_divider_page(title, deal_name):
    packet = BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    can.setFont("Helvetica-Bold", 24)
    can.drawCentredString(300, 500, title)

    can.setFont("Helvetica", 14)
    can.drawCentredString(300, 470, f"Deal: {deal_name}")

    can.save()
    packet.seek(0)
    return PdfReader(packet)

def combine_pdfs(folder):
    writer = PdfWriter()
    deal_name = os.path.basename(folder)
    current_page = 0

    for section in FOLDER_STRUCTURE:
        section_path = os.path.join(folder, section)
        if not os.path.exists(section_path):
            continue

        files = sorted([f for f in os.listdir(section_path) if f.lower().endswith(".pdf")])
        if not files:
            continue

        divider_pdf = create_divider_page(section.replace("_", " "), deal_name)
        writer.add_page(divider_pdf.pages[0])
        writer.add_outline_item(section.replace("_", " "), current_page)
        current_page += 1

        for file in files:
            pdf_path = os.path.join(section_path, file)
            reader = PdfReader(pdf_path)

            for page in reader.pages:
                writer.add_page(page)
                current_page += 1

    output_file = os.path.join(folder, deal_name + "_PRINT_PACKAGE.pdf")
    with open(output_file, "wb") as f:
        writer.write(f)

# ==============================
# GUI
# ==============================

def select_folder():
    global selected_folder
    selected_folder = filedialog.askdirectory()
    folder_label.config(text=selected_folder)

def process_deal():
    if not selected_folder:
        messagebox.showerror("Error", "Select a folder first.")
        return

    try:
        preconvert_all_files(selected_folder)
        organize_pdfs(selected_folder)
        combine_pdfs(selected_folder)
        messagebox.showinfo("Success", "Deal organized and print package created.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("Credit Deal Organizer")
root.geometry("500x250")

tk.Button(root, text="Select Deal Folder", command=select_folder).pack(pady=10)
folder_label = tk.Label(root, text="No folder selected", wraplength=450)
folder_label.pack()
tk.Button(root, text="Organize + Create Print Package", command=process_deal).pack(pady=20)

root.mainloop()