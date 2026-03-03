import os
import shutil
import pdfplumber
from docx import Document
from pypdf import PdfWriter, PdfReader
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from io import BytesIO
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox

# ===== FOLDER STRUCTURE =====
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

# ===== HELPER FUNCTIONS =====

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
    if file_name.lower().endswith(".pdf"):
        content_text = get_pdf_text(full_path)
    elif file_name.lower().endswith(".docx"):
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

# ===== CONVERSION FUNCTIONS =====

def convert_txt_to_pdf(txt_path, output_path):
    doc = SimpleDocTemplate(output_path, pagesize=letter)
    styles = getSampleStyleSheet()
    elements = []

    with open(txt_path, "r", encoding="utf-8") as f:
        for line in f:
            elements.append(Paragraph(line.strip(), styles["Normal"]))

    doc.build(elements)

def convert_image_to_pdf(image_path, output_path):
    image = Image.open(image_path)
    if image.mode == "RGBA":
        image = image.convert("RGB")
    image.save(output_path, "PDF")

def convert_docx_to_pdf(docx_path, output_path):
    # Requires Word installed on Windows
    subprocess.run([
        "powershell",
        "-Command",
        f"""$word = New-Object -ComObject Word.Application;
        $doc = $word.Documents.Open('{docx_path}');
        $doc.SaveAs('{output_path}', 17);
        $doc.Close();
        $word.Quit();"""
    ])

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

# ===== COMBINE FUNCTION WITH DIVIDERS & BOOKMARKS =====

def combine_pdfs(folder):
    writer = PdfWriter()
    deal_name = os.path.basename(folder)
    current_page = 0

    for section in FOLDER_STRUCTURE:
        section_path = os.path.join(folder, section)

        if not os.path.exists(section_path):
            continue

        files = sorted(os.listdir(section_path))
        section_has_files = any(
            f.lower().endswith((".pdf", ".txt", ".docx", ".jpg", ".jpeg", ".png"))
            for f in files
        )

        if not section_has_files:
            continue

        # Add Divider Page
        divider_pdf = create_divider_page(section.replace("_", " "), deal_name)
        writer.add_page(divider_pdf.pages[0])
        writer.add_outline_item(section.replace("_", " "), current_page)
        current_page += 1

        for file in files:
            file_path = os.path.join(section_path, file)
            temp_pdf = None

            if file.lower().endswith(".pdf"):
                temp_pdf = file_path

            elif file.lower().endswith(".txt"):
                temp_pdf = file_path.replace(".txt", "_converted.pdf")
                convert_txt_to_pdf(file_path, temp_pdf)

            elif file.lower().endswith(".docx"):
                temp_pdf = file_path.replace(".docx", "_converted.pdf")
                convert_docx_to_pdf(file_path, temp_pdf)

            elif file.lower().endswith((".jpg", ".jpeg", ".png")):
                temp_pdf = file_path + "_converted.pdf"
                convert_image_to_pdf(file_path, temp_pdf)

            if temp_pdf and os.path.exists(temp_pdf):
                reader = PdfReader(temp_pdf)
                for page in reader.pages:
                    writer.add_page(page)
                    current_page += 1

    output_file = os.path.join(folder, deal_name + "_PRINT_PACKAGE.pdf")

    with open(output_file, "wb") as f:
        writer.write(f)

# ===== GUI FUNCTIONS =====

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

# ===== GUI LAYOUT =====

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