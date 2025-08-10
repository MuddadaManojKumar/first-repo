import os
import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from rapidfuzz import fuzz

PDF_FOLDER = "certificates_folder"
EXCEL_PATH = "dataset.xlsx"
OUTPUT_FOLDER = "renamed_certificates"
FUZZY_THRESHOLD = 90

# Normalize for matching
def normalize(name):
    return " ".join(str(name).lower().replace("mr.", "").replace("ms.", "").replace("mrs.", "").split()).strip()

# Robust Excel header/column detection
def load_excel_with_flexible_headers(excel_path):
    header_row = None
    for i in range(10):
        df_try = pd.read_excel(excel_path, header=i)
        cols = [str(c).lower() for c in df_try.columns]
        if any('name' in c for c in cols) and any('roll' in c for c in cols):
            header_row = i
            df = df_try
            break
    if header_row is None:
        raise ValueError("❌ Could not find columns with 'name' and 'roll' in the first 10 rows. Please check your Excel file.")
    name_col = next(c for c in df.columns if 'name' in str(c).lower())
    roll_col = next(c for c in df.columns if 'roll' in str(c).lower())
    df['Name'] = df[name_col].apply(normalize)
    df['Roll Number'] = df[roll_col].astype(str)
    return df

def get_best_match(name_from_pdf, excel_names, threshold=FUZZY_THRESHOLD):
    best_score = 0
    best_match = None
    for name in excel_names:
        score = fuzz.ratio(name_from_pdf, name)
        if score > best_score:
            best_score = score
            best_match = name
    if best_score >= threshold:
        return best_match
    return None

df = load_excel_with_flexible_headers(EXCEL_PATH)

os.makedirs(OUTPUT_FOLDER, exist_ok=True)
pdf_files = [f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf")]
current_index = 0

# UI setup
root = tk.Tk()
root.title("Manual Certificate Renamer (with Auto-Suggest)")

pdf_label = tk.Label(root, text="", font=("Arial", 12))
pdf_label.pack(pady=10)

name_var = tk.StringVar()
name_combo = ttk.Combobox(root, textvariable=name_var, values=list(df['Name']), width=60)
name_combo.pack(pady=5)

suggest_label = tk.Label(root, text="", font=("Arial", 10), fg="blue")
suggest_label.pack(pady=2)

def open_pdf(path):
    try:
        os.startfile(path)  # For Windows
    except Exception as e:
        messagebox.showerror("Error", f"Could not open PDF: {e}")

def extract_name_from_pdf(pdf_path):
    try:
        pdf = fitz.open(pdf_path)
        text = pdf[0].get_text("text").lower()
        pdf.close()
    except Exception as e:
        print(f"❌ Failed to read {os.path.basename(pdf_path)}: {e}")
        return None
    import re
    name_line = next((line for line in text.split('\n') if 'we are glad to inform' in line), None)
    if name_line:
        match = re.search(r'inform\s+(mr\.|ms\.|mrs\.)?\s*([a-z\s]+)\(', name_line)
        if match:
            return normalize(match.group(2))
    return None

def rename_and_next():
    global current_index
    selected_name = normalize(name_var.get())
    match = df[df['Name'] == selected_name]
    if not match.empty:
        roll = match.iloc[0]['Roll Number']
        old_path = os.path.join(PDF_FOLDER, pdf_files[current_index])
        new_path = os.path.join(OUTPUT_FOLDER, f"{roll}.pdf")
        os.rename(old_path, new_path)
        print(f"✅ Renamed to {roll}.pdf")
    else:
        messagebox.showerror("Match Error", "No match found for selected name!")
        return
    current_index += 1
    load_next_pdf()

def load_next_pdf():
    if current_index >= len(pdf_files):
        messagebox.showinfo("Done", "All PDFs processed!")
        root.quit()
        return
    current_file = pdf_files[current_index]
    pdf_label.config(text=f"[{current_index+1}/{len(pdf_files)}] {current_file}")
    pdf_path = os.path.join(PDF_FOLDER, current_file)
    # Extract name and auto-suggest
    extracted_name = extract_name_from_pdf(pdf_path)
    if extracted_name:
        best_match = get_best_match(extracted_name, df['Name'])
        if best_match:
            name_var.set(best_match)
            suggest_label.config(text=f"Auto-suggested: {best_match}")
        else:
            name_var.set("")
            suggest_label.config(text="No confident match found. Please select manually.")
    else:
        name_var.set("")
        suggest_label.config(text="Could not extract name from PDF. Please select manually.")
    open_pdf(pdf_path)

next_button = tk.Button(root, text="Rename & Next", command=rename_and_next, font=("Arial", 12), bg="#4CAF50", fg="white")
next_button.pack(pady=10)

# Start
load_next_pdf()
root.mainloop() 