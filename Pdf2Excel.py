import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import PyPDF2
import csv
from openpyxl import Workbook

def extract_pdf(pdf_path, save_path, output_type, progress_bar, status_label):
    try:
        with open(pdf_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            total_pages = len(reader.pages)

            progress_bar["value"] = 0
            progress_bar["maximum"] = total_pages

            if output_type == "csv":
                with open(save_path, "w", newline="", encoding="utf-8") as out_file:
                    writer = csv.writer(out_file)
                    for i, page in enumerate(reader.pages):
                        text = page.extract_text()
                        if text:
                            for line in text.split("\n"):
                                writer.writerow(line.split())
                        progress_bar["value"] = i + 1
                        status_label.config(text=f"Processing page {i+1}/{total_pages}")
                        root.update_idletasks()
                messagebox.showinfo("Success", f"Saved as {save_path}")

            elif output_type == "excel":
                wb = Workbook()
                ws = wb.active
                for i, page in enumerate(reader.pages):
                    text = page.extract_text()
                    if text:
                        for line in text.split("\n"):
                            ws.append(line.split())
                    progress_bar["value"] = i + 1
                    status_label.config(text=f"Processing page {i+1}/{total_pages}")
                    root.update_idletasks()
                wb.save(save_path)
                messagebox.showinfo("Success", f"Saved as {save_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Failed: {e}")

def select_file(output_type):
    pdf_path = filedialog.askopenfilename(
        title="Select PDF",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not pdf_path:
        return

    if output_type == "csv":
        save_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
    else:
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

    if save_path:
        extract_pdf(pdf_path, save_path, output_type, progress_bar, status_label)

# ---------------- GUI ----------------
root = tk.Tk()
root.title("PDF to CSV/Excel")
root.geometry("400x350")
root.configure(bg="#1e1e1e")  # dark background

label = tk.Label(root, text="Choose a PDF and export:", font=("Arial", 14),
                 bg="#1e1e1e", fg="white")
label.pack(pady=20)

btn_csv = tk.Button(root, text="Export to CSV",
                    command=lambda: select_file("csv"),
                    bg="#3498DB", fg="white", font=("Arial", 12), relief="flat")
btn_csv.pack(pady=10)

btn_excel = tk.Button(root, text="Export to Excel",
                      command=lambda: select_file("excel"),
                      bg="#2ECC71", fg="white", font=("Arial", 12), relief="flat")
btn_excel.pack(pady=10)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=20)

status_label = tk.Label(root, text="", font=("Arial", 10),
                        bg="#1e1e1e", fg="#aaaaaa")
status_label.pack()

# Credit footer
credit_label = tk.Label(root, text="Created by MasterK", font=("Arial", 10, "italic"),
                        bg="#1e1e1e", fg="#888888")
credit_label.pack(side="bottom", pady=10)

root.mainloop()