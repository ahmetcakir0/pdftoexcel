import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os

def convert_pdf_to_excel(pdf_path, excel_path):
    try:
        all_data = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    all_data.extend(table)

        if all_data:
            df = pd.DataFrame(all_data[1:], columns=all_data[0])  # ilk satır başlık
            df.to_excel(excel_path, index=False)
            messagebox.showinfo("Başarılı", f"Excel dosyası oluşturuldu:\n{excel_path}")
        else:
            messagebox.showwarning("Uyarı", "PDF içinde tablo bulunamadı.")
    except Exception as e:
        messagebox.showerror("Hata", str(e))

def select_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF dosyaları", "*.pdf")])
    if pdf_path:
        default_name = os.path.splitext(os.path.basename(pdf_path))[0] + ".xlsx"
        excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                  initialfile=default_name,
                                                  filetypes=[("Excel dosyası", "*.xlsx")])
        if excel_path:
            thread = threading.Thread(target=convert_pdf_to_excel, args=(pdf_path, excel_path))
            thread.start()

# Ana arayüz
root = tk.Tk()
root.title("PDF ➜ Excel Dönüştürücü")
root.geometry("450x230")
root.resizable(False, False)
root.configure(bg="#f5f5f5")

# Başlık
title = tk.Label(root, text="📄 PDF ➜ Excel Dönüştürücü", font=("Segoe UI", 16, "bold"), bg="#f5f5f5")
title.pack(pady=20)

info = tk.Label(root, text="PDF dosyasındaki tabloyu Excel dosyasına çevirin", font=("Segoe UI", 11), bg="#f5f5f5")
info.pack()

# Buton
select_button = tk.Button(root, text="PDF Seç ve Dönüştür", command=select_pdf,
                          bg="#4CAF50", fg="white", font=("Segoe UI", 11), relief="flat", padx=10, pady=5)
select_button.pack(pady=20)

# Alt yazı
footer = tk.Label(root, text="Ahmet ÇAKIR", font=("Segoe UI", 9, "italic"), bg="#f5f5f5", fg="gray")
footer.pack(side="bottom", pady=10)

root.mainloop()
