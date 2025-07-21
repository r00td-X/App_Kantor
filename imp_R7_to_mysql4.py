import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from ttkbootstrap import Style
from ttkbootstrap.widgets import Button
from dotenv import load_dotenv
import mysql.connector
import fitz  # PyMuPDF
import logging
import traceback

# Setup logging
logging.basicConfig(filename="debug_log.txt", level=logging.DEBUG,
                    format="%(asctime)s - %(levelname)s - %(message)s")

# Load .env
load_dotenv()

DB_CONFIG = {
    "host": os.getenv("DB_HOST", ""),
    "user": os.getenv("DB_USER", ""),
    "password": os.getenv("DB_PASS", ""),
    "database": os.getenv("DB_NAME", ""),
    "use_pure": True
}
KODE_KANTOR = os.getenv("KODE_KANTOR", "")

# Init GUI
root = tk.Tk()
root.title("Import R7 ke Database")
root.geometry("800x600")
style = Style("flatly")

label_frame = ttk.Frame(root)
label_frame.pack(fill='x', padx=10, pady=5)

label1_var = tk.StringVar(value="Manifest Kantong: -")
label2_var = tk.StringVar(value="Kode: -")
label3_var = tk.StringVar(value=f"Kode Kantor (.env): {KODE_KANTOR}")
label_koneksi_var = tk.StringVar(value="Status Koneksi: ❓")

for lbl in (label1_var, label2_var, label3_var, label_koneksi_var):
    ttk.Label(label_frame, textvariable=lbl).pack(anchor='w')

cols = ("No", "No Kantong", "Produk", "Berat (Kg)", "Asal Bag")
tree = ttk.Treeview(root, columns=cols, show="headings", height=20)
for col in cols:
    tree.heading(col, text=col)
    tree.column(col, anchor='center')
tree.pack(fill='both', expand=True, padx=10, pady=5)

progress = ttk.Progressbar(root, mode="determinate")
progress.pack(fill='x', padx=10, pady=5)

btn_frame = ttk.Frame(root)
btn_frame.pack(pady=10)

log_text = tk.Text(root, height=5, bg="black", fg="lime", font=("Consolas", 9))
log_text.pack(fill='both', padx=10, pady=5)

def log(msg):
    log_text.insert('end', msg + "\n")
    log_text.see('end')
    logging.info(msg)

# Cek koneksi awal

def cek_koneksi():
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        if conn.is_connected():
            label_koneksi_var.set("🟢 Koneksi DB: Berhasil")
            conn.close()
        else:
            label_koneksi_var.set("🔴 Gagal koneksi DB")
    except Exception as e:
        label_koneksi_var.set(f"🔴 Koneksi Error: {e}")
        log(f"[ERROR] {e}")
        logging.exception("Terjadi error saat koneksi DB:")

def browse_pdf():
    filepath = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not filepath:
        return

    tree.delete(*tree.get_children())
    label1_var.set("Manifest Kantong: -")
    label2_var.set("Kode: -")
    label3_var.set(f"Kode Kantor (.env): {KODE_KANTOR}")
    log(f"📄 Membuka file: {filepath}")

    try:
        doc = fitz.open(filepath)
        full_text = "\n".join(page.get_text() for page in doc)

        manifest_match = re.search(r"Manifest Kantong\s*:\s*(KCP|KC)\s+.+?(\d{5}[A-Z0-9]*)", full_text)
        if manifest_match:
            label1_var.set(f"Manifest Kantong: {manifest_match.group(0).split(':')[-1].strip()}")
            label2_var.set(manifest_match.group(2))
        else:
            log("❗ Gagal deteksi kode manifest")

        table_lines = re.findall(r"^\d+\s+P\d+\s+\w+\s+[0-9.]+\s+-", full_text, re.MULTILINE)
        for i, line in enumerate(table_lines, 1):
            parts = line.split()
            no = str(i)
            no_kantong = parts[1]
            produk = parts[2]
            berat = parts[3]
            asal = parts[4] if len(parts) > 4 else "-"
            tree.insert('', 'end', values=(no, no_kantong, produk, berat, asal))
        log(f"✅ Berhasil parsing {len(table_lines)} baris dari PDF")
    except Exception as e:
        log(f"[ERROR] Gagal parsing PDF: {e}")
        logging.exception("Gagal parsing PDF")

def insert_ke_db():
    try:
        kode_label2 = label2_var.get()
        kode_label3 = KODE_KANTOR

        if kode_label2 != kode_label3:
            messagebox.showwarning("Kode Tidak Cocok", "Kode manifest tidak sama dengan KODE_KANTOR di .env")
            return

        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()

        total = 0
        pid_dilewati = 0
        duplikat = 0

        all_items = tree.get_children()
        progress["maximum"] = len(all_items)
        progress["value"] = 0

        for item in all_items:
            progress["value"] += 1
            root.update_idletasks()
            no_kantong, produk = tree.item(item)['values'][1:3]

            if no_kantong.startswith("PID"):
                pid_dilewati += 1
                continue

            cursor.execute("SELECT COUNT(*) FROM tbl_antrn WHERE connote=%s", (no_kantong,))
            if cursor.fetchone()[0] > 0:
                duplikat += 1
                continue

            cursor.execute("INSERT INTO tbl_antrn (connote, produk, ktr_antrn) VALUES (%s, %s, %s)",
                           (no_kantong, produk, kode_label2))
            total += 1

        conn.commit()
        conn.close()
        msg = f"✅ Insert Selesai. Baru: {total}, PID dilewati: {pid_dilewati}, Duplikat: {duplikat}"
        log(msg)
        messagebox.showinfo("Sukses", msg)
    except Exception as e:
        log(f"[ERROR] {e}")
        traceback.print_exc()
        logging.exception("Error insert ke DB")

# Tombol
Button(btn_frame, text="📂 Buka PDF", command=browse_pdf).pack(side='left', padx=10)
Button(btn_frame, text="⬇️ Insert ke Database", command=insert_ke_db).pack(side='left', padx=10)
Button(btn_frame, text="🔄 Cek Koneksi", command=cek_koneksi).pack(side='left', padx=10)

cek_koneksi()
root.mainloop()
