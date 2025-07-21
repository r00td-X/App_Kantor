"""
imp_R7_to_mysql.py – GUI untuk menampilkan data dari tabel pdf
----------------------------------------------------------------------------

"""

import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from ttkbootstrap import Style, ScrolledText
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

# --- GUI Setup ---
root = tk.Tk()
root.title("Import R7 ke Database")
root.geometry("850x700")
style = Style("cosmo")

main_frame = ttk.Frame(root, padding="10")
main_frame.pack(fill=tk.BOTH, expand=True)

# Header Frame
header_frame = ttk.Frame(main_frame)
header_frame.pack(fill=tk.X, pady=(0, 10))
header_frame.columnconfigure(0, weight=1)
header_frame.columnconfigure(1, weight=1)

info_frame = ttk.Labelframe(header_frame, text="Informasi PDF", padding="10")
info_frame.grid(row=0, column=0, sticky="ew", padx=(0, 5))

label1_var = tk.StringVar(value="Manifest Kantong: -")
label2_var = tk.StringVar(value="Kode: -")
ttk.Label(info_frame, textvariable=label1_var, font=["-size", "10"]).pack(anchor="w")
ttk.Label(info_frame, textvariable=label2_var, font=["-size", "10", "-weight", "bold"]).pack(anchor="w")

status_frame = ttk.Labelframe(header_frame, text="Status Sistem", padding="10")
status_frame.grid(row=0, column=1, sticky="ew", padx=(5, 0))

label3_var = tk.StringVar(value=f"Kode Kantor (.env): {KODE_KANTOR}")
label_koneksi_var = tk.StringVar(value="Status Koneksi: ❓")
koneksi_label = ttk.Label(status_frame, textvariable=label_koneksi_var, font=["-size", "10"])
koneksi_label.pack(anchor="w")
ttk.Label(status_frame, textvariable=label3_var, font=["-size", "10"]).pack(anchor="w")


# Table Frame
table_frame = ttk.Frame(main_frame)
table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

cols = ("No", "No Kantong", "Produk", "Berat (Kg)", "Asal Bag")
tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=15)
for col in cols:
    tree.heading(col, text=col)
    tree.column(col, anchor='center', width=100)
tree.column("No Kantong", width=150)

vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=vsb.set)

tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
vsb.pack(side=tk.RIGHT, fill=tk.Y)

# Progressbar
progress = ttk.Progressbar(main_frame, mode="determinate")
progress.pack(fill=tk.X, pady=(0, 10))

# Button Frame
btn_frame = ttk.Frame(main_frame)
btn_frame.pack(fill=tk.X, pady=(0, 10))
btn_frame.columnconfigure((0, 1, 2), weight=1)

Button(btn_frame, text="📂 Buka PDF", command=lambda: browse_pdf(), bootstyle="primary").grid(row=0, column=0, sticky="ew", padx=(0, 5))
Button(btn_frame, text="⬇️ Insert ke Database", command=lambda: insert_ke_db(), bootstyle="success").grid(row=0, column=1, sticky="ew", padx=5)
Button(btn_frame, text="🔄 Cek Koneksi", command=lambda: cek_koneksi(), bootstyle="info").grid(row=0, column=2, sticky="ew", padx=(5, 0))

# Log Frame
log_frame = ttk.Labelframe(main_frame, text="Log Aktivitas", padding="10")
log_frame.pack(fill=tk.BOTH, expand=True)

log_text = ScrolledText(log_frame, height=6, font=("Consolas", 9), wrap=tk.WORD)
log_text.pack(fill=tk.BOTH, expand=True)


def log(msg):
    log_text.insert('end', msg + "\n")
    log_text.see('end')
    logging.info(msg)

def cek_koneksi():
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        if conn.is_connected():
            label_koneksi_var.set("🟢 Koneksi DB: Berhasil")
            koneksi_label.config(bootstyle="success")
            conn.close()
        else:
            label_koneksi_var.set("🔴 Gagal koneksi DB")
            koneksi_label.config(bootstyle="danger")
    except Exception as e:
        label_koneksi_var.set(f"🔴 Koneksi Error")
        koneksi_label.config(bootstyle="danger")
        log(f"[ERROR] {e}")
        logging.exception("Terjadi error saat koneksi DB:")

def browse_pdf():
    filepath = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not filepath:
        return

    tree.delete(*tree.get_children())
    label1_var.set("Manifest Kantong: -")
    label2_var.set("Kode: -")
    log(f"📄 Membuka file: {filepath}")

    try:
        doc = fitz.open(filepath)
        full_text = "\n".join(page.get_text() for page in doc)
        
        manifest_match = re.search(r"Manifest Kantong(?:.|\n)*?:\s*([^\n]*)", full_text, re.IGNORECASE)
        if manifest_match:
            full_manifest_text = manifest_match.group(1).strip()
            if full_manifest_text and not full_manifest_text[0].isalnum():
                full_manifest_text = full_manifest_text[1:].strip()
            label1_var.set(f"Manifest Kantong: {full_manifest_text}")
            parts = full_manifest_text.split()
            if parts:
                label2_var.set(parts[-1])
            else:
                label2_var.set("Kode: -")
                log("❗ Kode tidak ditemukan di dalam manifest.")
        else:
            log("❗ Gagal deteksi baris 'Manifest Kantong'")

        table_lines = re.findall(r"^\d+\s+P\d+\s+\w+\s+[0-9.]+\s+-", full_text, re.MULTILINE)
        for i, line in enumerate(table_lines, 1):
            parts = line.split()
            tree.insert('', 'end', values=(str(i), parts[1], parts[2], parts[3], "-"))
        log(f"✅ Berhasil parsing {len(table_lines)} baris dari PDF")

    except Exception as e:
        log(f"[ERROR] Gagal parsing PDF: {e}")
        logging.exception("Gagal parsing PDF")

def insert_ke_db():
    if not tree.get_children():
        messagebox.showwarning("Data Kosong", "Tidak ada data untuk di-insert. Silakan buka file PDF terlebih dahulu.")
        return

    kode_label2 = label2_var.get()
    if not KODE_KANTOR or kode_label2 != KODE_KANTOR:
        msg = f"Kode manifest ({kode_label2}) tidak sama dengan KODE_KANTOR di .env ({KODE_KANTOR})."
        messagebox.showwarning("Kode Tidak Cocok", msg)
        return

    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()
        
        total = 0
        pid_dilewati = 0
        duplikat = 0
        
        all_items = tree.get_children()
        progress["maximum"] = len(all_items)
        
        for i, item in enumerate(all_items):
            progress["value"] = i + 1
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
        progress["value"] = 0
        
    except Exception as e:
        log(f"[ERROR] {e}")
        traceback.print_exc()
        logging.exception("Error insert ke DB")
        messagebox.showerror("Error Database", f"Terjadi kesalahan saat insert ke DB:\n{e}")

# Initial connection check
cek_koneksi()
root.mainloop()
