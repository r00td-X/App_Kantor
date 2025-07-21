import os
import warnings
warnings.filterwarnings("ignore")

import tkinter as tk
import mysql.connector
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from ttkbootstrap.widgets import Treeview
from dotenv import load_dotenv
import pdfplumber

# Load variabel dari .env
load_dotenv()
KODE_KANTOR = os.getenv("KODE_KANTOR", "")

DB_CONFIG = {
    "host": os.getenv("DB_HOST"),
    "user": os.getenv("DB_USER"),
    "password": os.getenv("DB_PASS"),
    "database": os.getenv("DB_NAME"),
    "use_pure": True
}

# Fungsi untuk ekstrak tabel dari PDF
def extract_table_from_pdf(pdf_path):
    data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    # Filter baris kosong dan baris ringkasan
                    if (
                        row
                        and any(cell.strip() for cell in row)
                        and not row[0].strip().lower().startswith(("produk", "ec3", "pkh", "pe", "total", "kantor", "agus", "nippos"))
                    ):
                        data.append(row)
    return data

# Fungsi untuk handle browse PDF
def browse_pdf():
    file_path = filedialog.askopenfilename(
        title="Pilih file PDF",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not file_path:
        return
    
    try:
        with pdfplumber.open(file_path) as pdf:
            first_page_text = pdf.pages[0].extract_text()
            # Cari baris "Manifest Kantong : ..."
            for line in first_page_text.splitlines():
                if "Manifest Kantong" in line:
                    # Contoh baris: "Manifest Kantong : KCP DRINGU 67271"
                    full_text = line.split(":")[-1].strip()
                    label1_var.set(full_text)
                    # Ambil digit belakang kode (dipisah dari nama kantor)
                    kode = full_text.split()[-1]
                    label2_var.set(kode)
                    break  # Stop setelah ditemukan

        # Proses data tabel seperti sebelumnya
        table_data = extract_table_from_pdf(file_path)
        if not table_data:
            messagebox.showwarning("Peringatan", "Tidak ditemukan tabel pada PDF.")
            return

        # Bersihkan Treeview
        for col in tree["columns"]:
            tree.heading(col, text="")
        tree.delete(*tree.get_children())

        headers = table_data[0]
        tree["columns"] = headers
        tree["show"] = "headings"

        for col in headers:
            tree.heading(col, text=col)
            tree.column(col, anchor="center")

        for row in table_data[1:]:
            tree.insert("", "end", values=row)

    except Exception as e:
        messagebox.showerror("Error", f"Gagal membaca PDF:\n{e}")

# Fungsi untuk insert No Kantong ke tbl_antrn field connote
def insert_to_db():
    if label2_var.get() != label3_var.get():
        messagebox.showwarning("Peringatan", "⚠️ Kode Manifest (Label2) dan Kode Kantor (.env) tidak sama!")
        return
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()

        inserted = 0
        pid_skipped = 0
        duplicate_skipped = 0
        ktr_antrn_value = label2_var.get()

        for row_id in tree.get_children():
            row = tree.item(row_id)["values"]
            no_kantong = str(row[1])
            produk = str(row[2])

            if no_kantong.startswith("PID"):
                pid_skipped += 1
                continue

            # Cek duplikat di database
            cursor.execute("SELECT COUNT(*) FROM tbl_antrn WHERE connote = %s", (no_kantong,))
            if cursor.fetchone()[0] > 0:
                duplicate_skipped += 1
                continue

            # Insert data valid
            cursor.execute(
                "INSERT INTO tbl_antrn (connote, produk, ktr_antrn) VALUES (%s, %s, %s)",
                (no_kantong, produk, ktr_antrn_value)
            )
            inserted += 1

        conn.commit()
        cursor.close()
        conn.close()

        messagebox.showinfo("Hasil Insert",
            f"✅ Data berhasil diinsert: {inserted}\n"
            f"⏭️ Diabaikan karena PID: {pid_skipped}\n"
            f"📦 Sudah ada di database: {duplicate_skipped}"
        )

    except Exception as e:
        messagebox.showerror("Error DB", str(e))

# === GUI SETUP ===
root = tk.Tk()
root.title("Baca Tabel PDF ke Treeview")
root.geometry("1000x500")

style = Style("flatly")

frame = tk.Frame(root)
frame.pack(fill="both", expand=True, padx=10, pady=10)

label1_var = tk.StringVar()
label2_var = tk.StringVar()
label3_var = tk.StringVar(value=KODE_KANTOR)

info_frame = tk.Frame(root)
info_frame.pack(pady=5)

tk.Label(info_frame, text="Manifest Kantong:").grid(row=0, column=0, sticky="w")
tk.Label(info_frame, textvariable=label1_var, font=("Segoe UI", 10, "bold")).grid(row=0, column=1, sticky="w", padx=10)

tk.Label(info_frame, text="Kode:").grid(row=1, column=0, sticky="w")
tk.Label(info_frame, textvariable=label2_var, font=("Segoe UI", 10, "bold")).grid(row=1, column=1, sticky="w", padx=10)

tk.Label(info_frame, text="Kode Kantor (.env):").grid(row=2, column=0, sticky="w")
tk.Label(info_frame, textvariable=label3_var, font=("Segoe UI", 10, "bold")).grid(row=2, column=1, sticky="w", padx=10)

btn_browse = tk.Button(root, text="📄 Browse PDF", command=browse_pdf)
btn_browse.pack(pady=10)

btn_insert = tk.Button(root, text="⬆️ Insert ke Database", command=insert_to_db)
btn_insert.pack(pady=5)

tree = Treeview(frame, show="headings")
tree.pack(side="left", fill="both", expand=True)

scroll = tk.Scrollbar(frame, orient="vertical", command=tree.yview)
scroll.pack(side="right", fill="y")
tree.configure(yscrollcommand=scroll.set)

root.mainloop()