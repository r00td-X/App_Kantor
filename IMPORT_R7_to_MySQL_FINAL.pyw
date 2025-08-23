

#
# imp_R7_to_mysql.py - GUI untuk menampilkan data dari tabel pdf
#----------------------------------------------------------------------------
#
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from ttkbootstrap import Style, ScrolledText
from ttkbootstrap.widgets import Button, DateEntry
from dotenv import load_dotenv
import mysql.connector
import pdfplumber
import traceback
from datetime import datetime
import requests
from bs4 import BeautifulSoup
import threading

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

user_name_map = {}

# --- GUI Setup ---
root = tk.Tk()
root.title("Import R7 ke Database")
root.geometry("850x650") # Adjusted size
style = Style("cosmo")

main_frame = ttk.Frame(root, padding="10")
main_frame.pack(fill=tk.BOTH, expand=True)

# Header Frame
header_frame = ttk.Frame(main_frame)
header_frame.pack(fill=tk.X, pady=(0, 10))
header_frame.columnconfigure(0, weight=1)
header_frame.columnconfigure(1, weight=1)

info_frame = ttk.Labelframe(header_frame, text="Informasi PDF", padding="10")
info_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

label1_var = tk.StringVar(value="Manifest Kantong: -")
label2_var = tk.StringVar(value="Kode: -")
ttk.Label(info_frame, textvariable=label1_var, font=["-size", "10"]).pack(anchor="w")
ttk.Label(info_frame, textvariable=label2_var, font=["-size", "10", "-weight", "bold"]).pack(anchor="w")

status_frame = ttk.Labelframe(header_frame, text="Status Sistem", padding="10")
status_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

label3_var = tk.StringVar(value=f"Kode Kantor (.env): {KODE_KANTOR}")
label_koneksi_var = tk.StringVar(value="Status Koneksi: ❓")

koneksi_label = ttk.Label(status_frame, textvariable=label_koneksi_var, font=["-size", "10"])
koneksi_label.pack(anchor="w")
ttk.Label(status_frame, textvariable=label3_var, font=["-size", "10"]).pack(anchor="w")

# Date Picker
date_picker_frame = ttk.Frame(status_frame)
date_picker_frame.pack(anchor="w", pady=(5,0))
ttk.Label(date_picker_frame, text="Tgl NRC:", font=["-size", "10"]).pack(side="left")
date_picker = DateEntry(date_picker_frame, bootstyle="primary", dateformat="%Y-%m-%d")
date_picker.pack(side="left", padx=(5,0))

# User Selection
user_selection_frame = ttk.Frame(status_frame)
user_selection_frame.pack(anchor="w", pady=(5,0))
ttk.Label(user_selection_frame, text="Pilih User:", font=["-size", "10"]).pack(side="left")
user_var = tk.StringVar()
user_combobox = ttk.Combobox(user_selection_frame, textvariable=user_var, state="readonly", bootstyle="primary")
user_combobox.pack(side="left", padx=(5,0))
user_display_var = tk.StringVar(value="")
ttk.Label(user_selection_frame, textvariable=user_display_var, font=["-size", "10", "-weight", "bold"]).pack(side="left", padx=(10,0))

# Table Frame
table_frame = ttk.Frame(main_frame)
table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

cols = ("No", "No Kantong", "Produk", "Berat (Kg)", "Asal Bag", "Tujuan Bag")
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
btn_frame.columnconfigure((0, 1, 2, 3), weight=1)

# Log Frame
log_frame = ttk.Labelframe(main_frame, text="Log Aktivitas", padding="10")
log_frame.pack(fill=tk.BOTH, expand=True)

log_text = ScrolledText(log_frame, height=3, font=("Consolas", 9), wrap=tk.WORD)
log_text.pack(fill=tk.BOTH, expand=True)

def log(msg):
    root.after(0, lambda: _log_to_widget(msg))

def _log_to_widget(msg):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_text.insert('end', f"[{now}] {msg}\n")
    log_text.see('end')

def parse_and_format_date(date_str):
    if not date_str:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M"):
        try:
            return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            pass
    return None

def _perform_scraping_and_update(connotes_to_scrap):
    updated_count = 0
    failed_count = 0
    
    if not connotes_to_scrap:
        log("✅ Tidak ada data untuk di-scrap pada mode ini.")
        return 0, 0

    log(f"🔍 Ditemukan {len(connotes_to_scrap)} connote untuk diproses.")
    progress["maximum"] = len(connotes_to_scrap)

    conn = mysql.connector.connect(**DB_CONFIG)

    for i, row in enumerate(connotes_to_scrap):
        progress["value"] = i + 1
        root.update_idletasks()
            
        connote = row['connote']
        log(f"🔄 Memproses connote: {connote}")
        
        url = f"https://kibana.posindonesia.co.id:4433/x123449/3.php?id={connote}&6f017f90-f299-11ec-988f-6f1763dc6f47xdsdkjshhsahsaksasjsaasldsllsdjldsjsbdaksdslssjasjaa"
        
        try:
            response = requests.get(url, timeout=15, verify=False)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")
            
            data_to_update = {}

            tgl_kirim_th = soup.find("th", string=re.compile(r"^\s*Tanggal Kirim\s*$"))
            if tgl_kirim_th and tgl_kirim_th.find_next_sibling("td"):
                tgl_kirim_raw = tgl_kirim_th.find_next_sibling("td").text.strip()
                data_to_update['tgl_kirim'] = parse_and_format_date(tgl_kirim_raw)

            pengirim_th = soup.find("th", string=re.compile(r"^\s*Pengirim\s*$"))
            if pengirim_th and pengirim_th.find_next_sibling("td"):
                pengirim_full = pengirim_th.find_next_sibling("td").text.strip()
                data_to_update['pgrm'] = pengirim_full.split(',')[0].strip()

            penerima_th = soup.find("th", string=re.compile(r"^\s*Penerima\s*$"))
            if penerima_th and penerima_th.find_next_sibling("td"):
                penerima_full = penerima_th.find_next_sibling("td").text.strip()
                data_to_update['pnrm'] = penerima_full.split(',')[0].strip()
                if 'Alamat :' in penerima_full:
                    data_to_update['al_pnrm'] = penerima_full.split('Alamat :')[-1].strip()

            status_akhir_th = soup.find("th", string=re.compile(r"^\s*STATUS AKHIR\s*$"))
            if status_akhir_th and status_akhir_th.find_next_sibling("td"):
                status_full = status_akhir_th.find_next_sibling("td").text.strip()
                status_text = status_full.split(' Di ')[0].strip()
                data_to_update['status'] = status_text
                
                if "DELIVERED" in status_text.upper():
                    data_to_update['st'] = '99'

                lokasi_part = status_full.split(',')[0]
                kodepos_match = re.search(r"\b(\d{5})\b", lokasi_part)
                if kodepos_match:
                    data_to_update['lok_akhir'] = kodepos_match.group(1)

                date_match = re.search(r"tgl\s*:\s*(\d{4}-\d{2}-\d{2}\s*\d{2}:\d{2}:\d{2})", status_full)
                if date_match:
                    data_to_update['tgl_proses'] = parse_and_format_date(date_match.group(1))
            
            cod_th = soup.find("th", string=re.compile(r"^\s*COD/NON COD\s*$"))
            if cod_th and cod_th.find_next_sibling("td"):
                cod_full = cod_th.find_next_sibling("td").text.strip()
                data_to_update['cod'] = cod_full.split('Nilai Cod :')[0].strip()
                match = re.search(r"Nilai Cod\s*:\s*([0-9,.]*)", cod_full)
                if match:
                    bsu_cod_raw = match.group(1).strip().replace(',', '').replace('.', '')
                    data_to_update['bsu_cod'] = int(bsu_cod_raw) if bsu_cod_raw.isdigit() else 0
                else:
                    data_to_update['bsu_cod'] = 0

            if data_to_update:
                if 'st' not in data_to_update:
                    data_to_update['st'] = '33'
                
                set_clauses = ", ".join([f"{key}=%s" for key in data_to_update.keys()])
                sql_update = f"UPDATE tbl_antrn SET {set_clauses} WHERE connote=%s"
                update_values = list(data_to_update.values()) + [connote]
                
                update_cursor = conn.cursor()
                update_cursor.execute(sql_update, tuple(update_values))
                conn.commit()
                update_cursor.close()
                updated_count += 1
                log(f"  ✅ Data untuk {connote} berhasil diupdate.")
            else:
                log(f"  ⚠️ Tidak ada data valid yang diekstrak untuk {connote}.")
                failed_count += 1

        except requests.exceptions.RequestException as e:
            log(f"  ❌ Gagal mengambil data untuk {connote}. Error: {e}")
            failed_count += 1
        except Exception as e:
            log(f"  ❌ Terjadi error saat memproses {connote}. Error: {e}")
            traceback.print_exc()
            failed_count += 1
    
    conn.close()
    progress["value"] = 0
    return updated_count, failed_count

def jalankan_scrap_awal():
    if messagebox.askyesno("Konfirmasi Scrap Awal", "Proses ini akan mengambil data baru (st=0) dari internet. Lanjutkan?"):
        def run():
            try:
                log("🚀 Memulai proses scraping awal...")
                conn = mysql.connector.connect(**DB_CONFIG)
                cursor = conn.cursor(dictionary=True)
                cursor.execute("SELECT connote FROM tbl_antrn WHERE st = '0'")
                connotes_to_scrap = cursor.fetchall()
                cursor.close()
                conn.close()
                
                if not connotes_to_scrap:
                    log("✅ Tidak ada data baru (st=0) untuk di-scrap.")
                    messagebox.showinfo("Selesai", "Tidak ada data baru untuk di-scrap.")
                    return

                updated, failed = _perform_scraping_and_update(connotes_to_scrap)
                
                log(f"🎉 Scraping Awal Selesai. Berhasil: {updated}, Gagal: {failed}")
                messagebox.showinfo("Scraping Selesai", f"Proses scraping awal selesai.\nBerhasil update: {updated}\nGagal: {failed}")

            except Exception as e:
                log(f"[ERROR] Terjadi kesalahan fatal saat scrap awal: {e}")
                traceback.print_exc()
                messagebox.showerror("Error Scraping", f"Terjadi kesalahan fatal:\n{e}")
            finally:
                progress["value"] = 0
        
        threading.Thread(target=run, daemon=True).start()

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

def populate_user_combobox(kode_kantor):
    global user_name_map
    user_name_map = {}
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()
        cursor.execute("SELECT nama, username FROM user WHERE status2 = 'Pengantar' AND ktr = %s", (kode_kantor,))
        users_data = cursor.fetchall()
        cursor.close()
        conn.close()
        
        display_names = []
        for nama, username in users_data:
            display_names.append(nama)
            user_name_map[nama] = username
            
        user_combobox['values'] = display_names
        if display_names:
            user_combobox.set(display_names[0])
            on_user_select(None)
        else:
            user_combobox.set("")
            user_display_var.set("")
        log(f"✅ Combobox user diisi dengan {len(display_names)} nama.")
    except Exception as e:
        log(f"[ERROR] Gagal mengisi combobox user: {e}")
        user_combobox['values'] = []
        user_combobox.set("")
        user_display_var.set("")

def on_user_select(event):
    selected_name = user_var.get()
    username = user_name_map.get(selected_name, "")
    user_display_var.set(username)
    log(f"User terpilih: {selected_name} (Username: {username})")

def browse_pdf():
    filepath = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not filepath:
        return

    tree.delete(*tree.get_children())
    label1_var.set("Manifest Kantong: -")
    label2_var.set("Kode: -")
    user_combobox.set("")
    log(f"📄 Membuka file: {filepath}")

    try:
        with pdfplumber.open(filepath) as pdf:
            first_page_text = pdf.pages[0].extract_text()
            for line in first_page_text.splitlines():
                if "Manifest Kantong" in line:
                    full_text = line.split(":")[-1].strip()
                    label1_var.set(f"Manifest Kantong: {full_text}")
                    kode = full_text.split()[-1]
                    label2_var.set(kode)
                    populate_user_combobox(kode)

            all_data = []
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if row and row[0] and isinstance(row[0], str) and row[0].strip().isdigit():
                            all_data.append(row)
            
            if not all_data:
                log("❗ Tidak ditemukan data tabel yang valid.")
                return

            for row_data in all_data:
                display_row = row_data
                while len(display_row) < len(cols):
                    display_row.append("-")
                tree.insert("", "end", values=display_row[:len(cols)])

            log(f"✅ Berhasil parsing {len(all_data)} baris dari PDF.")

    except Exception as e:
        log(f"[ERROR] Gagal parsing PDF: {e}")
        messagebox.showerror("Error Parsing", f"Gagal memproses file PDF.\nError: {e}")

def insert_ke_db():
    if not tree.get_children():
        messagebox.showwarning("Data Kosong", "Tidak ada data untuk di-insert.")
        return

    kode_label2 = label2_var.get()
    if not KODE_KANTOR or kode_label2 != KODE_KANTOR:
        msg = f"Kode manifest ({kode_label2}) tidak sama dengan KODE_KANTOR di .env ({KODE_KANTOR})."
        messagebox.showwarning("Kode Tidak Cocok", msg)
        return

    try:
        tgl_nrc_str = date_picker.entry.get()
        datetime.strptime(tgl_nrc_str, "%Y-%m-%d")
        tgl_nrc = tgl_nrc_str
    except (ValueError, Exception) as e:
        messagebox.showerror("Tanggal Tidak Valid", f"Format tanggal tidak valid (YYYY-MM-DD).\nError: {e}")
        return

    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor(dictionary=True) # Gunakan dictionary=True untuk akses kolom via nama
        
        total_insert, total_update, pid_dilewati, duplikat_tidak_diubah = 0, 0, 0, 0
        all_items = tree.get_children()
        progress["maximum"] = len(all_items)
        
        for i, item in enumerate(all_items):
            progress["value"] = i + 1
            root.update_idletasks()
            
            values = tree.item(item)['values']
            no_kantong, produk = values[1], values[2]

            if str(no_kantong).startswith("PID"):
                pid_dilewati += 1
                continue

            # Cek apakah connote sudah ada dan ambil status ktr_antrn
            cursor.execute("SELECT ktr_antrn FROM tbl_antrn WHERE connote=%s", (no_kantong,))
            result = cursor.fetchone()

            if result: # Jika connote ditemukan
                if str(result['ktr_antrn']) == '0':
                    # Update ktr_antrn menjadi KODE_KANTOR
                    update_sql = "UPDATE tbl_antrn SET ktr_antrn = %s WHERE connote = %s"
                    update_val = (KODE_KANTOR, no_kantong)
                    cursor.execute(update_sql, update_val)
                    total_update += 1
                    log(f"🔄 Connote {no_kantong} ditemukan, ktr_antrn diupdate menjadi {KODE_KANTOR}.")
                else:
                    # ktr_antrn bukan 0, lewati
                    duplikat_tidak_diubah += 1
                    log(f"⏭️ Connote {no_kantong} sudah ada dengan ktr_antrn != 0, dilewati.")
                continue
            
            # Jika connote tidak ditemukan, lakukan insert baru
            sql = "INSERT INTO tbl_antrn (connote, produk, ktr_antrn, tgl_nrc, pic) VALUES (%s, %s, %s, %s, %s)"
            val = (no_kantong, produk, kode_label2, tgl_nrc, user_display_var.get())
            cursor.execute(sql, val)
            total_insert += 1
        
        conn.commit()
        cursor.close()
        conn.close()
        
        msg = (f"✅ Proses Selesai.\n" 
               f"Data Baru: {total_insert}\n" 
               f"Data Diupdate: {total_update}\n" 
               f"Duplikat Dilewati: {duplikat_tidak_diubah}\n" 
               f"PID Dilewati: {pid_dilewati}")

        log(msg.replace('\n', ', '))
        messagebox.showinfo("Sukses", msg)
        progress["value"] = 0
        
    except Exception as e:
        log(f"[ERROR] {e}")
        traceback.print_exc()
        messagebox.showerror("Error Database", f"Terjadi kesalahan saat proses ke DB:\n{e}")

if __name__ == "__main__":
    Button(btn_frame, text="📂 Buka PDF", command=browse_pdf, bootstyle="primary").grid(row=0, column=0, sticky="ew", padx=(0, 5))
    Button(btn_frame, text="⬇️ Insert ke Database", command=insert_ke_db, bootstyle="success").grid(row=0, column=1, sticky="ew", padx=5)
    Button(btn_frame, text="▶️ Jalankan Scrap Awal", command=jalankan_scrap_awal, bootstyle="warning").grid(row=0, column=2, sticky="ew", padx=5)
    Button(btn_frame, text="🔄 Cek Koneksi", command=cek_koneksi, bootstyle="info").grid(row=0, column=3, sticky="ew", padx=(5, 0))
    
    user_combobox.bind("<<ComboboxSelected>>", on_user_select)

    cek_koneksi()
    root.mainloop()
