import os
import re
import tkinter as tk
from tkinter import messagebox, ttk
from ttkbootstrap import Style, ScrolledText
from ttkbootstrap.widgets import Button
from dotenv import load_dotenv
import mysql.connector
import traceback
from datetime import datetime, timedelta
import requests
from bs4 import BeautifulSoup
import threading
import time
from PIL import Image
import pystray

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

# --- Global variables for tray and background thread ---
icon = None
background_thread_stop = threading.Event()
auto_update_status_var = None

# --- GUI Setup ---
root = tk.Tk()
root.title("Antaran Updater Status v2") # Added v2 to title
root.geometry("700x550") # Adjusted height
style = Style("cosmo")

main_frame = ttk.Frame(root, padding="10")
main_frame.pack(fill=tk.BOTH, expand=True)

# Status Frame
status_frame = ttk.Labelframe(main_frame, text="Status Sistem", padding="10")
status_frame.pack(fill=tk.X, pady=(0, 10))

label_koneksi_var = tk.StringVar(value="Status Koneksi: ❓")
auto_update_status_var = tk.StringVar(value="Auto-Update: Menunggu...")
antrn_count_var = tk.StringVar(value="Data Siap Update (tbl_antrn): -")
retfs_count_var = tk.StringVar(value="Data Siap Update (antrn_tblretfs): -")

koneksi_label = ttk.Label(status_frame, textvariable=label_koneksi_var, font=["-size", "10"])
koneksi_label.pack(anchor="w")
ttk.Label(status_frame, textvariable=auto_update_status_var, font=["-size", "10"]).pack(anchor="w", pady=(5,0))
ttk.Label(status_frame, textvariable=antrn_count_var, font=["-size", "10", "-weight", "bold"]).pack(anchor="w", pady=(5,0))
ttk.Label(status_frame, textvariable=retfs_count_var, font=["-size", "10", "-weight", "bold"]).pack(anchor="w", pady=(5,0))


# Log Frame
log_frame = ttk.Labelframe(main_frame, text="Log Aktivitas", padding="10")
log_frame.pack(fill=tk.BOTH, expand=True)

log_text = ScrolledText(log_frame, height=10, font=("Consolas", 9), wrap=tk.WORD)
log_text.pack(fill=tk.BOTH, expand=True)

# Progress Bar Frame
progress_frame = ttk.Frame(main_frame)
progress_frame.pack(fill=tk.X, pady=(5, 5))

progress_label_var = tk.StringVar()
ttk.Label(progress_frame, textvariable=progress_label_var, font=["-size", "9"]).pack(side=tk.LEFT, anchor="w")

progress = ttk.Progressbar(progress_frame, mode="determinate")
progress.pack(fill=tk.X, expand=True, side=tk.LEFT, padx=5)


# Button Frame
btn_frame = ttk.Frame(main_frame)
btn_frame.pack(fill=tk.X)
btn_frame.columnconfigure((0, 1, 2), weight=1)

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

def _perform_scraping_and_update(connotes_to_scrap, is_manual_run=False):
    updated_count = 0
    failed_count = 0
    
    if not connotes_to_scrap:
        log("✅ Tidak ada data untuk di-scrap.")
        return 0, 0

    log(f"🔍 Ditemukan {len(connotes_to_scrap)} connote untuk diproses.")
    progress["maximum"] = len(connotes_to_scrap)

    conn = mysql.connector.connect(**DB_CONFIG)
    update_cursor = conn.cursor()

    for i, row in enumerate(connotes_to_scrap):
        progress["value"] = i + 1
        progress_label_var.set(f"Memproses {i + 1}/{len(connotes_to_scrap)}...")
        root.update_idletasks()

        if background_thread_stop.is_set():
            log("🛑 Proses update dihentikan.")
            break
            
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
                sql_update_antrn = (f"UPDATE tbl_antrn SET {set_clauses} "
                                    f"WHERE connote = %s AND ktr_antrn = %s AND "
                                    f"(lok_akhir = %s OR lok_akhir = '' OR lok_akhir IS NULL OR lok_akhir = '0') AND st = '33'")
                
                update_values = list(data_to_update.values())
                update_values.extend([connote, KODE_KANTOR, KODE_KANTOR])
                
                update_cursor.execute(sql_update_antrn, tuple(update_values))
                
                if update_cursor.rowcount > 0:
                    log(f"  ✅ Data untuk {connote} berhasil diupdate di tbl_antrn.")
                    updated_count += 1

                    if 'lok_akhir' in data_to_update and KODE_KANTOR:
                        select_sql_retfs = ( "SELECT 1 FROM antrn_tblretfs "
                                           "WHERE connote_rf = %s AND ktr_rf = %s AND "
                                           "(lok_akhir_rf = %s OR lok_akhir_rf = '' OR lok_akhir_rf IS NULL) LIMIT 1" )
                        select_values_retfs = (connote, KODE_KANTOR, KODE_KANTOR)
                        update_cursor.execute(select_sql_retfs, select_values_retfs)
                        row_to_update = update_cursor.fetchone()

                        if row_to_update:
                            log(f"  ℹ️ Connote {connote} ditemukan di antrn_tblretfs, mencoba update.")
                            sql_update_retfs = ( "UPDATE antrn_tblretfs SET lok_akhir_rf = %s "
                                               "WHERE connote_rf = %s AND ktr_rf = %s AND "
                                               "(lok_akhir_rf = %s OR lok_akhir_rf = '' OR lok_akhir_rf IS NULL)" )
                            
                            update_values_retfs = (data_to_update['lok_akhir'], connote, KODE_KANTOR, KODE_KANTOR)
                            update_cursor.execute(sql_update_retfs, update_values_retfs)

                            if update_cursor.rowcount > 0:
                                log(f"    ✅ lok_akhir_rf untuk {connote} berhasil diupdate menjadi {data_to_update['lok_akhir']}.")
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
    
    conn.commit()
    update_cursor.close()
    conn.close()
    progress["value"] = 0
    progress_label_var.set("")
    return updated_count, failed_count

def _background_update_task():
    log("⚙️ Auto-update thread dimulai.")
    while not background_thread_stop.is_set():
        try:
            root.after(0, lambda: auto_update_status_var.set("Auto-Update: 🏃‍♂️ Sedang berjalan..."))
            log(" ऑटो-अपडेट शुरू हो रहा है... (Memulai auto-update...)")
            
            conn = mysql.connector.connect(**DB_CONFIG)
            cursor = conn.cursor(dictionary=True)
            cursor.execute("SELECT connote FROM tbl_antrn WHERE st = '0' OR st = '33' OR status = 'FAILEDTODELIVERED'")
            connotes_to_update = cursor.fetchall()
            cursor.close()
            conn.close()

            if connotes_to_update:
                updated, failed = _perform_scraping_and_update(connotes_to_update)
                log(f"🤖 Auto-Update Selesai. Berhasil: {updated}, Gagal: {failed}")
            else:
                log("🤖 Auto-Update: Tidak ada data untuk diperbarui.")

            next_run_time = datetime.now() + timedelta(minutes=30)
            status_msg = f"Auto-Update: Idle. Cek berikutnya: {next_run_time.strftime('%H:%M:%S')}"
            root.after(0, lambda: auto_update_status_var.set(status_msg))
            log(f"😴 Menunggu 30 menit untuk siklus berikutnya.")
            
            background_thread_stop.wait(1800)

        except Exception as e:
            log(f"[ERROR] Kesalahan fatal di background thread: {e}")
            traceback.print_exc()
            root.after(0, lambda: auto_update_status_var.set("Auto-Update: ⚠️ Error!"))
            background_thread_stop.wait(60)

    log("🛑 Auto-update thread dihentikan.")

def run_manual_update():
    if messagebox.askyesno("Konfirmasi Update Manual", "Proses ini akan menjalankan update berdasarkan data yang ada di database. Lanjutkan?"):
        def run():
            try:
                log("🚀 Memulai proses update manual...")
                conn = mysql.connector.connect(**DB_CONFIG)
                cursor = conn.cursor(dictionary=True)
                cursor.execute("SELECT connote FROM tbl_antrn WHERE st = '0' OR st = '33' OR status = 'FAILEDTODELIVERED'")
                connotes_to_scrap = cursor.fetchall()
                cursor.close()
                conn.close()
                
                if not connotes_to_scrap:
                    log("✅ Tidak ada data baru untuk di-scrap.")
                    messagebox.showinfo("Selesai", "Tidak ada data baru untuk di-scrap.")
                    return

                updated, failed = _perform_scraping_and_update(connotes_to_scrap, is_manual_run=True)
                
                log(f"🎉 Update Manual Selesai. Berhasil: {updated}, Gagal: {failed}")
                messagebox.showinfo("Update Selesai", f"Proses update manual selesai.\nBerhasil update: {updated}\nGagal: {failed}")

            except Exception as e:
                log(f"[ERROR] Terjadi kesalahan fatal saat update manual: {e}")
                traceback.print_exc()
                messagebox.showerror("Error Update", f"Terjadi kesalahan fatal:\n{e}")
            finally:
                progress["value"] = 0
                progress_label_var.set("")
        
        threading.Thread(target=run, daemon=True).start()

def cek_koneksi():
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        if conn.is_connected():
            label_koneksi_var.set("🟢 Koneksi DB: Berhasil")
            koneksi_label.config(bootstyle="success")
            conn.close()
            log("✅ Koneksi ke database berhasil.")
        else:
            label_koneksi_var.set("🔴 Gagal koneksi DB")
            koneksi_label.config(bootstyle="danger")
            log(" Gagal terhubung ke database.")
    except Exception as e:
        label_koneksi_var.set(f"🔴 Koneksi Error")
        koneksi_label.config(bootstyle="danger")
        log(f"[ERROR] Gagal koneksi DB: {e}")

def check_data_to_update():
    log("🔍 Memeriksa data yang akan diupdate...")
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()

        # Hitung tbl_antrn
        antrn_query = "SELECT COUNT(*) FROM tbl_antrn WHERE st = '0' OR st = '33' OR status = 'FAILEDTODELIVERED'"
        cursor.execute(antrn_query)
        antrn_count = cursor.fetchone()[0]
        antrn_count_var.set(f"Data Siap Update (tbl_antrn): {antrn_count}")
        log(f"  📊 Ditemukan {antrn_count} data di tbl_antrn yang siap diupdate.")

        # Hitung antrn_tblretfs
        if not KODE_KANTOR:
            log("⚠️ KODE_KANTOR tidak diset, pengecekan antrn_tblretfs dilewati.")
            retfs_count_var.set("Data Siap Update (antrn_tblretfs): KODE_KANTOR kosong")
            retfs_count = "N/A"
        else:
            retfs_query = ( "SELECT COUNT(*) FROM antrn_tblretfs "
                          "WHERE ktr_rf = %s AND (lok_akhir_rf = %s OR lok_akhir_rf = '' OR lok_akhir_rf IS NULL)" )
            cursor.execute(retfs_query, (KODE_KANTOR, KODE_KANTOR))
            retfs_count = cursor.fetchone()[0]
            retfs_count_var.set(f"Data Siap Update (antrn_tblretfs): {retfs_count}")
            log(f"  📊 Ditemukan {retfs_count} data di antrn_tblretfs yang siap diupdate.")

        cursor.close()
        conn.close()
        messagebox.showinfo("Pengecekan Selesai", f"Pengecekan data selesai.\n- tbl_antrn: {antrn_count}\n- antrn_tblretfs: {retfs_count}")

    except Exception as e:
        log(f"[ERROR] Gagal memeriksa data: {e}")
        messagebox.showerror("Error Pengecekan", f"Gagal memeriksa data:\n{e}")

# --- System Tray Functions ---

def toggle_window(icon=None, item=None):
    if root.state() == "withdrawn":
        log("Membuka jendela utama.")
        root.deiconify()
    else:
        log("Menyembunyikan jendela ke tray.")
        root.withdraw()

def run_manual_update_thread():
    log("▶️ Update manual dari tray menu dijalankan.")
    threading.Thread(target=run_manual_update, daemon=True).start()

def check_data_to_update_thread():
    log("📊 Pengecekan data dari tray menu dijalankan.")
    threading.Thread(target=check_data_to_update, daemon=True).start()

def quit_app(icon, item):
    log("👋 Aplikasi akan ditutup.")
    background_thread_stop.set()
    if icon:
        icon.stop()
    root.quit()
    root.destroy()

def setup_tray():
    global icon
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(base_dir, "icon.png")
        image = Image.open(icon_path)
    except FileNotFoundError:
        log("⚠️ File 'icon.png' tidak ditemukan. Menggunakan ikon default.")
        image = Image.new('RGB', (64, 64), 'black')

    menu = (
        pystray.MenuItem('Tampilkan', lambda: toggle_window(), default=True),
        pystray.MenuItem('Jalankan Update Manual', run_manual_update_thread),
        pystray.MenuItem('Cek Data Update', check_data_to_update_thread),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('Keluar', quit_app)
    )
    
    icon = pystray.Icon("UpdateApp", image, "Aplikasi Updater Status", menu)
    icon.run()

def on_closing():
    if messagebox.askokcancel("Sembunyikan ke Tray", "Aplikasi akan tetap berjalan di system tray. Untuk keluar sepenuhnya, klik kanan ikon di tray dan pilih 'Keluar'. Lanjutkan?"):
        log("Aplikasi disembunyikan ke system tray.")
        root.withdraw()

if __name__ == "__main__":
    Button(btn_frame, text="🔄 Cek Koneksi", command=cek_koneksi, bootstyle="info").grid(row=0, column=0, sticky="ew", padx=(0, 5))
    Button(btn_frame, text="📊 Cek Data Update", command=check_data_to_update, bootstyle="secondary").grid(row=0, column=1, sticky="ew", padx=5)
    Button(btn_frame, text="▶️ Jalankan Update Manual", command=run_manual_update, bootstyle="success").grid(row=0, column=2, sticky="ew", padx=(5, 0))
    
    cek_koneksi()
    
    update_thread = threading.Thread(target=_background_update_task, daemon=True)
    update_thread.start()
    
    tray_thread = threading.Thread(target=setup_tray, daemon=True)
    tray_thread.start()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    root.mainloop()