import os
import re
import tkinter as tk
from tkinter import messagebox, ttk
from ttkbootstrap import Style, ScrolledText
from dotenv import load_dotenv
import mysql.connector
import traceback
from datetime import datetime
import requests
from bs4 import BeautifulSoup
import threading
import time
import pystray
from PIL import Image

# Load environment variables from .env file
load_dotenv()

# --- Database Configuration ---
DB_CONFIG = {
    "host": os.getenv("DB_HOST", "localhost"),
    "user": os.getenv("DB_USER", "root"),
    "password": os.getenv("DB_PASS", ""),
    "database": os.getenv("DB_NAME", ""),
    "use_pure": True
}

# --- Status criteria for fetching connotes ---
STATUS_KRITERIA = ['INLOCATION', 'DELIVERYRUNSHEET', 'inBag', 'ON', 'unBag', 'INVEHICLE']

class MileUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Mile Updater App")
        self.root.geometry("650x550")
        self.style = Style("cosmo")
        self.auto_update_stop_event = threading.Event()
        self.auto_update_thread = None
        self.tray_icon = None

        self.load_icons()
        try:
            self.root.iconbitmap(r'C:\Users\POS\Music\PYTHON\App_Kantor\icon_mile.png')
        except tk.TclError:
            self.log("PERINGATAN: File ikon untuk jendela tidak ditemukan.")

        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.setup_gui()
        self.check_db_connection()
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)

    def load_icons(self):
        try:
            self.icon_idle = Image.open(r'C:\Users\POS\Music\PYTHON\App_Kantor\icon_mile.png')
        except FileNotFoundError:
            self.log("PERINGATAN: File ikon 'icon_mile.png' tidak ditemukan. Membuat ikon default.")
            self.icon_idle = Image.new('RGB', (64, 64), 'black')
        try:
            self.icon_busy = Image.open(r'C:\Users\POS\Music\PYTHON\App_Kantor\icon_mile_busy.png')
        except FileNotFoundError:
            self.log("PERINGATAN: File ikon 'icon_mile_busy.png' tidak ditemukan. Membuat ikon default.")
            self.icon_busy = Image.new('RGB', (64, 64), 'red')

    def setup_gui(self):
        # ... GUI setup is unchanged ...
        control_frame = ttk.Labelframe(self.main_frame, text="Kontrol Proses", padding="10")
        control_frame.pack(fill=tk.X, pady=5)
        control_frame.columnconfigure(1, weight=1)
        self.manual_button = ttk.Button(control_frame, text="Jalankan Manual Sekali", command=self.start_manual_update_thread)
        self.manual_button.grid(row=0, column=0, columnspan=3, padx=5, pady=(5,10), sticky="ew")
        self.interval_label = ttk.Label(control_frame, text="Interval Auto-Update (detik):")
        self.interval_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.interval_var = tk.IntVar(value=60)
        self.interval_slider = ttk.Scale(control_frame, from_=10, to_=3600, orient=tk.HORIZONTAL, variable=self.interval_var, command=self.update_interval_display)
        self.interval_slider.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.interval_display_label = ttk.Label(control_frame, text="60 detik")
        self.interval_display_label.grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.start_auto_button = ttk.Button(control_frame, text="Mulai Auto-Update", command=self.start_auto_update_thread, bootstyle="success")
        self.start_auto_button.grid(row=2, column=0, columnspan=2, padx=5, pady=10, sticky="ew")
        self.stop_auto_button = ttk.Button(control_frame, text="Stop Auto-Update", command=self.stop_auto_update, bootstyle="danger", state=tk.DISABLED)
        self.stop_auto_button.grid(row=2, column=2, padx=5, pady=10, sticky="ew")
        log_frame = ttk.Labelframe(self.main_frame, text="Log Aktivitas", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.log_text = ScrolledText(log_frame, height=10, font=("Consolas", 9), wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def update_interval_display(self, value):
        self.interval_display_label.config(text=f"{int(float(value))} detik")

    def log(self, msg):
        self.root.after(0, self._log_to_widget, msg)

    def _log_to_widget(self, msg):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log_text.insert('end', f"[{now}] {msg}\n")
        self.log_text.see('end')

    def check_db_connection(self):
        # ... unchanged ...
        try:
            conn = mysql.connector.connect(**DB_CONFIG)
            if conn.is_connected(): self.log("✅ Koneksi ke database berhasil.")
            else: self.log("🔴 Gagal terhubung ke database.")
            conn.close()
        except Exception as e: self.log(f"🔴 ERROR Koneksi DB: {e}")

    def set_tray_status(self, status, message="", title="Mile Updater"):
        if not self.tray_icon:
            return
        if status == 'busy':
            self.tray_icon.icon = self.icon_busy
            self.tray_icon.notify(message or "Proses update dimulai...", title=title)
        elif status == 'idle':
            self.tray_icon.icon = self.icon_idle
            if message: # Only notify if a message is provided
                self.tray_icon.notify(message, title=title)

    def parse_date_from_status(self, status_text):
        """Mencoba mengekstrak dan memformat tanggal dari teks status."""
        try:
            # Mencari pola tanggal seperti 'YYYY-MM-DD HH:MM:SS' atau 'YYYY-MM-DD'
            match = re.search(r'(\d{4}-\d{2}-\d{2}(?: \d{2}:\d{2}:\d{2})?)', status_text)
            if not match:
                self.log(f"  ⚠️ Peringatan: Tidak dapat menemukan pola tanggal yang valid di '{status_text}'")
                return None

            date_str = match.group(1)

            # Coba parsing dengan format tanggal-waktu
            if ' ' in date_str:
                try:
                    # Formatnya sudah YYYY-MM-DD HH:MM:SS, jadi bisa langsung dikembalikan
                    # setelah validasi.
                    datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
                    return date_str
                except ValueError:
                    self.log(f"  ⚠️ Peringatan: Format tanggal-waktu tidak valid '{date_str}'")
                    return None
            # Coba parsing dengan format tanggal saja
            else:
                try:
                    # Formatnya sudah YYYY-MM-DD, tambahkan waktu default.
                    datetime.strptime(date_str, '%Y-%m-%d')
                    return f"{date_str} 00:00:00"
                except ValueError:
                    self.log(f"  ⚠️ Peringatan: Format tanggal tidak valid '{date_str}'")
                    return None

        except Exception as e:
            self.log(f"  ❌ Error tak terduga saat parsing tanggal: {e}")
            return None

    def run_batch_update(self, is_manual_run=False):
        if is_manual_run:
            self.log("🚀 Memulai proses update MANUAL...")
            self.root.after(0, lambda: self.manual_button.config(state=tk.DISABLED))
        else:
            self.log("🤖 Memulai siklus auto-update...")
        
        self.set_tray_status('busy', "Sedang mengambil dan mengupdate data...")
        updated_count, failed_count = 0, 0
        failed_connotes = [] # List to hold failed connotes for logging

        try:
            conn = mysql.connector.connect(**DB_CONFIG)
            cursor = conn.cursor(dictionary=True)
            status_placeholders = ', '.join(['%s'] * len(STATUS_KRITERIA))
            sql_select = f"SELECT connote FROM tbl_db WHERE status IN ({status_placeholders})"
            cursor.execute(sql_select, STATUS_KRITERIA)
            connotes_to_process = cursor.fetchall()

            if not connotes_to_process:
                self.log("✅ Tidak ada connote baru.")
            else:
                self.log(f"🔍 Ditemukan {len(connotes_to_process)} connote untuk diproses.")

            for i, row in enumerate(connotes_to_process):
                if self.auto_update_stop_event.is_set() and not is_manual_run:
                    self.log("🛑 Proses dihentikan."); break
                
                connote = row['connote']
                try:
                    url = f"https://kibana.posindonesia.co.id:4433/x123449/3.php?id={connote}&6f017f90-f299-11ec-988f-6f1763dc6f47xdsdkjshhsahsaksasjsaasldsllsdjldsjsbdaksdslssjasjaa"
                    response = requests.get(url, timeout=20, verify=False)
                    response.raise_for_status()
                    soup = BeautifulSoup(response.text, "html.parser")
                    status_akhir_th = soup.find("th", string=re.compile(r"^\s*STATUS AKHIR\s*$"))
                    
                    if not status_akhir_th or not status_akhir_th.find_next_sibling("td"):
                        self.log(f"  ❌ Gagal: Tidak dapat menemukan 'STATUS AKHIR' untuk {connote}.")
                        failed_count += 1; failed_connotes.append(connote); continue

                    status_full = status_akhir_th.find_next_sibling("td").text.strip()
                    status = status_full.split(' Di ')[0].strip()
                    tgl_receiving = self.parse_date_from_status(status_full)
                    lokasi_part = status_full.split(',')[0]
                    # Ambil kata terakhir dari lokasi_part sebagai kc_akhir
                    lokasi_words = lokasi_part.strip().split()
                    kc_akhir = lokasi_words[-1] if lokasi_words else None

                    if not all([status, tgl_receiving, kc_akhir]):
                        self.log(f"  ❌ Gagal: Data dari web tidak lengkap untuk {connote}.")
                        failed_count += 1; failed_connotes.append(connote); continue

                    sql_update = "UPDATE tbl_db SET status=%s, tgl_receiving=%s, kc_akhir=%s WHERE connote=%s"
                    update_values = (status, tgl_receiving, kc_akhir, connote)
                    cursor.execute(sql_update, update_values)

                    if cursor.rowcount > 0:
                        conn.commit(); updated_count += 1
                        self.log(f"  ✅ SUKSES: Connote {connote} berhasil diupdate.")
                    else:
                        conn.rollback(); failed_count += 1; failed_connotes.append(connote)
                        self.log(f"  ⚠️ GAGAL: Connote {connote} tidak ditemukan saat mencoba update.")

                except Exception as e:
                    self.log(f"  ❌ Error memproses {connote}. Error: {e}")
                    failed_count += 1; failed_connotes.append(connote)
                
                time.sleep(0.5)

            cursor.close(); conn.close()
            summary = f"Proses Selesai. Berhasil: {updated_count}, Gagal: {failed_count}"
            self.log(f"🎉 {summary}")
            if is_manual_run: messagebox.showinfo("Proses Selesai", summary)

        except Exception as e:
            self.log(f"❌ Error fatal: {e}")
            traceback.print_exc()
        finally:
            self.set_tray_status('idle', f"Update selesai. Berhasil: {updated_count}, Gagal: {failed_count}")
            if is_manual_run: self.root.after(0, lambda: self.manual_button.config(state=tk.NORMAL))
            
            # Write failed connotes to log file
            if failed_connotes:
                log_file_path = r'C:\Users\POS\Music\PYTHON\App_Kantor\mile_gagal.log'
                try:
                    with open(log_file_path, 'a') as f:
                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        f.write(f"---\n")
                        for cn in failed_connotes:
                            f.write(f"{cn}\n")
                    self.log(f"💾 {len(failed_connotes)} connote gagal dicatat di {log_file_path}")
                except Exception as e:
                    self.log(f"🔴 Gagal menulis ke file log: {e}")

    def auto_update_loop(self):
        # ... unchanged ...
        while not self.auto_update_stop_event.is_set():
            self.run_batch_update(is_manual_run=False)
            if self.auto_update_stop_event.is_set(): break
            interval = self.interval_var.get()
            self.log(f"😴 Menunggu {interval} detik...")
            self.auto_update_stop_event.wait(interval)
        self.log("Auto-update loop berhenti.")

    def start_manual_update_thread(self):
        # ... unchanged ...
        threading.Thread(target=self.run_batch_update, args=(True,), daemon=True).start()

    def start_auto_update_thread(self):
        # ... unchanged ...
        self.auto_update_stop_event.clear()
        self.start_auto_button.config(state=tk.DISABLED); self.stop_auto_button.config(state=tk.NORMAL)
        self.manual_button.config(state=tk.DISABLED); self.interval_slider.config(state=tk.DISABLED)
        self.log("▶️ Memulai auto-update...")
        self.auto_update_thread = threading.Thread(target=self.auto_update_loop, daemon=True)
        self.auto_update_thread.start()

    def stop_auto_update(self):
        # ... unchanged ...
        self.log("⏹️ Menghentikan auto-update...")
        self.auto_update_stop_event.set()
        self.start_auto_button.config(state=tk.NORMAL); self.stop_auto_button.config(state=tk.DISABLED)
        self.manual_button.config(state=tk.NORMAL); self.interval_slider.config(state=tk.NORMAL)

    def hide_to_tray(self):
        # ... unchanged ...
        self.root.withdraw()
        self.log("Aplikasi berjalan di system tray.")

    def show_from_tray(self, icon, item):
        # ... unchanged ...
        self.root.deiconify()

    def run_manual_from_tray(self, icon, item):
        # ... unchanged ...
        self.log("⚙️ Memulai update manual dari tray...")
        self.start_manual_update_thread()

    def quit_app(self, icon, item):
        # ... unchanged ...
        self.log("👋 Menutup aplikasi...")
        if self.auto_update_thread and self.auto_update_thread.is_alive():
            self.auto_update_stop_event.set()
        self.tray_icon.stop()
        self.root.destroy()

    def setup_tray(self):
        # ... uses the pre-loaded self.icon_idle ...
        menu = (
            pystray.MenuItem('Tampilkan', self.show_from_tray, default=True),
            pystray.MenuItem('Jalankan Update Manual', self.run_manual_from_tray),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem('Exit', self.quit_app)
        )
        self.tray_icon = pystray.Icon("MileUpdaterApp", self.icon_idle, "Mile Updater App", menu)
        self.tray_icon.run()

if __name__ == "__main__":
    requests.packages.urllib3.disable_warnings(requests.packages.urllib3.exceptions.InsecureRequestWarning)
    root = tk.Tk()
    app = MileUpdaterApp(root)
    tray_thread = threading.Thread(target=app.setup_tray, daemon=True)
    tray_thread.start()
    root.mainloop()
