import tkinter as tk
from tkinter import messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import mysql.connector
mysql.connector.connect = __import__('pymysql').connect
import pandas as pd
from dotenv import load_dotenv
import os
import traceback

# --- Console Logging ---
# Menambahkan fungsi ini untuk memastikan print() muncul di konsol
# bahkan jika dijalankan sebagai file .pyw dari command prompt.
import sys
sys.stdout = sys.__stdout__
sys.stderr = sys.__stderr__
print("---" + "-" * 10 + " Log Start " + "-" * 10 + "---")

load_dotenv(dotenv_path='conn_pens.env')

class ExportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Export Data Pensiun")
        self.root.geometry("800x600")
        self.style = ttk.Style(theme='cosmo')

        self.db_data = []

        # Main frame
        self.frame = ttk.Frame(self.root, padding="10")
        self.frame.pack(fill=BOTH, expand=YES)

        # Top frame for buttons
        self.top_frame = ttk.Frame(self.frame)
        self.top_frame.pack(fill=X, pady=5)

        self.show_data_button = ttk.Button(self.top_frame, text="Tampilkan Data", command=self.show_data, style="primary.TButton")
        self.show_data_button.pack(side=LEFT, padx=5)

        self.export_button = ttk.Button(self.top_frame, text="Export ke Excel", command=self.export_data, style="success.TButton", state=DISABLED)
        self.export_button.pack(side=LEFT, padx=5)

        # Treeview frame
        self.tree_frame = ttk.Frame(self.frame)
        self.tree_frame.pack(fill=BOTH, expand=YES, pady=10)

        self.columns = ("nP", "norek", "nmP", "ktb")
        self.tree = ttk.Treeview(self.tree_frame, columns=self.columns, show="headings", height=15)
        
        for col in self.columns:
            self.tree.heading(col, text=col.upper())
            self.tree.column(col, width=150)

        self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        self.vsb.pack(side=RIGHT, fill=Y)
        self.hsb.pack(side=BOTTOM, fill=X)
        self.tree.pack(fill=BOTH, expand=YES)
        
        # Bottom frame for progress bar
        self.bottom_frame = ttk.Frame(self.frame)
        self.bottom_frame.pack(fill=X, pady=5)
        
        self.progress = ttk.Progressbar(self.bottom_frame, orient=HORIZONTAL, length=300, mode='determinate')
        self.progress.pack(pady=5)

    def show_data(self):
        print("\n[LOG] Tombol 'Tampilkan Data' diklik.")
        conn = None
        cursor = None
        try:
            self.progress['value'] = 10
            self.root.update_idletasks()
            
            print("[LOG] Membersihkan data Treeview lama.")
            for i in self.tree.get_children():
                self.tree.delete(i)
            self.db_data = []
            self.export_button.config(state=DISABLED)

            db_host = os.getenv("DB_HOST")
            db_user = os.getenv("DB_USER")
            db_password = os.getenv("DB_PASSWORD")
            db_name = os.getenv("DB_NAME")
            
            print(f"[LOG] Kredensial DB: HOST={db_host}, USER={db_user}, DB={db_name}, PASS_LEN={len(db_password) if db_password else 0}")

            if not all([db_host, db_user, db_password, db_name]):
                print("[ERROR] File .env tidak lengkap.")
                messagebox.showerror("Error", "Pastikan file .env sudah terisi dengan benar (DB_HOST, DB_USER, DB_PASSWORD, DB_NAME)")
                self.progress['value'] = 0
                return

            self.progress['value'] = 20
            self.root.update_idletasks()

            print("[LOG] Mencoba koneksi ke database menggunakan PyMySQL...")
            conn = mysql.connector.connect(
                host=db_host, user=db_user, password=db_password, database=db_name,
                connect_timeout=10
            )
            cursor = conn.cursor()
            print("[LOG] Koneksi database berhasil.")
            
            self.progress['value'] = 40
            self.root.update_idletasks()

            query = "SELECT nP, norek, nmP, ktb FROM pens_tblpens"
            print(f"[LOG] Menjalankan query: {query}")
            cursor.execute(query)
            self.db_data = cursor.fetchall()
            print(f"[LOG] Query selesai. {len(self.db_data)} baris data diterima.")
            
            self.progress['value'] = 60
            self.root.update_idletasks()

            if self.db_data:
                print(f"[LOG] Contoh baris data pertama (sebelum konversi): {self.db_data[0]}")
                print("[LOG] Memulai proses memasukkan data ke Treeview...")
                for i, row in enumerate(self.db_data):
                    # Konversi setiap item di baris menjadi string untuk mencegah error
                    str_row = [str(item) if item is not None else "" for item in row]
                    self.tree.insert("", END, values=str_row)
                print(f"[LOG] Selesai memasukkan {len(self.db_data)} baris ke Treeview.")

            self.progress['value'] = 100
            self.root.update_idletasks()
            
            if self.db_data:
                self.export_button.config(state=NORMAL)
                print("[LOG] Menampilkan pesan sukses.")
                messagebox.showinfo("Sukses", f"{len(self.db_data)} baris data berhasil ditampilkan.")
            else:
                print("[LOG] Tidak ada data, menampilkan pesan info.")
                messagebox.showinfo("Info", "Tidak ada data untuk ditampilkan.")

        except Exception as e:
            print("\n--- TRACEBACK ERROR ---")
            traceback.print_exc()
            print("--- END TRACEBACK ---")
            messagebox.showerror("Error Kritis", f"Terjadi error yang tidak terduga. Lihat konsol untuk detail.\n\nError: {e}")
        finally:
            if cursor:
                print("[LOG] Menutup cursor.")
                cursor.close()
            if conn and conn.is_connected():
                print("[LOG] Menutup koneksi database.")
                conn.close()
            self.progress['value'] = 0
            print("[LOG] Proses selesai.")

    def export_data(self):
        print("\n[LOG] Tombol 'Export ke Excel' diklik.")
        if not self.db_data:
            messagebox.showwarning("Peringatan", "Tidak ada data untuk di-export. Silakan tampilkan data terlebih dahulu.")
            return
            
        try:
            output_filename = "data_pensiun.xlsx"
            print(f"[LOG] Memulai export ke {output_filename}")
            df = pd.DataFrame(self.db_data, columns=self.columns)
            df.to_excel(output_filename, index=False)
            print("[LOG] Export berhasil.")
            messagebox.showinfo("Sukses", f"Data berhasil di-export ke {output_filename}")

        except Exception as e:
            print("\n--- TRACEBACK ERROR (Export) ---")
            traceback.print_exc()
            print("--- END TRACEBACK ---")
            messagebox.showerror("Error Export", f"Gagal export data: {e}")


if __name__ == "__main__":
    root = ttk.Window(themename="cosmo")
    app = ExportApp(root)
    root.mainloop()