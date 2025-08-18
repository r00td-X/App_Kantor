
import tkinter as tk
from tkinter import ttk, messagebox
from ttkbootstrap.constants import *
import ttkbootstrap as tb
from datetime import datetime
import requests
import locale
import math
import re # Untuk regex
import webbrowser # Untuk membuka URL di browser

class R7App(tb.Window):
    def __init__(self):
        super().__init__(themename="cosmo")
        self.title("Aplikasi Data R7 - Filter & Pagination")
        self.geometry("950x750")

        self.all_data = []
        self.current_page = 1
        self.items_per_page = 100

        self.create_widgets()
        self.load_initial_data()

    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=BOTH, expand=True)

        # --- Kontrol ---
        control_frame = ttk.LabelFrame(main_frame, text="Kontrol & Filter", padding="10")
        control_frame.pack(fill=X, pady=5)

        self.date_format = '%m/%d/%y'
        ttk.Label(control_frame, text="Date From:").grid(row=0, column=0, padx=5, pady=5, sticky=W)
        self.date_from_entry = tb.DateEntry(control_frame, bootstyle="primary", dateformat=self.date_format)
        self.date_from_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(control_frame, text="Date To:").grid(row=0, column=2, padx=5, pady=5, sticky=W)
        self.date_to_entry = tb.DateEntry(control_frame, bootstyle="primary", dateformat=self.date_format)
        self.date_to_entry.grid(row=0, column=3, padx=5, pady=5)

        self.fetch_button = tb.Button(control_frame, text="Ambil Data", command=self.fetch_and_process_data, bootstyle="success")
        self.fetch_button.grid(row=0, column=4, padx=10, pady=5, ipady=5)

        # --- Filter Teks ---
        ttk.Label(control_frame, text="Filter Kantor Tujuan:").grid(row=1, column=0, padx=5, pady=5, sticky=W)
        self.filter_var = tk.StringVar()
        self.filter_entry = tb.Entry(control_frame, textvariable=self.filter_var, width=50)
        self.filter_entry.grid(row=1, column=1, columnspan=3, padx=5, pady=5, sticky=W+E)
        self.filter_var.trace_add("write", self.on_filter_change) # Panggil on_filter_change saat teks berubah
        
        self.status_label = tb.Label(control_frame, text="Status: Menunggu...", bootstyle="info")
        self.status_label.grid(row=2, column=0, columnspan=5, pady=5, sticky=W)

        # --- Treeview ---
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill=BOTH, expand=True, pady=10)
        self.tree = ttk.Treeview(tree_frame, columns=("No R7", "Kantor Tujuan", "Tgl Manifest"), show="headings", bootstyle="primary")
        self.tree.heading("No R7", text="No R7"); self.tree.heading("Kantor Tujuan", text="Kantor Tujuan"); self.tree.heading("Tgl Manifest", text="Tgl Manifest")
        self.tree.column("No R7", width=250); self.tree.column("Kantor Tujuan", width=350); self.tree.column("Tgl Manifest", width=200)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview); hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side=RIGHT, fill=Y); hsb.pack(side=BOTTOM, fill=X); self.tree.pack(fill=BOTH, expand=True)

        # Bind click event to treeview
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        # --- Pagination ---
        pagination_frame = ttk.Frame(main_frame)
        pagination_frame.pack(fill=X, pady=5)
        self.prev_button = tb.Button(pagination_frame, text="<< Sebelumnya", command=self.prev_page, bootstyle="secondary"); self.prev_button.pack(side=LEFT, padx=5)
        self.page_label = tb.Label(pagination_frame, text="Halaman 1 / 1"); self.page_label.pack(side=LEFT, padx=20)
        self.next_button = tb.Button(pagination_frame, text="Berikutnya >>", command=self.next_page, bootstyle="secondary"); self.next_button.pack(side=LEFT, padx=5)

    def fetch_api_data(self, date_from, date_to):
        self.status_label.config(text="Status: Menghubungi API...", bootstyle="warning"); self.update_idletasks()
        try: locale.setlocale(locale.LC_TIME, 'en_US.UTF-8')
        except locale.Error: locale.setlocale(locale.LC_TIME, 'English_United States.1252')
        date_from_str = date_from.strftime('%d %B %Y'); date_to_str = date_to.strftime('%d %B %Y')
        locale.setlocale(locale.LC_TIME, '')

        api_url = "https://posindo.mile.app/api/manifestR7-filter"
        initial_payload = {"length": 1, "date_from": date_from_str, "date_to": date_to_str, "draw": 1, "start": 0}
        try:
            self.status_label.config(text="Status: Mengambil jumlah total data...", bootstyle="info"); self.update_idletasks()
            initial_response = requests.post(api_url, data=initial_payload, timeout=30)
            initial_response.raise_for_status()
            records_total = initial_response.json().get('recordsTotal', 0)
            if records_total == 0: return None

            self.status_label.config(text=f"Status: Mengunduh {records_total} baris data...", bootstyle="info"); self.update_idletasks()
            full_payload = {"length": records_total, "date_from": date_from_str, "date_to": date_to_str, "draw": 2, "start": 0}
            full_response = requests.post(api_url, data=full_payload, timeout=180)
            full_response.raise_for_status()
            return full_response.json()
        except requests.exceptions.RequestException as e:
            messagebox.showerror("Koneksi Gagal", f"Tidak dapat mengambil data dari server: {e}"); return None

    def process_data(self, api_response):
        self.all_data = []
        if not api_response or 'data' not in api_response: return
        for item in api_response['data']:
            if len(item) > 12: # Pastikan ada kolom untuk taskId
                no_r7, kantor_tujuan, tgl_manifest_str = item[1], item[9], item[10]
                action_html = item[12] # Kolom terakhir berisi HTML dengan taskId
                
                # Ekstrak taskId menggunakan regex
                match = re.search(r'taskId=([a-f0-9]+)', action_html)
                task_id = match.group(1) if match else None

                if task_id:
                    try: 
                        tgl_manifest = datetime.strptime(tgl_manifest_str, '%Y-%m-%d %H:%M:%S')
                        self.all_data.append((no_r7, kantor_tujuan, tgl_manifest, task_id))
                    except (ValueError, IndexError): continue
        self.all_data.sort(key=lambda x: x[2], reverse=True)

    def on_filter_change(self, *args):
        self.current_page = 1
        self.display_data()

    def display_data(self):
        self.tree.delete(*self.tree.get_children())
        filter_text = self.filter_var.get().lower()
        
        if filter_text:
            filtered_data = [item for item in self.all_data if filter_text in item[1].lower()]
        else:
            filtered_data = self.all_data

        total_items = len(filtered_data)
        if total_items == 0:
            self.page_label.config(text="Halaman 0 / 0"); self.prev_button.config(state=DISABLED); self.next_button.config(state=DISABLED)
            self.tree.insert("", "end", values=("Tidak ada data yang cocok.", "", "")); return

        total_pages = math.ceil(total_items / self.items_per_page)
        self.page_label.config(text=f"Halaman {self.current_page} / {total_pages}")
        start_index = (self.current_page - 1) * self.items_per_page
        end_index = start_index + self.items_per_page
        page_data = filtered_data[start_index:end_index]

        for item in page_data:
            # item sekarang berisi (no_r7, kantor_tujuan, tgl_manifest, task_id)
            # Kita menampilkan 3 kolom pertama, dan menyimpan task_id sebagai tag
            self.tree.insert("", "end", values=(item[0], item[1], item[2].strftime('%d-%m-%Y %H:%M:%S')), tags=(item[3],))

        self.prev_button.config(state=NORMAL if self.current_page > 1 else DISABLED)
        self.next_button.config(state=NORMAL if self.current_page < total_pages else DISABLED)
        self.status_label.config(text=f"Status: Selesai. Menampilkan {total_items} dari {len(self.all_data)} total data.", bootstyle="success")

    def on_tree_select(self, event):
        selected_item = self.tree.focus()
        if selected_item:
            tags = self.tree.item(selected_item, "tags")
            if tags:
                task_id = tags[0] # taskId disimpan sebagai tag pertama
                url = f"https://posindo.mile.app/manifestR7/print?taskId={task_id}"
                webbrowser.open(url)

    def next_page(self):
        filter_text = self.filter_var.get().lower()
        total_items = len([item for item in self.all_data if filter_text in item[1].lower()]) if filter_text else len(self.all_data)
        total_pages = math.ceil(total_items / self.items_per_page)
        if self.current_page < total_pages: self.current_page += 1; self.display_data()

    def prev_page(self):
        if self.current_page > 1: self.current_page -= 1; self.display_data()

    def fetch_and_process_data(self):
        try:
            start_date = datetime.strptime(self.date_from_entry.entry.get(), self.date_format)
            end_date = datetime.strptime(self.date_to_entry.entry.get(), self.date_format)
        except ValueError: messagebox.showerror("Tanggal Salah", f"Format tanggal tidak valid."); return

        api_response = self.fetch_api_data(start_date, end_date)
        if api_response:
            self.process_data(api_response)
            self.filter_var.set("") # Hapus filter teks setelah data baru diambil
            self.display_data()

    def load_initial_data(self):
        self.fetch_and_process_data()

if __name__ == "__main__":
    app = R7App()
    app.mainloop()
