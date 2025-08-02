import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import pandas as pd
import mysql.connector
import os
from dotenv import load_dotenv
import logging
import threading
import queue

# --- Load Environment Variables ---
load_dotenv()

# --- Setup Logging ---
log_queue = queue.Queue()

# File logger
file_handler = logging.FileHandler('app_debug.log', 'w')
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

# Custom handler for the queue
class QueueHandler(logging.Handler):
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(record)

# Configure root logger
logging.getLogger().setLevel(logging.INFO)
logging.getLogger().addHandler(file_handler)
logging.getLogger().addHandler(QueueHandler(log_queue))

class App(tb.Window):
    def __init__(self):
        super().__init__(themename="cosmo")
        self.title("Upload Excel to MySQL")
        self.geometry("900x750")
        logging.info("Application started.")

        self.file_path = ""
        self.df = None

        self.create_widgets()
        self.after(100, self.process_log_queue)

    def create_widgets(self):
        main_frame = tb.Frame(self, padding=15)
        main_frame.pack(fill=BOTH, expand=True)

        top_frame = tb.Frame(main_frame)
        top_frame.pack(fill=BOTH, expand=True)

        # --- File Selection ---
        file_frame = tb.Labelframe(top_frame, text="Select Excel File", padding=15)
        file_frame.pack(fill=X, pady=(0, 10))

        self.file_entry = tb.Entry(file_frame, width=70)
        self.file_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 10))

        browse_button = tb.Button(file_frame, text="Browse", command=self.browse_file, bootstyle="info")
        browse_button.pack(side=LEFT)

        # --- Treeview Display ---
        tree_frame = tb.Labelframe(top_frame, text="Excel Data Preview", padding=15)
        tree_frame.pack(fill=BOTH, expand=True, pady=10)

        self.tree = ttk.Treeview(tree_frame, show="headings")
        self.tree.pack(fill=BOTH, expand=True)

        # --- Progress Bar and Status ---
        progress_frame = tb.Frame(top_frame, padding=(0, 10))
        progress_frame.pack(fill=X)

        self.progress_label = tb.Label(progress_frame, text="Ready to upload.")
        self.progress_label.pack(side=LEFT, padx=(0, 10))

        self.progress = tb.Progressbar(progress_frame, orient=HORIZONTAL, length=300, mode='determinate')
        self.progress.pack(fill=X, expand=True)

        # --- Upload Button ---
        self.upload_button = tb.Button(top_frame, text="Upload to MySQL", command=self.start_upload_thread, bootstyle="success")
        self.upload_button.pack(pady=(5, 15))

        # --- Log Viewer ---
        log_frame = tb.Labelframe(main_frame, text="Log", padding=10)
        log_frame.pack(fill=BOTH, expand=True, side=BOTTOM)
        self.log_view = scrolledtext.ScrolledText(log_frame, height=6, state='disabled', wrap=tk.WORD)
        self.log_view.pack(fill=BOTH, expand=True)

    def process_log_queue(self):
        while not log_queue.empty():
            record = log_queue.get()
            msg = file_handler.formatter.format(record)
            self.log_view.config(state='normal')
            self.log_view.insert(tk.END, msg + '\n')
            self.log_view.config(state='disabled')
            self.log_view.see(tk.END)
        self.after(100, self.process_log_queue)

    def browse_file(self):
        logging.info("Browsing for a file.")
        self.file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=(("Excel Files", "*.xlsx"), ("All files", "*.*" ))
        )
        if self.file_path:
            logging.info(f"File selected: {self.file_path}")
            self.file_entry.delete(0, END)
            self.file_entry.insert(0, self.file_path)
            self.load_excel_data()
        else:
            logging.info("File selection cancelled.")

    def load_excel_data(self):
        try:
            logging.info("Loading Excel data.")
            self.df = pd.read_excel(self.file_path)
            self.display_dataframe()
            logging.info("Excel data loaded and displayed successfully.")
        except Exception as e:
            logging.error(f"Failed to read Excel file: {e}", exc_info=True)
            messagebox.showerror("Error", f"Failed to read Excel file: {e}")

    def display_dataframe(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.tree["column"] = list(self.df.columns)
        self.tree["show"] = "headings"
        for col in self.tree["column"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        for index, row in self.df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def start_upload_thread(self):
        logging.info("Upload process initiated by user.")
        if self.df is None:
            messagebox.showwarning("Warning", "No data to upload. Please select an Excel file first.")
            logging.warning("Upload attempt with no data.")
            return
        self.upload_button.config(state=DISABLED)
        self.progress['value'] = 0
        self.progress_label.config(text="Starting upload...")
        thread = threading.Thread(target=self._upload_worker)
        thread.daemon = True
        thread.start()

    def _upload_worker(self):
        try:
            column_mapping = {
                'NOTAS': 'nP', 'NOMOR_REKENING_POS': 'norek2', 'NOREK BARU': 'norek',
                'NAMA PENERIMA': 'nmP', 'JNS': 'jP', 'ALAMAT': 'adP', 'KTB': 'ktb'
            }
            # Add the new fields to the mapping
            column_mapping_db = {
                'NOTAS': 'nP', 'NOMOR_REKENING_POS': 'norek2', 'NOREK BARU': 'norek',
                'NAMA PENERIMA': 'nmP', 'JNS': 'jP', 'ALAMAT': 'adP', 'KTB': 'ktb',
                'KTB': 'ktby' # Map KTB to ktby as well
            }

            required_columns = list(column_mapping.keys())
            missing_columns = [col for col in required_columns if col not in self.df.columns]
            if missing_columns:
                msg = f"Missing required columns in Excel file: {', '.join(missing_columns)}"
                logging.error(msg)
                messagebox.showerror("Error", msg)
                return

            logging.info("Connecting to database...")
            db_host = os.getenv("DB_HOST")
            db_user = os.getenv("DB_USER")
            db_password = os.getenv("DB_PASSWORD")
            db_name = os.getenv("DB_NAME")

            if not all([db_host, db_user, db_name]):
                logging.error("Database configuration is missing in .env file.")
                messagebox.showerror("Error", "Database configuration is missing in .env file.")
                return

            conn = mysql.connector.connect(
                host=db_host, user=db_user, password=db_password, database=db_name, connection_timeout=15, use_pure=True
            )
            cursor = conn.cursor()
            logging.info("Database connection successful.")

            table_name = "pens_tblpens"
            inserted_count = 0
            skipped_count = 0
            total_rows = len(self.df)
            self.progress['maximum'] = total_rows

            for index, row in self.df.iterrows():
                current_row_num = index + 1
                notas_value = row['NOTAS']
                self.progress['value'] = current_row_num
                self.progress_label.config(text=f"Processing {current_row_num}/{total_rows} | Inserted: {inserted_count} | Skipped: {skipped_count}")
                
                cursor.execute(f"SELECT nP FROM `{table_name}` WHERE nP = %s", (notas_value,))
                if cursor.fetchone():
                    skipped_count += 1
                    logging.info(f"Skipping NOTAS (duplicate): {notas_value}")
                    continue

                # Prepare data for insertion
                db_columns = list(column_mapping_db.values()) + ['sts']
                excel_values = [row[col] for col in column_mapping_db.keys()] + [1] # Add 1 for sts

                # Build the SQL query
                placeholders = ", ".join(["%s"] * len(db_columns))
                sql = f"INSERT INTO `{table_name}` (`" + "`, `".join(db_columns) + "`) VALUES ({placeholders})"
                
                cursor.execute(sql, tuple(excel_values))
                inserted_count += 1
                logging.info(f"Inserted NOTAS: {notas_value}")

            logging.info("Committing changes to the database.")
            conn.commit()
            cursor.close()
            conn.close()
            logging.info("Database connection closed.")

            final_msg = f"Upload finished.\n\nSuccessfully Inserted: {inserted_count} rows.\nSkipped (Duplicates): {skipped_count} rows."
            self.progress_label.config(text="Upload complete.")
            logging.info(f"Final result: Inserted={inserted_count}, Skipped={skipped_count}")
            messagebox.showinfo("Upload Complete", final_msg)

        except mysql.connector.Error as err:
            logging.error(f"Database Error: {err}", exc_info=True)
            self.progress_label.config(text="Database Error.")
            messagebox.showerror("Database Error", f"Error: {err}")
        except Exception as e:
            logging.error(f"An unexpected error occurred: {e}", exc_info=True)
            self.progress_label.config(text="An unexpected error occurred.")
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        finally:
            self.upload_button.config(state=NORMAL)
            logging.info("Upload process finished. Button re-enabled.")

if __name__ == "__main__":
    app = App()
    app.mainloop()
