
import tkinter as tk
from tkinter import scrolledtext
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import threading
import queue
import sys
import os
import pymysql
import requests
import re
import base64
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# --- Logic from update_sla.py ---

def get_db_connection(log_queue):
    """Establishes a database connection."""
    try:
        conn = pymysql.connect(
            host=os.getenv("DB_HOST"),
            user=os.getenv("DB_USER"),
            password=os.getenv("DB_PASS"),
            database=os.getenv("DB_NAME"),
            port=int(os.getenv("DB_PORT", 3306)),
            cursorclass=pymysql.cursors.DictCursor
        )
        log_queue.put("INFO: Database connection successful.")
        return conn
    except pymysql.MySQLError as e:
        log_queue.put(f"ERROR: Error connecting to MySQL: {e}")
        return None

def fetch_connotes_to_process(conn, log_queue):
    """Fetches connotes from tbl_antrn where st = 33."""
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT connote FROM tbl_antrn WHERE sla = '0'")
            results = cursor.fetchall()
            return [row['connote'] for row in results]
    except pymysql.MySQLError as e:
        log_queue.put(f"ERROR: Error fetching connotes: {e}")
        return []

def get_sla_from_web(connote, log_queue):
    """Scrapes the SLA value for a given connote."""
    try:
        encoded_connote = base64.urlsafe_b64encode(connote.encode()).decode('ascii')
        url = f"https://pid.posindonesia.co.id/lacak/admin/detail_lacak_banyak.php?id={encoded_connote}"
        
        response = requests.get(url, timeout=10)
        response.raise_for_status()

        match = re.search(r"SLA\s*:\s*(\d+)\s*hari", response.text)
        if match:
            sla_value = int(match.group(1))
            log_queue.put(f"SUCCESS: Connote {connote} -> SLA: {sla_value}")
            return sla_value
        else:
            log_queue.put(f"INFO: SLA value not found for connote {connote}")
            return None
    except requests.exceptions.RequestException as e:
        log_queue.put(f"ERROR: Could not fetch data for connote {connote}. Reason: {e}")
        return None

def update_sla_in_db(conn, connote, sla_value, log_queue):
    """Updates the SLA value for a connote in the database."""
    try:
        with conn.cursor() as cursor:
            sql = "UPDATE tbl_antrn SET sla = %s WHERE connote = %s"
            cursor.execute(sql, (sla_value, connote))
        conn.commit()
        # log_queue.put(f"DB_UPDATE: Updated SLA for {connote}")
    except pymysql.MySQLError as e:
        log_queue.put(f"ERROR: Error updating database for connote {connote}: {e}")
        conn.rollback()

def run_update_process(log_queue, status_var, progress_callback):
    """Main function to run the SLA update process."""
    try:
        status_var.set("Running...")
        log_queue.put("Starting SLA update process...")
        db_conn = get_db_connection(log_queue)
        if not db_conn:
            status_var.set("Finished with errors.")
            progress_callback(0, 1) # Reset progress
            return

        connotes = fetch_connotes_to_process(db_conn, log_queue)
        if not connotes:
            log_queue.put("No connotes to process with sla = '0'.")
            db_conn.close()
            status_var.set("Finished.")
            progress_callback(0, 1) # Reset progress
            return
            
        total_connotes = len(connotes)
        log_queue.put(f"Found {total_connotes} connote(s) to process.")
        progress_callback(0, total_connotes)

        for i, connote in enumerate(connotes):
            sla = get_sla_from_web(connote, log_queue)
            if sla is not None:
                update_sla_in_db(db_conn, connote, sla, log_queue)
            # Update progress after each item is processed
            progress_callback(i + 1, total_connotes)

        db_conn.close()
        log_queue.put("SLA update process finished.")
        status_var.set("Finished.")
    except Exception as e:
        log_queue.put(f"FATAL_ERROR: An unexpected error occurred: {e}")
        status_var.set("Finished with errors.")


# --- GUI Application ---

class SlaUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SLA Updater")
        self.root.geometry("700x550") # Increased height for progress bar

        self.log_queue = queue.Queue()

        # --- Widgets ---
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=BOTH, expand=YES)

        header_label = ttk.Label(main_frame, text="SLA Update Process", font=("Helvetica", 16, "bold"))
        header_label.pack(pady=5)

        self.log_text = scrolledtext.ScrolledText(main_frame, state='disabled', height=20, font=("Courier New", 9))
        self.log_text.pack(fill=BOTH, expand=YES, padx=10, pady=10)

        # Progress Bar and Label
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=X, padx=10, pady=5)

        self.progress_bar = ttk.Progressbar(progress_frame, orient='horizontal', mode='determinate')
        self.progress_bar.pack(fill=X, expand=YES, side=LEFT, padx=(0, 10))

        self.progress_text_var = tk.StringVar(value="N/A")
        self.progress_label = ttk.Label(progress_frame, textvariable=self.progress_text_var, font=("Helvetica", 9))
        self.progress_label.pack(side=RIGHT)
        
        self.start_button = ttk.Button(main_frame, text="Start Update Process", command=self.start_process_thread, style='success.TButton')
        self.start_button.pack(pady=10, fill=X, padx=10)

        self.status_var = tk.StringVar(value="Idle")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=SUNKEN, anchor=W, padding=5)
        status_bar.pack(side=BOTTOM, fill=X)

        self.root.after(100, self.process_log_queue)

    def log(self, message):
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.configure(state='disabled')
        self.log_text.see(tk.END)

    def process_log_queue(self):
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log(message)
        except queue.Empty:
            pass
        self.root.after(100, self.process_log_queue)

    def update_progress(self, current, total):
        if total > 0:
            percent = (current / total) * 100
            self.progress_bar['value'] = percent
            self.progress_text_var.set(f"{current}/{total} ({percent:.1f}%)")
        else:
            self.progress_bar['value'] = 0
            self.progress_text_var.set("N/A")
        self.root.update_idletasks()


    def start_process_thread(self):
        self.start_button.config(state=DISABLED)
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
        self.update_progress(0, 0) # Reset progress bar
        
        self.thread = threading.Thread(target=run_update_process, args=(self.log_queue, self.status_var, self.update_progress))
        self.thread.daemon = True
        self.thread.start()
        self.root.after(100, self.check_thread)

    def check_thread(self):
        if self.thread.is_alive():
            self.root.after(100, self.check_thread)
        else:
            self.start_button.config(state=NORMAL)


if __name__ == "__main__":
    root = ttk.Window(themename="cosmo")
    app = SlaUpdaterApp(root)
    root.mainloop()
