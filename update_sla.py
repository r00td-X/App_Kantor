
import os
import pymysql
import requests
import re
import base64

def get_env_vars(env_file=".env"):
    """Manually parses a .env file and returns a dictionary of variables."""
    vars = {}
    try:
        with open(env_file) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    key, value = line.split('=', 1)
                    vars[key.strip()] = value.strip()
    except FileNotFoundError:
        print(f"ERROR: .env file not found at {os.path.abspath(env_file)}")
    return vars

# Get database credentials by manually parsing .env
env_vars = get_env_vars()
db_host = env_vars.get("DB_HOST")
db_user = env_vars.get("DB_USER")
db_password = env_vars.get("DB_PASSWORD")
db_name = env_vars.get("DB_NAME")
db_port = int(env_vars.get("DB_PORT", 3306))

import os
import pymysql
import requests
import re
import base64
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Get database credentials by manually parsing .env
db_host = os.getenv("DB_HOST")
db_user = os.getenv("DB_USER")
db_password = os.getenv("DB_PASSWORD")
db_name = os.getenv("DB_NAME")
db_port = int(os.getenv("DB_PORT", 3306))

def get_db_connection():
    """Establishes a database connection."""
    try:
        conn = pymysql.connect(
            host=db_host,
            user=db_user,
            password=db_password,
            database=db_name,
            port=db_port,
            cursorclass=pymysql.cursors.DictCursor
        )
        return conn
    except pymysql.MySQLError as e:
        print(f"Error connecting to MySQL: {e}")
        return None

def fetch_connotes_to_process(conn):
    """Fetches connotes from tbl_antrn where st = 33."""
    try:
        with conn.cursor() as cursor:
            cursor.execute("SELECT connote FROM tbl_antrn WHERE st = 33")
            results = cursor.fetchall()
            return [row['connote'] for row in results]
    except pymysql.MySQLError as e:
        print(f"Error fetching connotes: {e}")
        return []

def get_sla_from_web(connote):
    """Scrapes the SLA value for a given connote."""
    try:
        # Encode connote to Base64 for the URL
        encoded_connote = base64.urlsafe_b64encode(connote.encode()).decode('ascii')
        url = f"https://pid.posindonesia.co.id/lacak/admin/detail_lacak_banyak.php?id={encoded_connote}"
        
        response = requests.get(url, timeout=10)
        response.raise_for_status()  # Raise an exception for bad status codes

        # Search for the SLA value using regex
        # Pattern: SLA : (\d+) hari
        match = re.search(r"SLA\s*:\s*(\d+)\s*hari", response.text)
        if match:
            sla_value = int(match.group(1))
            print(f"SUCCESS: Connote {connote} -> SLA: {sla_value}")
            return sla_value
        else:
            print(f"INFO: SLA value not found for connote {connote}")
            return None
    except requests.exceptions.RequestException as e:
        print(f"ERROR: Could not fetch data for connote {connote}. Reason: {e}")
        return None

def update_sla_in_db(conn, connote, sla_value):
    """Updates the SLA value for a connote in the database."""
    try:
        with conn.cursor() as cursor:
            sql = "UPDATE tbl_antrn SET sla = %s WHERE connote = %s"
            cursor.execute(sql, (sla_value, connote))
        conn.commit()
        # print(f"DB_UPDATE: Updated SLA for {connote}")
    except pymysql.MySQLError as e:
        print(f"Error updating database for connote {connote}: {e}")
        conn.rollback()

def main():
    """Main function to run the SLA update process."""
    print("Starting SLA update process...")
    db_conn = get_db_connection()
    if not db_conn:
        return

    connotes = fetch_connotes_to_process(db_conn)
    if not connotes:
        print("No connotes to process with st = 33.")
        db_conn.close()
        return
        
    print(f"Found {len(connotes)} connote(s) to process.")

    for connote in connotes:
        sla = get_sla_from_web(connote)
        if sla is not None:
            update_sla_in_db(db_conn, connote, sla)

    db_conn.close()
    print("SLA update process finished.")

if __name__ == "__main__":
    main()
