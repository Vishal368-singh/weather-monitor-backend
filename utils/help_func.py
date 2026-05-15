from datetime import datetime, timedelta
from user_agents import parse
import re, json
import pytz
from collections import defaultdict
import pandas as pd
from io import BytesIO
import yagmail
import tempfile
import os
import subprocess
import sys
from utils.db import get_db_conn
from dotenv import load_dotenv
load_dotenv()
import glob
import requests
from urllib.parse import urlparse, unquote
from psycopg2.extras import RealDictCursor

def circle_name_cover_page(name):
    match = re.search(r'\((.*?)\)', name)
    if match:
        return match.group(1).strip()
    return name.replace("Upper North", "").strip()

def format_hazard_records(items):
    expanded = []
    today = datetime.now().date()   # Current date (YYYY-MM-DD)

    for item in items:
        # Extract numeric part from "Day 1", "Day 2", etc.
        day_str = item.get("day", "")
        day_num = int(day_str.replace("Day", "").strip())

        # Calculate date based on day number
        record_date = today + timedelta(days=day_num - 1)
        record_date_str = record_date.strftime("%d-%m-%Y") 

        # Convert list → comma-separated string
        district_str = ",".join(item.get("districts", []))

        expanded.append({
            "circle": item["circle"],
            "day": item["day"],
            "date": item["date"],
            "description": item["description"],
            "district": district_str,
            "hazardValue": item["hazardValue"],
            "severity": item["severity"]
        })

    return expanded


def format_device_name(device_str: str) -> str:
    if not device_str:
        return "an unknown device"

    parts = device_str.split("-")
    if len(parts) >= 4:
        device_type = parts[0].capitalize()
        browser = parts[2].capitalize()
        os_version = " ".join(parts[3:]).replace("windows", "Windows").replace("-", " ")
        return f"{device_type} through {browser}"
    else:
        return device_str
    

def get_device_label(user_agent_string: str) -> str:
    if not user_agent_string:
        return "Unknown Device"

    ua = parse(user_agent_string)

    # Device type
    if ua.is_mobile:
        device_type = "Mobile"
    elif ua.is_tablet:
        device_type = "Tablet"
    elif ua.is_pc:
        device_type = "PC"
    else:
        device_type = "Device"

    # OS name
    os_name = ua.os.family

    # Browser name + major version
    browser_name = ua.browser.family
    browser_version = ua.browser.version[0] if ua.browser.version else ""
    return f"{os_name} {device_type} | {browser_name} {browser_version}"


# Generate excel file
IST = pytz.timezone("Asia/Kolkata")
def make_excel_safe(value):
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            value = value.astimezone(IST).replace(tzinfo=None)
    return value

def export_user_activity_excel(data, summary_data):
    output = BytesIO()

    # Make data excel safe
    safe_data = [
        {k: make_excel_safe(v) for k, v in rec.items()}
        for rec in data
    ]

    safe_summary = [
        {k: make_excel_safe(v) for k, v in rec.items()}
        for rec in summary_data
    ]

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        # Sheet 1: Usage Summary
        summary_df = pd.DataFrame(safe_summary)
        if not summary_df.empty:
            # Rename columns
            summary_df = summary_df.rename(columns={
                "login_date": "Login Date",
                "name": "Name",
                "userid": "User ID",
                "login_count": "Login Count",
                "duration": "Duration (HH:MM:SS)"
            })

            # Ensure correct column order
            desired_summary_columns = [
                "Login Date",
                "Name",
                "User ID",
                "Login Count",
                "Duration (HH:MM:SS)"
            ]

            summary_df = summary_df[
                [col for col in desired_summary_columns if col in summary_df.columns]
            ]

            summary_df.to_excel(writer, sheet_name="Usage Summary", index=False)

        # Sheet 2: Usage Report
        df = pd.DataFrame(safe_data)
        if not df.empty:
            # Convert list column to string
            if "action" in df.columns:
                df["action"] = df["action"].apply(
                    lambda x: ", ".join(x) if isinstance(x, list) else x
                )

            # df = df.sort_values(by="id")
            # Rename columns
            df = df.rename(columns={
                "id": "Record ID",
                "name": "Full Name",
                "userid": "User ID",
                "login_time": "Login Date & Time",
                "logout_time": "Logout Date & Time",
                "loggedin_device": "Device",
                "action": "User Actions"
            })

            desired_columns = [
                "Record ID",
                "Full Name",
                "User ID",
                "Login Date & Time",
                "Logout Date & Time",
                "Device",
                "User Actions"
            ]

            df = df[[col for col in desired_columns if col in df.columns]]

            df.to_excel(writer, sheet_name="Usage Report", index=False)

    output.seek(0)
    return output

# Send mail
def send_mail_with_excel_bytes(to_emails, excel_bytes):

    EMAIL = os.environ.get("EMAIL")
    PASSWORD = os.environ.get("PASSWORD")
    CC_RECIVERS = os.environ.get("CC_RECEIVERS")
    
    html_body = f"""
       Dear Receivers,\n
       Attached is the dashboard usage report.\n\n
    """
    html_body2 = f"""
       Regards,\n
       ML Infomap | Weather Services
    """
    
    yag = yagmail.SMTP(user=EMAIL, password=PASSWORD)
        
    tmp_dir = tempfile.gettempdir()
    file_path = os.path.join(tmp_dir, "User_Activity_Report.xlsx")

    with open(file_path, "wb") as f:
        f.write(excel_bytes.getvalue())

    try:
        yag.send(
            # to=['rizwan@mlinfomap.com'],
            to=to_emails,
            # cc=CC_RECIVERS,
            subject=f'Dashboard Usage Report',
            contents=[html_body, html_body2],
            # contents=[html_body],
            attachments=[file_path]
        )
    finally:
        os.remove(file_path)
    return True

def remove_lat_long(text):
    if not text:
        return text

    return re.sub(
        r'\s*,?\s*Lat:\s*[-+]?\d*\.?\d+\s*&\s*Long:\s*[-+]?\d*\.?\d+',
        '',
        text
    ).strip()

def insert_user_management_history(user_data, conn):
    try:
        name = user_data.get("name", None)
        userid = user_data.get("userid")
        modified_data = user_data.get("modified_data")
        modified_by = user_data.get("modified_by")
        modifier_role = user_data.get("modifier_role")
        user_type_flag = user_data.get("user_type_flag")
        
        for rec in modified_data:
            action_on = rec.get("key")
            old_value = rec.get("prev_value")
            new_value = rec.get("new_value")

            if action_on == "indus_circle" or action_on == "restore":
                old_value = ", ".join(old_value) if isinstance(old_value, list) else old_value
                new_value = ", ".join(new_value) if isinstance(new_value, list) else new_value

            with conn.cursor() as cursor:
                query = f"""
                INSERT INTO weatherdata.user_management_history
                    ("name", old_value, new_value, userid, modified_by, modifier_role, modified_on, action_on, user_type_flag)
                VALUES('{name}', '{old_value}', '{new_value}', '{userid}','{modified_by}', '{modifier_role}', now(), '{action_on}', '{user_type_flag}');
                """
                cursor.execute(query)
            conn.commit() 
    except Exception as e:
        conn.rollback()
        raise e
    
def insert_kpi_update_history(kpi_data, conn):
    try:
        kpi_name = kpi_data.get("kpi_name", None)
        indus_circle = kpi_data.get("indus_circle", None)
        modified_data = kpi_data.get("modified_data")

        modified_by = kpi_data.get("modified_by")
        modifier_role = kpi_data.get("modifier_role")
        
        for rec in modified_data:
            action_on = rec.get("key")
            
            old_value = rec.get("prev_value")
            new_value = rec.get("new_value")
            with conn.cursor() as cursor:
                query = f"""
                    INSERT INTO weatherdata.kpi_modify_history
                    (kpi_name, old_value, new_value, indus_circle, modified_by, modifier_role, modified_on, action_on)
                    VALUES( '{kpi_name}', '{old_value}', '{new_value}', '{indus_circle}', '{modified_by}', '{modifier_role}', now(), '{action_on.capitalize()}');
                """
                cursor.execute(query)
            conn.commit() 
    except Exception as e:
        conn.rollback()
        raise e
    
def get_indus_circle_location(indus_circle, conn):
    try:
        indus_circle = indus_circle if indus_circle != 'All Circle' else 'M&G'
        with conn.cursor() as cursor:
            query = """
                SELECT yy, xx, location_name
                FROM weatherdata.indus_circle_geomerty
                WHERE indus_circle = %s;
            """
            cursor.execute(query, (indus_circle,))
            location = cursor.fetchone()

            if not location:
                return None

            yy, xx, location_name = location
            location = {
                "location" : f"{yy},{xx}",
                "location_name" : location_name
            }
            return location
    except Exception as e:
        raise e

def update_trigger_status(execution_status, hazard_type):
    try:
        conn = get_db_conn()
        with conn.cursor() as cursor:
            query = """
                UPDATE weatherdata.trigger_master
                SET execution_status = %s,
                    last_update = now()
                WHERE trigger_name = %s;
            """
            cursor.execute(query, (execution_status, hazard_type))
        conn.commit() 
    except Exception as e:
        conn.rollback()
        raise e

def fetch_hazard_cyclone_data():
    try:
        conn = get_db_conn()
        cursor = conn.cursor(cursor_factory=RealDictCursor)

        circles = get_impacted_circles()
        indus_circle = circles.get("name").split(",")

        # Defensive check
        if not indus_circle:
            return []

        # Prepare placeholders for SQL IN clause
        placeholders = ", ".join(["%s"] * len(indus_circle))

        query = f"""
        SELECT *
        FROM weatherdata.hazard_cyclone
        WHERE insert_at = (
            SELECT MAX(insert_at)
            FROM weatherdata.hazard_cyclone
        )
        AND indus_circle IN ({placeholders})
        AND date >= CURRENT_DATE
        AND date < CURRENT_DATE + INTERVAL '3 day'
    """

        # ✅ Pass values here
        cursor.execute(query, indus_circle)

        results = cursor.fetchall()

        cursor.close()
        conn.close()

        return [dict(row) for row in results] if results else []

    except Exception as e:
        print("Error :", e)
        return []

def get_impacted_circles():
    conn = None
    try:
        query = """
            SELECT *
            FROM weatherdata.cyclone_impacted_circles
            ORDER BY inserted_at DESC
            LIMIT 1
        """

        conn = get_db_conn()
        df = pd.read_sql_query(query, conn)

        if df.empty:
            return None

        # Return first row as {field: value}
        return df.iloc[0].to_dict()

    except Exception as e:
        return None
    except Exception as e:
        return None

# NDMA Disaster helper
def normalize(text):
    text = text.lower()
    text = re.sub(r'[^a-z\s]', '', text) 
    return text

def get_ndma_hazards_events(hazard_list):
    import re

    keywords = [
        {"label": "Rainfall", "value": "Rain"},
        {"label": "Thunderstorm", "value": "Thunderstorm"},
        {"label": "Lightning", "value": "Lightning"},
        {"label": "Flood", "value": "Flood"},
        {"label": "Landslide", "value": "Landslide"},
        {"label": "Avalanche", "value": "Avalanche"},
        {"label": "Fog", "value": "Fog"},
        {"label": "Snowfall", "value": "Snowfall"},
        {"label": "Heat Wave", "value": "Heat Wave"},
        {"label": "Cold Wave", "value": "Cold Wave"},
        {"label": "Earthquake", "value": "Earthquake"},
    ]

    matched = set()

    for event in hazard_list:
        event_norm = normalize(event)
        for k in keywords:
            key = normalize(k["value"])
            # match singular or plural
            if key in event_norm or key + "s" in event_norm:
                matched.add(k["label"])

    # print(sorted(matched))
    return sorted(matched)

def get_latest_cyclone_pdf_file():

    directory = r"C:\inetpub\wwwroot\weather\reports\pdf_cyclone"
    pdf_files = glob.glob(os.path.join(directory, "*.pdf"))
    if not pdf_files:
        latest_pdf = None
    else:
        latest_pdf = max(pdf_files, key=os.path.getmtime)
    filename = os.path.basename(latest_pdf) 
    return filename

# Cyclone data insert using subprocess
def insert_cyclone_data():
    try:
        update_trigger_status("RUNNING", 'cyclone_data')
        subprocess.Popen(
            [
                os.environ["CYCLONE_PYTHON"],
                os.environ["CYCLONE_SCRIPT"]
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            start_new_session=True
        )
        update_trigger_status("SUCCESS", 'cyclone_data')
        return {"msg": "Cyclone data insert started"}
    except Exception as e:
        update_trigger_status("FAILED", 'cyclone_data')
        return {"error": str(e)}

def insert_hazard_data_subprocess():
    try:
        update_trigger_status("RUNNING", 'hazard_data')

        script_path = os.environ["HAZARD_SCRIPT"]
        python_path = os.environ["HAZARD_PYTHON"]

        result = subprocess.run(
            [python_path, script_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            cwd=os.path.dirname(script_path)
        )

        update_trigger_status("SUCCESS", 'hazard_data')
        return {"msg": "Cyclone data insert started"}
    except Exception as e:
        update_trigger_status("FAILED", 'hazard_data')
        return {"error": str(e)}

def generate_cyclone_report():
    try:
        script_path = os.environ["CYCLONE_REPORT_SCRIPT"]
        python_path = os.environ["CYCLONE_REPORT_PYTHON"]

        result = subprocess.run(
            [python_path, script_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            cwd=os.path.dirname(script_path)
        )

        pdf_file = get_latest_cyclone_pdf_file()

        return {
            "msg": "Report generated successfully",
            "file": pdf_file
        }

    except subprocess.CalledProcessError as e:
        return {
            "error": "Cyclone report generation failed",
            "details": e.stderr
        }

    except Exception as e:
        return {"error": str(e)}

def execute_send_report_district_hourly_weather():
    try:
        # update_trigger_status("RUNNING", 'cyclone_data')
        subprocess.Popen(
            [
                os.environ["HOURLY_REPORT_PYTHON"],
                os.environ["HOURLY_REPORT_SCRIPT"]
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            start_new_session=True
        )
        # update_trigger_status("SUCCESS", 'cyclone_data')
        return {"msg": "Report sending..."}
    except Exception as e:
        # update_trigger_status("FAILED", 'cyclone_data')
        return {"error": str(e)}

def get_districts_zero_rain_day1(circle, conn):
    sql = f"""select distinct district from weatherdata.district_wise_7dayfc_severity
            where days = 'day1' and rain_precip = 0 and indus_circle = '{circle}' """
    df = pd.read_sql_query(sql, conn)
    district_list = df['district'].tolist()
    return district_list

def get_circlewise_user_mails():
    try:
        conn = get_db_conn()
        cursor = conn.cursor()

        circles = get_impacted_circles()
        impacted_circles = circles.get("name").split(",")

        if not impacted_circles:
            return []

        circles_tuple = tuple(impacted_circles)

        query = f"""
            SELECT distinct mail
            FROM weatherdata.master_users
            WHERE status = 'active'
            AND team = 'mlinfo'
            AND indus_circle IN {circles_tuple}
            ORDER BY mail
        """

        cursor.execute(query)
        rows = cursor.fetchall()

        cursor.close()
        conn.close()

        mails = [row[0] for row in rows if row[0]]
        return mails

    except Exception as e:
        print("Error in get_circlewise_user_mails:", e)
        return []

def send_mail_with_url_attachment(pdf_path):

    to_mails = get_circlewise_user_mails()

    if len(to_mails) == 0:
        return

    # Format current date & time
    now = datetime.now()
    formatted_datetime = f"{now.strftime('%d %b %Y')}, {now.strftime('%I:%M %p')}"

    try:
        sender = os.environ.get("EMAIL")
        password = os.environ.get("PASSWORD")
        # to_mails = json.loads(os.environ.get("TO_RECEIVERS"))
        cc_mails = json.loads(os.environ.get("CC_RECEIVERS"))

        yag = yagmail.SMTP(user=sender, password=password)

        yag.send(
            to=to_mails,
            # cc=cc_mails,
            subject=f"Cyclone Alert – {formatted_datetime}",
            contents=[
                "Dear Team,",
                "",
                "Please find attached cyclone alert report. <br> <br>",
                "",
                "Regards,",
                "ML InfoMap | Weather Services",
            ],
            attachments=[pdf_path],
        )

        res = {"status": "success", "message": "Mail sent successfully", "mails":to_mails}
        return res
    except Exception as e:
        res = {"status": "error", "message": e}
        return res

# Computation on observation data for data accuracy
def calc_accuracy(parameter, variable, obs, fst, rainfall_threshold=0):
    # Pair and filter valid data
    valid_data = [
        (o, f) for o, f in zip(obs, fst)
        if isinstance(o, (int, float)) and isinstance(f, (int, float))
    ]

    n = len(valid_data)

    if n == 0:
        return {
            "parameter": parameter,
            "variable": variable,
            "mae": 0,
            "rmse": 0,
            "mape": 0,
            "accuracy": 0
        }

    # MAE
    mae = sum(abs(o - f) for o, f in valid_data) / n

    # RMSE (same logic as your TS: divide by n+1)
    rmse = (sum((o - f) ** 2 for o, f in valid_data) / (n + 1)) ** 0.5

    # MAPE
    if rainfall_threshold > 0:
        mape_rows = [(o, f) for o, f in valid_data if o >= rainfall_threshold]
    else:
        mape_rows = [(o, f) for o, f in valid_data if o != 0]

    mape_count = len(mape_rows)

    if mape_count > 0:
        mape = sum(abs((o - f) / o) for o, f in mape_rows) / mape_count
    else:
        mape = 0

    accuracy = max(0, (1 - mape) * 100)

    return {
        "parameter": parameter,
        "variable": variable,
        "mae": mae,
        "rmse": rmse,
        "mape": mape,
        "accuracy": accuracy
    }

def process_weather_data(df,month_year):
    # Normalize column names (important)
    df.columns = [col.strip() for col in df.columns]

    # Convert numeric columns safely
    numeric_cols = [
            'circle', 'district', 'date',
            't_max_obs', 't_min_obs', 'h_avg_obs',
            'win_m_obs', 'rain_t_obs', 'v_avg_obs',
            't_max_fst', 't_min_fst', 'h_avg_fst',
            'win_m_fst', 'rain_t_fst', 'v_avg_fst'
    ]

    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Extract arrays
    records = {
        "tMaxObs": df['t_max_obs'].tolist(),
        "tMinObs": df['t_min_obs'].tolist(),
        "hAvgObs": df['h_avg_obs'].tolist(),
        "winMObs": df['win_m_obs'].tolist(),
        "rainTObs": df['rain_t_obs'].tolist(),
        "vAvgObs": df['v_avg_obs'].tolist(),

        "tMaxFst": df['t_max_fst'].tolist(),
        "tMinFst": df['t_min_fst'].tolist(),
        "hAvgFst": df['h_avg_fst'].tolist(),
        "winMFst": df['win_m_fst'].tolist(),
        "rainTFst": df['rain_t_fst'].tolist(),
        "vAvgFst": df['v_avg_fst'].tolist(),
    }

    # Calculate accuracy rows
    accuracy_rows = [
        calc_accuracy('Maximum Temperature', 'T_Max', records["tMaxObs"], records["tMaxFst"]),
        calc_accuracy('Minimum Temperature', 'T_Min', records["tMinObs"], records["tMinFst"]),
        calc_accuracy('Avg Relative Humidity', 'H_Avg', records["hAvgObs"], records["hAvgFst"]),
        calc_accuracy('Wind Speed', 'Win_M', records["winMObs"], records["winMFst"]),
        calc_accuracy('Rainfall', 'Rain_T', records["rainTObs"], records["rainTFst"], rainfall_threshold=1.0),
        calc_accuracy('Visibility/Fog', 'V_Avg', records["vAvgObs"], records["vAvgFst"]),
    ]

    # Average accuracy
    average_accuracy = sum(r["accuracy"] for r in accuracy_rows) / len(accuracy_rows)
    
    insert_accuracy_rows(accuracy_rows, month_year)

    return {
        "accuracy_rows": accuracy_rows,
        "average_accuracy": average_accuracy
    }



def insert_accuracy_rows(accuracy_rows, month_year):
    conn = get_db_conn()
    cursor = conn.cursor()

    insert_query = """
    INSERT INTO weatherdata.monthly_weather_accuracy (
        parameter, variable, month_year,
        mae, rmse, mape, accuracy_perc
    )
    VALUES (%s, %s, %s, %s, %s, %s, %s)
    """

    data = [
        (
            row['parameter'],
            row['variable'],
            month_year,
            round(row['mae'], 6),
            round(row['rmse'], 6),
            round(row['mape'], 6),
            round(row['accuracy'], 6)
        )
        for row in accuracy_rows
    ]

    cursor.executemany(insert_query, data)

    conn.commit()
    cursor.close()
    conn.close()
