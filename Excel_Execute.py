import pandas as pd
import tkinter.messagebox as messagebox
import openpyxl
import re
from typing import List, Dict

from pyexpat.errors import messages

Required_columns = {'STT', 'Full_Name', 'Academic_Year', 'Email'}

Email_Regrex = re.compile(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$")

def validate_email(email: str) -> bool:
    return bool(Email_Regrex.fullmatch(email))

def read_recipients(path: str) -> List[Dict]:
    """
    Read recipients from an Excel file.
    Returns list of dicts with keys STT, Full_Name, Academic_Year, Email
    Raises ValueError if columns missing.
    """

    # ==========================
    # Đọc file Excel
    # ==========================
    try:
        df = pd.read_excel(path)
    except Exception as e:
        messagebox.showerror("File Error", f"Không đọc được file:\n{e}")
        return []

    # File trống
    if df.empty:
        messagebox.showerror("Empty File", "File không có dữ liệu.")
        return []

    # Chuẩn hóa tên cột
    df.columns = df.columns.str.strip()
    cols = set(df.columns)

    if not Required_columns.issubset(cols):
        missing = Required_columns - cols
        messagebox.showerror("Missing Columns", f"Các cột thiếu: {missing}")
        return []

    errors = []  # lưu lại danh sách lỗi
    emails_seen = set()  # kiểm tra trùng email
    result = []

    for idx, row in df.iterrows():   # cần iterrows để xử lý NaN
        stt = row.get("STT")
        name = row.get("Full_Name")
        A_Y = row.get("Academic_Year")
        mail = row.get("Email")

        # 1) Xử lý NaN
        if pd.isna(stt) or pd.isna(name) or pd.isna(A_Y) or pd.isna(mail):
            errors.append(f"Dòng {idx+2}: Thiếu dữ liệu (STT / Tên / Khóa / Email).")
            continue

        # 2) STT phải là số nguyên
        try:
            stt = int(float(stt))
        except:
            errors.append(f"Dòng {idx + 2}: STT không hợp lệ: {stt}")
            continue

        # 3) Chuẩn hóa chuỗi
        name = str(name).strip()
        A_Y = str(A_Y).strip()
        mail = str(mail).strip()

        # 4) Kiểm tra email hợp lệ
        if not validate_email(mail):
            errors.append(f"Dòng {idx + 2}: Email không hợp lệ: {mail}")
            continue

        # 5) Kiểm tra email trùng
        if mail in emails_seen:
            errors.append(f"Dòng {idx + 2}: Email bị trùng: {mail}")
            continue
        emails_seen.add(mail)

        # 6) Lưu dữ liệu sạch
        result.append({'STT':stt,
                       'Full_Name': name,
                       'Academic_Year': A_Y,
                       'Email': mail})

        # Nếu có lỗi → thông báo tổng hợp
        if errors:
            error_text = "\n".join(errors)
            messagebox.showerror("Data Validation Errors", error_text)
            return result

    return result





