import pandas as pd
import tkinter.messagebox as messagebox
import re
import os

from typing import List, Dict

# Navigate exactly where contains script
os.chdir(os.path.dirname(os.path.abspath(__file__)))

class RecipientReader:
    Required_columns = {'STT', 'Full_Name', 'Academic_Year', 'Email'}
    Email_Regrex = re.compile(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$")

    def __init__(self, file_path: str):
        self.file_path = file_path
        self.errors: List[str] = []
        self.emails_seen = set()

    # Validation Methods
    def validate_email(self,email: str) -> bool:
        return bool(self.Email_Regrex.fullmatch(email))

    def read_recipients(self) -> List[Dict]:
        """
        Read recipients from an Excel file.
        Returns list of dicts with keys STT, Full_Name, Academic_Year, Email
        Raises ValueError if columns missing.
        """

        # ==========================
        # Đọc file Excel
        # ==========================
        try:
            df = pd.read_excel(self.file_path)
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

        if not self.Required_columns.issubset(cols):
            missing = self.Required_columns - cols
            messagebox.showerror("Missing Columns", f"Các cột thiếu: {missing}")
            return []

        result = []

        for idx, row in df.iterrows():   # cần iterrows để xử lý NaN
            stt = row.get("STT")
            name = row.get("Full_Name")
            A_Y = row.get("Academic_Year")
            mail = row.get("Email")

            # 1) Xử lý NaN
            if pd.isna(stt) or pd.isna(name) or pd.isna(A_Y) or pd.isna(mail):
                self.errors.append(f"Dòng {idx+2}: Thiếu dữ liệu (STT / Tên / Khóa / Email).")
                continue

            # 2) STT phải là số nguyên
            try:
                stt = int(float(stt))
            except ValueError as VE:
                self.errors.append(f"Dòng {idx + 2}: STT không hợp lệ: {stt}")
                continue

            # 3) Chuẩn hóa chuỗi
            name = str(name).strip()
            A_Y = str(A_Y).strip()
            mail = str(mail).strip()

            # 4) Kiểm tra email hợp lệ
            if not self.validate_email(mail):
                self.errors.append(f"Dòng {idx + 2}: Email không hợp lệ: {mail}")
                continue

            # 5) Kiểm tra email trùng
            if mail in self.emails_seen:
                self.errors.append(f"Dòng {idx + 2}: Email bị trùng: {mail}")
                continue
            self.emails_seen.add(mail)

            # 6) Lưu dữ liệu sạch
            result.append({'STT':stt,
                           'Full_Name': name,
                           'Academic_Year': A_Y,
                           'Email': mail})

            # Nếu có lỗi → thông báo tổng hợp
        if self.errors:
            error_text = "\n".join(self.errors)
            messagebox.showerror("Data Validation Errors", error_text)
            return result

        return result
