import smtplib
import sys
import os
import tkinter.messagebox as messagebox
import Excel_Execute

from openpyxl import load_workbook
from email.message import EmailMessage
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding='utf-8')

# Load environment variables from secure.env
# These should include SMTP_SERVER, SMTP_PORT, EMAIL_USER, EMAIL_PASS
load_dotenv("secure.env")

# Manage sending emails
class EmailSender:
    def __init__(self):
        # Get SMTP server details and credentials from environment variables
        self.smtp_server = os.getenv("SMTP_SERVER")
        self.smtp_port = os.getenv("SMTP_PORT")
        self.email_user = os.getenv("EMAIL_USER")
        self.email_pass = os.getenv("EMAIL_PASS")
        self.Send_Errors = []

    # Method to send email
    def __Send_Emails__(self, Recip_Email, subject, html_content):
        # Create an EmailMessage object
        msg = EmailMessage()
        msg['Subject'] = subject # Email Subject
        msg['From'] = self.email_user # Sender Email
        msg['To'] = Recip_Email # Recipient Email


        # Add HTML content to the email
        # charset UTF-8 ensures Vietnamese or other Unicode characters are displayed correctly
        msg.add_alternative(html_content, subtype='html', charset='utf-8')

        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.ehlo()
                server.starttls()
                server.login(self.email_user, self.email_pass)
                server.send_message(msg)
                print(f"✅ Gửi mail thành công tới {msg['To']}")
        except Exception as e:
            print(f"❌ Lỗi khi gửi tới {msg['To']}: {e}")
            self.Send_Errors.append(f"{msg['To']} → {e}")

    # Method to send email to multiple recipients from CSV
    # CSV format: email, name (Optional)
    def __Send_Multiple_Emails__(self, Recipients_List, subject, html_template):
        # Reset Error
        self.Send_Errors = []

        for rec in Recipients_List:
            email = rec["Email"]
            name = rec["Full_Name"]
            
            html_content = html_template.replace("{{Full_Name}}", name)

            self.__Send_Emails__(email, subject, html_content)
        
        if self.Send_Errors:
            error = "\n".join(self.Send_Errors)
            messagebox.showerror("Error Sending!", error)
        
if __name__ == "__main__":
    email_sender = EmailSender()

    html_template = """
<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>{{subject}}</title>
</head>
<body style="margin:0; padding:0; background-color:#f4f4f6; font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif; color:#333333;">

  <!-- Wrapper -->
  <table width="100%" cellpadding="0" cellspacing="0" role="presentation">
    <tr>
      <td align="center" style="padding:20px 10px;">
        <!-- Container -->
        <table width="600" cellpadding="0" cellspacing="0" role="presentation" style="max-width:600px; width:100%; background:#ffffff; border-radius:6px; overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.08);">
          
          <!-- Header -->
          <tr>
            <td style="padding:20px 30px; background:#1a73e8; color:#ffffff; text-align:left;">
              <h1 style="margin:0; font-size:20px; font-weight:700;">Tên tổ chức / Công ty</h1>
              <p style="margin:4px 0 0 0; font-size:13px; opacity:0.95;">Tiêu đề phụ hoặc slogan</p>
            </td>
          </tr>

          <!-- Body -->
          <tr>
            <td style="padding:30px;">
              <!-- Greeting and intro -->
              <p style="margin:0 0 16px 0; font-size:16px;">
                Xin chào <strong>{{name}}</strong>,
              </p>

              <!-- Aligned text examples -->
              <p style="margin:0 0 12px 0; text-align:justify; line-height:1.5;">
                <strong>Đoạn văn căn đều (justify):</strong> Đây là ví dụ đoạn văn với <em>text-align: justify</em>. Bạn có thể dùng <strong>chữ đậm</strong>, <em>chữ nghiêng</em> hoặc <u>gạch chân</u> để nhấn mạnh.
              </p>

              <p style="margin:0 0 12px 0; text-align:center;">
                <strong style="font-size:16px;">Đoạn văn căn giữa (center)</strong>
              </p>

              <p style="margin:0 0 20px 0; text-align:right; color:#666;">
                <small>Đoạn văn căn phải (right-aligned)</small>
              </p>

              <!-- Bullet list -->
              <p style="margin:0 0 8px 0; font-weight:600;">Những điểm chính:</p>
              <ul style="margin:0 0 16px 20px; padding:0; color:#333;">
                <li style="margin:6px 0;">Điểm 1: <strong>Nội dung quan trọng</strong></li>
                <li style="margin:6px 0;">Điểm 2: <em>Thêm thông tin</em></li>
                <li style="margin:6px 0;">Điểm 3: Ghi chú <span style="text-decoration:underline;">quan trọng</span></li>
              </ul>

              <!-- Table -->
              <p style="margin:0 0 8px 0; font-weight:600;">Bảng thông tin:</p>
              <table width="100%" cellpadding="8" cellspacing="0" role="presentation" style="border-collapse:collapse; margin-bottom:18px;">
                <thead>
                  <tr>
                    <th align="left" style="background:#f1f5fb; border:1px solid #e6eefc; font-weight:700; font-size:13px;">Mục</th>
                    <th align="left" style="background:#f1f5fb; border:1px solid #e6eefc; font-weight:700; font-size:13px;">Giá trị</th>
                  </tr>
                </thead>
                <tbody>
                  <tr>
                    <td style="border:1px solid #eef6ff;">STT</td>
                    <td style="border:1px solid #eef6ff;">1</td>
                  </tr>
                  <tr>
                    <td style="border:1px solid #eef6ff;">Họ tên</td>
                    <td style="border:1px solid #eef6ff;">{{name}}</td>
                  </tr>
                  <tr>
                    <td style="border:1px solid #eef6ff;">Khóa</td>
                    <td style="border:1px solid #eef6ff;">2024</td>
                  </tr>
                </tbody>
              </table>

              <!-- Call to action -->
              <p style="margin:0 0 18px 0;">
                <a href="#" style="display:inline-block; text-decoration:none; padding:12px 18px; border-radius:6px; background:#28a745; color:#ffffff; font-weight:700;">Xác nhận ngay</a>
                <span style="margin-left:10px; color:#888; font-size:13px;">hoặc <a href="#" style="color:#1a73e8; text-decoration:underline;">xem thêm</a></span>
              </p>

              <!-- Small text / note -->
              <p style="margin:18px 0 0 0; font-size:13px; color:#666;">
                <strong>Lưu ý:</strong> Đây là email tự động. Vui lòng không trả lời email này. Nếu bạn cần hỗ trợ, liên hệ <a href="mailto:support@example.com" style="color:#1a73e8; text-decoration:underline;">support@example.com</a>.
              </p>
            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="padding:16px 30px; background:#f7f7fb; color:#666; font-size:12px; text-align:center;">
              <p style="margin:0;">© 2025 Công ty ABC - Địa chỉ: Số X, Đường Y, Thành phố Z</p>
              <p style="margin:6px 0 0 0;"><a href="#" style="color:#1a73e8; text-decoration:underline;">Chính sách bảo mật</a> • <a href="#" style="color:#1a73e8; text-decoration:underline;">Hủy đăng ký</a></p>
            </td>
          </tr>

        </table>
        <!-- /Container -->
      </td>
    </tr>
  </table>
  <!-- /Wrapper -->

</body>
</html>
"""
    xlsx_file = "File_Recipients.xlsx"
    xlsx_list = Excel_Execute.read_recipients(xlsx_file)

    email_sender.__Send_Multiple_Emails__(xlsx_list, "Test email from python", html_template)
    