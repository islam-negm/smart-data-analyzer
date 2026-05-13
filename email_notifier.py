"""
Email Notifier – Smart Data Analyzer
=====================================
يرسل التقارير النهائية تلقائياً عبر البريد الإلكتروني.

الإعداد:
    export EMAIL_SENDER="your@email.com"
    export EMAIL_PASSWORD="app_password"
    export EMAIL_RECIPIENT="recipient@email.com"
"""

import os, smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text      import MIMEText
from email.mime.base      import MIMEBase
from email               import encoders
from pathlib             import Path


def send_report(pdf_path: str, xlsx_path: str, pptx_path: str,
                subject: str = "تقرير تحليل البيانات – Smart Data Analyzer"):
    sender    = os.getenv("EMAIL_SENDER",    "sender@example.com")
    password  = os.getenv("EMAIL_PASSWORD",  "")
    recipient = os.getenv("EMAIL_RECIPIENT", "recipient@example.com")

    msg = MIMEMultipart()
    msg["From"]    = sender
    msg["To"]      = recipient
    msg["Subject"] = subject

    body = """
    السادة المحترمون،
    
    يُرفق طيّه تقرير التحليل الذكي للبيانات المُنشأ تلقائياً بواسطة نظام Smart Data Analyzer.
    
    المرفقات:
    • تقرير PDF احترافي
    • ملف Excel بالنتائج والتحليلات
    • عرض PowerPoint التقديمي
    
    مع التقدير،
    نظام Smart Data Analyzer
    """
    msg.attach(MIMEText(body, "plain", "utf-8"))

    for filepath in [pdf_path, xlsx_path, pptx_path]:
        if not filepath or not Path(filepath).exists():
            continue
        with open(filepath, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition",
                        f"attachment; filename={Path(filepath).name}")
        msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, password)
        server.sendmail(sender, recipient, msg.as_string())
    print(f"✅  تم إرسال التقرير إلى: {recipient}")


if __name__ == "__main__":
    import glob
    outputs = Path("outputs")
    pdfs  = sorted(glob.glob(str(outputs / "report_*.pdf")))
    xlsxs = sorted(glob.glob(str(outputs / "results_*.xlsx")))
    pptxs = sorted(glob.glob(str(outputs / "presentation_*.pptx")))
    if pdfs:
        send_report(pdfs[-1],
                    xlsxs[-1] if xlsxs else "",
                    pptxs[-1] if pptxs else "")
