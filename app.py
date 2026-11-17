import io
import re
import shutil
import subprocess
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import jdatetime
from flask import Flask, Response, render_template, request, send_file
from lxml import etree

BASE = Path(__file__).parent
TEMPLATE_DOCX = BASE / "assets" / "template.docx"
FONT_FILE = BASE / "assets" / "fonts" / "IRANSansWeb(FaNum).ttf"

app = Flask(__name__, static_folder="assets", static_url_path="/assets")
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024

def fa_digits(value):
    return value.translate(str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹"))

def today_jalali():
    now = datetime.now(ZoneInfo("Asia/Tehran"))
    return fa_digits(jdatetime.date.fromgregorian(date=now.date()).strftime("%Y/%m/%d"))

def fill_word(full_name, national_id, date_text, destination):
    # فقط متن placeholderهای فایل Word اصلی عوض می‌شود؛ استایل‌ها و پاراگراف‌ها دست‌نخورده‌اند.
    with zipfile.ZipFile(TEMPLATE_DOCX, "r") as source:
        parts = {name: source.read(name) for name in source.namelist()}
    xml = etree.fromstring(parts["word/document.xml"])
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    replacements = {
        "........": full_name,
        "...........": national_id,
        "......................": full_name,
        "تاریخ : ": "تاریخ : " + date_text,
    }
    for node in xml.xpath(".//w:t", namespaces=ns):
        if node.text in replacements:
            node.text = replacements[node.text]
            if node.text.startswith(" ") or node.text.endswith(" "):
                node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    parts["word/document.xml"] = etree.tostring(xml, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in parts.items():
            target.writestr(name, data)

def create_pdf(full_name, national_id):
    full_name = " ".join(full_name.strip().split())
    national_id = national_id.strip()
    if not full_name or len(full_name) > 80:
        raise ValueError("نام و نام خانوادگی را وارد کنید.")
    if not re.fullmatch(r"[0-9۰-۹]{10}", national_id):
        raise ValueError("شماره ملی باید دقیقاً ۱۰ رقم باشد.")
    national_id = fa_digits(national_id.translate(str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")))

    with tempfile.TemporaryDirectory(prefix="letter_") as temp:
        tmp = Path(temp)
        docx_path = tmp / "letter.docx"
        pdf_dir = tmp / "pdf"
        pdf_dir.mkdir()
        fill_word(full_name, national_id, today_jalali(), docx_path)
        office = shutil.which("libreoffice") or shutil.which("soffice") or "/usr/bin/libreoffice"
        result = subprocess.run(
            [office, "-env:UserInstallation=" + (tmp / "lo_profile").as_uri(),
             "--headless", "--convert-to", "pdf", "--outdir", str(pdf_dir), str(docx_path)],
            capture_output=True, text=True, timeout=90
        )
        pdf_path = pdf_dir / "letter.pdf"
        if result.returncode != 0 or not pdf_path.exists():
            raise RuntimeError("تبدیل نامه به PDF انجام نشد.")
        return pdf_path.read_bytes()

@app.get("/")
def home():
    return render_template("index.html")

@app.post("/generate")
def generate():
    try:
        pdf_bytes = create_pdf(request.form.get("full_name", ""), request.form.get("national_id", ""))
        return send_file(io.BytesIO(pdf_bytes), mimetype="application/pdf",
                         as_attachment=True, download_name="letter.pdf",
                         max_age=0)
    except ValueError as exc:
        return Response(str(exc), status=400, mimetype="text/plain; charset=utf-8")
    except Exception:
        app.logger.exception("PDF generation failed")
        return Response("در تولید فایل خطایی رخ داد. لطفاً بعداً دوباره تلاش کنید.",
                        status=500, mimetype="text/plain; charset=utf-8")

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
