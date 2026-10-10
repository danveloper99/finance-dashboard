"""
PDF 對帳單解鎖服務（Cloud Run）

GAS 的 callCloudRunToUnlock_ 會 POST：
  {"file_content": "<PDF base64>", "password": "<開啟密碼>"}
回傳：
  {"status": "success", "data": [[儲存格, ...], ...]}   # 每一列是一個 list
  {"status": "error",   "message": "..."}

隱私：只在記憶體裡處理，不存檔、不記錄 PDF 內容與密碼。
"""
import base64
import io
import os
import re

import pdfplumber
from flask import Flask, jsonify, request

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB

DATE_RE = re.compile(r"\d{3,4}\s*[/年]\s*\d{1,2}\s*[/月]\s*\d{1,2}")


def clean(v):
    return re.sub(r"\s+", " ", str(v or "")).strip()


def extract_rows(pdf_bytes, password):
    rows = []
    with pdfplumber.open(io.BytesIO(pdf_bytes), password=password or "") as pdf:
        # 先放幾行含日期的文字（GAS 會從前 20 列找對帳單日期，表格裡沒有日期時用得到）
        first_text = (pdf.pages[0].extract_text() or "") if pdf.pages else ""
        date_lines = [clean(l) for l in first_text.splitlines() if DATE_RE.search(l) and len(re.findall(r"\d+", l)) <= 4]
        rows.extend([[l] for l in date_lines[:3]])

        for page in pdf.pages:
            tables = page.extract_tables() or []
            if not tables:  # 沒有框線的表格改用文字對齊判斷
                tables = page.extract_tables({"vertical_strategy": "text", "horizontal_strategy": "text"}) or []
            for t in tables:
                for r in t:
                    cells = [clean(c) for c in r]
                    if any(cells):
                        rows.append(cells)
    return rows


@app.route("/", methods=["GET"])
def health():
    return "ok"


@app.route("/", methods=["POST"])
def unlock():
    try:
        body = request.get_json(force=True, silent=True) or {}
        b64 = body.get("file_content")
        if not b64:
            return jsonify(status="error", message="缺少 file_content")
        pdf_bytes = base64.b64decode(b64)
        rows = extract_rows(pdf_bytes, body.get("password", ""))
        return jsonify(status="success", data=rows)
    except Exception as e:  # 密碼錯誤、檔案損毀等
        detail = repr(e)
        msg = "密碼錯誤，無法開啟 PDF" if "Password" in detail else f"無法讀取 PDF：{detail}"
        return jsonify(status="error", message=msg)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
