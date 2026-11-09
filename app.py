from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, unquote
from datetime import datetime
import hashlib
import hmac
import json
import os
import tempfile

from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent
# Render: set DATA_DIR=/var/data and attach a Persistent Disk at /var/data.
# Local test: workbook is stored in ./data/Daily Production.xlsx.
DATA_DIR = Path(os.environ.get("DATA_DIR", str(BASE / "data"))).resolve()
XLSX = DATA_DIR / "Daily Production.xlsx"
PORT = int(os.environ.get("PORT", "8000"))
MAX_UPLOAD_BYTES = 20 * 1024 * 1024  # 20 MB
SYNC_TOKEN = os.environ.get("SYNC_TOKEN", "")

# Only normal website assets are served. Workbook, Python files, config files,
# BAT files, and hidden files are never made available through the web server.
PUBLIC_EXTENSIONS = {
    ".html", ".css", ".js", ".png", ".jpg", ".jpeg", ".gif", ".svg",
    ".ico", ".webp", ".woff", ".woff2", ".ttf", ".map"
}


def read_jobs():
    if not XLSX.exists():
        return {"error": "Chưa nhận được Daily Production.xlsx từ máy chủ xưởng."}

    wb = load_workbook(XLSX, data_only=True, read_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        jobs = []
        current = None

        for values in ws.iter_rows(min_row=3, max_col=6, values_only=True):
            no, process, person, _blank, performance, remain = values

            if no is not None and (process is not None or remain is not None):
                if current:
                    jobs.append(current)
                current = {
                    "no": no,
                    "process": str(process or "").strip(),
                    "person": str(person or "").strip(),
                    "performance_raw": str(performance or "").strip(),
                    "remain": [],
                }

            if current and remain is not None and str(remain).strip():
                current["remain"].append(str(remain).strip())

        if current:
            jobs.append(current)

        for job in jobs:
            raw = job.pop("performance_raw")
            import re
            match = re.search(r"(\d+(?:[.,]\d+)?)\s*%", raw)
            if match:
                pct = float(match.group(1).replace(",", "."))
            elif "hoàn thành" in raw.lower():
                pct = 100.0
            else:
                pct = 0.0

            pct = max(0, min(100, pct))
            job["progress"] = pct
            job["status"] = (
                "Hoàn thành" if pct >= 100 else
                ("Chưa bắt đầu" if pct <= 0 else "Đang thực hiện")
            )
            job["product"] = job["remain"][0] if job["remain"] else "—"
            job["remaining"] = job["remain"][1:] if len(job["remain"]) > 1 else []
            job["remain_text"] = " · ".join(job["remaining"][:3]) if job["remaining"] else (
                "Không có mục còn lại" if pct >= 100 else "Chưa ghi công việc còn lại"
            )
            job["person"] = job["person"] or "Chưa phân công"

        stat = XLSX.stat()
        stamp = datetime.fromtimestamp(stat.st_mtime).strftime("%d/%m/%Y %H:%M:%S")
        completed = sum(1 for job in jobs if job["progress"] >= 100)
        started = sum(1 for job in jobs if 0 < job["progress"] < 100)
        not_started = sum(1 for job in jobs if job["progress"] == 0)
        average = round(sum(job["progress"] for job in jobs) / len(jobs)) if jobs else 0

        return {
            "title": ws.cell(1, 1).value or "THEO DÕI TIẾN ĐỘ SẢN XUẤT",
            "jobs": jobs,
            "total": len(jobs),
            "completed": completed,
            "started": started,
            "not_started": not_started,
            "average": average,
            "updated": stamp,
            "source": XLSX.name,
        }
    finally:
        wb.close()


def json_response(handler, status, body):
    payload = json.dumps(body, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Cache-Control", "no-store, no-cache, must-revalidate")
    handler.send_header("Content-Length", str(len(payload)))
    handler.end_headers()
    handler.wfile.write(payload)


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE), **kwargs)

    def do_GET(self):
        route = urlparse(self.path).path

        if route == "/api/data":
            try:
                result = read_jobs()
                json_response(self, 200 if "error" not in result else 503, result)
            except Exception as exc:
                json_response(self, 500, {"error": "Không đọc được Excel: " + str(exc)})
            return

        if route == "/healthz":
            json_response(self, 200, {"ok": True, "excel_received": XLSX.exists()})
            return

        if route == "/":
            self.path = "/index.html"
            route = "/index.html"

        clean_path = unquote(route)
        parts = Path(clean_path.lstrip("/")).parts
        if (not parts or any(part.startswith(".") for part in parts)
                or Path(clean_path).suffix.lower() not in PUBLIC_EXTENSIONS):
            json_response(self, 404, {"error": "Không tìm thấy tài nguyên."})
            return

        super().do_GET()

    def do_POST(self):
        route = urlparse(self.path).path
        if route != "/api/upload":
            json_response(self, 404, {"error": "Không tìm thấy API."})
            return

        if not SYNC_TOKEN:
            json_response(self, 503, {"error": "Cloud chưa cấu hình SYNC_TOKEN."})
            return

        supplied = self.headers.get("Authorization", "")
        expected = "Bearer " + SYNC_TOKEN
        if not hmac.compare_digest(supplied, expected):
            json_response(self, 401, {"error": "Mã đồng bộ không hợp lệ."})
            return

        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            length = 0
        if length <= 0 or length > MAX_UPLOAD_BYTES:
            json_response(self, 413, {"error": "Dung lượng Excel không hợp lệ hoặc vượt quá 20 MB."})
            return

        DATA_DIR.mkdir(parents=True, exist_ok=True)
        temp_path = None
        try:
            with tempfile.NamedTemporaryFile(
                mode="wb", dir=str(DATA_DIR), prefix="excel_upload_", suffix=".xlsx", delete=False
            ) as temp_file:
                temp_path = Path(temp_file.name)
                remaining = length
                while remaining:
                    chunk = self.rfile.read(min(1024 * 1024, remaining))
                    if not chunk:
                        raise ValueError("Dữ liệu tải lên bị thiếu.")
                    temp_file.write(chunk)
                    remaining -= len(chunk)

            # Validate the workbook before replacing the current good copy.
            check_wb = load_workbook(temp_path, read_only=True, data_only=True)
            try:
                if not check_wb.sheetnames:
                    raise ValueError("Workbook không có sheet dữ liệu.")
                check_ws = check_wb[check_wb.sheetnames[0]]
                if check_ws.max_column < 6:
                    raise ValueError("Excel không đủ 6 cột theo cấu trúc Dashboard.")
            finally:
                check_wb.close()

            os.replace(str(temp_path), str(XLSX))
            temp_path = None
            json_response(self, 200, {
                "ok": True,
                "message": "Đã cập nhật Excel trên Cloud.",
                "updated": datetime.fromtimestamp(XLSX.stat().st_mtime).strftime("%d/%m/%Y %H:%M:%S"),
            })
        except Exception as exc:
            json_response(self, 400, {"error": "Không cập nhật được Excel: " + str(exc)})
        finally:
            if temp_path and temp_path.exists():
                try:
                    temp_path.unlink()
                except OSError:
                    pass

    def log_message(self, fmt, *args):
        print("[%s] %s" % (self.log_date_time_string(), fmt % args))


if __name__ == "__main__":
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    print("DASHBOARD SAN XUAT - CLOUD")
    print("Listening on 0.0.0.0:%s" % PORT)
    print("Excel storage: %s" % XLSX)
    print("Nhan Ctrl+C de dung.\n")
    ThreadingHTTPServer(("0.0.0.0", PORT), Handler).serve_forever()
