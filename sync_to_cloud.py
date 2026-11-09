"""Watch the factory Excel file and securely upload each changed version to Cloud."""
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import URLError, HTTPError
import hashlib
import json
import sys
import time

BASE = Path(__file__).resolve().parent
CONFIG = BASE / "sync_config.json"


def load_config():
    if not CONFIG.exists():
        raise FileNotFoundError(
            "Chua co sync_config.json. Hay copy sync_config.example.json thanh sync_config.json va dien thong tin."
        )
    with CONFIG.open("r", encoding="utf-8") as f:
        cfg = json.load(f)
    for key in ("dashboard_url", "sync_token", "excel_path"):
        if not cfg.get(key) or "YOUR_" in str(cfg[key]) or "PASTE_" in str(cfg[key]):
            raise ValueError("Hay dien truong '%s' trong sync_config.json." % key)
    cfg["check_interval_seconds"] = max(5, int(cfg.get("check_interval_seconds", 10)))
    cfg["stable_seconds"] = max(2, int(cfg.get("stable_seconds", 3)))
    cfg["dashboard_url"] = cfg["dashboard_url"].rstrip("/")
    cfg["excel_path"] = str(Path(cfg["excel_path"]).expanduser())
    return cfg


def sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def is_stable(path, stable_seconds):
    first = path.stat()
    time.sleep(stable_seconds)
    second = path.stat()
    return first.st_size == second.st_size and first.st_mtime_ns == second.st_mtime_ns


def upload(cfg, path):
    endpoint = cfg["dashboard_url"] + "/api/upload"
    with open(path, "rb") as f:
        data = f.read()
    request = Request(
        endpoint,
        data=data,
        headers={
            "Authorization": "Bearer " + cfg["sync_token"],
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        },
        method="POST",
    )
    with urlopen(request, timeout=60) as response:
        result = json.loads(response.read().decode("utf-8"))
        if response.status != 200 or not result.get("ok"):
            raise RuntimeError(result.get("error", "Cloud tu choi cap nhat Excel."))
        return result


def main():
    try:
        cfg = load_config()
    except Exception as exc:
        print("LOI CAU HINH: %s" % exc)
        return 2

    excel = Path(cfg["excel_path"])
    interval = cfg["check_interval_seconds"]
    stable = cfg["stable_seconds"]
    last_hash = None

    print("DONG BO EXCEL LEN CLOUD")
    print("Excel: %s" % excel)
    print("Dashboard: %s" % cfg["dashboard_url"])
    print("Kiem tra moi %s giay. Nhan Ctrl+C de dung.\n" % interval)

    while True:
        try:
            if not excel.is_file():
                print("[%s] Chua tim thay file Excel." % time.strftime("%H:%M:%S"))
            elif is_stable(excel, stable):
                current_hash = sha256_file(excel)
                if current_hash != last_hash:
                    result = upload(cfg, excel)
                    last_hash = current_hash
                    print("[%s] %s %s" % (
                        time.strftime("%H:%M:%S"),
                        result.get("message", "Da dong bo."),
                        result.get("updated", ""),
                    ))
        except (PermissionError, OSError) as exc:
            print("[%s] Chua doc duoc Excel (co the dang luu): %s" % (time.strftime("%H:%M:%S"), exc))
        except HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            print("[%s] Loi Cloud HTTP %s: %s" % (time.strftime("%H:%M:%S"), exc.code, detail))
        except URLError as exc:
            print("[%s] Khong ket noi duoc Cloud: %s" % (time.strftime("%H:%M:%S"), exc))
        except Exception as exc:
            print("[%s] Loi dong bo: %s" % (time.strftime("%H:%M:%S"), exc))

        time.sleep(interval)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nDa dung chuong trinh dong bo.")
