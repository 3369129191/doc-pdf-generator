import importlib.util
import io
import json
import re
import sys
import time
import urllib.parse
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
ROOT_DIR = BASE_DIR.parent
HOST = "127.0.0.1"
PORT = 8090


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    old_path = list(sys.path)
    sys.path.insert(0, str(path.parent))
    try:
        spec.loader.exec_module(module)
    finally:
        sys.path[:] = old_path
    return module


AGENCY = load_module("agency_contract_server", ROOT_DIR / "contract-generator" / "server.py")
TRADE = load_module("trade_contract_server", ROOT_DIR / "trade-contract-generator" / "server.py")
APPENDIX = load_module("appendix2_contract_server", ROOT_DIR / "appendix-generator" / "server.py")


def prefixed(prefix: str, data: dict) -> dict:
    marker = f"{prefix}__"
    return {k[len(marker):]: v for k, v in data.items() if k.startswith(marker)}


def combined_defaults() -> dict:
    return {
        "agency": AGENCY.DEFAULTS,
        "trade": TRADE.DEFAULTS,
        "appendix": APPENDIX.DEFAULTS,
        "agency_labels": AGENCY.FIELD_LABELS,
        "trade_labels": TRADE.FIELD_LABELS,
        "appendix_labels": APPENDIX.FIELD_LABELS,
    }


def make_zip(data: dict) -> tuple[bytes, str]:
    agency_data = prefixed("agency", data)
    trade_data = prefixed("trade", data)
    appendix_data = prefixed("appendix", data)

    agency_pdf, agency_name = AGENCY.generate_pdf(agency_data)
    trade_pdf, trade_name = TRADE.generate_pdf(trade_data)
    appendix_pdf, appendix_name = APPENDIX.generate_pdf(appendix_data)

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(agency_name, agency_pdf)
        zf.writestr(trade_name, trade_pdf)
        zf.writestr(appendix_name, appendix_pdf)

    safe_no = (
        trade_data.get("contract_no")
        or appendix_data.get("contract_no")
        or agency_data.get("trade_contract_no")
        or agency_data.get("agreement_no")
        or time.strftime("%Y%m%d")
    )
    safe_no = re.sub(r"[^\w\-一-龥]+", "_", safe_no)
    return buffer.getvalue(), f"三份合同PDF_{safe_no}.zip"


def load_file(path: Path) -> bytes:
    return path.read_bytes()


class CombinedHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        return

    def send_bytes(self, status: int, content: bytes, content_type: str, headers: dict | None = None):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        if headers:
            for key, value in headers.items():
                self.send_header(key, value)
        self.end_headers()
        self.wfile.write(content)

    def do_GET(self):
        route = urllib.parse.urlparse(self.path).path
        if route == "/":
            self.send_bytes(200, load_file(BASE_DIR / "index.html"), "text/html; charset=utf-8")
        elif route == "/style.css":
            self.send_bytes(200, load_file(BASE_DIR / "style.css"), "text/css; charset=utf-8")
        elif route == "/app.js":
            self.send_bytes(200, load_file(BASE_DIR / "app.js"), "application/javascript; charset=utf-8")
        elif route == "/defaults":
            payload = json.dumps(combined_defaults(), ensure_ascii=False).encode("utf-8")
            self.send_bytes(200, payload, "application/json; charset=utf-8")
        else:
            self.send_bytes(404, "未找到页面".encode("utf-8"), "text/plain; charset=utf-8")

    def do_POST(self):
        self.send_bytes(503, "合并生成页已临时暂停，请先使用 8088、8089、8091 三个单独页面。".encode("utf-8"), "text/plain; charset=utf-8")


def self_test():
    data = {}
    for key, value in AGENCY.DEFAULTS.items():
        data[f"agency__{key}"] = value
    for key, value in TRADE.DEFAULTS.items():
        data[f"trade__{key}"] = value
    for key, value in APPENDIX.DEFAULTS.items():
        data[f"appendix__{key}"] = value
    zip_bytes, filename = make_zip(data)
    out = BASE_DIR / filename
    out.write_bytes(zip_bytes)
    print(f"自检完成：{out}")


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        self_test()
    else:
        print(f"合并合同生成网站已启动：http://{HOST}:{PORT}")
        ThreadingHTTPServer((HOST, PORT), CombinedHandler).serve_forever()
