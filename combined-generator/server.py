"""合并生成网关 - 同时调用 7 个独立生成服务并打包为 ZIP。

端口 8090，将前端表单按前缀拆分为 7 份子表单，分别转发至
localhost:8088~8095 的 /generate 端点，收集全部 PDF 后打包为 ZIP 返回。
"""

from __future__ import annotations

import importlib.util
import io
import json
import sys
import time
import urllib.parse
import urllib.request
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

# ---------------------------------------------------------------------------
# 加载共享字段映射
# ---------------------------------------------------------------------------
_SHARED_MAP_PATH = Path(r"c:\Users\Administrator\.trae-cn\work\6a30ebda6e33e99cfb93bf68\shared_field_map.py")
_spec = importlib.util.spec_from_file_location("_shared_field_map", _SHARED_MAP_PATH)
_shared_map_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_shared_map_mod)
SHARED_FIELD_MAP = _shared_map_mod.SHARED_FIELD_MAP


BASE_DIR = Path(__file__).resolve().parent
ROOT_DIR = BASE_DIR.parent
HOST = "127.0.0.1"
PORT = 8090

SERVERS: list[tuple[str, int, str]] = [
    ("doc8088", 8088, "01-委托代理协议.pdf"),
    ("doc8089", 8089, "02-俄英贸易合同.pdf"),
    ("doc8091", 8091, "03-附件2.pdf"),
    ("doc8092", 8092, "04-代理出口结算协议.pdf"),
    ("doc8093", 8093, "05-应付账款垫付协议.pdf"),
    ("doc8094", 8094, "06-垫付申请书与承诺函.pdf"),
    ("doc8095", 8095, "07-付款保函.pdf"),
]

# ---------------------------------------------------------------------------
# 模块加载工具
# ---------------------------------------------------------------------------

def load_module(name: str, path: Path):
    """用唯一 spec 名称加载一个 .py 文件为模块。"""
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    old_path = list(sys.path)
    sys.path.insert(0, str(path.parent))
    try:
        spec.loader.exec_module(module)
    finally:
        sys.path[:] = old_path
    return module


# 预加载 trade-contract-generator，并注册为 _trade_module，
# 使得之后加载的其他模块（appendix、export-settlement 等）可以复用同一份。
_TRADE_PATH = ROOT_DIR / "trade-contract-generator" / "server.py"
_trade_spec = importlib.util.spec_from_file_location("_shared_trade_module", _TRADE_PATH)
SHARED_TRADE = importlib.util.module_from_spec(_trade_spec)
old_path = list(sys.path)
sys.path.insert(0, str(_TRADE_PATH.parent))
_trade_spec.loader.exec_module(SHARED_TRADE)
sys.path[:] = old_path
sys.modules["_trade_module"] = SHARED_TRADE

# 加载全部 7 个服务模块
AGENCY = load_module("doc8088_mod", ROOT_DIR / "contract-generator" / "server.py")
# TRADE 已由 SHARED_TRADE 提供
TRADE = SHARED_TRADE
APPENDIX = load_module("doc8091_mod", ROOT_DIR / "appendix-generator" / "server.py")
EXPORT = load_module("doc8092_mod", ROOT_DIR / "export-settlement-generator" / "server.py")
PAYABLE_PREPAY = load_module("doc8093_mod", ROOT_DIR / "payable-prepay-generator" / "server.py")
PAYABLE_APP = load_module("doc8094_mod", ROOT_DIR / "payable-application-generator" / "server.py")
GUARANTEE = load_module("doc8095_mod", ROOT_DIR / "payment-guarantee-generator" / "server.py")

ALL_MODULES = [
    ("doc8088", AGENCY),
    ("doc8089", TRADE),
    ("doc8091", APPENDIX),
    ("doc8092", EXPORT),
    ("doc8093", PAYABLE_PREPAY),
    ("doc8094", PAYABLE_APP),
    ("doc8095", GUARANTEE),
]

# ---------------------------------------------------------------------------
# 合并 defaults / labels
# ---------------------------------------------------------------------------

def combined_defaults() -> dict:
    """返回全部 7 份文档以 docXXXX_ 为前缀的 DEFAULTS。"""
    result: dict = {}
    for prefix, mod in ALL_MODULES:
        for key, value in mod.DEFAULTS.items():
            result[f"{prefix}_{key}"] = value
    return result


def combined_labels() -> dict:
    """返回全部 7 份文档以 docXXXX_ 为前缀的 FIELD_LABELS。

    payment-guarantee-generator（8095）没有 FIELD_LABELS，跳过即可。
    """
    result: dict = {}
    for prefix, mod in ALL_MODULES:
        labels = getattr(mod, "FIELD_LABELS", None)
        if labels:
            for key, value in labels.items():
                result[f"{prefix}_{key}"] = value
    return result

# ---------------------------------------------------------------------------
# 表单拆分与转发
# ---------------------------------------------------------------------------

def prefixed(prefix: str, data: dict) -> dict:
    """从 data 中提取以 ``{prefix}_`` 开头的字段，去掉前缀后返回。"""
    marker = f"{prefix}_"
    out = {}
    for key, value in data.items():
        if key.startswith(marker):
            out[key[len(marker):]] = value
    return out


def apply_shared_field_sync(form_data: dict) -> dict:
    """应用共享字段同步：对每个共享字段组，找到第一个非空值并填充到所有文档。

    form_data 中的键格式为 ``docXXXX_field_name``。
    """
    for shared_name, mappings in SHARED_FIELD_MAP.items():
        # 找到第一个非空值
        master_value = None
        for prefix, field_name in mappings:
            full_key = f"{prefix}_{field_name}"
            val = form_data.get(full_key, "")
            if val and val.strip():
                master_value = val
                break
        # 如果找到了非空值，填充到所有其他文档的对应字段
        if master_value is not None:
            for prefix, field_name in mappings:
                full_key = f"{prefix}_{field_name}"
                form_data[full_key] = master_value
    return form_data


# 模块到 ZIP 文件名的映射
_MODULE_MAP: list[tuple[str, object, str]] = [
    ("doc8088", AGENCY, "01-委托代理协议.pdf"),
    ("doc8089", TRADE, "02-俄英贸易合同.pdf"),
    ("doc8091", APPENDIX, "03-附件2.pdf"),
    ("doc8092", EXPORT, "04-代理出口结算协议.pdf"),
    ("doc8093", PAYABLE_PREPAY, "05-应付账款垫付协议.pdf"),
    ("doc8094", PAYABLE_APP, "06-垫付申请书与承诺函.pdf"),
    ("doc8095", GUARANTEE, "07-付款保函.pdf"),
]


def _generate_one(prefix: str, module, zip_filename: str, doc_data: dict) -> tuple[str, bytes | None, str | None]:
    """直接调用模块的 generate_pdf，返回 (zip_filename, pdf_bytes, error_msg)。"""
    try:
        pdf_bytes, _filename = module.generate_pdf(doc_data)
        return zip_filename, pdf_bytes, None
    except Exception as exc:
        err_msg = f"{zip_filename} 生成失败：{exc}"
        return zip_filename, None, err_msg


def generate_zip(form_data: dict) -> tuple[bytes, list[str]]:
    """串行调用 7 个已加载模块的 generate_pdf，收集 PDF 后打包为 ZIP。

    使用串行执行，因为 Word COM 在同一进程内并行会导致冲突/死锁。
    每个模块维护自己的 Word COM 实例，串行执行更安全可靠。
    返回 (zip_bytes, errors)，errors 为字符串列表，记录失败信息。
    """
    buffer = io.BytesIO()
    errors: list[str] = []
    results: dict[str, bytes | None] = {}

    # 串行执行 7 个 PDF 生成任务
    for prefix, module, zip_filename in _MODULE_MAP:
        doc_data = prefixed(prefix, form_data)
        zf_name, pdf_bytes, err_msg = _generate_one(prefix, module, zip_filename, doc_data)
        if err_msg:
            errors.append(err_msg)
        results[zf_name] = pdf_bytes

    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for _prefix, _module, zip_filename in _MODULE_MAP:
            pdf_bytes = results.get(zip_filename)
            if pdf_bytes:
                zf.writestr(zip_filename, pdf_bytes)
            else:
                txt_name = zip_filename.replace(".pdf", "-ERROR.txt")
                err_text = next((e for e in errors if zip_filename in e), "未知错误")
                zf.writestr(txt_name, err_text.encode("utf-8"))

    return buffer.getvalue(), errors

# ---------------------------------------------------------------------------
# 文件加载辅助
# ---------------------------------------------------------------------------

def load_file(path: Path) -> bytes:
    try:
        return path.read_bytes()
    except FileNotFoundError:
        raise  # 交给上层处理

# ---------------------------------------------------------------------------
# HTTP Handler
# ---------------------------------------------------------------------------

class CombinedHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        return

    def send_bytes(
        self,
        status: int,
        content: bytes,
        content_type: str,
        headers: dict | None = None,
        no_cache: bool = False,
    ):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        if no_cache:
            self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
            self.send_header("Pragma", "no-cache")
            self.send_header("Expires", "0")
        if headers:
            for key, value in headers.items():
                self.send_header(key, value)
        self.end_headers()
        self.wfile.write(content)

    def do_GET(self):
        route = urllib.parse.urlparse(self.path).path

        try:
            if route == "/":
                self.send_bytes(
                    200, load_file(BASE_DIR / "index.html"),
                    "text/html; charset=utf-8",
                    no_cache=True,
                )
            elif route == "/style.css":
                self.send_bytes(
                    200, load_file(BASE_DIR / "style.css"),
                    "text/css; charset=utf-8",
                    no_cache=True,
                )
            elif route == "/wizard.css":
                self.send_bytes(
                    200, load_file(BASE_DIR / "wizard.css"),
                    "text/css; charset=utf-8",
                    no_cache=True,
                )
            elif route == "/wizard.js":
                self.send_bytes(
                    200, load_file(BASE_DIR / "wizard.js"),
                    "application/javascript; charset=utf-8",
                    no_cache=True,
                )
            elif route == "/app.js":
                self.send_bytes(
                    200, load_file(BASE_DIR / "app.js"),
                    "application/javascript; charset=utf-8",
                    no_cache=True,
                )
            elif route == "/defaults":
                payload = json.dumps(
                    {"defaults": combined_defaults(), "labels": combined_labels()},
                    ensure_ascii=False,
                ).encode("utf-8")
                self.send_bytes(200, payload, "application/json; charset=utf-8", no_cache=True)
            else:
                self.send_bytes(404, b"Not Found", "text/plain")
        except FileNotFoundError:
            self.send_bytes(404, b"File not found", "text/plain")

    def do_POST(self):
        if urllib.parse.urlparse(self.path).path != "/generate":
            self.send_bytes(404, b"Not found", "text/plain")
            return

        try:
            length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(length).decode("utf-8")
            form_data: dict[str, str] = {
                k: v[0] for k, v in urllib.parse.parse_qs(
                    raw, keep_blank_values=True
                ).items()
            }

            # 应用共享字段同步：将公共信息填充到所有相关文档
            form_data = apply_shared_field_sync(form_data)

            zip_bytes, errors = generate_zip(form_data)

            # 用一个友好的文件名
            safe_no = time.strftime("%Y%m%d_%H%M%S")
            zip_filename = f"全部七份文档_{safe_no}.zip"

            quoted = urllib.parse.quote(zip_filename)
            self.send_bytes(
                200,
                zip_bytes,
                "application/zip",
                {"Content-Disposition": f"attachment; filename*=UTF-8''{quoted}"},
            )
        except Exception as exc:
            import traceback
            tb = traceback.format_exc()
            msg = f"合并生成失败：{exc}\n{tb}"
            self.send_bytes(500, msg.encode("utf-8"), "text/plain; charset=utf-8")


# ---------------------------------------------------------------------------
# 自检
# ---------------------------------------------------------------------------

def self_test():
    data: dict = {}
    for prefix, mod in ALL_MODULES:
        for key, value in mod.DEFAULTS.items():
            data[f"{prefix}_{key}"] = value

    zip_bytes, errors = generate_zip(data)
    out = BASE_DIR / f"全部七份文档_自检_{time.strftime('%Y%m%d_%H%M%S')}.zip"
    out.write_bytes(zip_bytes)

    lines = [f"自检完成：{out}（{len(zip_bytes)} 字节）"]
    if errors:
        lines.append("生成过程中出现以下错误：")
        lines.extend(f"  - {e}" for e in errors)
    else:
        lines.append("全部 7 份文档生成成功。")
    print("\n".join(lines))


def _warmup_word():
    """预热 Word COM：生成一份最小 PDF，让 Word 实例提前启动。
    _get_word_app 会自动打开空白文档保持 Word 进程活跃。"""
    try:
        print("正在预热 Word COM（首次启动约需 30-60 秒，请稍候）...")
        import time
        start = time.time()
        # 预热 contract-generator (8088) 和 trade-contract-generator (8089)
        # 因为这两个模块有独立的 Word COM 实例
        sample = {k: v for k, v in AGENCY.DEFAULTS.items()}
        sample["sign_date"] = "2026-01-01"
        AGENCY.generate_pdf(sample)

        trade_sample = {k: v for k, v in TRADE.DEFAULTS.items()}
        trade_sample["trade_contract_date_en"] = "January 1, 2026"
        trade_sample["trade_contract_date_ru"] = "1 января 2026 г."
        TRADE.generate_pdf(trade_sample)

        print(f"Word COM 预热完成（耗时 {time.time()-start:.1f} 秒），后续生成将大幅加速。")
    except Exception as exc:
        print(f"Word COM 预热失败（不影响服务启动）：{exc}")


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        self_test()
    else:
        _warmup_word()
        print(f"合并生成网关已启动：http://{HOST}:{PORT}")
        print("保持此窗口运行，在浏览器打开上面的地址。按 Ctrl+C 停止。")
        ThreadingHTTPServer((HOST, PORT), CombinedHandler).serve_forever()