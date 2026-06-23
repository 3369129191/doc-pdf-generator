"""付款保函 Demo PDF 生成服务（端口 8095）。"""

from __future__ import annotations

import copy
import importlib.util
import json
import re
import tempfile
import urllib.parse
import xml.etree.ElementTree as ET
try:
    from lxml import etree as LXML_ET
except ImportError:
    LXML_ET = None
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "templates" / "payment-guarantee-template.docx"
HOST = "127.0.0.1"
PORT = 8095

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

TRADE_SERVER_PY = BASE_DIR.parent / "trade-contract-generator" / "server.py"
spec = importlib.util.spec_from_file_location("_trade_module", TRADE_SERVER_PY)
TRADE = importlib.util.module_from_spec(spec)
assert spec.loader is not None
spec.loader.exec_module(TRADE)

DEFAULTS = {
    "guarantee_no": "",
    "issue_date": "",
    "beneficiary_name": "SMARTCHAIN NEXUS LIMITED",
    "beneficiary_address": "Room A27, 24/F, Hyde House, Prince Industrial Building, 706 Prince Edward Road East, San Po Kong, Kowloon, Hong Kong",
    "notice_bank": "VTB Moscow",
    "collecting_agent_account": "40807810104800000155",
    "collecting_agent_currency": "RUB",
    "applicant_name": "LLC ZAM",
    "applicant_address": "127486, Moscow, inner territory of the Western Degunino Municipal District, Deguninskaya Street, Building 9, Block 1, Room 138A",
    "guarantor_branch_full_address": "",
    "guarantor_swift": "",
    "finance_agreement_no": "",
    "finance_agreement_name": "Account Payable Financing Agreement",
    "base_amount": "145583.44",
    "guarantee_percent": "110",
    "guarantee_amount": "160141.78",
    "presentation_bank": "Bank of Asian-Pacific Bank",
    "presentation_swift": "ASANRU8X",
    "presentation_branch": "",
    "presentation_address": "",
    "presentation_department": "Guarantee Department",
    "presentation_tel": "",
    "presentation_fax": "",
    "presentation_swift_line": "",
    "account_payable_no": "",
    "expiry_days": "100",
    "expiry_place": "Guarantor’s counter",
    "claim_bank_address": "",
    "sign_bank_name": "",
    "sign_branch_name": "",
    "signatory_names_titles": "",
}

for prefix, uri in {
    "w": W_NS,
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}.items():
    try:
        ET.register_namespace(prefix, uri)
    except ValueError:
        pass


def _para_text(paragraph: ET.Element) -> str:
    return "".join((t.text or "") for t in paragraph.iter(W + "t"))


def _new_run(text: str, template_run: ET.Element | None = None, *, underline: bool = False) -> ET.Element:
    run = (LXML_ET.Element if LXML_ET else ET.Element)(W + "r")
    rpr_template = None
    if template_run is not None:
        rpr_template = template_run.find(W + "rPr")
    if rpr_template is not None:
        run.append(copy.deepcopy(rpr_template))
    else:
        (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "rPr")
    if underline:
        rpr = run.find(W + "rPr")
        if rpr is None:
            rpr = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "rPr")
        u = rpr.find(W + "u")
        if u is None:
            u = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(rpr, W + "u")
        u.set(W + "val", "single")
    t = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "t")
    t.text = text
    if text and (text[0].isspace() or text[-1].isspace()):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return run


def replace_text_in_paragraph(paragraph: ET.Element, old: str, new: str, *, underline: bool = False) -> bool:
    if not old or new is None:
        return False
    text = _para_text(paragraph)
    pos = text.find(old)
    if pos < 0:
        return False
    end = pos + len(old)
    runs = list(paragraph.iter(W + "r"))
    cur = 0
    touched: list[tuple[ET.Element, int, int, str]] = []
    for run in runs:
        run_text = "".join((t.text or "") for t in run.iter(W + "t"))
        run_start, run_end = cur, cur + len(run_text)
        cur = run_end
        if run_end <= pos or run_start >= end:
            continue
        touched.append((run, run_start, run_end, run_text))
    if not touched:
        return False

    first_run, first_start, _, first_text = touched[0]
    last_run, _, last_end, last_text = touched[-1]
    first_keep = pos - first_start
    last_keep = end - (last_end - len(last_text))

    first_nodes = list(first_run.iter(W + "t"))
    if first_nodes:
        first_nodes[0].text = first_text[:first_keep]
        for node in first_nodes[1:]:
            node.text = ""

    for run, *_ in touched[1:-1]:
        for node in run.iter(W + "t"):
            node.text = ""

    if last_run is not first_run:
        last_nodes = list(last_run.iter(W + "t"))
        if last_nodes:
            last_nodes[0].text = last_text[last_keep:]
            for node in last_nodes[1:]:
                node.text = ""
    else:
        suffix = first_text[last_keep:]
        if suffix:
            suffix_run = copy.deepcopy(first_run)
            nodes = list(suffix_run.iter(W + "t"))
            if nodes:
                nodes[0].text = suffix
                for node in nodes[1:]:
                    node.text = ""
            children = list(paragraph)
            paragraph.insert(children.index(first_run) + 1, suffix_run)

    children = list(paragraph)
    paragraph.insert(children.index(first_run) + 1, _new_run(new, first_run, underline=underline))
    return True


def replace_all(paragraphs: list[ET.Element], old: str, new: str, *, underline: bool = False) -> int:
    count = 0
    if not old or new is None or new == "" or old == new:
        return count
    for p in paragraphs:
        while old in _para_text(p):
            if not replace_text_in_paragraph(p, old, new, underline=underline):
                break
            count += 1
    return count


def append_to_label(paragraph: ET.Element, label: str, value: str, *, underline: bool = True) -> bool:
    if not value or label not in _para_text(paragraph):
        return False
    runs = list(paragraph.iter(W + "r"))
    template = runs[-1] if runs else None
    paragraph.append(_new_run(" " + value, template, underline=underline))
    return True


def rewrite_paragraph(paragraph: ET.Element, value: str, *, underline: bool = False) -> None:
    ppr = paragraph.find(W + "pPr")
    rpr_template = None
    first_run = next(paragraph.iter(W + "r"), None)
    if first_run is not None:
        rpr_template = first_run.find(W + "rPr")
    for child in list(paragraph):
        if child.tag != W + "pPr":
            paragraph.remove(child)
    run = _new_run(value, None, underline=underline)
    if rpr_template is not None:
        current = run.find(W + "rPr")
        if current is not None:
            run.remove(current)
        run.insert(0, copy.deepcopy(rpr_template))
    if ppr is None and paragraph.find(W + "pPr") is None:
        pass
    paragraph.append(run)


def fill_document(root: ET.Element, values: dict[str, str]) -> None:
    paragraphs = list(root.iter(W + "p"))

    beneficiary = f"[{values['beneficiary_name']} and address: {values['beneficiary_address']}]"
    applicant = f"The Buyer: [{values['applicant_name']} and Address: {values['applicant_address']}]"
    finance_no = values.get("finance_agreement_no", "").strip() or "***"
    finance_name = values.get("finance_agreement_name", "").strip() or "Account Payable Financing Agreement"
    account_payable_no = values.get("account_payable_no", "").strip() or finance_no
    guarantor_bank = values.get("presentation_bank", "").strip() or "Bank of Asian-Pacific Bank"
    guarantor_swift = values.get("presentation_swift", "").strip() or "ASANRU8X"
    sign_bank = values.get("sign_bank_name", "").strip() or guarantor_bank.upper()
    sign_branch = values.get("sign_branch_name", "").strip() or values.get("presentation_branch", "").strip()

    for p in paragraphs:
        text = _para_text(p)
        if text.startswith("1.Guarantee No.:"):
            append_to_label(p, "1.Guarantee No.:", values.get("guarantee_no", ""))
        elif text.startswith("2.Date of Issue:"):
            append_to_label(p, "2.Date of Issue:", values.get("issue_date", ""))
        elif text.startswith("[SMARTCHAIN"):
            rewrite_paragraph(p, beneficiary, underline=False)
        elif text.startswith("The Buyer:"):
            rewrite_paragraph(p, applicant, underline=False)
        elif text == "Branch Name and Full Address]":
            rewrite_paragraph(p, values.get("guarantor_branch_full_address", ""), underline=True)
        elif text.startswith("(Swift Code:"):
            swift = values.get("guarantor_swift", "").strip()
            if swift:
                rewrite_paragraph(p, f"(Swift Code: {swift})", underline=False)
        elif text.startswith("Bank of Asian-Pacific Bank SWIFT::"):
            branch = values.get("presentation_branch", "").strip()
            rewrite_paragraph(p, f"{guarantor_bank} SWIFT:{guarantor_swift}, {branch}", underline=False)
        elif text.startswith("Address:"):
            rewrite_paragraph(p, f"Address: {values.get('presentation_address', '').strip()}", underline=False)
        elif text.startswith("Attention:"):
            rewrite_paragraph(p, f"Attention: {values.get('presentation_department', '').strip()}", underline=False)
        elif text.startswith("Tel:"):
            rewrite_paragraph(p, f"Tel: {values.get('presentation_tel', '').strip()}", underline=False)
        elif text.startswith("Fax:"):
            rewrite_paragraph(p, f"Fax: {values.get('presentation_fax', '').strip()}", underline=False)
        elif text.startswith("SWIFT: XXX"):
            swift_line = values.get("presentation_swift_line", "").strip() or guarantor_swift
            rewrite_paragraph(p, f"SWIFT: {swift_line}", underline=False)
        elif text == "BANK OF XXX":
            rewrite_paragraph(p, sign_bank, underline=False)
        elif text == "[Branch Name]":
            rewrite_paragraph(p, sign_branch, underline=False)
        elif text == "(Name(s) and Title(s))" and values.get("signatory_names_titles", "").strip():
            rewrite_paragraph(p, f"({values['signatory_names_titles'].strip()})", underline=False)

    replace_all(paragraphs, "VTB Moscow", values.get("notice_bank", ""))
    replace_all(paragraphs, "40807810104800000155", values.get("collecting_agent_account", ""))
    replace_all(paragraphs, "（RUB）", f"（{values.get('collecting_agent_currency', '').strip()}）")

    replace_all(paragraphs, "*** AP Finance Agreement", f"{finance_no} {finance_name}")
    replace_all(paragraphs, "*** Account Payable Financing Agreement", f"{finance_no} {finance_name}")
    replace_all(paragraphs, "No.** Account Payable", f"No.{account_payable_no} Account Payable")

    replace_all(paragraphs, "110% of", f"{values.get('guarantee_percent', '').strip() or '110'}% of")
    replace_all(paragraphs, "RMB 145583.44", f"RMB {values.get('base_amount', '').strip()}")
    replace_all(paragraphs, "RMB 160141.78", f"RMB {values.get('guarantee_amount', '').strip()}")

    replace_all(paragraphs, "100 days", f"{values.get('expiry_days', '').strip() or '100'} days")
    if values.get("expiry_place", "").strip():
        replace_all(paragraphs, "Guarantor’s counter", values["expiry_place"].strip())
    claim_addr = values.get("claim_bank_address", "").strip() or values.get("presentation_address", "").strip()
    if claim_addr:
        replace_all(paragraphs, "[bank’s address]", claim_addr)
    replace_all(paragraphs, "Bank of Asian-Pacific Bank", guarantor_bank)
    replace_all(paragraphs, "ASANRU8X", guarantor_swift)


def process_word_xml(content: bytes, values: dict[str, str]) -> bytes:
    import sys as _sys
    _lib_dir = Path(__file__).resolve().parent.parent
    if str(_lib_dir) not in _sys.path:
        _sys.path.insert(0, str(_lib_dir))
    from docx_processor import process_document_xml

    # 通用后处理（接受修订、删除批注、移除高亮/底纹、统一字体、压缩间距等）
    content = process_document_xml(content)

    # 服务特有的后处理：段落级精确填写
    try:
        from lxml import etree as lxml_etree
        parser = lxml_etree.XMLParser(remove_blank_text=False)
        root = lxml_etree.fromstring(content, parser)
    except Exception:
        return content
    fill_document(root, values)
    if LXML_ET:
        return LXML_ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def process_docx(input_path: Path, output_path: Path, data: dict[str, str]) -> None:
    values = {**DEFAULTS, **{k: v.strip() for k, v in data.items() if isinstance(v, str)}}
    with zipfile.ZipFile(input_path, "r") as zin, zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            if item.filename == "word/document.xml":
                content = process_word_xml(content, values)
            zout.writestr(item, content)


def generate_pdf(data: dict[str, str]) -> tuple[bytes, str]:
    with tempfile.TemporaryDirectory(prefix="payment_guarantee_") as tmp:
        tmp_dir = Path(tmp)
        docx_path = tmp_dir / "payment_guarantee.docx"
        process_docx(TEMPLATE_PATH, docx_path, data)
        pdf_path = TRADE.convert_to_pdf(docx_path, tmp_dir)
        pdf_bytes = pdf_path.read_bytes()
    safe_no = re.sub(r"[^\w\-一-龥]+", "_", data.get("guarantee_no") or "PaymentGuarantee")
    return pdf_bytes, f"付款保函_{safe_no}.pdf"


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        return

    def send_bytes(self, status: int, content: bytes, content_type: str, headers: dict | None = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        if headers:
            for k, v in headers.items():
                self.send_header(k, v)
        self.end_headers()
        self.wfile.write(content)

    def do_GET(self):
        route = urllib.parse.urlparse(self.path).path
        if route == "/":
            self.send_bytes(200, (BASE_DIR / "index.html").read_bytes(), "text/html; charset=utf-8")
        elif route == "/style.css":
            self.send_bytes(200, (BASE_DIR / "style.css").read_bytes(), "text/css; charset=utf-8")
        elif route == "/wizard.css":
            self.send_bytes(200, (BASE_DIR / "wizard.css").read_bytes(), "text/css; charset=utf-8")
        elif route == "/app.js":
            self.send_bytes(200, (BASE_DIR / "app.js").read_bytes(), "application/javascript; charset=utf-8")
        elif route == "/wizard.js":
            self.send_bytes(200, (BASE_DIR / "wizard.js").read_bytes(), "application/javascript; charset=utf-8")
        elif route == "/defaults":
            self.send_bytes(200, json.dumps({"defaults": DEFAULTS}, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")
        else:
            self.send_bytes(404, "未找到页面".encode("utf-8"), "text/plain; charset=utf-8")

    def do_POST(self):
        if urllib.parse.urlparse(self.path).path != "/generate":
            self.send_bytes(404, b"Not found", "text/plain")
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(length).decode("utf-8")
            data = {k: v[0] for k, v in urllib.parse.parse_qs(raw, keep_blank_values=True).items()}
            pdf, filename = generate_pdf(data)
            quoted = urllib.parse.quote(filename)
            self.send_bytes(200, pdf, "application/pdf", {"Content-Disposition": f"attachment; filename*=UTF-8''{quoted}"})
        except Exception as exc:
            import traceback
            self.send_bytes(500, f"生成失败：{exc}\n{traceback.format_exc()}".encode("utf-8"), "text/plain; charset=utf-8")


if __name__ == "__main__":
    print(f"付款保函生成网站已启动：http://{HOST}:{PORT}")
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
