import copy
import importlib.util
import io
import json
import re
import sys
import tempfile
import time
import urllib.parse
import zipfile
import xml.etree.ElementTree as ET
try:
    from lxml import etree as LXML_ET
except ImportError:
    LXML_ET = None
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
ROOT_DIR = BASE_DIR.parent
TEMPLATE_PATH = BASE_DIR / "templates" / "appendix2-template.docx"
HOST = "127.0.0.1"
PORT = 8091


def load_trade_module():
    path = ROOT_DIR / "trade-contract-generator" / "server.py"
    spec = importlib.util.spec_from_file_location("trade_contract_server_for_appendix", path)
    module = importlib.util.module_from_spec(spec)
    old_path = list(sys.path)
    sys.path.insert(0, str(path.parent))
    try:
        spec.loader.exec_module(module)
    finally:
        sys.path[:] = old_path
    return module


TRADE = load_trade_module()
W_NS = TRADE.W_NS


DEFAULTS = {
    **TRADE.DEFAULTS,
    "appendix_no": "2",
    "form_appendix_no": "2",
    "delivery_time_ru": "Поставка осуществляется в срок не позднее 45 (сорок пять) календарных дней с даты оплаты аванса согласно п.6.1. настоящего Приложения без учета национальных праздников. Срок поставки Товара продлевается на срок задержки исполнения Покупателем своих обязательств по оплате согласно п.6.1. настоящего Приложения.",
    "delivery_time_en": "Goods should be delivered up in time not later than 45 (forty-five) calendar days from the date of payment of the advance payment according to the paragraph 6.1. of this Appendix excluding national holidays. Delivery of Goods shall be extended for a period of delay of the Buyer's payment obligations under the paragraph 6.1 of the present Appendix.",
    "payment_advance_percent": "20%",
    "payment_advance_amount": "410 580",
    "payment_advance_amount_ru": "четыреста десять тысяч пятьсот восемьдесят",
    "payment_advance_amount_en": "four hundred and ten thousand five hundred eighty",
    "payment_balance_percent": "80%",
    "payment_balance_amount": "1 642 320",
    "payment_balance_amount_ru": "один миллион шестьсот сорок две тысячи триста двадцать",
    "payment_balance_amount_en": "one million six hundred forty-two thousand three hundred twenty",
    "payment_bank_days": "60",
    "payment_bank_days_ru": "шестьдесят",
    "payment_bank_days_en": "sixty",
    "delivery_notice_days": "10",
    "delivery_notice_days_ru": "десять",
    "delivery_notice_days_en": "ten",
}


FIELD_LABELS = {
    **TRADE.FIELD_LABELS,
    "payment_advance_percent": "附件2：预付款比例",
    "payment_advance_amount": "附件2：预付款金额（数字）",
    "payment_advance_amount_ru": "附件2：预付款金额大写（俄文）",
    "payment_advance_amount_en": "附件2：预付款金额大写（英文）",
    "payment_balance_percent": "附件2：尾款比例",
    "payment_balance_amount": "附件2：尾款金额（数字）",
    "payment_balance_amount_ru": "附件2：尾款金额大写（俄文）",
    "payment_balance_amount_en": "附件2：尾款金额大写（英文）",
    "payment_bank_days": "附件2：预付款支付期限（数字）",
    "payment_bank_days_ru": "附件2：预付款支付期限大写（俄文）",
    "payment_bank_days_en": "附件2：预付款支付期限大写（英文）",
    "delivery_notice_days": "附件2：发货前通知天数（数字）",
    "delivery_notice_days_ru": "附件2：发货前通知天数大写（俄文）",
    "delivery_notice_days_en": "附件2：发货前通知天数大写（英文）",
}


def build_values(data: dict) -> dict:
    values = {**DEFAULTS, **{k: v.strip() for k, v in data.items() if isinstance(v, str) and v.strip()}}
    return values


def goods_total_amount(values: dict) -> str:
    items = TRADE.parse_goods_items(values)
    if not items:
        return values.get("goods_total", "")
    amounts = []
    for item in items:
        raw = str(item.get("goods_total", "")).replace("\xa0", " ").replace(" ", "").replace(",", ".")
        try:
            amounts.append(float(raw))
        except ValueError:
            pass
    if not amounts:
        return items[0].get("goods_total", "")
    total = sum(amounts)
    text = f"{total:,.2f}".replace(",", " ").replace(".", ",")
    return text


def payment_terms_ru(values: dict) -> str:
    advance = values.get("payment_advance_amount") or ""
    balance = values.get("payment_balance_amount") or ""
    return (
        f"6.1 Аванс в размере {values.get('payment_advance_percent', '')} {advance} "
        f"({values.get('payment_advance_amount_ru', '')}) юаней КНР 00 фэней Покупатель оплачивает прямым банковским переводом "
        f"на счет Продавца в течение {values.get('payment_bank_days', '')} ({values.get('payment_bank_days_ru', '')}) банковских дней "
        "с даты подписания уполномоченными представителями обеих Сторон настоящего приложения на основании счета Продавца.\n"
        f"6.2 Сумму в размере {values.get('payment_balance_percent', '')} {balance} "
        f"({values.get('payment_balance_amount_ru', '')}) юаней КНР 00 фэней Покупатель оплачивает прямым банковским переводом "
        "на счет Продавца перед отгрузкой Товара.\n"
        "6.3 Датой исполнения Покупателем обязательств по оплате считается дата поступления денежных средств в полном объеме "
        "на расчетный счет Продавца, указанный в п. 15 Контракта."
    )


def payment_terms_en(values: dict) -> str:
    advance = values.get("payment_advance_amount") or ""
    balance = values.get("payment_balance_amount") or ""
    return (
        f"6.1 The advance payment in the amount of {values.get('payment_advance_percent', '')} {advance} "
        f"({values.get('payment_advance_amount_en', '')}) Chinese yuan 00 fen is paid by the Buyer by direct bank transfer "
        f"to the Seller's account within {values.get('payment_bank_days', '')} ({values.get('payment_bank_days_en', '')}) banking days "
        "from the date of signing by authorized representatives of both Parties of this Appendix on the basis of the Seller’s Invoice.\n"
        f"6.2 The amount of {values.get('payment_balance_percent', '')} {balance} "
        f"({values.get('payment_balance_amount_en', '')}) Chinese yuan 00 fen shall be paid by the Buyer by direct bank transfer "
        "to the Seller’s account before shipment of the Goods.\n"
        "6.3 The date of the Buyer's performance of its payment obligations under this Appendix shall be deemed the date of full receipt "
        "of the funds in the Seller's bank account specified in Clause 15 of the Contract."
    )


def delivery_time_ru(values: dict) -> str:
    base = values.get("delivery_time_ru") or values.get("form_delivery_time_ru") or ""
    notice = (
        f"Продавец извещает Покупателя о готовности Товара к отгрузке посредством факсимильной связи "
        f"{values.get('buyer_notice_tel', '')} или электронной почты {values.get('buyer_notice_email', '')}, "
        f"за {values.get('delivery_notice_days', '')} ({values.get('delivery_notice_days_ru', '')}) календарных дней до начала отгрузки."
    )
    return "\n".join([part for part in [base, notice] if part.strip()])


def delivery_time_en(values: dict) -> str:
    base = values.get("delivery_time_en") or values.get("form_delivery_time_en") or ""
    notice = (
        f"The Seller shall notify the Buyer of the readiness to supply the Goods through facsimile "
        f"{values.get('buyer_notice_tel', '')} or e-mail {values.get('buyer_notice_email', '')}, "
        f"{values.get('delivery_notice_days', '')} calendar days before the beginning of the shipment."
    )
    return "\n".join([part for part in [base, notice] if part.strip()])


def shipping_docs_ru() -> str:
    return (
        "- Упаковочный лист - оригинал (3 экз.);\n"
        "- Инвойс - оригинал (3 экз.);\n"
        "- Прайс-лист - оригинал (2 экз.);\n"
        "- Сертификат происхождения товара – оригинал (2 экз.);\n"
        "- Сертификат качества товара - оригинал (1экз.)."
    )


def shipping_docs_en() -> str:
    return (
        "- Packing list – 3 original copies;\n"
        "- Invoice – 3 original copies;\n"
        "- Price list – 2 original copies;\n"
        "- Certificate of Origin – 2 original copies;\n"
        "- Certificate of Conformity – 1 original copy."
    )


def build_replacements(values: dict) -> dict:
    replacements = TRADE.build_replacements(values)
    total_number = goods_total_amount(values) or values.get("goods_total", "")
    if total_number:
        replacements["2 052\xa0200,00"] = total_number
        replacements["2 052 200,00"] = total_number
    if values.get("total_price_ru"):
        replacements["2 052\xa0200,00 (два миллиона пятьдесят две тысячи двести) юаней КНР 00 фэней. "] = (
            f"{total_number} ({values['total_price_ru']}) юаней КНР 00 фэней. "
        )
    if values.get("total_price_en"):
        replacements["2 052\xa0200,00 (two millions fifty two thousand two hundred) CNY and 00 fens only."] = (
            f"{total_number} ({values['total_price_en']}) CNY and 00 fens only."
        )
    replacements["Приложение № 2"] = f"Приложение № {values.get('appendix_no', '2')}"
    replacements["Appendix № 2"] = f"Appendix № {values.get('appendix_no', '2')}"
    replacements["Appendix №2"] = f"Appendix №{values.get('appendix_no', '2')}"
    return replacements


def split_lines_to_cell(cell: ET.Element, value: str, keep_first: bool = False) -> None:
    paragraphs = cell.findall(f"{{{W_NS}}}p")
    if not paragraphs:
        paragraphs = [(LXML_ET.SubElement if LXML_ET else ET.SubElement)(cell, f"{{{W_NS}}}p")]
    first_index = 1 if keep_first and len(paragraphs) > 1 else 0
    for p in paragraphs[first_index:]:
        cell.remove(p)
    for idx, line in enumerate(str(value or "").splitlines() or [""]):
        if idx == 0 and first_index < len(paragraphs):
            paragraph = paragraphs[first_index]
        else:
            paragraph = copy.deepcopy(paragraphs[0])
            cell.append(paragraph)
        TRADE.set_paragraph_text(paragraph, line)


def find_total_row(rows: list[ET.Element]) -> int:
    for i, row in enumerate(rows):
        cells = row.findall(f"{{{W_NS}}}tc")
        text = "".join((t.text or "") for t in row.iter(f"{{{W_NS}}}t"))
        if len(cells) == 8 and "Итого" in text and "Total" in text:
            return i
    return -1


def apply_appendix_structured_fields(root: ET.Element, values: dict) -> None:
    tables = list(root.iter(f"{{{W_NS}}}tbl"))
    if not tables:
        return
    table = tables[0]
    rows = table.findall(f"{{{W_NS}}}tr")
    total_index = find_total_row(rows)
    if total_index > 4:
        goods_items = TRADE.parse_goods_items(values)
        if goods_items:
            base_row = rows[4]
            for row in reversed(rows[5:total_index]):
                table.remove(row)
            TRADE.fill_goods_row(base_row, goods_items[0])
            insert_pos = list(table).index(base_row) + 1
            for offset, item in enumerate(goods_items[1:]):
                new_row = copy.deepcopy(base_row)
                TRADE.fill_goods_row(new_row, item)
                table.insert(insert_pos + offset, new_row)
            rows = table.findall(f"{{{W_NS}}}tr")
            total_index = find_total_row(rows)
    total_number = goods_total_amount(values)
    if total_index >= 0 and total_number:
        cells = rows[total_index].findall(f"{{{W_NS}}}tc")
        if len(cells) >= 8:
            TRADE.set_cell_first_paragraph(cells[3], total_number)
            TRADE.set_cell_first_paragraph(cells[7], total_number)

    rows = table.findall(f"{{{W_NS}}}tr")
    for row in rows:
        text = "".join((t.text or "") for t in row.iter(f"{{{W_NS}}}t"))
        cells = row.findall(f"{{{W_NS}}}tc")
        if len(cells) != 2:
            continue
        if "2.Package Includes" in text or "2. Package Includes" in text:
            continue
        if "Особенности поставки" in text:
            if values.get("package_includes_ru"):
                split_lines_to_cell(cells[0], values["package_includes_ru"])
            if values.get("package_includes_en"):
                split_lines_to_cell(cells[1], values["package_includes_en"])
        elif "3. Итого общая цена" in text:
            if values.get("total_price_ru"):
                split_lines_to_cell(cells[0], f"3. Итого общая цена Товара по настоящему Приложению:\n{total_number} ({values['total_price_ru']}) юаней КНР 00 фэней. ")
            if values.get("total_price_en"):
                split_lines_to_cell(cells[1], f"3. The total price of the Goods under this Appendix:\n{total_number} ({values['total_price_en']}) CNY and 00 fens only.")
        elif "4. Условия поставки" in text:
            if values.get("delivery_terms_ru") or values.get("form_delivery_terms_ru"):
                split_lines_to_cell(cells[0], f"4. Условия поставки:\n{values.get('delivery_terms_ru') or values.get('form_delivery_terms_ru')}")
            if values.get("delivery_terms_en") or values.get("form_delivery_terms_en"):
                split_lines_to_cell(cells[1], f"4. Conditions of delivery:\n{values.get('delivery_terms_en') or values.get('form_delivery_terms_en')}")
        elif "5. Срок поставки Товара" in text:
            split_lines_to_cell(cells[0], f"5. Срок поставки Товара:\n{delivery_time_ru(values)}")
            split_lines_to_cell(cells[1], f"5. Goods delivery time:\n{delivery_time_en(values)}")
        elif "6. Условия платежа" in text:
            split_lines_to_cell(cells[0], payment_terms_ru(values))
            split_lines_to_cell(cells[1], payment_terms_en(values))
        elif "7. Комплект товаросопроводительных документов" in text:
            split_lines_to_cell(cells[0], "7. Комплект товаросопроводительных документов, необходимых для таможенной очистки Товара \n" + (values.get("form_shipping_docs_ru") or shipping_docs_ru()))
            split_lines_to_cell(cells[1], "7. The set of shipping documents required for customs clearance of the Goods:\n" + (values.get("form_shipping_docs_en") or shipping_docs_en()))
        elif "8. Условия проведения технической приемки" in text:
            if values.get("form_acceptance_ru"):
                split_lines_to_cell(cells[0], f"8. Условия проведения технической приемки:\n{values['form_acceptance_ru']}")
            if values.get("form_acceptance_en"):
                split_lines_to_cell(cells[1], f"8. The terms of carrying out technical acceptance:\n{values['form_acceptance_en']}")


def process_word_xml(content: bytes, replacements: dict, fillable_values: set[str], values: dict) -> bytes:
    import sys as _sys
    _lib_dir = Path(__file__).resolve().parent.parent
    if str(_lib_dir) not in _sys.path:
        _sys.path.insert(0, str(_lib_dir))
    from docx_processor import process_document_xml, replace_text_in_xml

    # 通用后处理
    content = process_document_xml(content)
    # 文本替换
    content = replace_text_in_xml(content, replacements)

    # 服务特有的后处理：附件结构化字段、可填写值下划线
    try:
        from lxml import etree as lxml_etree
        parser = lxml_etree.XMLParser(remove_blank_text=False)
        root = lxml_etree.fromstring(content, parser)
    except Exception:
        return content
    apply_appendix_structured_fields(root, values)
    TRADE.underline_fillable_values(root, fillable_values)
    if LXML_ET:
        return LXML_ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def build_fillable_values(data: dict) -> set[str]:
    values = build_values(data)
    fillable = {v for v in values.values() if isinstance(v, str) and v.strip()}
    fillable.update(v for item in TRADE.parse_goods_items(values) for v in item.values() if v)
    fillable.update({payment_terms_ru(values), payment_terms_en(values), delivery_time_ru(values), delivery_time_en(values)})
    fillable.discard("")
    return fillable


def process_docx(input_path: Path, output_path: Path, data: dict) -> None:
    values = build_values(data)
    replacements = build_replacements(values)
    fillable_values = build_fillable_values(data)
    with zipfile.ZipFile(input_path, "r") as zin, zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            # 只处理 word/document.xml，其他 XML 文件原样复制
            # 避免 ET.tostring 重写导致命名空间声明丢失
            if item.filename == "word/document.xml":
                content = process_word_xml(content, replacements, fillable_values, values)
            zout.writestr(item, content)


def generate_pdf(data: dict) -> tuple[bytes, str]:
    if not TEMPLATE_PATH.exists():
        raise RuntimeError("缺少 Word 模板文件 templates/appendix2-template.docx")
    values = build_values(data)
    with tempfile.TemporaryDirectory(prefix="appendix2_pdf_") as tmp:
        tmp_dir = Path(tmp)
        docx_path = tmp_dir / "appendix2.docx"
        process_docx(TEMPLATE_PATH, docx_path, data)
        pdf_path = TRADE.convert_to_pdf(docx_path, tmp_dir)
        pdf_bytes = pdf_path.read_bytes()
    safe_no = re.sub(r"[^\w\-一-龥]+", "_", values.get("contract_no") or time.strftime("%Y%m%d"))
    return pdf_bytes, f"附件2_{safe_no}.pdf"


def load_file(path: Path) -> bytes:
    return path.read_bytes()


class AppendixHandler(BaseHTTPRequestHandler):
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
        elif route == "/wizard.css":
            self.send_bytes(200, load_file(BASE_DIR / "wizard.css"), "text/css; charset=utf-8")
        elif route == "/wizard.js":
            self.send_bytes(200, load_file(BASE_DIR / "wizard.js"), "application/javascript; charset=utf-8")
        elif route == "/defaults":
            payload = json.dumps({"defaults": DEFAULTS, "labels": FIELD_LABELS}, ensure_ascii=False).encode("utf-8")
            self.send_bytes(200, payload, "application/json; charset=utf-8")
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
            pdf_bytes, filename = generate_pdf(data)
            quoted = urllib.parse.quote(filename)
            self.send_bytes(200, pdf_bytes, "application/pdf", {"Content-Disposition": f"attachment; filename*=UTF-8''{quoted}"})
        except Exception as exc:
            self.send_bytes(500, f"生成失败：{exc}".encode("utf-8"), "text/plain; charset=utf-8")


def self_test():
    pdf_bytes, filename = generate_pdf({**DEFAULTS})
    out = BASE_DIR / filename
    out.write_bytes(pdf_bytes)
    print(f"自检完成：{out}")


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        self_test()
    else:
        print(f"附件2生成服务已启动：http://{HOST}:{PORT}")
        ThreadingHTTPServer((HOST, PORT), AppendixHandler).serve_forever()
