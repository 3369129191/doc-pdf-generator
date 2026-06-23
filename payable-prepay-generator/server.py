"""进口应付账款委托垫付框架协议 PDF 生成服务（端口 8093）。

模板：templates/payable-prepay-template.docx
特点：
1. 模板里大量字段已经有默认值（甲方智链通达、乙方 AIRFLOT、合同号 SM-TH01 等），
   客户填写的内容会替换已有默认值或空白占位。
2. 替换方式：① 标签后空白占位（_____ / ----）；② 段落内文本替换（公司名、邮箱等）。
3. PDF 转换复用 trade-contract-generator 模块。
"""

from __future__ import annotations

import copy
import importlib.util
import json
import re
import sys
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
TEMPLATE_PATH = BASE_DIR / "templates" / "payable-prepay-template.docx"
HOST = "127.0.0.1"
PORT = 8093

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

TRADE_SERVER_PY = BASE_DIR.parent / "trade-contract-generator" / "server.py"
_spec = importlib.util.spec_from_file_location("_trade_module", TRADE_SERVER_PY)
TRADE = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(TRADE)


DEFAULTS = {
    # ① 协议头
    "contract_no": "SM-TH01",
    "sign_date_zh": "2026年6月17日",
    "sign_date_ru": "17 июня 2026 г.",
    "sign_place_zh": "上海",
    "sign_place_ru": "Шанхай",

    # ② 甲方
    "partyA_name_zh": "智鏈通達供應鏈管理有限公司",
    "partyA_name_ru": "СМАРТЧЭЙН НЕКСУС ЛИМИТЕД",
    "partyA_inn": "9909721327",
    "partyA_kpp": "784287001",
    "partyA_address_zh": "香港新蒲崗太子道東 706 號太子工業大廈海德匯 24 樓 A27 室",
    "partyA_address_ru": "ПОМЕЩЕНИЕ/КОМНАТА А27, ЭТАЖ 24, РИДЖЕНТС ПАРК ПРИНС ИНДАСТРИАЛ БИЛДИНГ №706, ПРИНС ЭДВАРД РОАД, ОСТОЧНЫЙ КОУЛУН, Гонконг, КНР",
    "partyA_contact_zh": "黎健文",
    "partyA_contact_ru": "Ли Цзяньвэнь",
    "partyA_tel": "+852-21349717",
    "partyA_tel_ru": "+852-21349717",
    "partyA_email": "info@aftproject.ru",
    "partyA_email_ru": "info@aftproject.ru",

    # ③ 乙方
    "partyB_name_zh": "AIRFLOT TECHNICS TH Ltd.",
    "partyB_name_ru": "OOO «Торговый Дом ЭЙРФЛОТ ТЕХНИКС»",
    "partyB_inn": "9909721327",
    "partyB_kpp": "784287001",
    "partyB_credit_code": "",
    "partyB_address_zh": "Polkovaya Street 3, bld.5, floor T, section V, office 15, Moscow, Russia. Postcode: 127018.",
    "partyB_address_ru": "127018, Россия, г. Москва, улица Полковая, дом 3, строение 5, этаж Т, пом. V, ком. 15.",
    "partyB_contact_zh": "Zhuravlev Sergey Vladimirovich",
    "partyB_contact_ru": "Журавлева Сергея Владимировича",
    "partyB_tel": "+7 (495) 221-8026",
    "partyB_tel_ru": "+7 (495) 221-8026",
    "partyB_email": "info@aftproject.ru",
    "partyB_email_ru": "info@aftproject.ru",

    # ④ 正文条款
    "trade_seller_name": "GUANGDONG OSHUJIAN FURNITURE CO.,LTD",
    "partyB_contact_email": "106442819@qq.com",
    "partyA_contact_email": "info@aftproject.ru",

    # ⑤ 附件一
    "annex1_contract_no": "SM-TH01",
    "annex1_order_no": "",
    "annex1_client_no": "",
    "fee_annual_percent": "8",
    "fee_overdue_percent": "18",
    "contact_china_name": "",

    # ⑥ 附件二
    "annex2_contract_no": "SM-TH01",
    "annex2_order_no": "",
    "service_start_year": "",
    "service_start_month": "",
    "service_start_day": "",
    "service_end_year": "",
    "service_end_month": "",
    "service_end_day": "",
    "total_cny": "",

    # ⑦ 签署
    "sign_year": "2026",
    "sign_month": "6",
    "sign_day": "17",
    "partyA_signer": "黎健文",
    "partyA_signer_ru": "Ли Цзяньвэнь",
    "partyB_signer": "Zhuravlev Sergey Vladimirovich",
    "partyB_signer_ru": "Журавлева Сергея Владимировича",
}


FIELD_LABELS = {
    "contract_no": "协议编号",
    "sign_date_zh": "签订日期（中文）",
    "sign_date_ru": "签订日期（俄文）",
    "sign_place_zh": "签订地点（中文）",
    "sign_place_ru": "签订地点（俄文）",
    "partyA_name_zh": "甲方公司名（中文）",
    "partyA_name_ru": "甲方公司名（俄文）",
    "partyA_inn": "甲方 INN",
    "partyA_kpp": "甲方 KPP",
    "partyA_address_zh": "甲方地址",
    "partyA_contact_zh": "甲方联系人",
    "partyA_tel": "甲方电话",
    "partyA_email": "甲方邮箱",
    "partyB_name_zh": "乙方公司名",
    "partyB_name_ru": "乙方公司名（俄）",
    "partyB_inn": "乙方 INN",
    "partyB_kpp": "乙方 KPP",
    "trade_seller_name": "贸易出口商",
    "fee_annual_percent": "年化费率（%）",
    "fee_overdue_percent": "逾期年化（%）",
    "contact_china_name": "中国联络人",
    "total_cny": "合计金额",
    "sign_year": "签署年份",
}


for prefix, uri in {
    "w": W_NS,
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
}.items():
    try:
        ET.register_namespace(prefix, uri)
    except ValueError:
        pass


# ---------------------------------------------------------------------------
# 段落工具
# ---------------------------------------------------------------------------

def _para_text(paragraph: ET.Element) -> str:
    return "".join((t.text or "") for t in paragraph.iter(W + "t"))


def _set_run_underline(run: ET.Element) -> None:
    rpr = run.find(W + "rPr")
    if rpr is None:
        rpr = (LXML_ET.Element if LXML_ET else ET.Element)(W + "rPr")
        run.insert(0, rpr)
    for child in list(rpr):
        tag = child.tag.split("}")[-1]
        if tag == "highlight":
            rpr.remove(child)
    u = rpr.find(W + "u")
    if u is None:
        u = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(rpr, W + "u")
    u.set(W + "val", "single")
    color = rpr.find(W + "color")
    if color is None:
        color = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(rpr, W + "color")
    color.set(W + "val", "000000")


def _new_run(text: str) -> ET.Element:
    run = (LXML_ET.Element if LXML_ET else ET.Element)(W + "r")
    rpr = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "rPr")
    u = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(rpr, W + "u")
    u.set(W + "val", "single")
    color = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(rpr, W + "color")
    color.set(W + "val", "000000")
    t = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "t")
    t.text = text
    if text and (text[0].isspace() or text[-1].isspace()):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return run


def _new_plain_run(text: str) -> ET.Element:
    run = (LXML_ET.Element if LXML_ET else ET.Element)(W + "r")
    t = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "t")
    t.text = text
    if text and (text[0].isspace() or text[-1].isspace()):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return run


def find_paragraphs(root: ET.Element) -> list[ET.Element]:
    return list(root.iter(W + "p"))


def find_paragraph_with(paragraphs: list[ET.Element], keyword: str, *, start: int = 0) -> int:
    for i in range(start, len(paragraphs)):
        if keyword in _para_text(paragraphs[i]):
            return i
    return -1


def overwrite_after_label(paragraph: ET.Element, label: str, value: str) -> bool:
    if not value:
        return False
    text = _para_text(paragraph)
    pos = text.find(label)
    if pos < 0:
        return False
    after_label = pos + len(label)

    runs = list(paragraph.iter(W + "r"))
    cur = 0
    insert_anchor = None
    for run in runs:
        run_text = "".join((t.text or "") for t in run.iter(W + "t"))
        run_len = len(run_text)
        run_start, run_end = cur, cur + run_len
        if run_end <= after_label:
            cur = run_end
            continue
        if run_start >= after_label:
            for tn in run.iter(W + "t"):
                tn.text = ""
            cur = run_end
            continue
        local_keep = after_label - run_start
        text_nodes = list(run.iter(W + "t"))
        if text_nodes:
            kept = run_text[:local_keep]
            text_nodes[0].text = kept
            for tn in text_nodes[1:]:
                tn.text = ""
        insert_anchor = run
        cur = run_end

    value_run = _new_run(" " + value)
    if insert_anchor is not None:
        children = list(paragraph)
        idx = children.index(insert_anchor)
        paragraph.insert(idx + 1, value_run)
    else:
        paragraph.append(value_run)
    return True


def replace_text_in_paragraph(paragraph: ET.Element, old: str, new: str, *, underline: bool = True) -> bool:
    if not old or not new:
        return False
    text = _para_text(paragraph)
    pos = text.find(old)
    if pos < 0:
        return False
    end = pos + len(old)

    runs = list(paragraph.iter(W + "r"))
    cur = 0
    first_run_meta = None
    last_run_meta = None
    middle_runs: list[ET.Element] = []
    for run in runs:
        run_text = "".join((t.text or "") for t in run.iter(W + "t"))
        run_len = len(run_text)
        run_start, run_end = cur, cur + run_len
        cur = run_end
        if run_end <= pos or run_start >= end:
            continue
        if run_start <= pos < run_end:
            first_run_meta = (run, pos - run_start, run_text)
        if run_start < end <= run_end:
            last_run_meta = (run, end - run_start, run_text)
        if first_run_meta and run is not first_run_meta[0] and run is not (last_run_meta[0] if last_run_meta else None):
            middle_runs.append(run)

    if not first_run_meta:
        return False

    first_run, first_offset, first_text = first_run_meta
    last_run, last_offset, last_text = last_run_meta if last_run_meta else first_run_meta

    text_nodes = list(first_run.iter(W + "t"))
    if text_nodes:
        text_nodes[0].text = first_text[:first_offset]
        for tn in text_nodes[1:]:
            tn.text = ""

    for r in middle_runs:
        for tn in r.iter(W + "t"):
            tn.text = ""

    if last_run is not first_run:
        last_text_nodes = list(last_run.iter(W + "t"))
        if last_text_nodes:
            last_text_nodes[0].text = last_text[last_offset:]
            for tn in last_text_nodes[1:]:
                tn.text = ""
    else:
        suffix = first_text[last_offset:]
        if suffix:
            suffix_run = copy.deepcopy(first_run)
            sr_text_nodes = list(suffix_run.iter(W + "t"))
            if sr_text_nodes:
                sr_text_nodes[0].text = suffix
                for tn in sr_text_nodes[1:]:
                    tn.text = ""
            children = list(paragraph)
            idx = children.index(first_run)
            paragraph.insert(idx + 1, suffix_run)
            last_run = suffix_run

    children = list(paragraph)
    insert_idx = children.index(first_run) + 1
    if underline:
        value_run = _new_run(new)
    else:
        value_run = _new_plain_run(new)
    paragraph.insert(insert_idx, value_run)
    return True


def rewrite_paragraph(paragraph: ET.Element, segments: list[tuple[str, bool]]) -> None:
    pPr = paragraph.find(W + "pPr")
    for child in list(paragraph):
        if child.tag != W + "pPr":
            paragraph.remove(child)
    for text, underlined in segments:
        if underlined:
            paragraph.append(_new_run(text))
        else:
            paragraph.append(_new_plain_run(text))


def _replace_blank_after(paragraph: ET.Element, label: str, value: str) -> bool:
    if not value:
        return False
    text = _para_text(paragraph)
    pos = text.find(label)
    if pos < 0:
        return False
    after = pos + len(label)
    rest = text[after:]
    m = re.search(r"_{3,}|-{3,}|\s{2,}", rest)
    if not m:
        runs = list(paragraph.iter(W + "r"))
        cur = 0
        for run in runs:
            run_text = "".join((t.text or "") for t in run.iter(W + "t"))
            run_len = len(run_text)
            cur += run_len
            if cur >= after:
                children = list(paragraph)
                idx = children.index(run)
                paragraph.insert(idx + 1, _new_run(" " + value + " "))
                return True
        paragraph.append(_new_run(" " + value + " "))
        return True

    placeholder_start = after + m.start()
    placeholder_end = after + m.end()
    runs = list(paragraph.iter(W + "r"))
    cur = 0
    insert_anchor = None
    for run in runs:
        run_text = "".join((t.text or "") for t in run.iter(W + "t"))
        run_len = len(run_text)
        run_start, run_end = cur, cur + run_len
        cur = run_end
        if run_end <= placeholder_start or run_start >= placeholder_end:
            continue
        local_start = max(0, placeholder_start - run_start)
        local_end = min(run_len, placeholder_end - run_start)
        kept = run_text[:local_start] + run_text[local_end:]
        text_nodes = list(run.iter(W + "t"))
        if text_nodes:
            text_nodes[0].text = kept
            for tn in text_nodes[1:]:
                tn.text = ""
        if insert_anchor is None:
            insert_anchor = run
    if insert_anchor is None:
        paragraph.append(_new_run(" " + value + " "))
    else:
        children = list(paragraph)
        idx = children.index(insert_anchor)
        paragraph.insert(idx + 1, _new_run(" " + value + " "))
    return True


# ---------------------------------------------------------------------------
# 业务填写
# ---------------------------------------------------------------------------

def _combine(zh: str, ru: str) -> str:
    zh = (zh or "").strip()
    ru = (ru or "").strip()
    if zh and ru:
        return f"{zh} / {ru}"
    return zh or ru


def fill_header(paragraphs: list[ET.Element], values: dict) -> None:
    if values.get("contract_no"):
        idx = find_paragraph_with(paragraphs, "Договор №")
        if idx >= 0:
            overwrite_after_label(paragraphs[idx], "Договор №:", values["contract_no"])

    combined_date = _combine(values.get("sign_date_zh", ""), values.get("sign_date_ru", ""))
    if combined_date:
        idx = find_paragraph_with(paragraphs, "Дата подписания")
        if idx >= 0:
            overwrite_after_label(paragraphs[idx], "Дата подписания:", combined_date)

    combined_place = _combine(values.get("sign_place_zh", ""), values.get("sign_place_ru", ""))
    if combined_place:
        idx = find_paragraph_with(paragraphs, "Место подписания")
        if idx >= 0:
            overwrite_after_label(paragraphs[idx], "Место подписания:", combined_place)


def _find_labeled_paragraph(paragraphs, label, exclude_kw, start=0):
    """在全文中找包含 label 且不包含 exclude_kw 的第一个段落。"""
    for j in range(start, len(paragraphs)):
        text = _para_text(paragraphs[j])
        if label in text and not any(kw in text for kw in (exclude_kw if isinstance(exclude_kw, (list, tuple)) else [exclude_kw])):
            return j
    return -1


def fill_party_a(paragraphs: list[ET.Element], values: dict) -> None:
    """甲方信息块（中文 + 俄文）。用全文替换策略。"""
    used = set()

    # 甲方中文标题行
    if values.get("partyA_name_zh"):
        idx = find_paragraph_with(paragraphs, "甲方（垫付方）")
        if idx >= 0:
            overwrite_after_label(paragraphs[idx], "甲方（垫付方）", values["partyA_name_zh"])
            used.add(idx)

    # 甲方独立公司名行（紧跟标题后的那行）
    if values.get("partyA_name_zh"):
        old_a_zh = "智鏈通達供應鏈管理有限公司"
        for i, p in enumerate(paragraphs):
            if i not in used and old_a_zh in _para_text(p):
                replace_text_in_paragraph(p, old_a_zh, values["partyA_name_zh"])
                used.add(i)

    # 甲方俄文公司名
    if values.get("partyA_name_ru"):
        old_a_ru = "СМАРТЧЭЙН НЕКСУС ЛИМИТЕД"
        for i, p in enumerate(paragraphs):
            if i not in used and old_a_ru in _para_text(p):
                replace_text_in_paragraph(p, old_a_ru, values["partyA_name_ru"])
                used.add(i)

    # 甲方中文信息字段：在"甲方（垫付方）"之后 15 段内搜索
    party_a_start = find_paragraph_with(paragraphs, "甲方（垫付方）")
    party_a_end = len(paragraphs)
    if party_a_start >= 0:
        party_a_end = min(party_a_start + 15, len(paragraphs))

    label_pairs_zh = [
        ("INN：", values.get("partyA_inn")),
        ("KPP：", values.get("partyA_kpp")),
        ("地址：", values.get("partyA_address_zh")),
        ("联系人：", values.get("partyA_contact_zh")),
        ("联系电话：", values.get("partyA_tel")),
        ("邮箱：", values.get("partyA_email")),
    ]
    for label, value in label_pairs_zh:
        if not value:
            continue
        for j in range(max(0, party_a_start + 1) if party_a_start >= 0 else 0, party_a_end):
            text = _para_text(paragraphs[j])
            if label in text:
                overwrite_after_label(paragraphs[j], label, value)
                used.add(j)
                break

    # 甲方俄文信息字段：在"Сторона А"之后 15 段内搜索
    party_a_ru_start = find_paragraph_with(paragraphs, "Сторона А(Платящая сторона)")
    if party_a_ru_start < 0:
        party_a_ru_start = find_paragraph_with(paragraphs, "Сторона А")
    party_a_ru_end = len(paragraphs)
    if party_a_ru_start >= 0:
        party_a_ru_end = min(party_a_ru_start + 15, len(paragraphs))

    label_pairs_ru = [
        ("ИНН:", values.get("partyA_inn")),
        ("КПП:", values.get("partyA_kpp")),
        ("Юридический адрес:", values.get("partyA_address_ru")),
        ("Контактное лицо:", values.get("partyA_contact_ru")),
        ("Телефон:", values.get("partyA_tel_ru") or values.get("partyA_tel")),
        ("Электронная почта:", values.get("partyA_email_ru") or values.get("partyA_email")),
    ]
    for label, value in label_pairs_ru:
        if not value:
            continue
        for j in range(max(0, party_a_ru_start + 1) if party_a_ru_start >= 0 else 0, party_a_ru_end):
            text = _para_text(paragraphs[j])
            if label in text:
                overwrite_after_label(paragraphs[j], label, value)
                used.add(j)
                break


def fill_party_b(paragraphs: list[ET.Element], values: dict) -> None:
    """乙方信息块（中文 + 俄文）。"""
    used = set()

    # 乙方中文标题行
    if values.get("partyB_name_zh"):
        idx = find_paragraph_with(paragraphs, "乙方（垫付委托方）")
        if idx >= 0:
            overwrite_after_label(paragraphs[idx], "乙方（垫付委托方）", values["partyB_name_zh"])
            used.add(idx)

    # 乙方俄文标题行：整段替换
    if values.get("partyB_name_ru"):
        for i, p in enumerate(paragraphs):
            if i not in used and "ЭЙРФЛОТ ТЕХНИКС" in _para_text(p):
                rewrite_paragraph(p, [(values["partyB_name_ru"], True)])
                used.add(i)
                break

    # 乙方中文信息字段：在"乙方（垫付委托方）"之后 15 段内搜索
    party_b_start = find_paragraph_with(paragraphs, "乙方（垫付委托方）")
    party_b_zh_end = len(paragraphs)
    if party_b_start >= 0:
        party_b_zh_end = min(party_b_start + 15, len(paragraphs))

    label_pairs_zh = [
        ("INN：", values.get("partyB_inn")),
        ("统一社会信用代码：", values.get("partyB_credit_code")),
        ("地址：", values.get("partyB_address_zh")),
        ("联系人：", values.get("partyB_contact_zh")),
        ("联系电话：", values.get("partyB_tel")),
        ("邮箱：", values.get("partyB_email")),
    ]
    cursor = max(0, party_b_start + 1) if party_b_start >= 0 else 0
    for label, value in label_pairs_zh:
        if not value:
            continue
        for j in range(cursor, party_b_zh_end):
            text = _para_text(paragraphs[j])
            if label in text:
                overwrite_after_label(paragraphs[j], label, value)
                cursor = j + 1
                break

    # 乙方俄文信息字段：在"Сторона Б (Заказчик"之后 15 段内搜索
    party_b_ru_start = find_paragraph_with(paragraphs, "Сторона Б (Заказчик")
    party_b_ru_end = len(paragraphs)
    if party_b_ru_start >= 0:
        party_b_ru_end = min(party_b_ru_start + 15, len(paragraphs))

    label_pairs_ru = [
        ("КПП:", values.get("partyB_kpp")),
        ("ИНН:", values.get("partyB_inn_ru") or values.get("partyB_inn")),
        ("Единый код социального кредита:", values.get("partyB_credit_code")),
        ("Юридический адрес:", values.get("partyB_address_ru")),
        ("Контактное лицо:", values.get("partyB_contact_ru")),
        ("Телефон:", values.get("partyB_tel_ru") or values.get("partyB_tel")),
        ("Электронная почта:", values.get("partyB_email_ru") or values.get("partyB_email")),
    ]
    cursor = max(0, party_b_ru_start + 1) if party_b_ru_start >= 0 else 0
    for label, value in label_pairs_ru:
        if not value:
            continue
        for j in range(cursor, party_b_ru_end):
            text = _para_text(paragraphs[j])
            if label in text:
                overwrite_after_label(paragraphs[j], label, value)
                cursor = j + 1
                break


def fill_body_text(paragraphs: list[ET.Element], values: dict) -> None:
    """正文条款中的可变文本：贸易出口商名、联系邮箱。"""
    if values.get("trade_seller_name"):
        old_name = "GUANGDONG OSHUJIAN FURNITURE CO.,LTD"
        for p in paragraphs:
            replace_text_in_paragraph(p, old_name, values["trade_seller_name"])

    if values.get("partyB_contact_email"):
        old_email = "106442819@qq.com"
        for p in paragraphs:
            replace_text_in_paragraph(p, old_email, values["partyB_contact_email"])

    if values.get("partyA_contact_email"):
        old_email_a = "info@aftproject.ru"
        for p in paragraphs:
            replace_text_in_paragraph(p, old_email_a, values["partyA_contact_email"])


def fill_annex1(paragraphs: list[ET.Element], values: dict) -> None:
    """附件一：垫付资金占用费率确认单。用全文搜索。"""
    annex1_idx = find_paragraph_with(paragraphs, "垫付资金占用费率确认单")
    if annex1_idx < 0:
        annex1_idx = 0
    end = len(paragraphs)  # 不再限制范围

    # 协议编号
    if values.get("annex1_contract_no"):
        for j in range(annex1_idx, end):
            text = _para_text(paragraphs[j])
            if "Соглашение №" in text:
                overwrite_after_label(paragraphs[j], "Соглашение №", " : " + values["annex1_contract_no"])
                break

    # 订单编号
    if values.get("annex1_order_no"):
        for j in range(annex1_idx, end):
            text = _para_text(paragraphs[j])
            if "订单编号：" in text and "Заказ" not in text:
                overwrite_after_label(paragraphs[j], "订单编号：", values["annex1_order_no"])
                break
        for j in range(annex1_idx, end):
            text = _para_text(paragraphs[j])
            if "Заказ №" in text:
                overwrite_after_label(paragraphs[j], "Заказ №", " : " + values["annex1_order_no"])
                break

    # 客户编号
    if values.get("annex1_client_no"):
        for j in range(annex1_idx, end):
            text = _para_text(paragraphs[j])
            if "客户编号：" in text:
                overwrite_after_label(paragraphs[j], "客户编号：", values["annex1_client_no"])
                break
        for j in range(annex1_idx, end):
            text = _para_text(paragraphs[j])
            if "Клиент №" in text:
                overwrite_after_label(paragraphs[j], "Клиент №", " : " + values["annex1_client_no"])
                break

    # 乙方名称（附件一抬头）
    if values.get("partyB_name_zh"):
        for j in range(annex1_idx, end):
            text = _para_text(paragraphs[j])
            if "乙方名称：" in text:
                overwrite_after_label(paragraphs[j], "乙方名称：", values["partyB_name_zh"])
                break

    # 中文费率段
    for j in range(annex1_idx, end):
        text = _para_text(paragraphs[j])
        if "中国联络人名称" in text and "年化" in text:
            para = paragraphs[j]
            # 模板格式："乙方的中国联络人名称：资金占用费率：年化_______%（逾期按年化____%）"
            # 先填费率（替换空白），再插入联络人名（在"名称："和"资金占用费率"之间）
            if values.get("fee_annual_percent"):
                _replace_blank_after(para, "年化", values["fee_annual_percent"])
            if values.get("fee_overdue_percent"):
                _replace_blank_after(para, "逾期按年化", values["fee_overdue_percent"])
            if values.get("contact_china_name"):
                # 在"名称："和"资金占用费率"之间插入联络人名
                replace_text_in_paragraph(para, "名称：资金占用费率",
                                          "名称：" + values["contact_china_name"] + "资金占用费率",
                                          underline=False)
            break

    # 俄文费率段
    for j in range(annex1_idx, end):
        text = _para_text(paragraphs[j])
        if "Ставка платы за использование средств" in text:
            para = paragraphs[j]
            if values.get("fee_annual_percent"):
                _replace_blank_after(para, "Ставка платы за использование средств:", values["fee_annual_percent"])
            if values.get("fee_overdue_percent"):
                _replace_blank_after(para, "ставка при просрочке –", values["fee_overdue_percent"])
            break

    # 俄文联络人
    if values.get("contact_china_name"):
        for j in range(annex1_idx, end):
            text = _para_text(paragraphs[j])
            if "Контактное лицо в Китае" in text:
                _replace_blank_after(paragraphs[j], "Контактное лицо в Китае от Стороны Б:", values["contact_china_name"])
                break

    # 附件一中的公司名全文替换
    if values.get("partyB_name_zh"):
        for j in range(annex1_idx, end):
            replace_text_in_paragraph(paragraphs[j], "AIRFLOT TECHNICS TH Ltd.", values["partyB_name_zh"])
    if values.get("partyB_name_ru"):
        for j in range(annex1_idx, end):
            replace_text_in_paragraph(paragraphs[j], "ООО «ТД ЭЙРФЛОТ ТЕХНИКС»", values["partyB_name_ru"])
    if values.get("partyA_name_zh"):
        for j in range(annex1_idx, end):
            replace_text_in_paragraph(paragraphs[j], "智鏈通達供應鏈管理有限公司", values["partyA_name_zh"])
    if values.get("partyA_name_ru"):
        for j in range(annex1_idx, end):
            replace_text_in_paragraph(paragraphs[j], "СМАРТЧЭЙН НЕКСУС ЛИМИТЕД", values["partyA_name_ru"])


def fill_annex2(paragraphs: list[ET.Element], values: dict) -> None:
    """附件二：垫付资金占用费确认单。用全文搜索。"""
    # 找到真正的附件二正文标题（跳过目录行和附件一标题）
    annex2_idx = -1
    for i, p in enumerate(paragraphs):
        text = _para_text(p)
        if "垫付资金占用费确认单" in text and "率" not in text and "附件" not in text and "《" not in text:
            annex2_idx = i
            break
    if annex2_idx < 0:
        return
    end = len(paragraphs)

    # 协议编号
    if values.get("annex2_contract_no"):
        for j in range(annex2_idx, end):
            text = _para_text(paragraphs[j])
            if "Соглашение №" in text:
                overwrite_after_label(paragraphs[j], "Соглашение №", " : " + values["annex2_contract_no"])
                break

    # 订单编号
    if values.get("annex2_order_no"):
        for j in range(annex2_idx, end):
            text = _para_text(paragraphs[j])
            if "订单编号：" in text and "Заказ" not in text:
                overwrite_after_label(paragraphs[j], "订单编号：", values["annex2_order_no"])
                break
        for j in range(annex2_idx, end):
            text = _para_text(paragraphs[j])
            if "Заказ №" in text:
                overwrite_after_label(paragraphs[j], "Заказ №", " : " + values["annex2_order_no"])
                break

    # 服务起止日期 + 公司名（中文段：含"自...至...年月日"的段落）
    sy = values.get("service_start_year") or ""
    sm = values.get("service_start_month") or ""
    sd = values.get("service_start_day") or ""
    ey = values.get("service_end_year") or ""
    em = values.get("service_end_month") or ""
    ed = values.get("service_end_day") or ""

    if any([sy, sm, sd, ey, em, ed]):
        for j in range(annex2_idx, end):
            text = _para_text(paragraphs[j])
            if "自" in text and "至" in text and "年" in text and "月" in text and "日" in text and "期间" in text:
                segments = [("   兹乙方：", False)]
                segments.append((values.get("partyB_name_zh") or "AIRFLOT TECHNICS TH Ltd.", True))
                segments.extend([
                    ("委托甲方：", False),
                    (values.get("partyA_name_zh") or "智鏈通達供應鏈管理有限公司", True),
                    ("承担的项目：进口应付账款垫付服务 ，自 ", False),
                    (sy or "    ", True), (" 年 ", False),
                    (sm or "  ", True), (" 月 ", False),
                    (sd or "  ", True), (" 日至 ", False),
                    (ey or "    ", True), (" 年 ", False),
                    (em or "  ", True), (" 月 ", False),
                    (ed or "  ", True), (" 日期间，乙方已提供甲方所需材料和相关服务，并根据《代理出口垫付框架协议》结算如下：", False),
                ])
                rewrite_paragraph(paragraphs[j], segments)
                break

    # 俄文服务期
    if sy and sm and sd:
        start_str = f"{sd}.{sm}.{sy}"
    else:
        start_str = ""
    if ey and em and ed:
        end_str = f"{ed}.{em}.{ey}"
    else:
        end_str = ""
    if start_str or end_str:
        for j in range(annex2_idx, end):
            text = _para_text(paragraphs[j])
            if "В период с" in text:
                if start_str:
                    replace_text_in_paragraph(paragraphs[j], "..___", start_str, underline=True)
                if end_str:
                    replace_text_in_paragraph(paragraphs[j], "..___", end_str, underline=True)
                break

    # 俄文段落："Сторона Б ... поручает Стороне А ... — _________ г. Шэньян ..."
    # 该段落中的空白处应填入服务开始日期（dd.mm.yyyy 格式，与"В период с..."保持一致）
    if start_str:
        for j in range(annex2_idx, end):
            text = _para_text(paragraphs[j])
            if "поручает Стороне" in text and "Шэньян" in text:
                # 模板中是连续下划线 "__________________________"
                # 用正则匹配并替换
                replace_text_in_paragraph(paragraphs[j], "__________________________", start_str, underline=True)
                # 兼容下划线长度不同的情况
                import re as _re
                para_text = _para_text(paragraphs[j])
                if start_str not in para_text:
                    # 仍未替换，尝试匹配较短的下划线
                    for length in (24, 20, 18, 16, 14, 12, 10, 8, 6):
                        underscores = "_" * length
                        if underscores in para_text:
                            replace_text_in_paragraph(paragraphs[j], underscores, start_str, underline=True)
                            break
                break

    # 合计金额
    if values.get("total_cny"):
        for j in range(annex2_idx, end):
            text = _para_text(paragraphs[j])
            if "合计：" in text and "Total" in text:
                _replace_blank_after(paragraphs[j], "合计：", values["total_cny"])
                _replace_blank_after(paragraphs[j], "Total amount:", values["total_cny"])
                _replace_blank_after(paragraphs[j], "Итого:", values["total_cny"])
                break

    # 附件二中的公司名全文替换
    if values.get("partyB_name_zh"):
        for j in range(annex2_idx, end):
            replace_text_in_paragraph(paragraphs[j], "AIRFLOT TECHNICS TH Ltd.", values["partyB_name_zh"])
    if values.get("partyB_name_ru"):
        for j in range(annex2_idx, end):
            replace_text_in_paragraph(paragraphs[j], "ООО «ТД ЭЙРФЛОТ ТЕХНИКС»", values["partyB_name_ru"])
    if values.get("partyA_name_zh"):
        for j in range(annex2_idx, end):
            replace_text_in_paragraph(paragraphs[j], "智鏈通達供應鏈管理有限公司", values["partyA_name_zh"])
    if values.get("partyA_name_ru"):
        for j in range(annex2_idx, end):
            replace_text_in_paragraph(paragraphs[j], "СМАРТЧЭЙН НЕКСУС ЛИМИТЕД", values["partyA_name_ru"])


def fill_service_table(root: ET.Element, items: list[dict[str, str]]) -> None:
    """附件二中的服务费用明细表（3 列：编号 / 类别 / 金额）。"""
    if not items:
        return
    tables = list(root.iter(W + "tbl"))
    if not tables:
        return

    # 找到包含 "服务相关编号" + "Номер по обслуживанию" 的表格
    target_table = None
    for tbl in tables:
        text = "".join((t.text or "") for t in tbl.iter(W + "t"))
        if "服务相关编号" in text and "Номер по" in text:
            target_table = tbl
            break
    if target_table is None:
        return

    rows = target_table.findall(W + "tr")
    if len(rows) < 2:
        return

    body_rows = rows[1:]  # 第 0 行是表头
    while len(body_rows) < len(items):
        new_row = copy.deepcopy(body_rows[-1])
        for cell in new_row.findall(W + "tc"):
            for p in cell.findall(W + "p"):
                for t in p.iter(W + "t"):
                    t.text = ""
        target_table.append(new_row)
        body_rows.append(new_row)

    for idx, item in enumerate(items):
        row = body_rows[idx]
        cells = row.findall(W + "tc")
        col_values = [
            str(idx + 1),
            item.get("category", ""),
            item.get("amount", ""),
        ]
        for cell, value in zip(cells, col_values):
            _set_cell_text(cell, value)

    for j in range(len(items), len(body_rows)):
        row = body_rows[j]
        for cell in row.findall(W + "tc"):
            for p in cell.findall(W + "p"):
                for t in p.iter(W + "t"):
                    t.text = ""


def _set_cell_text(cell: ET.Element, value: str) -> None:
    paragraphs = cell.findall(W + "p")
    if not paragraphs:
        p = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(cell, W + "p")
        paragraphs = [p]
    target = paragraphs[0]
    runs = list(target.iter(W + "r"))
    if not runs:
        # 从段落级 pPr/rPr 复制运行属性（字体、字号等），保证文字可见
        ppr = target.find(W + "pPr")
        rpr_template = None
        if ppr is not None:
            rpr_template = ppr.find(W + "rPr")
        run = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(target, W + "r")
        if rpr_template is not None:
            new_rpr = copy.deepcopy(rpr_template)
            run.insert(0, new_rpr)
        (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "t").text = value
    else:
        text_nodes = list(runs[0].iter(W + "t"))
        if not text_nodes:
            t = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(runs[0], W + "t")
            text_nodes = [t]
        text_nodes[0].text = value
        for tn in text_nodes[1:]:
            tn.text = ""
        _set_run_underline(runs[0])
    for extra in paragraphs[1:]:
        for run in extra.findall(W + "r"):
            extra.remove(run)


def fill_signature_dates(paragraphs: list[ET.Element], values: dict) -> None:
    """签字区：日期 + 授权代表姓名 + 公司名。全文搜索。"""
    y = (values.get("sign_year") or "").strip()
    m = (values.get("sign_month") or "").strip()
    d = (values.get("sign_day") or "").strip()
    if not (y or m or d):
        return
    y = y or "____"
    m = m or "__"
    d = d or "__"

    # 中文 "年    月    日"（纯日期行）
    for p in paragraphs:
        text = _para_text(p)
        if "年" in text and "月" in text and "日" in text and "签订" not in text and "年化" not in text:
            stripped = text.replace(" ", "").replace("\u3000", "")
            if stripped == "年月日":
                rewrite_paragraph(p, [
                    ("    ", False), (f" {y} ", True), ("  年  ", False),
                    (f" {m} ", True), ("  月  ", False), (f" {d} ", True), ("  日", False),
                ])

    # 俄文 "Дата：______г.______мес.______день"
    ru_month = _RU_MONTHS.get(str(m).lstrip("0") or m, m)
    for p in paragraphs:
        text = _para_text(p)
        if "Дата：" in text and "г." in text and "мес." in text and "день" in text:
            rewrite_paragraph(p, [
                ("Дата：", False), (f" {y} ", True), ("г.", False),
                (f" {ru_month} ", True), ("мес.", False), (f" {d} ", True), ("день", False),
            ])

    # 签字区甲方代表姓名（中文）— 兼容全角/半角冒号
    if values.get("partyA_signer"):
        for p in paragraphs:
            text = _para_text(p)
            if ("甲方：" in text or "甲方:" in text) and ("章" in text or "盖章" in text or "签字" in text):
                label = "甲方：" if "甲方：" in text else "甲方:"
                _replace_blank_after(p, label, values["partyA_signer"])
                break

    # 签字区乙方代表姓名（中文）— 兼容全角/半角冒号
    if values.get("partyB_signer"):
        for p in paragraphs:
            text = _para_text(p)
            if ("乙方：" in text or "乙方:" in text) and ("章" in text or "盖章" in text or "签字" in text):
                label = "乙方：" if "乙方：" in text else "乙方:"
                _replace_blank_after(p, label, values["partyB_signer"])
                break

    # 签字区甲方代表姓名（俄文）— 兼容 Сторона А: 和 Сторона А：
    if values.get("partyA_signer_ru"):
        for p in paragraphs:
            text = _para_text(p)
            if ("Сторона А：" in text or "Сторона А:" in text) and "Печать" in text:
                label = "Сторона А：" if "Сторона А：" in text else "Сторона А:"
                _replace_blank_after(p, label, values["partyA_signer_ru"])
                break

    # 签字区乙方代表姓名（俄文）— 兼容 Сторона Б: 和 Сторона Б：，以及分两段的情况
    # 限制搜索范围到正文区域（附件一之前），避免匹配到附件内的签字区
    if values.get("partyB_signer_ru"):
        _sig_end = find_paragraph_with(paragraphs, "垫付资金占用费率确认单")
        if _sig_end < 0:
            _sig_end = len(paragraphs)
        # 先尝试在同一行找到
        for p in paragraphs[:_sig_end]:
            text = _para_text(p)
            if ("Сторона Б：" in text or "Сторона Б:" in text) and "Печать" in text:
                label = "Сторона Б：" if "Сторона Б：" in text else "Сторона Б:"
                _replace_blank_after(p, label, values["partyB_signer_ru"])
                break
        else:
            # 分两段：找到 "Сторона Б:" 段落后，在该段落末尾追加签名值
            for i, p in enumerate(paragraphs[:_sig_end]):
                text = _para_text(p)
                if "Сторона Б" in text and ("Сторона Б:" in text or "Сторона Б：" in text) and "Печать" not in text:
                    p.append(_new_run(" " + values["partyB_signer_ru"]))
                    break

    # 公司名替换已由 fill_party_a / fill_party_b / fill_annex1 / fill_annex2 处理，此处不再重复


def _global_fallback(paragraphs: list[ET.Element], values: dict) -> None:
    """兜底：只处理确实遗漏的签字区和附件编号，不重复已替换的内容。

    不再做公司名全文替换（已由 fill_party_a/b + fill_annex1/2 完成）。
    """
    # 定位附件一和附件二的边界（跳过正文引用和目录行）
    annex1_title = -1
    for i, p in enumerate(paragraphs):
        text = _para_text(p)
        if "垫付资金占用费率确认单" in text and "附件" not in text and "《" not in text:
            annex1_title = i
            break
    # 找到真正的附件二正文标题（跳过目录行和附件一标题）
    annex2_title = -1
    start_search = annex1_title + 1 if annex1_title >= 0 else 0
    for i in range(start_search, len(paragraphs)):
        text = _para_text(paragraphs[i])
        if "垫付资金占用费确认单" in text and "率" not in text and "附件" not in text and "《" not in text:
            annex2_title = i
            break
    annex1_mid = (annex1_title + annex2_title) // 2 if annex1_title >= 0 and annex2_title >= 0 else len(paragraphs)

    # ---- 1. 正文签字区乙方名（在附件一之前，带"已存在"检查，兼容全角/半角冒号）----
    if values.get("partyB_signer"):
        end = annex1_title if annex1_title >= 0 else len(paragraphs)
        for p in paragraphs[:end]:
            text = _para_text(p)
            if "乙方" in text and ("章" in text or "盖章" in text) and "垫付" not in text:
                # 如果已经包含签名值，跳过
                if values["partyB_signer"] in text:
                    continue
                b_pos = text.find("乙方")
                if b_pos >= 0:
                    # 兼容全角：和半角:
                    colon_pos = text.find("：", b_pos)
                    colon_char = "："
                    if colon_pos < 0:
                        colon_pos = text.find(":", b_pos)
                        colon_char = ":"
                    if colon_pos >= 0:
                        after = text[colon_pos + 1:]
                        stripped = after.replace("_", "").replace(" ", "").replace("-", "").replace("（", "").replace("）", "").replace("章", "").replace("\n", "")
                        if not stripped:
                            _replace_blank_after(p, "乙方" + colon_char, values["partyB_signer"])

    # ---- 2. 正文俄文签字区乙方名（在附件一之前，带"已存在"检查，兼容分两段）----
    if values.get("partyB_signer_ru"):
        end = annex1_title if annex1_title >= 0 else len(paragraphs)
        filled = False
        for p in paragraphs[:end]:
            text = _para_text(p)
            if ("Сторона Б：" in text or "Сторона Б:" in text) and "Печать" in text:
                if values["partyB_signer_ru"] in text:
                    filled = True
                    break
                after = text[text.find("Сторона Б:") + len("Сторона Б:"):]
                stripped = after.replace("_", "").replace(" ", "").replace("-", "")
                if not stripped:
                    label = "Сторона Б：" if "Сторона Б：" in text else "Сторона Б:"
                    _replace_blank_after(p, label, values["partyB_signer_ru"])
                    filled = True
                    break
        if not filled:
            # 分两段的情况：Сторона Б: 和 Печать: 在不同段落
            for p in paragraphs[:end]:
                text = _para_text(p)
                if "Сторона Б" in text and ("Сторона Б:" in text or "Сторона Б：" in text) and "Печать" not in text:
                    if values["partyB_signer_ru"] not in text:
                        p.append(_new_run(" " + values["partyB_signer_ru"]))
                    break

    # ---- 3. 附件一协议编号（中文"协议编号："后为空或有默认值）----
    if values.get("annex1_contract_no"):
        start = max(0, annex1_title) if annex1_title >= 0 else 0
        end = annex1_mid if annex2_title >= 0 else len(paragraphs)
        for p in paragraphs[start:end]:
            text = _para_text(p)
            if "协议编号：" in text and "Соглашение" not in text:
                after = text[text.find("协议编号：") + len("协议编号："):]
                # 可能和"订单编号："在同一行，截取到下一个标签
                next_label_pos = after.find("订单编号：")
                if next_label_pos >= 0:
                    after = after[:next_label_pos]
                stripped = after.strip().replace("_", "").replace("-", "")
                if not stripped or stripped == "SM-TH01":
                    overwrite_after_label(p, "协议编号：", values["annex1_contract_no"])
                    break

    # ---- 4. 附件二协议编号 ----
    if values.get("annex2_contract_no"):
        start = annex2_title if annex2_title >= 0 else 0
        for p in paragraphs[start:]:
            text = _para_text(p)
            if "协议编号：" in text and "Соглашение" not in text:
                after = text[text.find("协议编号：") + len("协议编号："):]
                next_label_pos = after.find("订单编号：")
                if next_label_pos >= 0:
                    after = after[:next_label_pos]
                stripped = after.strip().replace("_", "").replace("-", "")
                if not stripped or stripped == "SM-TH01":
                    overwrite_after_label(p, "协议编号：", values["annex2_contract_no"])
                    break

    # ---- 5. 附件一俄文乙方名称（Наименование Стороны Б: 后为空或有默认值）----
    if values.get("partyB_name_ru"):
        start = max(0, annex1_title) if annex1_title >= 0 else 0
        end = annex1_mid if annex2_title >= 0 else len(paragraphs)
        for p in paragraphs[start:end]:
            text = _para_text(p)
            if "Наименование Стороны Б:" in text:
                after = text[text.find("Наименование Стороны Б:") + len("Наименование Стороны Б:"):]
                # 如果为空或还是默认值，替换
                if not after.strip() or "OOO" in after or "ТД" in after or "ЭЙРФЛОТ" in after:
                    overwrite_after_label(p, "Наименование Стороны Б:", values["partyB_name_ru"])
                    break

    # ---- 6. 附件一订单编号（中文"订单编号："后为空）----
    if values.get("annex1_order_no"):
        start = max(0, annex1_title) if annex1_title >= 0 else 0
        end = annex1_mid if annex2_title >= 0 else len(paragraphs)
        for p in paragraphs[start:end]:
            text = _para_text(p)
            if "订单编号：" in text and "Заказ" not in text:
                after = text[text.find("订单编号：") + len("订单编号："):]
                if not after.strip():
                    p.append(_new_run(" " + values["annex1_order_no"]))
                    break

    # ---- 7. 附件二订单编号 ----
    if values.get("annex2_order_no"):
        start = annex2_title if annex2_title >= 0 else 0
        for p in paragraphs[start:]:
            text = _para_text(p)
            if "订单编号：" in text and "Заказ" not in text:
                after = text[text.find("订单编号：") + len("订单编号："):]
                if not after.strip():
                    p.append(_new_run(" " + values["annex2_order_no"]))
                    break

    # ---- 8. 附件一签字区乙方名（中文）----
    if values.get("partyB_signer") and annex1_title >= 0:
        end = annex2_title if annex2_title >= 0 else len(paragraphs)
        for p in paragraphs[annex1_title:end]:
            text = _para_text(p)
            if ("乙方：" in text or "乙方:" in text) and ("章" in text):
                if values["partyB_signer"] in text:
                    break
                label = "乙方：" if "乙方：" in text else "乙方:"
                # 用 overwrite_after_label 直接覆盖"乙方："后的所有内容
                overwrite_after_label(p, label, values["partyB_signer"])
                break

    # ---- 9. 附件一签字区乙方名（俄文）----
    if values.get("partyB_signer_ru") and annex1_title >= 0:
        end = annex2_title if annex2_title >= 0 else len(paragraphs)
        for p in paragraphs[annex1_title:end]:
            text = _para_text(p)
            if ("Сторона Б：" in text or "Сторона Б:" in text) and "Печать" in text:
                if values["partyB_signer_ru"] in text:
                    break
                label = "Сторона Б：" if "Сторона Б：" in text else "Сторона Б:"
                _replace_blank_after(p, label, values["partyB_signer_ru"])
                break
        else:
            for p in paragraphs[annex1_title:end]:
                text = _para_text(p)
                if "Сторона Б" in text and ("Сторона Б:" in text or "Сторона Б：" in text) and "Печать" not in text:
                    if values["partyB_signer_ru"] not in text:
                        p.append(_new_run(" " + values["partyB_signer_ru"]))
                    break

    # ---- 10. 附件二签字区乙方名（中文）----
    if values.get("partyB_signer") and annex2_title >= 0:
        for p in paragraphs[annex2_title:]:
            text = _para_text(p)
            if ("乙方：" in text or "乙方:" in text) and ("章" in text):
                if values["partyB_signer"] in text:
                    break
                label = "乙方：" if "乙方：" in text else "乙方:"
                _replace_blank_after(p, label, values["partyB_signer"])
                break

    # ---- 11. 附件二签字区乙方名（俄文）----
    if values.get("partyB_signer_ru") and annex2_title >= 0:
        for p in paragraphs[annex2_title:]:
            text = _para_text(p)
            if ("Сторона Б：" in text or "Сторона Б:" in text) and "Печать" in text:
                if values["partyB_signer_ru"] in text:
                    break
                label = "Сторона Б：" if "Сторона Б：" in text else "Сторона Б:"
                _replace_blank_after(p, label, values["partyB_signer_ru"])
                break
        else:
            # 分两段的情况
            for p in paragraphs[annex2_title:]:
                text = _para_text(p)
                if "Сторона Б" in text and ("Сторона Б:" in text or "Сторона Б：" in text) and "Печать" not in text:
                    if values["partyB_signer_ru"] not in text:
                        p.append(_new_run(" " + values["partyB_signer_ru"]))
                    break


_RU_MONTHS = {
    "1": "января", "01": "января",
    "2": "февраля", "02": "февраля",
    "3": "марта", "03": "марта",
    "4": "апреля", "04": "апреля",
    "5": "мая", "05": "мая",
    "6": "июня", "06": "июня",
    "7": "июля", "07": "июля",
    "8": "августа", "08": "августа",
    "9": "сентября", "09": "сентября",
    "10": "октября", "11": "ноября", "12": "декабря",
}


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def parse_service_items(values: dict) -> list[dict[str, str]]:
    raw = values.get("service_items_json", "")
    items: list[dict[str, str]] = []
    if raw:
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            parsed = []
        if isinstance(parsed, list):
            for item in parsed:
                if isinstance(item, dict):
                    normalized = {
                        "category": str(item.get("category", "")).strip(),
                        "amount": str(item.get("amount", "")).strip(),
                    }
                    if any(normalized.values()):
                        items.append(normalized)
    return items


def process_word_xml(content: bytes, values: dict) -> bytes:
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

    paragraphs = find_paragraphs(root)

    fill_header(paragraphs, values)
    fill_party_a(paragraphs, values)
    fill_party_b(paragraphs, values)
    fill_body_text(paragraphs, values)
    fill_annex1(paragraphs, values)
    fill_annex2(paragraphs, values)
    fill_signature_dates(paragraphs, values)

    # ---- 兜底：全文精确替换（处理上面遗漏的边界情况）----
    _global_fallback(paragraphs, values)

    items = parse_service_items(values)
    fill_service_table(root, items)

    if LXML_ET:
        return LXML_ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def process_docx(input_path: Path, output_path: Path, data: dict) -> None:
    values = {**DEFAULTS, **{k: v.strip() for k, v in data.items() if isinstance(v, str)}}
    if "service_items_json" in data:
        values["service_items_json"] = data["service_items_json"]
    with zipfile.ZipFile(input_path, "r") as zin, zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            # 只处理 word/document.xml，其他 XML 文件原样复制
            # 避免 ET.tostring 重写导致命名空间声明丢失
            if item.filename == "word/document.xml":
                content = process_word_xml(content, values)
            zout.writestr(item, content)


def generate_pdf(data: dict) -> tuple[bytes, str]:
    if not TEMPLATE_PATH.exists():
        raise RuntimeError("缺少 Word 模板 templates/payable-prepay-template.docx")
    with tempfile.TemporaryDirectory(prefix="payable_prepay_pdf_") as tmp:
        tmp_dir = Path(tmp)
        docx_path = tmp_dir / "payable_prepay.docx"
        process_docx(TEMPLATE_PATH, docx_path, data)
        pdf_path = TRADE.convert_to_pdf(docx_path, tmp_dir)
        pdf_bytes = pdf_path.read_bytes()
    safe_no = re.sub(r"[^\w\-一-龥]+", "_", data.get("contract_no") or "应付账款垫付协议")
    return pdf_bytes, f"应付账款垫付协议_{safe_no}.pdf"


def load_file(path: Path) -> bytes:
    return path.read_bytes()


class PayablePrepayHandler(BaseHTTPRequestHandler):
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
            self.send_bytes(
                200,
                pdf_bytes,
                "application/pdf",
                {"Content-Disposition": f"attachment; filename*=UTF-8''{quoted}"},
            )
        except Exception as exc:
            message = f"生成失败：{exc}"
            self.send_bytes(500, message.encode("utf-8"), "text/plain; charset=utf-8")


if __name__ == "__main__":
    print(f"应付账款垫付协议生成网站已启动：http://{HOST}:{PORT}")
    print("保持此窗口运行，在浏览器打开上面的地址。按 Ctrl+C 停止。")
    ThreadingHTTPServer((HOST, PORT), PayablePrepayHandler).serve_forever()
