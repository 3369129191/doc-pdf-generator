"""进口应付账款垫付申请书与承诺函 PDF 生成服务（端口 8094）。

模板：templates/payable-application-template.docx
特点：
1. 文档为申请方写给智链通达的"申请书 + 承诺函"，主要由表格组成。
2. 需要填写：申请方/买方信息、卖方信息、垫付金额/期限、贸易合同信息、
   日期、费率、收款行信息、签署。
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
TEMPLATE_PATH = BASE_DIR / "templates" / "payable-application-template.docx"
HOST = "127.0.0.1"
PORT = 8094

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

TRADE_SERVER_PY = BASE_DIR.parent / "trade-contract-generator" / "server.py"
_spec = importlib.util.spec_from_file_location("_trade_module", TRADE_SERVER_PY)
TRADE = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(TRADE)


DEFAULTS = {
    # 收件方（垫付方，模板默认）
    "recipient_zh": "智鏈通達供應鏈管理有限公司",
    "recipient_ru": "СМАРТЧЭЙН НЕКСУС ЛИМИТЕД",

    # 申请方 / 买方
    "buyer_name_en": "AIRFLOT TECHNICS TH Ltd.",
    "buyer_name_ru": "ООО «ТД ЭЙРФЛОТ ТЕХНИКС»",

    # 卖方
    "seller_name": "GUANGDONG OSHUJIAN FURNITURE CO.,LTD",

    # 申请详情
    "advance_amount": "",                # 申请垫付金额（人民币元）
    "advance_term_days": "",             # 申请垫付期限（天）
    "trade_contract_no": "",             # 贸易合同号
    "trade_contract_date": "",           # 贸易合同日期
    "trade_contract_amount": "",         # 贸易合同金额（CNY）
    "invoice_amount": "",                # 商业发票金额（CNY）
    "shipment_port_zh": "",              # 中国装运港名称（中文，例：上海/深圳/广州）
    "shipment_port_ru": "",              # 装运港名称（俄文转写，例：Шанхай）
    "buyer_confirmed_amount": "",        # 买方确认应付账款金额（CNY）
    "interest_start_date": "",           # 买方确认原应付账款到期日（用作计息开始日）
    "fee_annual_percent": "",            # 资金占用费率（年化%）
    "fee_overdue_percent": "18",         # 逾期年化（默认 18）
    "advance_due_date": "",              # 买方确认垫付到期日
    "remarks": "",                       # 备注

    # 收款方（垫付款汇入的供应商账户 - 模板默认 GUANGDONG OSHUJIAN）
    "supplier_payee_name": "GUANGDONG OSHUJIAN FURNITURE CO.,LTD",
    "supplier_bank_zh": "VTB Bank (PJSC) Shanghai Branch",
    "supplier_bank_ru": "ВТБ Банк (ПАО) Шанхайский филиал",
    "supplier_account_no": "40807156200610027979",

    # 代理收款单位（贸易项下垫付款的代理收款方）
    "agent_payee_name": "",
    "agent_account_no": "",
    "agent_bank": "",

    # 索偿期限（在收到通知后 X 天内偿付）
    "claim_days": "",

    # 申请方签署区
    "applicant_company_name": "AIRFLOT TECHNICS TH Ltd.",  # 盖章公司名（默认买方）
    "legal_rep_name": "",                # 法定代表人/授权代表姓名
    "actual_controller_name": "",        # 实际控制人姓名
    "contact_phone": "",                 # 联系电话

    # 申请承诺日期（同时用于中文和俄文）
    "sign_year": "",
    "sign_month": "",
    "sign_day": "",
}


FIELD_LABELS = {
    "buyer_name_en": "买方公司名",
    "buyer_name_ru": "买方公司名（俄文）",
    "seller_name": "卖方公司名",
    "advance_amount": "申请垫付金额",
    "advance_term_days": "申请垫付期限（天）",
    "trade_contract_no": "贸易合同号",
    "trade_contract_amount": "贸易合同金额",
    "fee_annual_percent": "年化费率（%）",
    "interest_start_date": "原应付到期日 / 起息日",
    "advance_due_date": "垫付到期日",
    "applicant_company_name": "签署公司",
    "legal_rep_name": "法定代表人/授权代表",
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
}.items():
    try:
        ET.register_namespace(prefix, uri)
    except ValueError:
        pass


# ---------------------------------------------------------------------------
# 工具函数
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


def _new_run(text: str, *, underline: bool = True, rpr_template: ET.Element | None = None) -> ET.Element:
    run = (LXML_ET.Element if LXML_ET else ET.Element)(W + "r")
    if rpr_template is not None:
        rpr = copy.deepcopy(rpr_template)
        run.append(rpr)
    else:
        rpr = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "rPr")
    if underline:
        u = rpr.find(W + "u")
        if u is None:
            u = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(rpr, W + "u")
        u.set(W + "val", "single")
        color = rpr.find(W + "color")
        if color is None:
            color = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(rpr, W + "color")
        color.set(W + "val", "000000")
    t = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "t")
    t.text = text
    if text and (text[0].isspace() or text[-1].isspace()):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return run


def find_paragraphs(root: ET.Element) -> list[ET.Element]:
    return list(root.iter(W + "p"))


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
    # 尝试从 first_run 复制 rPr（字体等），保证新 run 有正确格式
    rpr_template = None
    first_rpr = first_run.find(W + "rPr")
    if first_rpr is not None:
        rpr_template = first_rpr
    paragraph.insert(insert_idx, _new_run(new, underline=underline, rpr_template=rpr_template))
    return True


def append_to_paragraph(paragraph: ET.Element, value: str, *, underline: bool = True, leading_space: bool = True) -> None:
    """在段落末尾追加运行"""
    text = " " + value if leading_space else value
    paragraph.append(_new_run(text, underline=underline))


def overwrite_after_label(paragraph: ET.Element, label: str, value: str, *, underline: bool = True) -> bool:
    """删除 label 后的所有内容，并在其后插入新的下划线运行（保留 label 之前的所有内容）。"""
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
            if run_end == after_label or insert_anchor is None and run_end > pos:
                insert_anchor = run
            continue
        if run_start >= after_label:
            for tn in run.iter(W + "t"):
                tn.text = ""
            cur = run_end
            continue
        # run spans the label boundary - keep prefix
        local_keep = after_label - run_start
        text_nodes = list(run.iter(W + "t"))
        if text_nodes:
            kept = run_text[:local_keep]
            text_nodes[0].text = kept
            for tn in text_nodes[1:]:
                tn.text = ""
        insert_anchor = run
        cur = run_end

    value_run = _new_run(" " + value, underline=underline)
    if insert_anchor is not None:
        children = list(paragraph)
        idx = children.index(insert_anchor)
        paragraph.insert(idx + 1, value_run)
    else:
        paragraph.append(value_run)
    return True


def _set_cell_paragraph_text(paragraph: ET.Element, value: str, *, underline: bool = True) -> None:
    """把段落里所有内容清空，写入一个新的 run（保留 pPr 和 rPr 模板）。
    
    用于填写表格里默认为空的单元格 <w:p w:pPr.../>，避免 PDF 渲染时无字体。
    """
    if not value:
        return
    ppr = paragraph.find(W + "pPr")
    rpr_template = None
    if ppr is not None:
        rpr_template = ppr.find(W + "rPr")
    # 清空除 pPr 之外的所有子元素
    for child in list(paragraph):
        if child.tag != W + "pPr":
            paragraph.remove(child)
    run = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(paragraph, W + "r")
    if rpr_template is not None:
        new_rpr = copy.deepcopy(rpr_template)
        run.insert(0, new_rpr)
    if underline:
        # 给 run 加下划线
        rpr = run.find(W + "rPr")
        if rpr is None:
            rpr = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "rPr")
            run.insert(0, rpr)
        u = rpr.find(W + "u")
        if u is None:
            u = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(rpr, W + "u")
        u.set(W + "val", "single")
        color = rpr.find(W + "color")
        if color is None:
            color = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(rpr, W + "color")
        color.set(W + "val", "000000")
    t = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "t")
    t.text = value
    if value and (value[0].isspace() or value[-1].isspace()):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _fill_empty_cell(cell: ET.Element, value: str, *, underline: bool = True) -> None:
    """填写一个空白单元格的第一个段落。"""
    if not value:
        return
    paragraphs = cell.findall(W + "p")
    if not paragraphs:
        p = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(cell, W + "p")
        paragraphs = [p]
    target = paragraphs[0]
    # 检查段落是否已有内容
    existing = "".join((t.text or "") for t in target.iter(W + "t"))
    if existing.strip():
        # 段落已有内容，附加到末尾
        append_to_paragraph(target, value, underline=underline)
    else:
        _set_cell_paragraph_text(target, value, underline=underline)


# ---------------------------------------------------------------------------
# 业务填写：在表格单元格 / 段落里逐项替换
# ---------------------------------------------------------------------------

def fill_recipient(paragraphs: list[ET.Element], values: dict) -> None:
    """收件方（致：智链通达）"""
    if values.get("recipient_zh"):
        old = "智鏈通達供應鏈管理有限公司"
        new = values["recipient_zh"]
        if new != old:
            for p in paragraphs:
                replace_text_in_paragraph(p, old, new, underline=False)
    if values.get("recipient_ru"):
        old = "СМАРТЧЭЙН НЕКСУС ЛИМИТЕД"
        new = values["recipient_ru"]
        if new != old:
            for p in paragraphs:
                replace_text_in_paragraph(p, old, new, underline=False)


def fill_buyer_seller(paragraphs: list[ET.Element], values: dict) -> None:
    """全文替换买方、卖方公司名。"""
    # 买方英文
    new_buyer_en = values.get("buyer_name_en", "").strip()
    if new_buyer_en and new_buyer_en != "AIRFLOT TECHNICS TH Ltd.":
        for p in paragraphs:
            replace_text_in_paragraph(p, "AIRFLOT TECHNICS TH Ltd.", new_buyer_en, underline=False)

    # 买方俄文
    new_buyer_ru = values.get("buyer_name_ru", "").strip()
    if new_buyer_ru and new_buyer_ru != "ООО «ТД ЭЙРФЛОТ ТЕХНИКС»":
        for p in paragraphs:
            replace_text_in_paragraph(p, "ООО «ТД ЭЙРФЛОТ ТЕХНИКС»", new_buyer_ru, underline=False)

    # 卖方
    new_seller = values.get("seller_name", "").strip()
    if new_seller and new_seller != "GUANGDONG OSHUJIAN FURNITURE CO.,LTD":
        for p in paragraphs:
            replace_text_in_paragraph(p, "GUANGDONG OSHUJIAN FURNITURE CO.,LTD", new_seller, underline=False)


def fill_table(root: ET.Element, values: dict) -> None:
    """填写主表格：根据行内左侧文字定位右侧 cell，写入对应字段。"""
    tables = list(root.iter(W + "tbl"))
    if not tables:
        return
    main_table = tables[0]
    rows = main_table.findall(W + "tr")

    def cell_text(cell: ET.Element) -> str:
        return "".join((t.text or "") for t in cell.iter(W + "t"))

    for row in rows:
        cells = row.findall(W + "tc")
        if len(cells) < 2:
            continue
        left_text = cell_text(cells[0])
        right_cell = cells[1]
        right_text = cell_text(right_cell)

        # 申请垫付额度（CNY）
        if "申请垫付额度" in left_text or "Запрашиваемая сумма" in left_text:
            amt = values.get("advance_amount", "").strip()
            if amt:
                # 右单元格已有 "CNY"，附加金额
                first_p = right_cell.find(W + "p")
                if first_p is not None:
                    append_to_paragraph(first_p, amt, underline=True)

        # 申请垫付期限
        elif "申请垫付期限" in left_text or "Срок предварительной оплаты" in left_text:
            days = values.get("advance_term_days", "").strip()
            if days:
                # 右单元格："_____天/дней"
                for p in right_cell.findall(W + "p"):
                    text = _para_text(p)
                    if "_____" in text or "_____天" in text:
                        replace_text_in_paragraph(p, "_____", days, underline=True)
                        break

        # 贸易合同号 + 日期
        elif "贸易合同号" in left_text or "Номер торгового контракта" in left_text:
            no = values.get("trade_contract_no", "").strip()
            date = values.get("trade_contract_date", "").strip()
            ps = right_cell.findall(W + "p")
            # 第一段：Contract NO. ___，第二段：Date (дата): ___
            if no and ps:
                for p in ps:
                    text = _para_text(p)
                    if "Contract NO" in text or "Contract NO." in text:
                        # 附加在 "Contract NO. " 后
                        append_to_paragraph(p, no, underline=True)
                        break
            if date and len(ps) >= 1:
                for p in ps:
                    text = _para_text(p)
                    if "Date" in text and "дата" in text:
                        append_to_paragraph(p, date, underline=True)
                        break

        # 贸易合同金额（CNY）
        elif "贸易合同金额" in left_text or "Сумма по торговому контракту" in left_text:
            amt = values.get("trade_contract_amount", "").strip()
            if amt:
                first_p = right_cell.find(W + "p")
                if first_p is not None:
                    append_to_paragraph(first_p, amt, underline=True)

        # 商业发票金额
        elif "商业发票金额" in left_text or "Сумма по коммерческому счету" in left_text:
            amt = values.get("invoice_amount", "").strip()
            if amt:
                _fill_empty_cell(right_cell, "CNY " + amt, underline=True)

        # 货物装运日期
        elif "货物装运日期" in left_text or "Дата отгрузки товара" in left_text:
            port_zh = values.get("shipment_port_zh", "").strip()
            port_ru = values.get("shipment_port_ru", "").strip()
            for p in right_cell.findall(W + "p"):
                text = _para_text(p)
                if "**港" in text and port_zh:
                    replace_text_in_paragraph(p, "**", port_zh, underline=True)
                elif "____" in text and port_ru:
                    replace_text_in_paragraph(p, "____", port_ru, underline=True)

        # 买方确认应付账款金额（这一格右单元格是空的）
        elif "买方确认应付账款金额" in left_text or "Подтвержденная покупателем" in left_text:
            amt = values.get("buyer_confirmed_amount", "").strip()
            if amt:
                _fill_empty_cell(right_cell, "CNY " + amt, underline=True)

        # 单独行 "- задолженности:" / 右单元格 "CNY"
        elif left_text.strip().startswith("- задолженности"):
            amt = values.get("buyer_confirmed_amount", "").strip()
            if amt:
                first_p = right_cell.find(W + "p")
                if first_p is not None:
                    append_to_paragraph(first_p, amt, underline=True)

        # 买方确认原应付账款到期日 / Дата начала начисления процентов（起息日）
        elif "买方确认原应付账款到期日" in left_text or "Дата начала начисления процентов" in left_text:
            d = values.get("interest_start_date", "").strip()
            if d:
                _fill_empty_cell(right_cell, d, underline=True)

        # 资金占用费率
        elif "资金占用费率" in left_text or "Ставка за пользование" in left_text:
            annual = values.get("fee_annual_percent", "").strip()
            overdue = values.get("fee_overdue_percent", "").strip() or "18"
            for p in right_cell.findall(W + "p"):
                text = _para_text(p)
                if "**" in text and annual:
                    replace_text_in_paragraph(p, "**", annual, underline=True)
                if overdue and "18%" in _para_text(p) and overdue != "18":
                    replace_text_in_paragraph(p, "18%", overdue + "%", underline=True)
                if "____% годовых" in _para_text(p) and annual:
                    replace_text_in_paragraph(p, "____", annual, underline=True)

        # 买方确认垫付到期日
        elif "买方确认垫付到期日" in left_text or "Дата погашения аванса" in left_text:
            d = values.get("advance_due_date", "").strip()
            if d:
                _fill_empty_cell(right_cell, d, underline=True)

        # 备注 / Примечание
        elif left_text.strip() in ("备注", "备注 Примечание", "Примечание") or (
            "备注" in left_text and "Примечание" in left_text
        ):
            r = values.get("remarks", "").strip()
            if r:
                _fill_empty_cell(right_cell, r, underline=False)


def fill_supplier_account(paragraphs: list[ET.Element], values: dict) -> None:
    """中文长段落里的供应商收款账户：收款人/开户银行 + 俄文行的获取者/银行/账号。"""
    new_payee = values.get("supplier_payee_name", "").strip()
    new_bank_zh = values.get("supplier_bank_zh", "").strip()
    new_bank_ru = values.get("supplier_bank_ru", "").strip()
    new_account = values.get("supplier_account_no", "").strip()

    # 中文段：在长段落里替换默认值（注意此段落重复有 "GUANGDONG OSHUJIAN" 等，已经由 fill_buyer_seller 处理；这里专门处理收款行）
    for p in paragraphs:
        text = _para_text(p)
        if "VTB Bank (PJSC) Shanghai Branch" in text and new_bank_zh and new_bank_zh != "VTB Bank (PJSC) Shanghai Branch":
            replace_text_in_paragraph(p, "VTB Bank (PJSC) Shanghai Branch", new_bank_zh, underline=False)
        if "40807156200610027979" in text and new_account and new_account != "40807156200610027979":
            replace_text_in_paragraph(p, "40807156200610027979", new_account, underline=False)
        # 俄文银行
        if "ВТБ Банк (ПАО) Шанхайский филиал" in text and new_bank_ru and new_bank_ru != "ВТБ Банк (ПАО) Шанхайский филиал":
            replace_text_in_paragraph(p, "ВТБ Банк (ПАО) Шанхайский филиал", new_bank_ru, underline=False)


def fill_agent_payee(paragraphs: list[ET.Element], values: dict) -> None:
    """中文长段："收款人名称  账号  开户银行" — 在标签后插入；俄文 "Получатель:/Номер счета:/Банк:" 三段。"""
    name = values.get("agent_payee_name", "").strip()
    account = values.get("agent_account_no", "").strip()
    bank = values.get("agent_bank", "").strip()

    # 中文版：在长段落里替换 "收款人名称             账号             开户银行                   ，"
    if name or account or bank:
        for p in paragraphs:
            text = _para_text(p)
            # 关键标识："只能将货款汇入贵方指定的代理收款单位的账户上"
            if "代理收款单位" in text and "收款人名称" in text and "账号" in text and "开户银行" in text:
                # 替换三个连续空白为填写值
                # 模式："收款人名称             账号             开户银行                   "
                # 用正则匹配三段空白
                full_text = text
                # 找到模板段
                m = re.search(r"收款人名称(\s+)账号(\s+)开户银行(\s+)", full_text)
                if m:
                    old_block = m.group(0)
                    new_block = f"收款人名称 {name or '____'} 账号 {account or '____'} 开户银行 {bank or '____'} "
                    replace_text_in_paragraph(p, old_block, new_block, underline=False)
                    # 进一步把填的值加下划线 — 用追加替换（仅限非默认值）
                break

    # 俄文版：三段独立行（Получатель:、Номер счета:、Банк:）出现在 "уполномоченного сборочного органа" 之后
    # 用段落定位
    n = len(paragraphs)
    for i, p in enumerate(paragraphs):
        text = _para_text(p)
        if "уполномоченного сборочного органа" in text:
            # 接下来 3 段应该是 Получатель: / Номер счета: / Банк:
            for j in range(i + 1, min(i + 8, n)):
                pt = _para_text(paragraphs[j])
                if pt.strip().startswith("Получатель:") and pt.strip() == "Получатель:" and name:
                    append_to_paragraph(paragraphs[j], name, underline=True)
                elif pt.strip().startswith("Номер счета:") and pt.strip() == "Номер счета:" and account:
                    append_to_paragraph(paragraphs[j], account, underline=True)
                elif pt.strip().startswith("Банк:") and pt.strip() == "Банк:" and bank:
                    append_to_paragraph(paragraphs[j], bank, underline=True)
            break


def fill_claim_days(paragraphs: list[ET.Element], values: dict) -> None:
    """索偿期限：中文段 '通知后     天内'，俄文段 'в течение ____ дней'。"""
    days = values.get("claim_days", "").strip()
    if not days:
        return

    # 中文：在长段落里 "在收到通知后     天内"
    for p in paragraphs:
        text = _para_text(p)
        if "在收到通知后" in text and "天内" in text:
            # 替换 "通知后     天内" 中间的空白
            m = re.search(r"在收到通知后(\s+)天内", text)
            if m:
                old = m.group(0)
                new = f"在收到通知后 {days} 天内"
                replace_text_in_paragraph(p, old, new, underline=False)
                # 把 days 周围加下划线（已通过 new 内嵌但 replace_text_in_paragraph underline=False 不带）
                # 重新加下划线版本：
            break

    # 俄文：'в течение ____ дней'
    for p in paragraphs:
        text = _para_text(p)
        if "в течение" in text and "дней" in text and "____" in text:
            replace_text_in_paragraph(p, "____", days, underline=True)
            break


def fill_signature(paragraphs: list[ET.Element], values: dict) -> None:
    """签署区：申请方公司名/盖章、法定代表人、实控人、电话、日期。"""
    company = values.get("applicant_company_name", "").strip()
    legal = values.get("legal_rep_name", "").strip()
    actual = values.get("actual_controller_name", "").strip()
    phone = values.get("contact_phone", "").strip()

    # 1. 申请方[公司名称]（盖章）—— 把 [公司名称] 替换为公司名
    if company:
        for p in paragraphs:
            text = _para_text(p)
            if "申请方" in text and "公司名称" in text and "盖章" in text:
                replace_text_in_paragraph(p, "[公司名称]", company, underline=True)
                break

    # 2. [法定代表人/授权代表]（签字）：______
    if legal:
        for p in paragraphs:
            text = _para_text(p)
            if "法定代表人" in text and "签字" in text:
                # 在末尾追加（标签后是冒号）
                if "：" in text:
                    append_to_paragraph(p, legal, underline=True)
                else:
                    append_to_paragraph(p, ": " + legal, underline=True)
                break

    # 3. [实际控制人]（签字）：______
    if actual:
        for p in paragraphs:
            text = _para_text(p)
            if "实际控制人" in text and "签字" in text:
                if "：" in text:
                    append_to_paragraph(p, actual, underline=True)
                else:
                    append_to_paragraph(p, ": " + actual, underline=True)
                break

    # 4. Контактный телефон: __________________
    if phone:
        for p in paragraphs:
            text = _para_text(p)
            if "Контактный телефон" in text:
                # 替换默认下划线
                replace_text_in_paragraph(p, "__________________", phone, underline=True)
                # 兼容更短的下划线
                if phone not in _para_text(p):
                    for length in (16, 14, 12, 10, 8, 6, 4):
                        underscores = "_" * length
                        if underscores in _para_text(p):
                            replace_text_in_paragraph(p, underscores, phone, underline=True)
                            break
                break

    # 5. 申请承诺日期：    年    月    日
    y = values.get("sign_year", "").strip()
    m = values.get("sign_month", "").strip()
    d = values.get("sign_day", "").strip()
    if y or m or d:
        # 中文行
        for p in paragraphs:
            text = _para_text(p)
            if "申请承诺日期" in text and "年" in text and "月" in text and "日" in text:
                # 在 年/月/日 前的空白替换
                # 模式："申请承诺日期：    年    月    日"
                _replace_yyyymmdd_zh(p, y, m, d)
                break

        # 俄文行：Дата подачи заявки и гарантийных обязательств: _____ г. _____ мес. _____ д.
        for p in paragraphs:
            text = _para_text(p)
            if "Дата подачи заявки" in text and "г." in text and "мес." in text:
                _replace_yyyymmdd_ru(p, y, m, d)
                break


def _replace_yyyymmdd_zh(p: ET.Element, y: str, m: str, d: str) -> None:
    """中文 '    年    月    日' 替换。支持多种空白模式。"""
    text = _para_text(p)
    # 尝试多种空白模式（从严格到宽松）
    patterns = [
        re.compile(r"(\s{2,})年(\s{2,})月(\s{2,})日"),
        re.compile(r"(\s+)年(\s+)月(\s+)日"),
        re.compile(r"年\s+月\s+日"),
    ]
    for pat in patterns:
        match = pat.search(text)
        if match:
            old = match.group(0)
            new = f" {y or '    '} 年 {m or '  '} 月 {d or '  '} 日"
            if replace_text_in_paragraph(p, old, new, underline=False):
                return
    # 兜底：直接追加到段落末尾（保留原文，追加填写值）
    if y or m or d:
        append_to_paragraph(p, f"{y or ''}年{m or ''}月{d or ''}日", underline=False)


def _replace_yyyymmdd_ru(p: ET.Element, y: str, m: str, d: str) -> None:
    """俄文 '_____ г. _____ мес. _____ д.' 替换。支持多种下划线模式。"""
    text = _para_text(p)
    patterns = [
        re.compile(r"(_{3,})\s*г\.\s*(_{3,})\s*мес\.\s*(_{3,})\s*д\."),
        re.compile(r"(_+)\s*г\.\s*(_+)\s*мес\.\s*(_+)\s*д\."),
        re.compile(r"г\.\s*_+\s*мес\.\s*_+\s*д\."),
    ]
    for pat in patterns:
        match = pat.search(text)
        if match:
            old = match.group(0)
            new = f"{y or '_____'} г. {m or '_____'} мес. {d or '_____'} д."
            if replace_text_in_paragraph(p, old, new, underline=True):
                return
    # 兜底：直接追加
    if y or m or d:
        append_to_paragraph(p, f"{y or ''} г. {m or ''} мес. {d or ''} д.", underline=True)


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

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

    fill_recipient(paragraphs, values)
    fill_buyer_seller(paragraphs, values)
    fill_table(root, values)
    fill_supplier_account(paragraphs, values)
    fill_agent_payee(paragraphs, values)
    fill_claim_days(paragraphs, values)
    fill_signature(paragraphs, values)

    if LXML_ET:
        return LXML_ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def process_docx(input_path: Path, output_path: Path, data: dict) -> None:
    values = {**DEFAULTS, **{k: v.strip() for k, v in data.items() if isinstance(v, str)}}
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
        raise RuntimeError("缺少 Word 模板 templates/payable-application-template.docx")
    with tempfile.TemporaryDirectory(prefix="payable_app_pdf_") as tmp:
        tmp_dir = Path(tmp)
        docx_path = tmp_dir / "payable_application.docx"
        process_docx(TEMPLATE_PATH, docx_path, data)
        pdf_path = TRADE.convert_to_pdf(docx_path, tmp_dir)
        pdf_bytes = pdf_path.read_bytes()
    safe_no = re.sub(r"[^\w\-一-龥]+", "_", data.get("trade_contract_no") or data.get("buyer_name_en") or "申请书")
    return pdf_bytes, f"应付账款垫付申请书_{safe_no}.pdf"


def load_file(path: Path) -> bytes:
    return path.read_bytes()


class PayableApplicationHandler(BaseHTTPRequestHandler):
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
            import traceback
            tb = traceback.format_exc()
            message = f"生成失败：{exc}\n{tb}"
            self.send_bytes(500, message.encode("utf-8"), "text/plain; charset=utf-8")


if __name__ == "__main__":
    print(f"应付账款垫付申请书与承诺函生成网站已启动：http://{HOST}:{PORT}")
    print("保持此窗口运行，在浏览器打开上面的地址。按 Ctrl+C 停止。")
    ThreadingHTTPServer((HOST, PORT), PayableApplicationHandler).serve_forever()
