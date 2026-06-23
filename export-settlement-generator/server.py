"""代理出口结算框架协议 PDF 生成服务（端口 8092）。

模板：templates/export-settlement-template.docx
特点：
1. 模板里大量空白由"标签 + 多个空格 / `____` / `---`"组成，没有黄色高亮统一标记。
2. 后端按"段落 + 标签关键词"逐段精确替换；用户填写值导出后自带下划线。
3. PDF 转换复用 trade-contract-generator 模块（pywin32 → PowerShell COM → soffice）。
"""

from __future__ import annotations

import copy
import importlib.util
import json
import re
import subprocess
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
TEMPLATE_PATH = BASE_DIR / "templates" / "export-settlement-template.docx"
HOST = "127.0.0.1"
PORT = 8092

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


# 复用 trade-contract-generator 里的工具函数（PDF 转换、修订接受、样式归一化等）。
TRADE_SERVER_PY = BASE_DIR.parent / "trade-contract-generator" / "server.py"
_spec = importlib.util.spec_from_file_location("_trade_module", TRADE_SERVER_PY)
TRADE = importlib.util.module_from_spec(_spec)
assert _spec.loader is not None
_spec.loader.exec_module(TRADE)


DEFAULTS = {
    # ① 基础信息
    "contract_no": "",
    "sign_date_zh": "",
    "sign_date_ru": "",
    "sign_place_zh": "",
    "sign_place_ru": "",

    # ② 委托方（甲方） —— 中文
    "partyA_name_zh": "",
    "partyA_credit_code": "",
    "partyA_address_zh": "",
    "partyA_contact_zh": "",
    "partyA_tel": "",
    "partyA_email": "",
    # 委托方（甲方） —— 俄文
    "partyA_name_ru": "",
    "partyA_credit_code_ru": "",
    "partyA_address_ru": "",
    "partyA_contact_ru": "",
    "partyA_tel_ru": "",
    "partyA_email_ru": "",

    # ③ 出口代理方（乙方） —— 中文
    "partyB_name_zh": "",
    "partyB_credit_code": "",
    "partyB_address_zh": "",
    "partyB_contact_zh": "",
    "partyB_tel": "",
    "partyB_email": "",
    # 出口代理方（乙方） —— 俄文
    "partyB_name_ru": "",
    "partyB_credit_code_ru": "",
    "partyB_address_ru": "",
    "partyB_contact_ru": "",
    "partyB_tel_ru": "",
    "partyB_email_ru": "",

    # ④ 货物品名（顶部框架）
    "goods_summary_zh": "",
    "goods_summary_ru": "",

    # ⑤ 第 IV 节代理费率
    "fee_percent": "",
    "fee_min_cny": "",
    "annual_rate_percent": "",

    # ⑥ 甲方收款账户（中俄共用同一信息）
    "account_company": "",
    "account_bank": "",
    "account_no": "",

    # ⑦ 乙方实际控制人
    "partyB_controller": "",

    # ⑧ 签署日期
    "sign_year": "",
    "sign_month": "",
    "sign_day": "",
}


FIELD_LABELS = {
    "contract_no": "协议编号",
    "sign_date_zh": "签订日期（中文）",
    "sign_date_ru": "签订日期（俄文）",
    "sign_place_zh": "签订地点（中文）",
    "sign_place_ru": "签订地点（俄文）",
    "partyA_name_zh": "委托方甲方公司名称（中文）",
    "partyA_credit_code": "甲方统一社会信用代码",
    "partyA_address_zh": "甲方联系地址（中文）",
    "partyA_contact_zh": "甲方联系人（中文）",
    "partyA_tel": "甲方电话",
    "partyA_email": "甲方邮箱",
    "partyA_name_ru": "甲方公司名称（俄文）",
    "partyA_credit_code_ru": "甲方统一社会信用代码（俄文版）",
    "partyA_address_ru": "甲方联系地址（俄文）",
    "partyA_contact_ru": "甲方联系人（俄文）",
    "partyA_tel_ru": "甲方电话（俄文版）",
    "partyA_email_ru": "甲方邮箱（俄文版）",
    "partyB_name_zh": "出口代理方乙方公司名称（中文）",
    "partyB_credit_code": "乙方统一社会信用代码",
    "partyB_address_zh": "乙方联系地址（中文）",
    "partyB_contact_zh": "乙方联系人（中文）",
    "partyB_tel": "乙方电话",
    "partyB_email": "乙方邮箱",
    "partyB_name_ru": "乙方公司名称（俄文）",
    "partyB_credit_code_ru": "乙方统一社会信用代码（俄文版）",
    "partyB_address_ru": "乙方联系地址（俄文）",
    "partyB_contact_ru": "乙方联系人（俄文）",
    "partyB_tel_ru": "乙方电话（俄文版）",
    "partyB_email_ru": "乙方邮箱（俄文版）",
    "goods_summary_zh": "货物品名（中文）",
    "goods_summary_ru": "货物品名（俄文）",
    "fee_percent": "代理费比例（%）",
    "fee_min_cny": "单笔最低代理费（CNY）",
    "annual_rate_percent": "代垫资金占用年化（%）",
    "account_company": "甲方收款账户：公司名称",
    "account_bank": "甲方收款账户：开户行",
    "account_no": "甲方收款账户：账号",
    "partyB_controller": "乙方实际控制人姓名",
    "sign_year": "签署日期：年",
    "sign_month": "签署日期：月",
    "sign_day": "签署日期：日",
}


# 注册命名空间，避免 ET 输出时把 w 改成 ns0。
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


# ----------------------------------------------------------------------------
# 段落工具
# ----------------------------------------------------------------------------

def _para_text(paragraph: ET.Element) -> str:
    return "".join((t.text or "") for t in paragraph.iter(W + "t"))


def _has_text(paragraph: ET.Element) -> bool:
    return bool(_para_text(paragraph).strip())


def _set_run_underline(run: ET.Element) -> None:
    """给一个 run 加单线下划线 + 黑色。"""
    rpr = run.find(W + "rPr")
    if rpr is None:
        rpr = (LXML_ET.Element if LXML_ET else ET.Element)(W + "rPr")
        run.insert(0, rpr)
    # 删除已有的 highlight，避免黄色保留
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


def _make_run_like(template_run: ET.Element, text: str) -> ET.Element:
    """复制一个 run 的样式，文本替换为 text，并加下划线。"""
    new_run = copy.deepcopy(template_run)
    # 清空文本 + 仅保留首个 t 节点
    text_nodes = list(new_run.iter(W + "t"))
    if not text_nodes:
        t = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(new_run, W + "t")
        text_nodes = [t]
    text_nodes[0].text = text
    if text and (text[0].isspace() or text[-1].isspace()):
        text_nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for extra in text_nodes[1:]:
        extra.text = ""
    _set_run_underline(new_run)
    return new_run


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


PLACEHOLDER_RE = re.compile(r"_{3,}|-{3,}")


def fill_label_blank(paragraph: ET.Element, label: str, value: str, *, append_if_no_blank: bool = True) -> bool:
    """在段落中找到 label 后的空白填写位，把它替换成 value 并加下划线。

    支持三种空白形式：
    1. 多个连续下划线  ____
    2. 多个连续短横线  ----
    3. 标签后紧跟的多个空格（≥3 个）

    若都没有，且 append_if_no_blank=True，则在段落末尾追加 ' value'（带下划线）。
    """
    if not value:
        return False
    text = _para_text(paragraph)
    if label not in text:
        return False

    runs = list(paragraph.iter(W + "r"))
    # 用 (run_index, text_within_run_offset) 跟踪段落字符位置
    layout: list[tuple[ET.Element, str]] = []
    for run in runs:
        merged = "".join((t.text or "") for t in run.iter(W + "t"))
        layout.append((run, merged))

    label_pos = text.find(label)
    if label_pos < 0:
        return False
    after_label = label_pos + len(label)

    # 1. 在 label 之后的纯文本中匹配下划线 / 短横线占位
    rest_text = text[after_label:]
    m = PLACEHOLDER_RE.search(rest_text)
    placeholder_start = placeholder_end = -1
    if m:
        placeholder_start = after_label + m.start()
        placeholder_end = after_label + m.end()
    else:
        # 2. 标签紧跟的多空格占位（≥2 个空格也算）
        space_match = re.match(r"\s{2,}", rest_text)
        if space_match:
            placeholder_start = after_label
            placeholder_end = after_label + space_match.end()

    if placeholder_start < 0:
        if not append_if_no_blank:
            return False
        # 把 value 追加到段落最后一个 run 之后
        new_run = _new_run(" " + value)
        paragraph.append(new_run)
        return True

    # 3. 把 [placeholder_start, placeholder_end) 的字符替换成 value，
    # 并尽量把 value 放在第一个被覆盖的 run 中并加下划线。
    pos = 0
    target_run = None
    new_layout: list[str] = []
    for idx, (run, run_text) in enumerate(layout):
        run_len = len(run_text)
        run_start, run_end = pos, pos + run_len
        if run_end <= placeholder_start or run_start >= placeholder_end:
            new_layout.append(run_text)
        else:
            # 这个 run 至少与占位区间相交
            local_start = max(0, placeholder_start - run_start)
            local_end = min(run_len, placeholder_end - run_start)
            if target_run is None:
                # 首个相交 run：保留前缀 + 写入 value，后缀转为单独 run（保留样式）
                prefix = run_text[:local_start]
                suffix = run_text[local_end:]
                # 写入：前缀
                _set_run_text(run, prefix)
                target_run = run
                # 在 prefix 占位 run 之后插入填值 run（带下划线）
                value_run = _make_run_like(run, value)
                _insert_after(paragraph, run, value_run)
                if suffix:
                    suffix_run = copy.deepcopy(run)
                    _set_run_text(suffix_run, suffix)
                    # 移除可能误加在源 run 上的下划线（源 run 前缀部分本没有下划线）
                    _insert_after(paragraph, value_run, suffix_run)
                new_layout.append(prefix + value + suffix)
            else:
                # 后续相交 run：占位之外的部分保留，占位部分清空。
                local_start = max(0, placeholder_start - run_start)
                local_end = min(run_len, placeholder_end - run_start)
                kept = run_text[:local_start] + run_text[local_end:]
                _set_run_text(run, kept)
                new_layout.append(kept)
        pos = run_end
    return True


def _set_run_text(run: ET.Element, value: str) -> None:
    text_nodes = list(run.iter(W + "t"))
    if not text_nodes:
        if not value:
            return
        t = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "t")
        t.text = value
        if value and (value[0].isspace() or value[-1].isspace()):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return
    text_nodes[0].text = value
    if value and (value[0].isspace() or value[-1].isspace()):
        text_nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    elif "{http://www.w3.org/XML/1998/namespace}space" in text_nodes[0].attrib:
        del text_nodes[0].attrib["{http://www.w3.org/XML/1998/namespace}space"]
    for extra in text_nodes[1:]:
        extra.text = ""


def _insert_after(paragraph: ET.Element, anchor: ET.Element, new_child: ET.Element) -> None:
    children = list(paragraph)
    idx = children.index(anchor)
    paragraph.insert(idx + 1, new_child)


def find_paragraphs(root: ET.Element) -> list[ET.Element]:
    return list(root.iter(W + "p"))


def find_paragraph_with(paragraphs: list[ET.Element], keyword: str, *, start: int = 0) -> int:
    for i in range(start, len(paragraphs)):
        if keyword in _para_text(paragraphs[i]):
            return i
    return -1


# ----------------------------------------------------------------------------
# 业务填写
# ----------------------------------------------------------------------------

def fill_party_block_zh(paragraphs: list[ET.Element], header_keyword: str, fields: dict[str, str]) -> None:
    """中文甲乙方信息块按段落顺序填写：标题段 + 5 行（统一信用代码/地址/联系人/电话/邮箱）。

    fields 的 key 必须按顺序：name / credit / address / contact / tel / email。
    name 写入标题段（如果模板里有空白）；其余 5 个字段写入紧随其后的 5 个标签段尾。
    """
    idx = find_paragraph_with(paragraphs, header_keyword)
    if idx < 0:
        return
    # 标题段尾追加公司名称
    if fields.get("name"):
        fill_label_blank(paragraphs[idx], header_keyword, fields["name"])

    label_targets = [
        ("统一社会信用代码：", fields.get("credit")),
        ("联系地址：", fields.get("address")),
        ("联系人名称：", fields.get("contact")),
        ("电话：", fields.get("tel")),
        ("邮箱：", fields.get("email")),
    ]
    cursor = idx + 1
    for label, value in label_targets:
        if not value:
            continue
        for j in range(cursor, min(cursor + 6, len(paragraphs))):
            text = _para_text(paragraphs[j])
            if label in text:
                fill_label_blank(paragraphs[j], label, value)
                cursor = j + 1
                break


def fill_party_block_ru(paragraphs: list[ET.Element], header_keyword: str, fields: dict[str, str]) -> None:
    """俄文甲乙方信息块。标题包含 'Сторона А' 或 'Сторона Б'。"""
    idx = find_paragraph_with(paragraphs, header_keyword)
    if idx < 0:
        return
    # 标题段尾追加公司名（俄文标题以 ":" 结尾，没有现成空白）
    if fields.get("name"):
        # 在标题后追加 ' name'
        para = paragraphs[idx]
        new_run = _new_run(" " + fields["name"])
        para.append(new_run)

    label_targets = [
        ("Единый код социального кредита:", fields.get("credit")),
        ("Юридический адрес:", fields.get("address")),
        ("Контактное лицо:", fields.get("contact")),
        ("Телефон:", fields.get("tel")),
        ("Электронная почта:", fields.get("email")),
    ]
    cursor = idx + 1
    for label, value in label_targets:
        if not value:
            continue
        for j in range(cursor, min(cursor + 8, len(paragraphs))):
            text = _para_text(paragraphs[j])
            if label in text:
                fill_label_blank(paragraphs[j], label, value)
                cursor = j + 1
                break


def fill_basic_info(paragraphs: list[ET.Element], values: dict) -> None:
    """协议头三行：编号 / 签订日期 / 签订地点。

    模板里这 3 段的中文标签和俄文标签处于同一段，共用一个尾部空白占位，
    因此每段只追加一次"中文值 / 俄文值"组合，避免重复填值挤在一起。
    """

    def _combine(zh: str | None, ru: str | None) -> str:
        zh = (zh or "").strip()
        ru = (ru or "").strip()
        if zh and ru:
            return f"{zh} / {ru}"
        return zh or ru

    # P3: 协议编号 / Договор №
    idx = find_paragraph_with(paragraphs, "协议编号：")
    if idx >= 0 and values.get("contract_no"):
        # 把段尾的连续空白先清掉，再追加值。
        para = paragraphs[idx]
        _trim_trailing_whitespace(para)
        para.append(_new_run(" " + values["contract_no"]))

    # P4: 签订日期 / Дата подписания
    idx = find_paragraph_with(paragraphs, "签订日期：")
    if idx >= 0:
        combined = _combine(values.get("sign_date_zh"), values.get("sign_date_ru"))
        if combined:
            para = paragraphs[idx]
            _trim_trailing_whitespace(para)
            para.append(_new_run(" " + combined))

    # P5: 签订地点 / Место подписания（无现成空白，需追加）
    idx = find_paragraph_with(paragraphs, "签订地点：")
    if idx >= 0:
        combined = _combine(values.get("sign_place_zh"), values.get("sign_place_ru"))
        if combined:
            para = paragraphs[idx]
            _trim_trailing_whitespace(para)
            para.append(_new_run(" " + combined))


def _trim_trailing_whitespace(paragraph: ET.Element) -> None:
    """把段落末尾连续的纯空白（≥2 个空格 / 制表符）run 文本清空。"""
    runs = list(paragraph.iter(W + "r"))
    for run in reversed(runs):
        text = "".join((t.text or "") for t in run.iter(W + "t"))
        if not text:
            continue
        if text.strip() == "":
            for t in run.iter(W + "t"):
                t.text = ""
        else:
            break



def fill_goods_summary(paragraphs: list[ET.Element], values: dict) -> None:
    # 中文 P38: "货物品名" + 空白
    idx = find_paragraph_with(paragraphs, "货物品名")
    if idx >= 0 and values.get("goods_summary_zh"):
        fill_label_blank(paragraphs[idx], "货物品名", values["goods_summary_zh"])

    # 俄文 P44: "Наименование товара: ____________"
    idx = find_paragraph_with(paragraphs, "Наименование товара:")
    if idx >= 0 and values.get("goods_summary_ru"):
        fill_label_blank(paragraphs[idx], "Наименование товара:", values["goods_summary_ru"])


def fill_section_iv(paragraphs: list[ET.Element], values: dict) -> None:
    # 中文 P125: "1、按代理出口收汇金额的 ____％ 收取出口代理费，单笔最低不少于人民币 ____元"
    idx = find_paragraph_with(paragraphs, "代理出口收汇金额的")
    if idx >= 0:
        para = paragraphs[idx]
        if values.get("fee_percent"):
            fill_label_blank(para, "代理出口收汇金额的", values["fee_percent"])
        if values.get("fee_min_cny"):
            fill_label_blank(para, "最低不少于人民币", values["fee_min_cny"])

    # 俄文 P127: "_____% ... не менее _______китайских юаней"
    idx = find_paragraph_with(paragraphs, "Агентское вознаграждение за экспорт")
    if idx >= 0:
        para = paragraphs[idx]
        # 该段两个占位都是 _____ 形式：先匹配第一个（在 "размере " 后），再匹配第二个（在 "менее " 后）
        if values.get("fee_percent"):
            fill_label_blank(para, "в размере", values["fee_percent"])
        if values.get("fee_min_cny"):
            fill_label_blank(para, "не менее", values["fee_min_cny"])

    # 中文 P140: "...实际垫付资金的年化 __ %/计收"
    idx = find_paragraph_with(paragraphs, "实际垫付资金的年化")
    if idx >= 0 and values.get("annual_rate_percent"):
        fill_label_blank(paragraphs[idx], "实际垫付资金的年化", values["annual_rate_percent"])

    # 中文 P142: "公司名称__开户行：__账号：__"
    idx = find_paragraph_with(paragraphs, "乙方结算款项汇入甲方以下账户")
    if idx >= 0:
        para = paragraphs[idx]
        if values.get("account_company"):
            fill_label_blank(para, "公司名称", values["account_company"])
        if values.get("account_bank"):
            fill_label_blank(para, "开户行：", values["account_bank"])
        if values.get("account_no"):
            fill_label_blank(para, "账号：", values["account_no"])

    # 俄文 P145/P146/P147 各自一行
    for label, key in [
        ("Наименование компании:", "account_company"),
        ("Банк:", "account_bank"),
        ("Номер счета:", "account_no"),
    ]:
        idx = find_paragraph_with(paragraphs, label)
        if idx >= 0 and values.get(key):
            fill_label_blank(paragraphs[idx], label, values[key])

    # 中文 P149: "...乙方实际控制人 __ 对全部的损失..."
    idx = find_paragraph_with(paragraphs, "乙方实际控制人")
    if idx >= 0 and values.get("partyB_controller"):
        fill_label_blank(paragraphs[idx], "乙方实际控制人", values["partyB_controller"])

    # 俄文 P151: "Фактический контролирующий субъект __ несет"
    idx = find_paragraph_with(paragraphs, "Фактический контролирующий субъект")
    if idx >= 0 and values.get("partyB_controller"):
        fill_label_blank(paragraphs[idx], "Фактический контролирующий субъект", values["partyB_controller"])


def fill_signature_date(paragraphs: list[ET.Element], values: dict) -> None:
    if not (values.get("sign_year") or values.get("sign_month") or values.get("sign_day")):
        return

    # 中文签署日期段：与"签订日期"区分（前者只包含"日期："且伴随"年""月""日"标签，
    # 出现在签字盖章行之后）。采用：找包含"年"和"月"和"日"且不含"签订日期"的段落。
    sig_zh_idx = -1
    for i, p in enumerate(paragraphs):
        txt = _para_text(p)
        if "签订日期" in txt:
            continue
        # 签字日期段的形态："日期：    年  月  日                  日期：   年   月  日"
        if "日期：" in txt and "年" in txt and "月" in txt and "日" in txt and txt.count("日期：") >= 2:
            sig_zh_idx = i
            break
    if sig_zh_idx >= 0:
        para = paragraphs[sig_zh_idx]
        y = values.get("sign_year") or "____"
        m = values.get("sign_month") or "__"
        d = values.get("sign_day") or "__"
        _rebuild_signature_zh(para, y, m, d)

    # 俄文签署日期段：包含两次 "Дата:"
    sig_ru_idx = -1
    for i, p in enumerate(paragraphs):
        txt = _para_text(p)
        if txt.count("Дата:") >= 2:
            sig_ru_idx = i
            break
    if sig_ru_idx >= 0:
        para = paragraphs[sig_ru_idx]
        y = values.get("sign_year") or "____"
        m = _ru_month(values.get("sign_month"))
        d = values.get("sign_day") or "__"
        _rebuild_signature_ru(para, d, m, y)



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
    "10": "октября",
    "11": "ноября",
    "12": "декабря",
}


def _ru_month(value: str | None) -> str:
    if not value:
        return "____"
    return _RU_MONTHS.get(str(value).strip(), str(value).strip())


def _replace_paragraph_keep_first_run(paragraph: ET.Element, text: str) -> None:
    runs = list(paragraph.iter(W + "r"))
    if not runs:
        run = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(paragraph, W + "r")
        (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "t").text = text
        return
    # 保留第一个 run，写入完整文本；后续 run 全部清空。
    text_nodes = list(runs[0].iter(W + "t"))
    if not text_nodes:
        t = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(runs[0], W + "t")
    else:
        t = text_nodes[0]
    t.text = text
    if text and (text[0].isspace() or text[-1].isspace()):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for run in runs[1:]:
        for tn in run.iter(W + "t"):
            tn.text = ""


def _rebuild_signature_zh(paragraph: ET.Element, y: str, m: str, d: str) -> None:
    """把"日期：__年__月__日"段落改造为带下划线的填值结构。

    简单做法：清空当前段落 run，按"标签 + 带下划线 run + 后续标签"的方式重建。
    保留段落 pPr。
    """
    pPr = paragraph.find(W + "pPr")
    # 清空所有 run
    for child in list(paragraph):
        if child.tag != W + "pPr":
            paragraph.remove(child)

    def add_plain(text: str) -> None:
        run = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(paragraph, W + "r")
        t = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "t")
        t.text = text
        if text and (text[0].isspace() or text[-1].isspace()):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    def add_underlined(text: str) -> None:
        paragraph.append(_new_run(text))

    # 第一个日期块
    add_plain("日期：  ")
    add_underlined(f" {y} ")
    add_plain(" 年 ")
    add_underlined(f" {m} ")
    add_plain(" 月 ")
    add_underlined(f" {d} ")
    add_plain(" 日                  ")
    # 第二个日期块
    add_plain("日期：  ")
    add_underlined(f" {y} ")
    add_plain(" 年 ")
    add_underlined(f" {m} ")
    add_plain(" 月 ")
    add_underlined(f" {d} ")
    add_plain(" 日")


def _rebuild_signature_ru(paragraph: ET.Element, d: str, m: str, y: str) -> None:
    for child in list(paragraph):
        if child.tag != W + "pPr":
            paragraph.remove(child)

    def add_plain(text: str) -> None:
        run = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(paragraph, W + "r")
        t = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "t")
        t.text = text
        if text and (text[0].isspace() or text[-1].isspace()):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    def add_underlined(text: str) -> None:
        paragraph.append(_new_run(text))

    add_plain("Дата:  ")
    add_underlined(f" {d} ")
    add_plain(" ")
    add_underlined(f" {m} ")
    add_plain(" ")
    add_underlined(f" {y} ")
    add_plain(" г.         ")
    add_plain("Дата:  ")
    add_underlined(f" {d} ")
    add_plain(" ")
    add_underlined(f" {m} ")
    add_plain(" ")
    add_underlined(f" {y} ")
    add_plain(" г.")


# ----------------------------------------------------------------------------
# 商品表（动态行）
# ----------------------------------------------------------------------------

def parse_table_items(values: dict) -> list[dict[str, str]]:
    raw = values.get("table_items_json", "")
    items: list[dict[str, str]] = []
    if raw:
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            parsed = []
        if isinstance(parsed, list):
            for item in parsed:
                if not isinstance(item, dict):
                    continue
                normalized = {
                    "customs_code": str(item.get("customs_code", "")).strip(),
                    "name_zh": str(item.get("name_zh", "")).strip(),
                    "name_ru": str(item.get("name_ru", "")).strip(),
                    "quantity": str(item.get("quantity", "")).strip(),
                    "amount": str(item.get("amount", "")).strip(),
                }
                if any(normalized.values()):
                    items.append(normalized)
    return items


def fill_goods_table(root: ET.Element, items: list[dict[str, str]]) -> None:
    """模板末尾的商品表是一个 5 列 × 6 行（表头+5空行）的 w:tbl。

    用户提供的 items 按顺序写入空行的 5 个单元格：序号 / 海关编码 / 中文品名 / 数量 / 金额。
    """
    if not items:
        return
    tables = list(root.iter(W + "tbl"))
    if not tables:
        return

    # 找到包含 "序号" + "Сумма" 的表格
    target_table = None
    for tbl in tables:
        text = "".join((t.text or "") for t in tbl.iter(W + "t"))
        if "序号" in text and "Сумма" in text:
            target_table = tbl
            break
    if target_table is None:
        return

    rows = target_table.findall(W + "tr")
    if len(rows) < 2:
        return

    body_rows = rows[1:]  # 第 0 行是表头
    # 为可填行的数量做扩展：默认 5 行，不够时复制最后一行
    while len(body_rows) < len(items):
        new_row = copy.deepcopy(body_rows[-1])
        # 清空所有单元格
        for cell in new_row.findall(W + "tc"):
            for p in cell.findall(W + "p"):
                for t in p.iter(W + "t"):
                    t.text = ""
        target_table.append(new_row)
        body_rows.append(new_row)

    # 写入数据
    for idx, item in enumerate(items):
        row = body_rows[idx]
        cells = row.findall(W + "tc")
        # 默认列序：序号 / 海关编码 / 中文品名 / 数量 / 金额
        col_values = [
            str(idx + 1),
            item["customs_code"],
            item["name_zh"] + ("\n" + item["name_ru"] if item["name_ru"] else ""),
            item["quantity"],
            item["amount"],
        ]
        for cell, value in zip(cells, col_values):
            _set_cell_text(cell, value)

    # 清空多余的空行（如果用户只填了 3 行，剩下两行保持空）
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
    # 保留第一个段落，其它段落清空
    target = paragraphs[0]
    runs = list(target.iter(W + "r"))
    if not runs:
        run = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(target, W + "r")
        (LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, W + "t").text = value
    else:
        text_nodes = list(runs[0].iter(W + "t"))
        if not text_nodes:
            t = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(runs[0], W + "t")
            text_nodes = [t]
        text_nodes[0].text = value
        for tn in text_nodes[1:]:
            tn.text = ""
        # 给值加下划线
        _set_run_underline(runs[0])
    for extra in paragraphs[1:]:
        for run in extra.findall(W + "r"):
            extra.remove(run)


# ----------------------------------------------------------------------------
# 主流程
# ----------------------------------------------------------------------------

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
    fill_basic_info(paragraphs, values)

    fill_party_block_zh(paragraphs, "委托方（以下称甲方）", {
        "name": values.get("partyA_name_zh"),
        "credit": values.get("partyA_credit_code"),
        "address": values.get("partyA_address_zh"),
        "contact": values.get("partyA_contact_zh"),
        "tel": values.get("partyA_tel"),
        "email": values.get("partyA_email"),
    })
    fill_party_block_zh(paragraphs, "出口代理方（以下称乙方）", {
        "name": values.get("partyB_name_zh"),
        "credit": values.get("partyB_credit_code"),
        "address": values.get("partyB_address_zh"),
        "contact": values.get("partyB_contact_zh"),
        "tel": values.get("partyB_tel"),
        "email": values.get("partyB_email"),
    })

    fill_party_block_ru(paragraphs, "Экспортный агент (далее именуемый Сторона А)", {
        "name": values.get("partyA_name_ru"),
        "credit": values.get("partyA_credit_code_ru") or values.get("partyA_credit_code"),
        "address": values.get("partyA_address_ru"),
        "contact": values.get("partyA_contact_ru"),
        "tel": values.get("partyA_tel_ru") or values.get("partyA_tel"),
        "email": values.get("partyA_email_ru") or values.get("partyA_email"),
    })
    fill_party_block_ru(paragraphs, "Доверитель (далее именуемый Сторона Б)", {
        "name": values.get("partyB_name_ru"),
        "credit": values.get("partyB_credit_code_ru") or values.get("partyB_credit_code"),
        "address": values.get("partyB_address_ru"),
        "contact": values.get("partyB_contact_ru"),
        "tel": values.get("partyB_tel_ru") or values.get("partyB_tel"),
        "email": values.get("partyB_email_ru") or values.get("partyB_email"),
    })

    fill_goods_summary(paragraphs, values)
    fill_section_iv(paragraphs, values)
    fill_signature_date(paragraphs, values)

    items = parse_table_items(values)
    fill_goods_table(root, items)

    if LXML_ET:
        return LXML_ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def process_docx(input_path: Path, output_path: Path, data: dict) -> None:
    values = {**DEFAULTS, **{k: v.strip() for k, v in data.items() if isinstance(v, str)}}
    # 把动态商品列表也保留下来，process_word_xml 内部解析
    if "table_items_json" in data:
        values["table_items_json"] = data["table_items_json"]
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
        raise RuntimeError("缺少 Word 模板文件 templates/export-settlement-template.docx")
    with tempfile.TemporaryDirectory(prefix="export_settlement_pdf_") as tmp:
        tmp_dir = Path(tmp)
        docx_path = tmp_dir / "export_settlement.docx"
        process_docx(TEMPLATE_PATH, docx_path, data)
        pdf_path = TRADE.convert_to_pdf(docx_path, tmp_dir)
        pdf_bytes = pdf_path.read_bytes()
    safe_no = re.sub(r"[^\w\-一-龥]+", "_", data.get("contract_no") or "代理出口结算协议")
    return pdf_bytes, f"代理出口结算协议_{safe_no}.pdf"


def load_file(path: Path) -> bytes:
    return path.read_bytes()


class ExportSettlementHandler(BaseHTTPRequestHandler):
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
        except Exception as exc:  # pylint: disable=broad-except
            message = f"生成失败：{exc}"
            self.send_bytes(500, message.encode("utf-8"), "text/plain; charset=utf-8")


if __name__ == "__main__":
    print(f"代理出口结算协议生成网站已启动：http://{HOST}:{PORT}")
    print("保持此窗口运行，在浏览器打开上面的地址。按 Ctrl+C 停止。")
    ThreadingHTTPServer((HOST, PORT), ExportSettlementHandler).serve_forever()
