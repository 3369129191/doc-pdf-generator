import copy
import html
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.parse
import zipfile
import xml.etree.ElementTree as ET
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

# ---------------------------------------------------------------------------
# Word COM 进程锁（避免多个线程同时创建 Word 实例导致冲突）
# ---------------------------------------------------------------------------
_word_lock = threading.Lock()


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "templates" / "agreement-template.docx"
HOST = "127.0.0.1"
PORT = 8088


DEFAULTS = {
    "agreement_no": "THOS-SMA",
    "sign_date": "",
    "sign_place": "中国香港 ГОНКОНГ, КНР",
    "party_a_name_en": "AIRFLOT TECHNICS TH Ltd.",
    "party_a_name_ru": "OOO «Торговый Дом ЭЙРФЛОТ ТЕХНИКС»",
    "party_a_address": "Polkovaya Street 3, bld.5, floor T, section V, office 15, Moscow, Russia. Postcode: 127018",
    "party_a_address_ru": "127018, Россия, г. Москва, улица Полковая, дом 3, строение 5, этаж Т, пом. V, ком. 15.",
    "party_a_representative": "Zhuravlev Sergey Vladimirovich",
    "party_a_representative_ru": "Журавлева Сергея Владимировича",
    "party_b_name_en": "GUANGDONG OSHUJIAN FURNITURE CO., LTD",
    "party_b_name_ru": "GUANGDONG OSHUJIAN FURNITURE CO., LTD",
    "party_b_address": "Sicun Industrial, Heqing, Jiujiang Nanhai, Foshan, Guangdong, China",
    "party_b_address_ru": "Sicun Industrial, Heqing, Jiujiang Nanhai, Foshan, Guangdong, China",
    "party_b_representative": "Xu Yuan",
    "party_b_representative_ru": "Сюй Юань",
    "party_c_name_cn": "智鏈通達供應鏈管理有限公司",
    "party_c_name_ru": "СМАРТЧЭЙН НЕКСУС ЛИМИТЕД",
    "party_c_address": "： 香港新蒲崗太子道東 706 號太子工業大廈海德匯 24 樓 A27 室",
    "party_c_address_ru": "ПОМЕЩЕНИЕ/КОМНАТА А27, ЭТАЖ 24, РИДЖЕНТС ПАРК ПРИНС ИНДАСТРИАЛ БИЛДИНГ №706, ПРИНС ЭДВАРД РОАД, ВОСТОЧНЫЙ КОУЛУН, Гонконг, КНР",
    "party_c_representative": "黎健文",
    "party_c_representative_ru": "Ли Цзяньвэнь",
    "trade_contract_no": "GOF-04/06-26-CHN",
    "trade_contract_no_with_symbol": "№ GOF-04/06-26-CHN",
    "trade_contract_date_en": "June 4, 2026",
    "trade_contract_date_ru": "«04» июня 2026 г",
    "appendix_no": "№ 2",
    "goods_cn": "采购贸易货款",
    "goods_ru": "Закупка товаров",
    "agent_fee_percent": "1",
    "party_b_bank": "VTB Bank (PJSC) Shanghai Branch",
    "party_b_bank_ru": "ВТБ Банк (ПАО) Шанхайский филиал",
    "party_b_swift": "VTBRCNSH",
    "party_b_bank_address": "Office 2503-2505, Shanghai Tower, 501 Middle Yincheng Road, Shanghai, China, 20012",
    "party_b_account_name": "GUANGDONG OSHUJIAN FURNITURE CO.,LTD",
    "party_b_account_no": "40807156200610027979",
}


for prefix, uri in {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
}.items():
    try:
        ET.register_namespace(prefix, uri)
    except ValueError:
        pass


FIELD_LABELS = {
    "agreement_no": "协议号",
    "sign_date": "签订日期",
    "sign_place": "签订地点",
    "party_a_name_en": "甲方英文名称",
    "party_a_name_ru": "甲方俄文名称",
    "party_a_address": "甲方法律地址",
    "party_a_address_ru": "甲方俄文地址",
    "party_a_representative": "甲方代表",
    "party_a_representative_ru": "甲方俄文代表",
    "party_b_name_en": "乙方英文名称",
    "party_b_name_ru": "乙方俄文名称",
    "party_b_address": "乙方法律地址",
    "party_b_address_ru": "乙方俄文地址",
    "party_b_representative": "乙方代表",
    "party_b_representative_ru": "乙方俄文代表",
    "party_c_name_cn": "丙方中文名称",
    "party_c_name_ru": "丙方俄文名称",
    "party_c_address": "丙方法律地址",
    "party_c_address_ru": "丙方俄文地址",
    "party_c_representative": "丙方代表",
    "party_c_representative_ru": "丙方俄文代表",
    "trade_contract_no": "贸易合同号",
    "trade_contract_date_en": "贸易合同日期",
    "trade_contract_date_ru": "贸易合同俄文日期",
    "appendix_no": "附件编号",
    "goods_cn": "支付用途/货物中文",
    "goods_ru": "支付用途/货物俄文",
    "agent_fee_percent": "代理费比例",
    "party_b_bank": "乙方开户银行",
    "party_b_bank_ru": "乙方开户银行俄文",
    "party_b_swift": "乙方 SWIFT",
    "party_b_bank_address": "乙方银行地址",
    "party_b_account_name": "乙方收款人名称",
    "party_b_account_no": "乙方账户号",
}


def _remove_balanced_tags(xml_text: str, tag_name: str, keep_content: bool = True) -> str:
    """移除正确匹配的嵌套标签。keep_content=True 时保留 <w:ins> 内容，False 时删除 <w:del> 内容。"""
    open_pat = f"<{tag_name}\\b[^>]*>"
    close_pat = f"</{tag_name}>"
    result = []
    i = 0
    while i < len(xml_text):
        open_match = re.search(open_pat, xml_text[i:])
        if open_match is None:
            result.append(xml_text[i:])
            break
        start = i + open_match.start()
        result.append(xml_text[i:start])
        # 找到匹配的关闭标签
        depth = 1
        pos = start + len(open_match.group(0))
        last_close_end = pos
        while depth > 0 and pos < len(xml_text):
            next_open = re.search(open_pat, xml_text[pos:])
            next_close = re.search(close_pat, xml_text[pos:])
            next_open_pos = pos + next_open.start() if next_open else len(xml_text)
            next_close_pos = pos + next_close.start() if next_close else len(xml_text)
            if next_close_pos >= len(xml_text):
                # 没有关闭标签，保留原样
                result.append(xml_text[start:])
                return "".join(result)
            if next_open_pos < next_close_pos:
                depth += 1
                pos = next_open_pos + len(next_open.group(0))
            else:
                depth -= 1
                last_close_end = next_close_pos + len(next_close.group(0))
                pos = last_close_end
        if keep_content:
            result.append(xml_text[start + len(open_match.group(0)):last_close_end - len(next_close.group(0))])
        i = pos
    return "".join(result)


def clean_revision_markup(xml_text: str) -> str:
    # 先检查标签是否平衡，不平衡则跳过处理（避免破坏XML结构）
    ins_open = len(re.findall(r"<w:ins\b", xml_text))
    ins_close = len(re.findall(r"</w:ins>", xml_text))
    del_open = len(re.findall(r"<w:del\b", xml_text))
    del_close = len(re.findall(r"</w:del>", xml_text))
    if ins_open != ins_close or del_open != del_close:
        # 标签不平衡，跳过修订标记清理（Word可以正常处理）
        pass
    else:
        xml_text = re.sub(r"<w:commentRangeStart\b[^>]*/>", "", xml_text)
        xml_text = re.sub(r"<w:commentRangeEnd\b[^>]*/>", "", xml_text)
        xml_text = re.sub(r"<w:r\b[^>]*>.*?<w:commentReference\b[^>]*/>.*?</w:r>", "", xml_text, flags=re.S)
        for _ in range(200):
            new_xml = _remove_balanced_tags(xml_text, "w:ins", keep_content=True)
            new_xml = _remove_balanced_tags(new_xml, "w:del", keep_content=False)
            if new_xml == xml_text:
                break
            xml_text = new_xml
    return xml_text


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_RFONTS_KEYS = ("ascii", "hAnsi", "cs")
LATIN_FONT = "Times New Roman"
EASTASIA_FONT = "SimSun"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def apply_format_overrides(root: ET.Element) -> None:
    """Force Latin/Russian to Times New Roman and prevent table overflow."""
    split_aligned_rows(root)
    for el in list(root.iter()):
        name = local_name(el.tag)
        if name == "rFonts":
            for key in W_RFONTS_KEYS:
                el.set(f"{{{W_NS}}}{key}", LATIN_FONT)
            el.set(f"{{{W_NS}}}eastAsia", EASTASIA_FONT)
            if el.get(f"{{{W_NS}}}hint"):
                el.set(f"{{{W_NS}}}hint", "eastAsia")
        elif name == "trHeight":
            rule = el.get(f"{{{W_NS}}}hRule")
            if rule == "exact":
                el.set(f"{{{W_NS}}}hRule", "atLeast")
        elif name == "tblLayout":
            el.set(f"{{{W_NS}}}type", "fixed")

    # Drop <w:noWrap/> inside table cell properties so long text wraps inside cells.
    for tcPr in root.iter(f"{{{W_NS}}}tcPr"):
        for child in list(tcPr):
            if local_name(child.tag) == "noWrap":
                tcPr.remove(child)

    for table in root.iter(f"{{{W_NS}}}tbl"):
        tbl_pr = ensure_child(table, "tblPr")
        tbl_layout = tbl_pr.find(f"{{{W_NS}}}tblLayout")
        if tbl_layout is None:
            tbl_layout = ET.SubElement(tbl_pr, f"{{{W_NS}}}tblLayout")
        tbl_layout.set(f"{{{W_NS}}}type", "fixed")

    normalize_table_paragraphs(root)
    normalize_table_row_heights(root)
    normalize_run_spacing(root)
    remove_blank_underlines(root)

    # Inject Times New Roman into the docDefaults run properties (styles.xml).
    if local_name(root.tag) == "styles":
        ensure_default_font(root)


def ensure_default_font(styles_root: ET.Element) -> None:
    doc_defaults = styles_root.find(f"{{{W_NS}}}docDefaults")
    if doc_defaults is None:
        return
    rpr_default = doc_defaults.find(f"{{{W_NS}}}rPrDefault")
    if rpr_default is None:
        rpr_default = ET.SubElement(doc_defaults, f"{{{W_NS}}}rPrDefault")
    rpr = rpr_default.find(f"{{{W_NS}}}rPr")
    if rpr is None:
        rpr = ET.SubElement(rpr_default, f"{{{W_NS}}}rPr")
    rfonts = rpr.find(f"{{{W_NS}}}rFonts")
    if rfonts is None:
        rfonts = ET.SubElement(rpr, f"{{{W_NS}}}rFonts")
    for key in W_RFONTS_KEYS:
        rfonts.set(f"{{{W_NS}}}{key}", LATIN_FONT)
    rfonts.set(f"{{{W_NS}}}eastAsia", EASTASIA_FONT)


def ensure_child(parent: ET.Element, tag_name: str) -> ET.Element:
    child = parent.find(f"{{{W_NS}}}{tag_name}")
    if child is None:
        child = ET.Element(f"{{{W_NS}}}{tag_name}")
        if tag_name in {"pPr", "rPr", "tblPr"}:
            parent.insert(0, child)
        else:
            parent.append(child)
    return child


SECTION_MARKER_RE = re.compile(
    r"^\s*("
    r"\d+\.\s*\d+(?:\.\d+)*\.?"  # 1.1 / 1.1.1 / 5.6.
    r"|\d+\.(?!\d)"               # 1. / 5.
    r"|[•\u2022\u25CF\-]"          # bullet markers
    r")"
)

PARTY_BLOCK_RE = re.compile(
    r"^\s*(甲方|乙方|丙方|Сторона\s*[ABCАВС]|Сторона\s+А|Сторона\s+В|Сторона\s+С)\b",
    re.I,
)


def _para_text(paragraph: ET.Element) -> str:
    return "".join((t.text or "") for t in paragraph.iter(f"{{{W_NS}}}t"))


def _has_text(paragraph: ET.Element) -> bool:
    return bool(_para_text(paragraph).strip())


def _section_groups(paragraphs: list[ET.Element]) -> list[list[ET.Element]]:
    """Cluster consecutive paragraphs into section groups.

    A group starts at a paragraph that begins a new top-level marker
    (e.g. `5.6.`, `1.`) and stops just before the next such marker.
    Bullet `•` lines and indented continuations stay with the previous group.
    """
    groups: list[list[ET.Element]] = []
    current: list[ET.Element] = []
    for paragraph in paragraphs:
        text = _para_text(paragraph)
        stripped = text.strip()
        starts_new = False
        if stripped:
            m = SECTION_MARKER_RE.match(stripped)
            if m:
                marker = m.group(1).strip()
                # Treat real numbered markers (1.1 / 5.) as section starts;
                # bullets/dashes belong to the existing group.
                if marker[0].isdigit():
                    starts_new = True
        if starts_new and current:
            groups.append(current)
            current = []
        current.append(paragraph)
    if current:
        groups.append(current)
    return groups


def _blank_split_groups(paragraphs: list[ET.Element]) -> list[list[ET.Element]]:
    groups: list[list[ET.Element]] = []
    current: list[ET.Element] = []
    for paragraph in paragraphs:
        if not _has_text(paragraph):
            if current:
                groups.append(current)
                current = []
            continue
        current.append(paragraph)
    if current:
        groups.append(current)
    return groups


def _party_split_groups(paragraphs: list[ET.Element]) -> list[list[ET.Element]]:
    groups: list[list[ET.Element]] = []
    current: list[ET.Element] = []
    for paragraph in paragraphs:
        text = _para_text(paragraph).strip()
        if PARTY_BLOCK_RE.match(text) and current:
            groups.append(current)
            current = []
        if text:
            current.append(paragraph)
    if current:
        groups.append(current)
    return groups


def _refine_groups(groups: list[list[ET.Element]]) -> list[list[ET.Element]]:
    refined: list[list[ET.Element]] = []
    for group in groups:
        text = "".join(_para_text(p) for p in group)
        blank_groups = _blank_split_groups(group)
        party_groups = _party_split_groups(group)
        if len(party_groups) >= 3 and ("各方信息" in text or "РЕКВИЗИТ" in text):
            refined.extend(party_groups)
        elif len(blank_groups) >= 2 and any(keyword in text for keyword in ("账户信息", "Информация по счетам", "Сторона В", "乙方（收款人）")):
            refined.extend(blank_groups)
        else:
            refined.append(group)
    return refined


def split_aligned_rows(root: ET.Element) -> None:
    """Split bilingual two-column rows so each subsection is its own row.

    The template stores all of section 5 (or section 4, etc.) inside a single
    table row. When the Russian wrapping pushes the row taller than the
    Chinese content, the Chinese cell is left with a large blank tail. By
    breaking each row into one row per numbered subsection, the left and right
    cells stay vertically synchronised and the blank padding disappears.
    """
    for table in list(root.iter(f"{{{W_NS}}}tbl")):
        rows = table.findall(f"{{{W_NS}}}tr")
        new_rows: list[ET.Element] = []
        for row in rows:
            cells = row.findall(f"{{{W_NS}}}tc")
            if len(cells) != 2:
                new_rows.append(row)
                continue
            left_paras = cells[0].findall(f"{{{W_NS}}}p")
            right_paras = cells[1].findall(f"{{{W_NS}}}p")
            left_groups = _section_groups(left_paras)
            right_groups = _section_groups(right_paras)
            left_groups = _refine_groups(left_groups)
            right_groups = _refine_groups(right_groups)
            if len(left_groups) <= 1 or len(right_groups) <= 1:
                new_rows.append(row)
                continue
            # Only split when both sides yield the same number of groups so
            # we do not accidentally misalign content.
            if len(left_groups) != len(right_groups):
                new_rows.append(row)
                continue

            table.remove(row)

            for li, ri in zip(left_groups, right_groups):
                new_row = ET.Element(row.tag, dict(row.attrib))
                # Carry over row-level properties (trPr) if present.
                for child in row:
                    if local_name(child.tag) == "trPr":
                        new_row.append(copy.deepcopy(child))
                # Build left cell.
                new_left = copy.deepcopy(cells[0])
                for p in new_left.findall(f"{{{W_NS}}}p"):
                    new_left.remove(p)
                for p in li:
                    new_left.append(copy.deepcopy(p))
                # Build right cell.
                new_right = copy.deepcopy(cells[1])
                for p in new_right.findall(f"{{{W_NS}}}p"):
                    new_right.remove(p)
                for p in ri:
                    new_right.append(copy.deepcopy(p))
                new_row.append(new_left)
                new_row.append(new_right)
                new_rows.append(new_row)

        # Re-attach rows preserving non-row children (tblPr / tblGrid).
        for child in list(table):
            if local_name(child.tag) == "tr":
                table.remove(child)
        for r in new_rows:
            table.append(r)


def normalize_table_paragraphs(root: ET.Element) -> None:
    """Avoid stretched Russian spacing and keep text inside table cells."""
    for tc in root.iter(f"{{{W_NS}}}tc"):
        # Drop trailing empty paragraphs in each cell so the cell shrinks to its
        # content and does not pad the row height with blank lines.
        paragraphs = tc.findall(f"{{{W_NS}}}p")
        while len(paragraphs) > 1:
            last = paragraphs[-1]
            text = "".join((t.text or "") for t in last.iter(f"{{{W_NS}}}t"))
            if text.strip():
                break
            tc.remove(last)
            paragraphs.pop()
        for paragraph in tc.iter(f"{{{W_NS}}}p"):
            ppr = ensure_child(paragraph, "pPr")
            jc = ppr.find(f"{{{W_NS}}}jc")
            if jc is None:
                jc = ET.SubElement(ppr, f"{{{W_NS}}}jc")
            jc.set(f"{{{W_NS}}}val", "left")
            spacing = ppr.find(f"{{{W_NS}}}spacing")
            if spacing is None:
                spacing = ET.Element(f"{{{W_NS}}}spacing")
                ppr.insert(0, spacing)
            spacing.set(f"{{{W_NS}}}line", "285")
            spacing.set(f"{{{W_NS}}}lineRule", "atLeast")
            # Tighten before/after spacing inside table cells so the Chinese
            # column does not accumulate large blank gaps when the Russian
            # column wraps to many lines.
            spacing.set(f"{{{W_NS}}}before", "0")
            spacing.set(f"{{{W_NS}}}after", "0")
            spacing.set(f"{{{W_NS}}}beforeLines", "0")
            spacing.set(f"{{{W_NS}}}afterLines", "0")
            spacing.set(f"{{{W_NS}}}beforeAutospacing", "0")
            spacing.set(f"{{{W_NS}}}afterAutospacing", "0")


def normalize_table_row_heights(root: ET.Element) -> None:
    """Remove inherited row-height locks so split rows shrink to their text."""
    for tr_pr in root.iter(f"{{{W_NS}}}trPr"):
        for child in list(tr_pr):
            if local_name(child.tag) == "trHeight":
                tr_pr.remove(child)


def normalize_run_spacing(root: ET.Element) -> None:
    """Remove negative character spacing that causes glyph overlap after PDF conversion."""
    for rpr in root.iter(f"{{{W_NS}}}rPr"):
        for child in list(rpr):
            name = local_name(child.tag)
            if name == "spacing":
                try:
                    value = int(child.get(f"{{{W_NS}}}val", "0"))
                except ValueError:
                    value = 0
                if value < 0:
                    child.set(f"{{{W_NS}}}val", "0")
            elif name == "fitText":
                rpr.remove(child)


def remove_blank_underlines(root: ET.Element) -> None:
    """Remove underline formatting from blank runs that can spill outside table borders."""
    for run in root.iter(f"{{{W_NS}}}r"):
        text = "".join((t.text or "") for t in run.iter(f"{{{W_NS}}}t"))
        if text and text.strip():
            continue
        rpr = run.find(f"{{{W_NS}}}rPr")
        if rpr is None:
            continue
        for child in list(rpr):
            if local_name(child.tag) == "u":
                rpr.remove(child)


def underline_fillable_values(root: ET.Element, fillable_values: set[str]) -> None:
    values = sorted((v for v in fillable_values if v and len(v.strip()) >= 2), key=len, reverse=True)
    if not values:
        return
    for run in root.iter(f"{{{W_NS}}}r"):
        text = "".join((t.text or "") for t in run.iter(f"{{{W_NS}}}t"))
        if not text or not text.strip():
            continue
        if any(value in text for value in values):
            rpr = ensure_child(run, "rPr")
            underline = rpr.find(f"{{{W_NS}}}u")
            if underline is None:
                underline = ET.SubElement(rpr, f"{{{W_NS}}}u")
            underline.set(f"{{{W_NS}}}val", "single")


def remove_party_c_signature_names(root: ET.Element) -> None:
    """删除最后丙方签章区中盖章行和签字行之间的公司名称。"""
    for parent in root.iter():
        children = list(parent)
        paragraph_indexes = [i for i, child in enumerate(children) if local_name(child.tag) == "p"]
        if len(paragraph_indexes) < 3:
            continue
        paragraph_texts = {
            i: "".join((t.text or "") for t in children[i].iter(f"{{{W_NS}}}t")).strip()
            for i in paragraph_indexes
        }
        for pos, start_index in enumerate(paragraph_indexes):
            start_text = paragraph_texts[start_index]
            if "丙方（盖章）" not in start_text and "Сторона С" not in start_text and "Сторона C" not in start_text:
                continue
            for end_index in paragraph_indexes[pos + 1:]:
                end_text = paragraph_texts[end_index]
                if "代表人（签字）" in end_text or "Представитель" in end_text:
                    for remove_index in reversed([i for i in paragraph_indexes if start_index < i < end_index]):
                        parent.remove(children[remove_index])
                    return


def accept_revisions_tree(element: ET.Element) -> None:
    children = list(element)
    index = 0
    for child in children:
        name = local_name(child.tag)
        if name in {"commentRangeStart", "commentRangeEnd"}:
            element.remove(child)
            continue
        if name == "r" and any(local_name(grand.tag) == "commentReference" for grand in list(child)):
            element.remove(child)
            continue
        if name == "ins":
            insert_at = list(element).index(child)
            element.remove(child)
            for grand in list(child):
                element.insert(insert_at, grand)
                insert_at += 1
            continue
        if name == "del":
            element.remove(child)
            continue
        accept_revisions_tree(child)
        index += 1


def replace_text_elements(text_elements: list[ET.Element], replacements: dict) -> None:
    parts = [el.text or "" for el in text_elements]

    def apply_one(old: str, new: str) -> None:
        if not old or old == new:
            return
        replace_once = old in new
        while True:
            full = "".join(parts)
            start = full.find(old)
            start = full.find(old)
            if start < 0:
                break
            end = start + len(old)
            pos = 0
            first_idx = last_idx = None
            first_offset = last_offset = 0
            for idx, text in enumerate(parts):
                next_pos = pos + len(text)
                if first_idx is None and start <= next_pos:
                    first_idx = idx
                    first_offset = max(0, start - pos)
                if first_idx is not None and end <= next_pos:
                    last_idx = idx
                    last_offset = max(0, end - pos)
                    break
                pos = next_pos
            if first_idx is None or last_idx is None:
                break
            parts[first_idx] = parts[first_idx][:first_offset] + new + parts[last_idx][last_offset:]
            for idx in range(first_idx + 1, last_idx + 1):
                parts[idx] = ""
            if replace_once:
                break

    for old_value, new_value in replacements.items():
        apply_one(str(old_value), str(new_value))

    for element, text in zip(text_elements, parts):
        element.text = text
        if text[:1].isspace() or text[-1:].isspace():
            element.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def process_word_xml(content: bytes, replacements: dict, fillable_values: set[str]) -> bytes:
    """使用 lxml 处理 word/document.xml，正确处理修订标记和命名空间。"""
    try:
        from lxml import etree as lxml_etree
    except ImportError:
        # 回退到纯字符串替换
        text = content.decode("utf-8")
        text = clean_revision_markup(text)
        text = replace_across_text_nodes(text, replacements)
        return text.encode("utf-8")

    # 使用 lxml 解析，保留所有命名空间
    parser = lxml_etree.XMLParser(remove_blank_text=False)
    root = lxml_etree.fromstring(content, parser)

    # 接受所有修订（将 <w:ins> 内容提升，删除 <w:del> 内容）
    _accept_revisions_lxml(root)

    # 替换文本
    _replace_text_lxml(root, replacements)

    # 删除丙方签章区中的公司名称
    _remove_party_c_signature_lxml(root)

    # 下划线可填写值
    _underline_fillable_lxml(root, fillable_values)

    # 应用格式覆盖
    _apply_format_overrides_lxml(root)

    # 序列化回字节，保留原始 XML 声明格式
    # lxml 会自动保留所有命名空间声明
    result = lxml_etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    return result


def _accept_revisions_lxml(root):
    """接受所有修订标记。"""
    from lxml import etree as lxml_etree
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    # 删除 <w:del> 元素
    for del_el in root.xpath("//w:del", namespaces=ns):
        parent = del_el.getparent()
        if parent is not None:
            parent.remove(del_el)

    # 将 <w:ins> 的子元素提升到父元素中，然后删除 <w:ins>
    for ins_el in root.xpath("//w:ins", namespaces=ns):
        parent = ins_el.getparent()
        if parent is not None:
            index = parent.index(ins_el)
            for child in list(ins_el):
                parent.insert(index, child)
                index += 1
            parent.remove(ins_el)

    # 删除 comment 相关元素
    for el in root.xpath("//w:commentRangeStart | //w:commentRangeEnd | //w:commentReference", namespaces=ns):
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)


def _replace_text_lxml(root, replacements):
    """在 <w:t> 元素中替换文本。"""
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for t_el in root.xpath("//w:t", namespaces=ns):
        if t_el.text:
            for old_val, new_val in replacements.items():
                if old_val and old_val != new_val:
                    t_el.text = t_el.text.replace(old_val, new_val)


def _remove_party_c_signature_lxml(root):
    """删除丙方签章区中盖章行和签字行之间的公司名称。"""
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for p in root.xpath("//w:p", namespaces=ns):
        text = "".join(t.text or "" for t in p.xpath(".//w:t", namespaces=ns))
        if "丙方（盖章）" in text or "Сторона С" in text or "Сторона C" in text:
            # 找到下一个包含 "代表人（签字）" 或 "Представитель" 的段落
            # 删除中间的段落
            current = p.getnext()
            while current is not None:
                current_text = "".join(t.text or "" for t in current.xpath(".//w:t", namespaces=ns))
                if "代表人（签字）" in current_text or "Представитель" in current_text:
                    break
                to_remove = current
                current = current.getnext()
                parent = to_remove.getparent()
                if parent is not None:
                    parent.remove(to_remove)


def _underline_fillable_lxml(root, fillable_values):
    """为可填写值添加下划线。"""
    from lxml import etree as lxml_etree
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    values = sorted((v for v in fillable_values if v and len(v.strip()) >= 2), key=len, reverse=True)
    if not values:
        return
    for run in root.xpath("//w:r", namespaces=ns):
        text = "".join(t.text or "" for t in run.xpath(".//w:t", namespaces=ns))
        if not text or not text.strip():
            continue
        if any(value in text for value in values):
            rpr = run.find(f"{{{W_NS}}}rPr")
            if rpr is None:
                rpr = lxml_etree.SubElement(run, f"{{{W_NS}}}rPr")
            u = rpr.find(f"{{{W_NS}}}u")
            if u is None:
                u = lxml_etree.SubElement(rpr, f"{{{W_NS}}}u")
            u.set(f"{{{W_NS}}}val", "single")


def _apply_format_overrides_lxml(root):
    """应用格式覆盖。"""
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    LATIN_FONT = "Times New Roman"
    EASTASIA_FONT = "SimSun"

    # 设置字体
    for rfonts in root.xpath("//w:rFonts", namespaces=ns):
        for key in ("ascii", "hAnsi", "cs"):
            rfonts.set(f"{{{ns['w']}}}{key}", LATIN_FONT)
        rfonts.set(f"{{{ns['w']}}}eastAsia", EASTASIA_FONT)
        if rfonts.get(f"{{{ns['w']}}}hint"):
            rfonts.set(f"{{{ns['w']}}}hint", "eastAsia")

    # 设置表格单元格对齐
    for jc in root.xpath("//w:tc//w:p/w:pPr/w:jc", namespaces=ns):
        jc.set(f"{{{ns['w']}}}val", "left")

    # 移除表格中的 noWrap
    for nowrap in root.xpath("//w:tcPr/w:noWrap", namespaces=ns):
        parent = nowrap.getparent()
        if parent is not None:
            parent.remove(nowrap)


def replace_across_text_nodes(xml_text: str, replacements: dict) -> str:
    pattern = re.compile(r"(<w:t\b[^>]*>)(.*?)(</w:t>)", re.S)
    matches = list(pattern.finditer(xml_text))
    if not matches:
        return xml_text

    parts = [html.unescape(match.group(2)) for match in matches]

    def apply_one(old: str, new: str) -> None:
        if not old or old == new:
            return
        replace_once = old in new
        while True:
            full = "".join(parts)
            start = full.find(old)
            if start < 0:
                break
            end = start + len(old)
            pos = 0
            first_idx = last_idx = None
            first_offset = last_offset = 0
            for idx, text in enumerate(parts):
                next_pos = pos + len(text)
                if first_idx is None and start <= next_pos:
                    first_idx = idx
                    first_offset = max(0, start - pos)
                if first_idx is not None and end <= next_pos:
                    last_idx = idx
                    last_offset = max(0, end - pos)
                    break
                pos = next_pos
            if first_idx is None or last_idx is None:
                break
            prefix = parts[first_idx][:first_offset]
            suffix = parts[last_idx][last_offset:]
            parts[first_idx] = prefix + new + suffix
            for idx in range(first_idx + 1, last_idx + 1):
                parts[idx] = ""
            if replace_once:
                break

    for old_value, new_value in replacements.items():
        apply_one(str(old_value), str(new_value))

    out = []
    last = 0
    for match, text in zip(matches, parts):
        out.append(xml_text[last:match.start()])
        escaped = html.escape(text, quote=False)
        tag = match.group(1)
        if (text[:1].isspace() or text[-1:].isspace()) and "xml:space" not in tag:
            tag = tag[:-1] + ' xml:space="preserve">'
        out.append(tag + escaped + match.group(3))
        last = match.end()
    out.append(xml_text[last:])
    return "".join(out)


def build_replacements(data: dict) -> dict:
    values = build_values(data)
    replacements = {}

    simple_keys = [
        "agreement_no",
        "sign_place",
        "party_a_name_en",
        "party_a_name_ru",
        "party_a_address",
        "party_a_address_ru",
        "party_a_representative",
        "party_a_representative_ru",
        "party_b_name_en",
        "party_b_name_ru",
        "party_b_address",
        "party_b_address_ru",
        "party_b_representative",
        "party_b_representative_ru",
        "party_c_name_cn",
        "party_c_name_ru",
        "party_c_address",
        "party_c_address_ru",
        "party_c_representative",
        "party_c_representative_ru",
        "trade_contract_no_with_symbol",
        "trade_contract_date_en",
        "trade_contract_date_ru",
        "appendix_no",
        "goods_cn",
        "goods_ru",
        "party_b_bank",
        "party_b_bank_ru",
        "party_b_swift",
        "party_b_bank_address",
        "party_b_account_name",
        "party_b_account_no",
    ]
    for key in simple_keys:
        if DEFAULTS.get(key) and values.get(key):
            replacements[DEFAULTS[key]] = values[key]

    if values.get("trade_contract_no"):
        replacements["GOF-04/06-26-CHN"] = values["trade_contract_no"]
        replacements["№ GOF-04/06-26-CHN"] = values.get("trade_contract_no_with_symbol") or "№ " + values["trade_contract_no"]
    if values.get("agent_fee_percent"):
        replacements["的 1 %"] = f"的 {values['agent_fee_percent']} %"
        replacements["составит 1 %"] = f"составит {values['agent_fee_percent']} %"
    if values.get("sign_date"):
        replacements["签订日期 Дата подписания:"] = f"签订日期 Дата подписания: {values['sign_date']}"
    return replacements


def build_values(data: dict) -> dict:
    return {**DEFAULTS, **{k: v.strip() for k, v in data.items() if isinstance(v, str) and v.strip()}}


def build_fillable_values(data: dict) -> set[str]:
    values = build_values(data)
    keys = {
        "agreement_no",
        "sign_date",
        "sign_place",
        "party_a_name_en",
        "party_a_name_ru",
        "party_a_address",
        "party_a_address_ru",
        "party_a_representative",
        "party_a_representative_ru",
        "party_b_name_en",
        "party_b_name_ru",
        "party_b_address",
        "party_b_address_ru",
        "party_b_representative",
        "party_b_representative_ru",
        "party_c_name_cn",
        "party_c_name_ru",
        "party_c_address",
        "party_c_address_ru",
        "party_c_representative",
        "party_c_representative_ru",
        "trade_contract_no",
        "trade_contract_no_with_symbol",
        "trade_contract_date_en",
        "trade_contract_date_ru",
        "appendix_no",
        "goods_cn",
        "goods_ru",
        "agent_fee_percent",
        "party_b_bank",
        "party_b_bank_ru",
        "party_b_swift",
        "party_b_bank_address",
        "party_b_account_name",
        "party_b_account_no",
    }
    return {values[key] for key in keys if values.get(key)}


def process_docx(input_path: Path, output_path: Path, data: dict) -> None:
    # 导入通用后处理模块
    import sys
    _lib_dir = Path(__file__).resolve().parent.parent
    if str(_lib_dir) not in sys.path:
        sys.path.insert(0, str(_lib_dir))
    from docx_processor import process_document_xml, replace_text_in_xml

    replacements = build_replacements(data)
    fillable_values = build_fillable_values(data)
    with zipfile.ZipFile(input_path, "r") as zin, zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            if item.filename == "word/document.xml":
                # 仅清除高亮、批注、底纹
                content = process_document_xml(content)
                # 文本替换（只替换填写占位符）
                content = replace_text_in_xml(content, replacements)
                # 下划线标记可填写值
                content = _underline_fillable(content, fillable_values)
                # 删除丙方签章区中的公司名称（保留盖章行和签字行）
                content = _remove_party_c_signature(content)
            zout.writestr(item, content)


def _remove_party_c_signature(content: bytes) -> bytes:
    """删除丙方签章区中盖章行和签字行之间的公司名称。
    
    只处理签章区（盖章行和签字行之间的段落），不处理文档正文中
    的丙方信息段落。
    """
    from lxml import etree as lxml_etree
    ns = {"w": W_NS}
    root = lxml_etree.fromstring(content)
    
    # 在文档正文中查找签章区段落（不在表格中）
    body = root.find(f"{{{W_NS}}}body")
    if body is not None:
        paragraphs = body.findall(f"{{{W_NS}}}p")
        for i, p in enumerate(paragraphs):
            text = "".join(t.text or "" for t in p.xpath(".//w:t", namespaces=ns))
            # 匹配签章区的"丙方（盖章）"行（包含盖章/печать关键词）
            if ("丙方（盖章）" in text or "Сторона С (Печать)" in text or "Сторона C (Печать)" in text):
                # 删除该段落和"代表人（签字）"段落之间的所有段落
                j = i + 1
                while j < len(paragraphs):
                    next_text = "".join(t.text or "" for t in paragraphs[j].xpath(".//w:t", namespaces=ns))
                    if "代表人（签字）" in next_text or "Представитель (Подпись)" in next_text:
                        break
                    to_remove = paragraphs[j]
                    j += 1
                    parent = to_remove.getparent()
                    if parent is not None:
                        parent.remove(to_remove)
    
    return lxml_etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def _underline_fillable(content: bytes, fillable_values: set) -> bytes:
    """为可填写值添加下划线。"""
    from lxml import etree as lxml_etree
    ns = {"w": W_NS}
    root = lxml_etree.fromstring(content)
    values = sorted((v for v in fillable_values if v and len(v.strip()) >= 2), key=len, reverse=True)
    if not values:
        return content
    for run in root.xpath("//w:r", namespaces=ns):
        text = "".join(t.text or "" for t in run.xpath(".//w:t", namespaces=ns))
        if not text or not text.strip():
            continue
        if any(value in text for value in values):
            rpr = run.find(f"{{{W_NS}}}rPr")
            if rpr is None:
                rpr = lxml_etree.SubElement(run, f"{{{W_NS}}}rPr")
            u = rpr.find(f"{{{W_NS}}}u")
            if u is None:
                u = lxml_etree.SubElement(rpr, f"{{{W_NS}}}u")
            u.set(f"{{{W_NS}}}val", "single")
    return lxml_etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def find_office_binary() -> str:
    candidates = [
        "soffice",
        "libreoffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if shutil.which(candidate) or Path(candidate).exists():
            return candidate
    return "soffice"


def convert_to_pdf(docx_path: Path, out_dir: Path) -> Path:
    pdf_path = out_dir / (docx_path.stem + ".pdf")

    # 使用 Word COM 转换，加进程锁避免多线程冲突
    with _word_lock:
        try:
            import pythoncom
            pythoncom.CoInitialize()
            try:
                import win32com.client
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                doc = word.Documents.Open(str(docx_path), False, True)
                # 接受所有修订，确保输出与"最终版"一致
                doc.AcceptAllRevisions()
                doc.ExportAsFixedFormat(str(pdf_path), 17)
                doc.Close(False)
                word.Quit()
                if pdf_path.exists():
                    return pdf_path
            finally:
                pythoncom.CoUninitialize()
        except Exception:
            pass

    # Word COM 失败时回退到 PowerShell 脚本
    ps_path = out_dir / "convert_word_to_pdf.ps1"
    docx_literal = str(docx_path).replace("'", "''")
    pdf_literal = str(pdf_path).replace("'", "''")
    ps_path.write_text(
        "$word = $null\n"
        "$doc = $null\n"
        "try {\n"
        "  $word = New-Object -ComObject Word.Application\n"
        "  $word.Visible = $false\n"
        f"  $doc = $word.Documents.Open('{docx_literal}', $false, $true)\n"
        f"  $doc.ExportAsFixedFormat('{pdf_literal}', 17)\n"
        "  exit 0\n"
        "} catch {\n"
        "  Write-Error $_.Exception.Message\n"
        "  exit 1\n"
        "} finally {\n"
        "  if ($doc) { $doc.Close($false) }\n"
        "  if ($word) { $word.Quit() }\n"
        "}\n",
        encoding="utf-8",
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps_path)],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=90,
        )
        if result.returncode == 0 and pdf_path.exists():
            return pdf_path
    except Exception:
        pass

    raise RuntimeError("Word COM 和 PowerShell 均无法转换 PDF")


def generate_pdf(data: dict) -> tuple[bytes, str]:
    if not TEMPLATE_PATH.exists():
        raise RuntimeError("缺少 Word 模板文件 templates/agreement-template.docx")
    with tempfile.TemporaryDirectory(prefix="contract_pdf_") as tmp:
        tmp_dir = Path(tmp)
        docx_path = tmp_dir / "委托代理协议.docx"
        process_docx(TEMPLATE_PATH, docx_path, data)
        pdf_path = convert_to_pdf(docx_path, tmp_dir)
        pdf_bytes = pdf_path.read_bytes()
    safe_no = re.sub(r"[^\w\-一-龥]+", "_", data.get("agreement_no") or DEFAULTS["agreement_no"])
    return pdf_bytes, f"委托代理协议_{safe_no}.pdf"


def load_file(path: Path) -> bytes:
    return path.read_bytes()


class ContractHandler(BaseHTTPRequestHandler):
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


def self_test():
    sample = {**DEFAULTS, "sign_date": time.strftime("%Y-%m-%d")}
    pdf_bytes, filename = generate_pdf(sample)
    out = BASE_DIR / filename
    out.write_bytes(pdf_bytes)
    print(f"自检完成：{out}")


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        self_test()
    else:
        print(f"合同生成网站已启动：http://{HOST}:{PORT}")
        print("保持此窗口运行，在浏览器打开上面的地址。按 Ctrl+C 停止。")
        ThreadingHTTPServer((HOST, PORT), ContractHandler).serve_forever()
