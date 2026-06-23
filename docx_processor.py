"""通用 DOCX XML 后处理模块。

仅执行必要操作：
1. 移除高亮标签（<w:highlight>）— 黄色提示标记
2. 删除批注标签（comment）— 批注气泡
3. 移除可见底纹（<w:shd>）— 彩色背景
4. 不移除修订标记（Word原生处理）
5. 不修改字体（保留模板原始字体）
6. 不修改间距（保留模板原始布局）
"""

from lxml import etree as lxml_etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def process_document_xml(content: bytes) -> bytes:
    """对 word/document.xml 执行后处理。"""
    parser = lxml_etree.XMLParser(remove_blank_text=False)
    root = lxml_etree.fromstring(content, parser)

    # 移除高亮标签
    for el in root.xpath("//w:rPr/w:highlight", namespaces=NS):
        parent = el.getparent()
        if parent is not None:
            parent.remove(el)

    # 删除批注相关元素
    for xpath in [
        "//w:commentRangeStart", "//w:commentRangeEnd",
        "//w:commentReference",
    ]:
        for el in root.xpath(xpath, namespaces=NS):
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)

    # 移除可见底纹（非 auto 的 shd）
    for el in root.xpath("//w:shd", namespaces=NS):
        fill = el.get(f"{{{W_NS}}}fill", "auto")
        if fill and fill.lower() != "auto" and fill.lower() != "ffffff":
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)

    return lxml_etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def replace_text_in_xml(content: bytes, replacements: dict) -> bytes:
    """在 XML 中执行文本替换。"""
    parser = lxml_etree.XMLParser(remove_blank_text=False)
    root = lxml_etree.fromstring(content, parser)

    for t_el in root.xpath("//w:t", namespaces=NS):
        if t_el.text:
            for old_val, new_val in replacements.items():
                if old_val and old_val != new_val:
                    t_el.text = t_el.text.replace(old_val, new_val)

    return lxml_etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)