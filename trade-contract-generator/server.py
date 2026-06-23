import copy
import html
import json
import os
import re
import shutil
import subprocess
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

# ---------------------------------------------------------------------------
# Word COM 实例复用（避免每次启动 Word 的 100+ 秒开销）
# ---------------------------------------------------------------------------
_word_app = None
_word_keepalive_doc = None

def _is_word_running():
    """检查 WINWORD 进程是否还在运行。"""
    try:
        import subprocess
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE"],
            capture_output=True, text=True, encoding="utf-8", errors="replace"
        )
        return "WINWORD" in result.stdout
    except Exception:
        return False

def _get_word_app():
    global _word_app, _word_keepalive_doc
    if _word_app is None or not _is_word_running():
        import win32com.client
        # 使用 DispatchEx 创建独立实例，避免 Word 自动退出
        _word_app = win32com.client.DispatchEx("Word.Application")
        _word_app.Visible = False
        _word_app.DisplayAlerts = 0
        # 打开一个空白文档保持 Word 进程活跃
        _word_keepalive_doc = _word_app.Documents.Add()
    return _word_app

def _quit_word_app():
    global _word_app, _word_keepalive_doc
    if _word_keepalive_doc:
        try:
            _word_keepalive_doc.Close(False)
        except Exception:
            pass
        _word_keepalive_doc = None
    if _word_app:
        try:
            _word_app.Quit()
        except Exception:
            pass
        _word_app = None


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "templates" / "trade-contract-template.docx"
HOST = "127.0.0.1"
PORT = 8089


DEFAULTS = {
    "contract_no": "GOF-04/06-26-CHN",
    "contract_date_ru": "«04» июня 2026 г.",
    "contract_date_en": "04 of June 2026",
    "appendix_no": "1",
    "appendix_date_ru": "«04» июня 2026 г.",
    "appendix_date_en": "4th of June 2026",
    "buyer_name_ru": "ООО «ТД ЭЙРФЛОТ ТЕХНИКС»",
    "buyer_name_en": "AIRFLOT TECHNICS TH Ltd.",
    "buyer_full_name_ru": "Общество с ограниченной ответственностью «Торговый Дом ЭЙРФЛОТ ТЕХНИКС»",
    "buyer_address_ru": "127018, Россия, г. Москва, улица Полковая, дом 3, строение 5, этаж Т, пом. V, ком. 15.",
    "buyer_address_en": "Polkovaya Street 3, bld.5, floor T, section V, office 15, Moscow, Russia. Postcode: 127018.",
    "buyer_tel": "+7 (495) 221-8026",
    "buyer_email": "info@aftproject.ru",
    "buyer_notice_tel": "+7 (495) 221-80-26",
    "buyer_notice_email": "info@ftproject.ru",
    "buyer_postal_ru": "ООО «ТД ЭЙРФЛОТ ТЕХНИКС», 127018, г. Москва, а/я 60.",
    "buyer_postal_en": "AIRFLOT TECHNICS TH Ltd., P.O. Box 60, Moscow, Russia, 127018.",
    "buyer_bank_ru": "Банк ВТБ (ПАО) (Центральный филиал, Москва)",
    "buyer_bank_en": "VTB Bank (PJSC) (TSENTRALNYI BRANCH, MOSCOW)",
    "buyer_bank_address_ru": "119121, 2-й Неопалимовский переулок, д.10, Москва, Россия",
    "buyer_bank_address_en": "119121, 2nd Neopalimovskiy pereulok, 10, Moscow, Russia",
    "buyer_account_no": "40702156824790000036",
    "buyer_swift": "VTBRRUM2MS2",
    "buyer_corr_account": "30101810145250000411",
    "buyer_bic": "044525411",
    "buyer_representative_ru": "Журавлева Сергея Владимировича",
    "buyer_representative_en": "Sergey ZHURAVLEV",
    "buyer_signature_ru": "Журавлёв С.В.",
    "buyer_signature_en": "Sergey Zhuravlev",
    "seller_name": "GUANGDONG OSHUJIAN FURNITURE CO.,LTD",
    "seller_address_ru": "Sicun  Industrial, Heqing, Jiujiang Nanhai Foshan, Guangdong, China",
    "seller_address_en": "Sicun  Industrial, Heqing, Jiujiang Nanhai, Foshan, Guangdong, China",
    "seller_tel": "757-81869861",
    "seller_email": "sale8@oshujian.com",
    "seller_bank_ru": "ВТБ Банк (ПАО) Шанхайский филиал",
    "seller_bank_en": "VTB Bank (PJSC) Shanghai Branch",
    "seller_bank_address_ru": "Китай, г. Шанхай, Сильвер-Сентрал Тауэр, 501, ул. Иньчэн Лу (Средняя), офисы 2503-2505",
    "seller_bank_address_en": "Office 2503-2505, Shanghai Tower, 501 Middle Yincheng Road, Shanghai, China, 20012",
    "seller_account_no": "40807156200610027979",
    "seller_swift": "VTBRCNSH",
    "seller_beneficiary": "GUANGDONG OSHUJIAN FURNITURE CO.,LTD",
    "seller_representative_ru": "Сюй Юань",
    "seller_representative_en": "Xu Yuan",
    "goods_name_ru": "",
    "goods_name_en": "",
    "goods_qty": "",
    "goods_unit_price": "",
    "goods_total": "",
    "package_includes_ru": "",
    "package_includes_en": "",
    "total_price_ru": "",
    "total_price_en": "",
    "delivery_terms_ru": "",
    "delivery_terms_en": "",
    "delivery_time_ru": "Поставка осуществляется в срок не позднее 45 (сорок пять) календарных дней с даты оплаты аванса согласно п.6.1. настоящего Приложения без учета национальных праздников. Срок поставки Товара продлевается на срок задержки исполнения Покупателем своих обязательств по оплате согласно п.6.1. настоящего Приложения.",
    "delivery_time_en": "Goods should be delivered up in time not later than 45 (forty-five) calendar days from the date of payment of the advance payment according to the paragraph 6.1. of this Appendix excluding national holidays. Delivery of Goods shall be extended for a period of delay of the Buyer's payment obligations under the paragraph 6.1 of the present Appendix.",
    # —— 附件表单（"开始 / Form start"之后的空白填写位）——
    "form_appendix_no": "",
    "form_appendix_day_ru": "",
    "form_appendix_month_ru": "",
    "form_appendix_year_ru": "",
    "form_contract_no": "",
    "form_contract_day_ru": "",
    "form_contract_month_ru": "",
    "form_contract_year_ru": "",
    "form_appendix_date_en": "",
    "form_appendix_year_en": "",
    "form_contract_no_en": "",
    "form_contract_date_en": "",
    "form_contract_year_en": "",
    "form_seller_name_ru": "",
    "form_seller_country_ru": "",
    "form_seller_director_ru": "",
    "form_seller_name_en": "",
    "form_seller_country_en": "",
    "form_seller_director_en": "",
    # —— 第 4 节交付条件、第 5 节交付时间的空白填写行 ——
    "form_delivery_terms_ru": "",
    "form_delivery_terms_en": "",
    "form_delivery_time_ru": "",
    "form_delivery_time_en": "",
    # —— 第 6 / 7 / 8 节自由填写段 ——
    "form_payment_terms_ru": "",
    "form_payment_terms_en": "",
    "form_shipping_docs_ru": "",
    "form_shipping_docs_en": "",
    "form_acceptance_ru": "",
    "form_acceptance_en": "",
    # —— 双方签字位 ——
    "form_buyer_signer_ru": "",
    "form_buyer_signature_ru": "",
    "form_seller_signer_ru": "",
    "form_seller_signature_ru": "",
    # —— 与新附件2一致的付款/通知拆分字段 ——
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
    "contract_no": "合同编号",
    "contract_date_ru": "合同日期（俄文）",
    "contract_date_en": "合同日期（英文）",
    "appendix_no": "附件编号",
    "appendix_date_ru": "附件日期（俄文）",
    "appendix_date_en": "附件日期（英文）",
    "buyer_name_ru": "买方名称（俄文）",
    "buyer_name_en": "买方名称（英文）",
    "buyer_full_name_ru": "买方完整法定名称（俄文）",
    "buyer_address_ru": "买方地址（俄文）",
    "buyer_address_en": "买方地址（英文）",
    "buyer_tel": "买方电话",
    "buyer_email": "买方邮箱",
    "buyer_notice_tel": "装运通知电话/传真",
    "buyer_notice_email": "装运通知邮箱",
    "buyer_postal_ru": "买方邮寄地址（俄文）",
    "buyer_postal_en": "买方邮寄地址（英文）",
    "buyer_bank_ru": "买方银行（俄文）",
    "buyer_bank_en": "买方银行（英文）",
    "buyer_bank_address_ru": "买方银行地址（俄文）",
    "buyer_bank_address_en": "买方银行地址（英文）",
    "buyer_account_no": "买方账户号",
    "buyer_swift": "买方 SWIFT",
    "buyer_corr_account": "买方银行对应账户",
    "buyer_bic": "买方银行 BIC",
    "buyer_representative_ru": "买方代表（俄文）",
    "buyer_representative_en": "买方代表（英文）",
    "buyer_signature_ru": "买方签字名（俄文）",
    "buyer_signature_en": "买方签字名（英文）",
    "seller_name": "卖方名称",
    "seller_address_ru": "卖方地址（俄文）",
    "seller_address_en": "卖方地址（英文）",
    "seller_tel": "卖方电话",
    "seller_email": "卖方邮箱",
    "seller_bank_ru": "卖方银行（俄文）",
    "seller_bank_en": "卖方银行（英文）",
    "seller_bank_address_ru": "卖方银行地址（俄文）",
    "seller_bank_address_en": "卖方银行地址（英文）",
    "seller_account_no": "卖方账户号",
    "seller_swift": "卖方 SWIFT",
    "seller_beneficiary": "卖方收款人名称",
    "seller_representative_ru": "卖方代表（俄文）",
    "seller_representative_en": "卖方代表（英文）",
    "goods_name_ru": "商品名称（俄文）",
    "goods_name_en": "商品名称（英文）",
    "goods_qty": "数量",
    "goods_unit_price": "单价（CNY）",
    "goods_total": "总价（CNY）",
    "package_includes_ru": "供货清单（俄文）",
    "package_includes_en": "供货清单（英文）",
    "total_price_ru": "总价大写（俄文）",
    "total_price_en": "总价大写（英文）",
    "delivery_terms_ru": "交付条件（俄文）",
    "delivery_terms_en": "交付条件（英文）",
    "delivery_time_ru": "交付时间（俄文）",
    "delivery_time_en": "交付时间（英文）",
    "form_appendix_no": "附件表单：附件编号",
    "form_appendix_day_ru": "附件表单：附件日（俄文）",
    "form_appendix_month_ru": "附件表单：附件月（俄文）",
    "form_appendix_year_ru": "附件表单：附件年（俄文）",
    "form_contract_no": "附件表单：合同编号",
    "form_contract_day_ru": "附件表单：合同日（俄文）",
    "form_contract_month_ru": "附件表单：合同月（俄文）",
    "form_contract_year_ru": "附件表单：合同年（俄文）",
    "form_appendix_date_en": "附件表单：附件日期（英文）",
    "form_appendix_year_en": "附件表单：附件年（英文）",
    "form_contract_no_en": "附件表单：合同编号（英文）",
    "form_contract_date_en": "附件表单：合同日期（英文）",
    "form_contract_year_en": "附件表单：合同年（英文）",
    "form_seller_name_ru": "附件表单：卖方名称（俄文）",
    "form_seller_country_ru": "附件表单：卖方国家（俄文）",
    "form_seller_director_ru": "附件表单：卖方法人（俄文）",
    "form_seller_name_en": "附件表单：卖方名称（英文）",
    "form_seller_country_en": "附件表单：卖方国家（英文）",
    "form_seller_director_en": "附件表单：卖方法人（英文）",
    "form_delivery_terms_ru": "附件表单：交付条件（俄文）",
    "form_delivery_terms_en": "附件表单：交付条件（英文）",
    "form_delivery_time_ru": "附件表单：交付时间（俄文）",
    "form_delivery_time_en": "附件表单：交付时间（英文）",
    "form_payment_terms_ru": "附件表单：付款条件（俄文）",
    "form_payment_terms_en": "附件表单：付款条件（英文）",
    "form_shipping_docs_ru": "附件表单：单据清单（俄文）",
    "form_shipping_docs_en": "附件表单：单据清单（英文）",
    "form_acceptance_ru": "附件表单：技术验收（俄文）",
    "form_acceptance_en": "附件表单：技术验收（英文）",
    "form_buyer_signer_ru": "附件表单：买方签字代表",
    "form_buyer_signature_ru": "附件表单：买方签字名",
    "form_seller_signer_ru": "附件表单：卖方签字代表",
    "form_seller_signature_ru": "附件表单：卖方签字名",
    "payment_advance_percent": "附件表单：预付款比例",
    "payment_advance_amount": "附件表单：预付款金额（数字）",
    "payment_advance_amount_ru": "附件表单：预付款金额大写（俄文）",
    "payment_advance_amount_en": "附件表单：预付款金额大写（英文）",
    "payment_balance_percent": "附件表单：尾款比例",
    "payment_balance_amount": "附件表单：尾款金额（数字）",
    "payment_balance_amount_ru": "附件表单：尾款金额大写（俄文）",
    "payment_balance_amount_en": "附件表单：尾款金额大写（英文）",
    "payment_bank_days": "附件表单：预付款支付期限（数字）",
    "payment_bank_days_ru": "附件表单：预付款支付期限大写（俄文）",
    "payment_bank_days_en": "附件表单：预付款支付期限大写（英文）",
    "delivery_notice_days": "附件表单：发货前通知天数（数字）",
    "delivery_notice_days_ru": "附件表单：发货前通知天数大写（俄文）",
    "delivery_notice_days_en": "附件表单：发货前通知天数大写（英文）",
}


def seller_name_with_comma(values: dict) -> str:
    name = values.get("seller_name") or DEFAULTS["seller_name"]
    return name.rstrip(" ,") + ","


def company_buyer_en_with_comma(values: dict) -> str:
    name = values.get("buyer_name_en") or DEFAULTS["buyer_name_en"]
    return "Company " + name.rstrip(" ,") + ","


def highlighted_replacements(values: dict) -> dict[str, str]:
    """黄色标注字段映射：用户标黄的内容全部视为可填写字段。"""
    return {
        "GOF-04/06-26-CHN": values["contract_no"],
        "«04» июня 2026 г.": values["contract_date_ru"],
        "June 4, 2026": values["contract_date_en"],
        "04 of June 2026": values["contract_date_en"],
        "1 от «04» июня 2026 г.": f'{values["appendix_no"]} от {values["appendix_date_ru"]}',
        "1 dtd. 4th of June 2026": f'{values["appendix_no"]} dtd. {values["appendix_date_en"]}',
        "dated 04 of June 2026": f'dated {values["contract_date_en"]}',
        "GUANGDONG OSHUJIAN FURNITURE CO., LTD,": seller_name_with_comma(values),
        "GUANGDONG OSHUJIAN FURNITURE CO.,LTD,": seller_name_with_comma(values),
        "GUANGDONG OSHUJIAN FURNITURE CO., LTD.": values["seller_name"].rstrip(" .") + ".",
        "GUANGDONG OSHUJIAN FURNITURE CO.,LTD": values["seller_name"],
        "Guangdong OSHUJIAN Furniture Co.,Ltd": values.get("seller_beneficiary") or values["seller_name"],
        "Сюй Юань": values["seller_representative_ru"],
        "Xu Yuan": values["seller_representative_en"],
        "Общество с ограниченной ответственностью «Торговый Дом ЭЙРФЛОТ ТЕХНИКС»": values["buyer_full_name_ru"],
        "ООО «ТД\xa0ЭЙРФЛОТ ТЕХНИКС»": values["buyer_name_ru"],
        "ООО «ТД ЭЙРФЛОТ ТЕХНИКС»": values["buyer_name_ru"],
        "AIRFLOT TECHNICS TH LTD": values["buyer_name_en"].upper(),
        "AIRFLOT TECHNICS TH Ltd.": values["buyer_name_en"],
        "Company AIRFLOT TECHNICS TH Ltd.,": company_buyer_en_with_comma(values),
        "Журавлева Сергея Владимировича": values["buyer_representative_ru"],
        "Sergey ZHURAVLEV": values["buyer_representative_en"],
        "Sergey Zhuravlev": values["buyer_signature_en"],
        "Журавлёв С.В.": values["buyer_signature_ru"],
        "info@aftproject.ru": values["buyer_email"],
        "info@ftproject.ru": values["buyer_notice_email"],
        "+7 (495) 221-80-26": values["buyer_notice_tel"],
        "+7\xa0(495) 221-8026": values["buyer_tel"],
        "+7 (495) 221-8026": values["buyer_tel"],
        "127018, Россия, г. Москва, улица Полковая, дом 3, строение 5, этаж Т, пом. V, ком. 15.": values["buyer_address_ru"],
        "Polkovaya Street 3, bld.5, floor T, section V, office 15, Moscow, Russia. Postcode: 127018.": values["buyer_address_en"],
        "ООО «ТД ЭЙРФЛОТ ТЕХНИКС», 127018, г. Москва, а/я 60.": values["buyer_postal_ru"],
        "AIRFLOT TECHNICS TH Ltd., P.O. Box 60, Moscow, Russia, 127018.": values["buyer_postal_en"],
        "Банк ВТБ (ПАО) (Центральный филиал, Москва)": values["buyer_bank_ru"],
        "VTB Bank (PJSC) (TSENTRALNYI BRANCH, MOSCOW)": values["buyer_bank_en"],
        "119121, 2-й Неопалимовский переулок, д.10, Москва, Россия": values["buyer_bank_address_ru"],
        "119121, 2nd Neopalimovskiy pereulok, 10, Moscow, Russia": values["buyer_bank_address_en"],
        "40702156824790000036": values["buyer_account_no"],
        "VTBRRUM2MS2": values["buyer_swift"],
        "30101810145250000411": values["buyer_corr_account"],
        "044525411": values["buyer_bic"],
        "Sicun  Industrial, Heqing, Jiujiang Nanhai Foshan, Guangdong, China": values["seller_address_ru"],
        "Sicun  Industrial, Heqing, Jiujiang Nanhai, Foshan, Guangdong, China": values["seller_address_en"],
        "757-81869861": values["seller_tel"],
        "sale8@oshujian.com": values["seller_email"],
        "ВТБ Банк (ПАО) Шанхайский филиал": values["seller_bank_ru"],
        "VTB Bank (PJSC) Shanghai Branch": values["seller_bank_en"],
        "Китай, г. Шанхай, Сильвер-Сентрал Тауэр, 501, ул. Иньчэн Лу (Средняя), офисы 2503-2505": values["seller_bank_address_ru"],
        "Office 2503-2505, Shanghai Tower, 501 Middle Yincheng Road, Shanghai, China, 20012": values["seller_bank_address_en"],
        "40807156200610027979": values["seller_account_no"],
        "VTBRCNSH": values["seller_swift"],
    }


HIGHLIGHT_FILLABLE_KEYS = {
    "contract_no", "contract_date_ru", "contract_date_en", "appendix_no", "appendix_date_ru", "appendix_date_en",
    "buyer_name_ru", "buyer_name_en", "buyer_full_name_ru", "buyer_address_ru", "buyer_address_en", "buyer_tel",
    "buyer_email", "buyer_notice_tel", "buyer_notice_email", "buyer_postal_ru", "buyer_postal_en",
    "buyer_bank_ru", "buyer_bank_en", "buyer_bank_address_ru", "buyer_bank_address_en", "buyer_account_no",
    "buyer_swift", "buyer_corr_account", "buyer_bic", "buyer_representative_ru", "buyer_representative_en",
    "buyer_signature_ru", "buyer_signature_en", "seller_name", "seller_address_ru", "seller_address_en",
    "seller_tel", "seller_email", "seller_bank_ru", "seller_bank_en", "seller_bank_address_ru",
    "seller_bank_address_en", "seller_account_no", "seller_swift", "seller_beneficiary",
    "seller_representative_ru", "seller_representative_en", "goods_name_ru", "goods_name_en", "goods_qty",
    "goods_unit_price", "goods_total", "package_includes_ru", "package_includes_en", "total_price_ru",
    "total_price_en", "delivery_terms_ru", "delivery_terms_en", "delivery_time_ru", "delivery_time_en",
    "form_appendix_no", "form_appendix_day_ru", "form_appendix_month_ru", "form_appendix_year_ru",
    "form_contract_no", "form_contract_day_ru", "form_contract_month_ru", "form_contract_year_ru",
    "form_appendix_date_en", "form_appendix_year_en", "form_contract_no_en", "form_contract_date_en",
    "form_contract_year_en", "form_seller_name_ru", "form_seller_country_ru", "form_seller_director_ru",
    "form_seller_name_en", "form_seller_country_en", "form_seller_director_en", "form_delivery_terms_ru",
    "form_delivery_terms_en", "form_delivery_time_ru", "form_delivery_time_en", "form_payment_terms_ru",
    "form_payment_terms_en", "form_shipping_docs_ru", "form_shipping_docs_en", "form_acceptance_ru",
    "form_acceptance_en", "form_buyer_signer_ru", "form_buyer_signature_ru", "form_seller_signer_ru",
    "form_seller_signature_ru", "payment_advance_percent", "payment_advance_amount", "payment_advance_amount_ru",
    "payment_advance_amount_en", "payment_balance_percent", "payment_balance_amount", "payment_balance_amount_ru",
    "payment_balance_amount_en", "payment_bank_days", "payment_bank_days_ru", "payment_bank_days_en",
    "delivery_notice_days", "delivery_notice_days_ru", "delivery_notice_days_en",
}


def clean_revision_markup(xml_text: str) -> str:
    # 先检查标签是否平衡，不平衡则跳过处理（避免破坏XML结构）
    ins_open = len(re.findall(r"<w:ins\b", xml_text))
    ins_close = len(re.findall(r"</w:ins>", xml_text))
    del_open = len(re.findall(r"<w:del\b", xml_text))
    del_close = len(re.findall(r"</w:del>", xml_text))
    if ins_open != ins_close or del_open != del_close:
        # 标签不平衡，跳过修订标记清理（Word可以正常处理）
        return xml_text
    xml_text = re.sub(r"<w:commentRangeStart\b[^>]*/>", "", xml_text)
    xml_text = re.sub(r"<w:commentRangeEnd\b[^>]*/>", "", xml_text)
    xml_text = re.sub(r"<w:r\b[^>]*>.*?<w:commentReference\b[^>]*/>.*?</w:r>", "", xml_text, flags=re.S)
    for _ in range(20):
        new_xml = re.sub(r"<w:ins\b[^>]*>(.*?)</w:ins>", r"\1", xml_text, flags=re.S)
        new_xml = re.sub(r"<w:del\b[^>]*>.*?</w:del>", "", new_xml, flags=re.S)
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
            tbl_layout = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(tbl_pr, f"{{{W_NS}}}tblLayout")
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
        rpr_default = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(doc_defaults, f"{{{W_NS}}}rPrDefault")
    rpr = rpr_default.find(f"{{{W_NS}}}rPr")
    if rpr is None:
        rpr = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(rpr_default, f"{{{W_NS}}}rPr")
    rfonts = rpr.find(f"{{{W_NS}}}rFonts")
    if rfonts is None:
        rfonts = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(rpr, f"{{{W_NS}}}rFonts")
    for key in W_RFONTS_KEYS:
        rfonts.set(f"{{{W_NS}}}{key}", LATIN_FONT)
    rfonts.set(f"{{{W_NS}}}eastAsia", EASTASIA_FONT)


def ensure_child(parent: ET.Element, tag_name: str) -> ET.Element:
    child = parent.find(f"{{{W_NS}}}{tag_name}")
    if child is None:
        child = (LXML_ET.Element if LXML_ET else ET.Element)(f"{{{W_NS}}}{tag_name}")
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
                new_row = (LXML_ET.Element if LXML_ET else ET.Element)(row.tag, dict(row.attrib))
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
                # 段落原本未指定对齐方式时，单元格内默认左对齐。
                jc = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(ppr, f"{{{W_NS}}}jc")
                jc.set(f"{{{W_NS}}}val", "left")
            # 保留模板中已显式声明的 center / right / both 等对齐方式，不再强制改写为 left。
            spacing = ppr.find(f"{{{W_NS}}}spacing")
            if spacing is None:
                spacing = (LXML_ET.Element if LXML_ET else ET.Element)(f"{{{W_NS}}}spacing")
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


def remove_yellow_highlights(root: ET.Element) -> None:
    """清除用户用于标记填写位的黄色底色，导出 PDF 只保留下划线。"""
    for rpr in root.iter(f"{{{W_NS}}}rPr"):
        for child in list(rpr):
            if local_name(child.tag) == "highlight":
                rpr.remove(child)
            elif local_name(child.tag) == "shd":
                fill = (child.get(f"{{{W_NS}}}fill") or "").upper()
                if fill in {"FFFF00", "FFF200", "FFFF99", "FFE699"}:
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
                underline = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(rpr, f"{{{W_NS}}}u")
            underline.set(f"{{{W_NS}}}val", "single")
            color = rpr.find(f"{{{W_NS}}}color")
            if color is None:
                color = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(rpr, f"{{{W_NS}}}color")
            color.set(f"{{{W_NS}}}val", "000000")


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
            if start < 0:
                break
            end = start + len(old)
            pos = 0
            first_idx = last_idx = None
            first_offset = last_offset = 0
            for idx, text in enumerate(parts):
                tlen = len(text)
                next_pos = pos + tlen
                # 关键：first_idx 必须落在"实际包含 start 字符"的节点上。
                # 若 start == next_pos，说明匹配从下一个节点开始，不应归到当前节点；
                # 同样跳过 tlen == 0 的空节点。
                if first_idx is None and tlen > 0 and start < next_pos:
                    first_idx = idx
                    first_offset = max(0, start - pos)
                if first_idx is not None and end <= next_pos and tlen > 0:
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


def replace_text_per_paragraph(root: ET.Element, replacements: dict) -> None:
    """按段落作用域执行替换，避免跨段落/跨单元格"借文本"导致排版错位。"""
    for paragraph in root.iter(f"{{{W_NS}}}p"):
        text_elements = [el for el in paragraph.iter() if local_name(el.tag) == "t"]
        if not text_elements:
            continue
        replace_text_elements(text_elements, replacements)



def set_paragraph_text(paragraph: ET.Element, value: str) -> None:
    texts = list(paragraph.iter(f"{{{W_NS}}}t"))
    if not texts:
        run = (LXML_ET.SubElement if LXML_ET else ET.SubElement)(paragraph, f"{{{W_NS}}}r")
        texts = [(LXML_ET.SubElement if LXML_ET else ET.SubElement)(run, f"{{{W_NS}}}t")]
    texts[0].text = value
    for text in texts[1:]:
        text.text = ""
    if value[:1].isspace() or value[-1:].isspace():
        texts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def set_cell_first_paragraph(cell: ET.Element, value: str) -> None:
    paragraphs = cell.findall(f"{{{W_NS}}}p")
    if not paragraphs:
        paragraphs = [(LXML_ET.SubElement if LXML_ET else ET.SubElement)(cell, f"{{{W_NS}}}p")]
    set_paragraph_text(paragraphs[0], value)


def set_cell_paragraph(cell: ET.Element, index: int, value: str) -> None:
    paragraphs = cell.findall(f"{{{W_NS}}}p")
    while len(paragraphs) <= index:
        paragraphs.append((LXML_ET.SubElement if LXML_ET else ET.SubElement)(cell, f"{{{W_NS}}}p"))
    set_paragraph_text(paragraphs[index], value)


def parse_goods_items(values: dict) -> list[dict[str, str]]:
    """解析页面提交的动态商品列表，兼容旧的单商品字段。"""
    raw = values.get("goods_items_json", "")
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
                    "goods_name_ru": str(item.get("goods_name_ru", "")).strip(),
                    "goods_name_en": str(item.get("goods_name_en", "")).strip(),
                    "goods_qty": str(item.get("goods_qty", "")).strip(),
                    "goods_unit_price": str(item.get("goods_unit_price", "")).strip(),
                    "goods_total": str(item.get("goods_total", "")).strip(),
                }
                if any(normalized.values()):
                    items.append(normalized)
    if not items:
        legacy = {
            "goods_name_ru": values.get("goods_name_ru", ""),
            "goods_name_en": values.get("goods_name_en", ""),
            "goods_qty": values.get("goods_qty", ""),
            "goods_unit_price": values.get("goods_unit_price", ""),
            "goods_total": values.get("goods_total", ""),
        }
        if any(legacy.values()):
            items.append(legacy)
    return items


def fill_goods_row(row: ET.Element, item: dict[str, str]) -> None:
    cells = row.findall(f"{{{W_NS}}}tc")
    goods_values = [
        item.get("goods_name_ru", ""),
        item.get("goods_qty", ""),
        item.get("goods_unit_price", ""),
        item.get("goods_total", ""),
        item.get("goods_name_en", ""),
        item.get("goods_qty", ""),
        item.get("goods_unit_price", ""),
        item.get("goods_total", ""),
    ]
    for cell, value in zip(cells, goods_values):
        set_cell_first_paragraph(cell, value or "")


FORM_PLACEHOLDER = "____"


def make_payment_terms_ru(values: dict) -> str:
    if not values.get("payment_advance_amount") and not values.get("payment_balance_amount"):
        return ""
    return (
        f"6.1 Аванс в размере {values.get('payment_advance_percent', '')} {values.get('payment_advance_amount', '')} "
        f"({values.get('payment_advance_amount_ru', '')}) юаней КНР 00 фэней Покупатель оплачивает прямым банковским переводом "
        f"на счет Продавца в течение {values.get('payment_bank_days', '')} ({values.get('payment_bank_days_ru', '')}) банковских дней "
        "с даты подписания уполномоченными представителями обеих Сторон настоящего приложения на основании счета Продавца.\n"
        f"6.2 Сумму в размере {values.get('payment_balance_percent', '')} {values.get('payment_balance_amount', '')} "
        f"({values.get('payment_balance_amount_ru', '')}) юаней КНР 00 фэней Покупатель оплачивает прямым банковским переводом "
        "на счет Продавца перед отгрузкой Товара.\n"
        "6.3 Датой исполнения Покупателем обязательств по оплате считается дата поступления денежных средств в полном объеме "
        "на расчетный счет Продавца, указанный в п. 15 Контракта."
    )


def make_payment_terms_en(values: dict) -> str:
    if not values.get("payment_advance_amount") and not values.get("payment_balance_amount"):
        return ""
    return (
        f"6.1 The advance payment in the amount of {values.get('payment_advance_percent', '')} {values.get('payment_advance_amount', '')} "
        f"({values.get('payment_advance_amount_en', '')}) Chinese yuan 00 fen is paid by the Buyer by direct bank transfer "
        f"to the Seller's account within {values.get('payment_bank_days', '')} ({values.get('payment_bank_days_en', '')}) banking days "
        "from the date of signing by authorized representatives of both Parties of this Appendix on the basis of the Seller’s Invoice.\n"
        f"6.2 The amount of {values.get('payment_balance_percent', '')} {values.get('payment_balance_amount', '')} "
        f"({values.get('payment_balance_amount_en', '')}) Chinese yuan 00 fen shall be paid by the Buyer by direct bank transfer "
        "to the Seller’s account before shipment of the Goods.\n"
        "6.3 The date of the Buyer's performance of its payment obligations under this Appendix shall be deemed the date of full receipt "
        "of the funds in the Seller's bank account specified in Clause 15 of the Contract."
    )


def compose_delivery_time_ru(values: dict) -> str:
    base = values.get("form_delivery_time_ru") or values.get("delivery_time_ru") or ""
    if not values.get("delivery_notice_days") and not values.get("buyer_notice_tel") and not values.get("buyer_notice_email"):
        return base
    notice = (
        f"Продавец извещает Покупателя о готовности Товара к отгрузке посредством факсимильной связи "
        f"{values.get('buyer_notice_tel', '')} или электронной почты {values.get('buyer_notice_email', '')}, "
        f"за {values.get('delivery_notice_days', '')} ({values.get('delivery_notice_days_ru', '')}) календарных дней до начала отгрузки."
    )
    if "извещает Покупателя" in base:
        return base
    return "\n".join([part for part in [base, notice] if part.strip()])


def compose_delivery_time_en(values: dict) -> str:
    base = values.get("form_delivery_time_en") or values.get("delivery_time_en") or ""
    if not values.get("delivery_notice_days") and not values.get("buyer_notice_tel") and not values.get("buyer_notice_email"):
        return base
    notice = (
        f"The Seller shall notify the Buyer of the readiness to supply the Goods through facsimile "
        f"{values.get('buyer_notice_tel', '')} or e-mail {values.get('buyer_notice_email', '')}, "
        f"{values.get('delivery_notice_days', '')} calendar days before the beginning of the shipment."
    )
    if "notify the Buyer" in base:
        return base
    return "\n".join([part for part in [base, notice] if part.strip()])


def _set_paragraph_blanks(paragraph: ET.Element, replacements: list[tuple[str, str]]) -> None:
    """按出现顺序把段落中"___"占位逐个替换为给定值，保留原 run 样式。

    replacements 是 (label, value) 列表，label 仅用于调试。"""
    runs = list(paragraph.iter(f"{{{W_NS}}}r"))
    queue = [value for _, value in replacements if value]
    if not queue:
        return
    pi = 0
    for run in runs:
        if pi >= len(queue):
            break
        text_nodes = list(run.iter(f"{{{W_NS}}}t"))
        if not text_nodes:
            continue
        merged = "".join((t.text or "") for t in text_nodes)
        if "_" not in merged:
            continue
        new_text = []
        i = 0
        while i < len(merged) and pi < len(queue):
            if merged[i] == "_":
                # 连续下划线整体替换为一个值
                j = i
                while j < len(merged) and merged[j] == "_":
                    j += 1
                new_text.append(queue[pi])
                pi += 1
                i = j
            else:
                new_text.append(merged[i])
                i += 1
        new_text.append(merged[i:])
        # 把替换结果合并写回首个 t 节点；其它 t 节点清空
        text_nodes[0].text = "".join(new_text)
        if text_nodes[0].text and (text_nodes[0].text[:1].isspace() or text_nodes[0].text[-1:].isspace()):
            text_nodes[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        for tnode in text_nodes[1:]:
            tnode.text = ""


def _find_paragraph(paragraphs: list[ET.Element], keyword: str, start_index: int = 0) -> int:
    for i in range(start_index, len(paragraphs)):
        text = "".join((t.text or "") for t in paragraphs[i].iter(f"{{{W_NS}}}t"))
        if keyword in text:
            return i
    return -1


def apply_form_blanks(root: ET.Element, values: dict) -> None:
    """把模板"Начало формы / Form start"段落里的下划线占位替换为用户填写的值。"""
    paragraphs = list(root.iter(f"{{{W_NS}}}p"))

    # 俄文附件抬头三段（"Приложение №__ ..."）
    idx = _find_paragraph(paragraphs, "Приложение №__")
    if idx >= 0:
        _set_paragraph_blanks(paragraphs[idx], [
            ("appendix_no", values.get("form_appendix_no")),
            ("day", values.get("form_appendix_day_ru")),
            ("month", values.get("form_appendix_month_ru")),
            ("year", values.get("form_appendix_year_ru")),
        ])
        if idx + 1 < len(paragraphs):
            _set_paragraph_blanks(paragraphs[idx + 1], [
                ("contract_no", values.get("form_contract_no")),
            ])
        if idx + 2 < len(paragraphs):
            _set_paragraph_blanks(paragraphs[idx + 2], [
                ("day", values.get("form_contract_day_ru")),
                ("month", values.get("form_contract_month_ru")),
                ("year", values.get("form_contract_year_ru")),
            ])

    # 英文附件抬头三段（"Appendix №__ ..."）
    idx_en = _find_paragraph(paragraphs, "Appendix №__")
    if idx_en >= 0:
        _set_paragraph_blanks(paragraphs[idx_en], [
            ("appendix_no", values.get("form_appendix_no")),
            ("date", values.get("form_appendix_date_en")),
            ("year", values.get("form_appendix_year_en")),
        ])
        if idx_en + 1 < len(paragraphs):
            _set_paragraph_blanks(paragraphs[idx_en + 1], [
                ("contract_no", values.get("form_contract_no_en") or values.get("form_contract_no")),
            ])
        if idx_en + 2 < len(paragraphs):
            _set_paragraph_blanks(paragraphs[idx_en + 2], [
                ("date", values.get("form_contract_date_en")),
                ("year", values.get("form_contract_year_en")),
            ])

    # 俄文卖方陈述句
    idx_party = _find_paragraph(paragraphs, "Компания ______________")
    if idx_party >= 0:
        _set_paragraph_blanks(paragraphs[idx_party], [
            ("seller_name", values.get("form_seller_name_ru")),
            ("country", values.get("form_seller_country_ru")),
            ("director", values.get("form_seller_director_ru")),
        ])

    # 英文卖方陈述句
    idx_party_en = _find_paragraph(paragraphs, "Company ______________")
    if idx_party_en >= 0:
        _set_paragraph_blanks(paragraphs[idx_party_en], [
            ("seller_name", values.get("form_seller_name_en") or values.get("form_seller_name_ru")),
            ("country", values.get("form_seller_country_en")),
            ("director", values.get("form_seller_director_en") or values.get("form_seller_director_ru")),
        ])

    # 4 / 5 节（俄文 + 英文）填写下划线行
    def fill_after(keyword: str, value: str) -> None:
        idx = _find_paragraph(paragraphs, keyword)
        if idx < 0 or not value:
            return
        for j in range(idx + 1, min(idx + 4, len(paragraphs))):
            text = "".join((t.text or "") for t in paragraphs[j].iter(f"{{{W_NS}}}t"))
            if "_" in text and text.replace("_", "").strip() == "":
                set_paragraph_text(paragraphs[j], value)
                return
            if not text.strip():
                continue
            break

    fill_after("4. Условия поставки", values.get("form_delivery_terms_ru"))
    fill_after("4. Conditions of delivery", values.get("form_delivery_terms_en"))
    fill_after("5. Срок поставки Товара", compose_delivery_time_ru(values))
    fill_after("5. Goods delivery time", compose_delivery_time_en(values))

    def fill_blank_paragraph_after(keyword: str, value: str) -> None:
        """在 keyword 段后第一个空白段写入 value。"""
        idx = _find_paragraph(paragraphs, keyword)
        if idx < 0 or not value:
            return
        for j in range(idx + 1, min(idx + 5, len(paragraphs))):
            text = "".join((t.text or "") for t in paragraphs[j].iter(f"{{{W_NS}}}t"))
            if not text.strip():
                set_paragraph_text(paragraphs[j], value)
                return

    fill_blank_paragraph_after("6. Условия платежа", values.get("form_payment_terms_ru") or make_payment_terms_ru(values))
    fill_blank_paragraph_after("6. Terms of payment", values.get("form_payment_terms_en") or make_payment_terms_en(values))
    fill_blank_paragraph_after("7. Комплект товаросопроводительных документов", values.get("form_shipping_docs_ru"))
    fill_blank_paragraph_after("7. The set of shipping documents", values.get("form_shipping_docs_en"))
    fill_blank_paragraph_after("8. Условия проведения технической приемки", values.get("form_acceptance_ru"))
    fill_blank_paragraph_after("8. The terms of carrying out technical acceptance", values.get("form_acceptance_en"))

    # 英文第 8 节原模板没有空白行，直接把标题后的第一段默认说明替换为填写内容。
    if values.get("form_acceptance_en"):
        idx = _find_paragraph(paragraphs, "8. The terms of carrying out technical acceptance")
        if idx >= 0 and idx + 1 < len(paragraphs):
            set_paragraph_text(paragraphs[idx + 1], values["form_acceptance_en"])

    # 双方签字位
    idx_buyer = _find_paragraph(paragraphs, "For and on behalf of the BUYER/")
    if idx_buyer >= 0:
        for j in range(idx_buyer + 1, min(idx_buyer + 5, len(paragraphs))):
            text = "".join((t.text or "") for t in paragraphs[j].iter(f"{{{W_NS}}}t"))
            if "/" in text and "_" in text:
                _set_paragraph_blanks(paragraphs[j], [
                    ("signer", values.get("form_buyer_signer_ru")),
                    ("signature", values.get("form_buyer_signature_ru")),
                ])
                break

    idx_seller = _find_paragraph(paragraphs, "For and on behalf of the SELLER/")
    if idx_seller >= 0:
        for j in range(idx_seller + 1, min(idx_seller + 5, len(paragraphs))):
            text = "".join((t.text or "") for t in paragraphs[j].iter(f"{{{W_NS}}}t"))
            if "/" in text and "_" in text:
                _set_paragraph_blanks(paragraphs[j], [
                    ("signer", values.get("form_seller_signer_ru")),
                    ("signature", values.get("form_seller_signature_ru")),
                ])
                break


def apply_structured_fields(root: ET.Element, values: dict) -> None:
    tables = list(root.iter(f"{{{W_NS}}}tbl"))
    if len(tables) < 2:
        return
    appendix_rows = tables[1].findall(f"{{{W_NS}}}tr")
    if len(appendix_rows) >= 13:
        goods_items = parse_goods_items(values)
        extra_goods_rows = max(0, len(goods_items) - 1)
        if goods_items:
            base_goods_row = appendix_rows[6]
            fill_goods_row(base_goods_row, goods_items[0])
            insert_pos = list(tables[1]).index(base_goods_row) + 1
            for offset, item in enumerate(goods_items[1:]):
                new_row = copy.deepcopy(base_goods_row)
                fill_goods_row(new_row, item)
                tables[1].insert(insert_pos + offset, new_row)
            appendix_rows = tables[1].findall(f"{{{W_NS}}}tr")
        package_cells = appendix_rows[9 + extra_goods_rows].findall(f"{{{W_NS}}}tc")
        for cell, value in zip(package_cells, [values.get("package_includes_ru"), values.get("package_includes_en")]):
            if value:
                set_cell_first_paragraph(cell, value)
        price_cells = appendix_rows[10 + extra_goods_rows].findall(f"{{{W_NS}}}tc")
        for cell, value in zip(price_cells, [values.get("total_price_ru"), values.get("total_price_en")]):
            if value:
                set_cell_paragraph(cell, 1, value)
        delivery_cells = appendix_rows[11 + extra_goods_rows].findall(f"{{{W_NS}}}tc")
        for cell, value in zip(delivery_cells, [values.get("delivery_terms_ru"), values.get("delivery_terms_en")]):
            if value:
                set_cell_paragraph(cell, 1, value)
        time_cells = appendix_rows[12 + extra_goods_rows].findall(f"{{{W_NS}}}tc")
        for cell, value in zip(time_cells, [compose_delivery_time_ru(values), compose_delivery_time_en(values)]):
            if value:
                set_cell_paragraph(cell, 1, value)


def process_word_xml(content: bytes, replacements: dict, fillable_values: set[str], values: dict) -> bytes:
    import sys as _sys
    _lib_dir = Path(__file__).resolve().parent.parent
    if str(_lib_dir) not in _sys.path:
        _sys.path.insert(0, str(_lib_dir))
    from docx_processor import process_document_xml, replace_text_in_xml

    # 通用后处理（接受修订、删除批注、移除高亮/底纹、统一字体、压缩间距等）
    content = process_document_xml(content)
    # 文本替换
    content = replace_text_in_xml(content, replacements)

    # 服务特有的后处理：结构化字段、表单空白、可填写值下划线
    try:
        from lxml import etree as lxml_etree
        parser = lxml_etree.XMLParser(remove_blank_text=False)
        root = lxml_etree.fromstring(content, parser)
    except Exception:
        return content
    apply_structured_fields(root, values)
    apply_form_blanks(root, values)
    underline_fillable_values(root, fillable_values)
    if LXML_ET:
        return LXML_ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


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
    replacements.update(highlighted_replacements(values))
    for key, default in DEFAULTS.items():
        if key == "appendix_no" or len(str(default).strip()) < 2:
            continue
        if default and values.get(key) and values[key] != default:
            replacements[default] = values[key]

    contract_no = values.get("contract_no") or DEFAULTS["contract_no"]
    replacements["GOF-04/06-26-CHN"] = contract_no
    replacements["Контракт № GOF-04/06-26-CHN"] = f"Контракт № {contract_no}"
    replacements["Contract № GOF-04/06-26-CHN"] = f"Contract № {contract_no}"

    if values.get("contract_date_ru"):
        replacements["«04» июня 2026 г."] = values["contract_date_ru"]
    if values.get("contract_date_en"):
        replacements["04 of June 2026"] = values["contract_date_en"]
    if values.get("appendix_date_en"):
        replacements["4th of June 2026"] = values["appendix_date_en"]
    if values.get("appendix_no"):
        replacements["Приложение № 1"] = f"Приложение № {values['appendix_no']}"
        replacements["Appendix №1"] = f"Appendix №{values['appendix_no']}"
    return replacements


def build_values(data: dict) -> dict:
    return {**DEFAULTS, **{k: v.strip() for k, v in data.items() if isinstance(v, str) and v.strip()}}


def build_fillable_values(data: dict) -> set[str]:
    values = build_values(data)
    underline_values = {values[k] for k in HIGHLIGHT_FILLABLE_KEYS if values.get(k)}
    underline_values.update(v for k, v in values.items() if k != "goods_items_json" and v and DEFAULTS.get(k) != v)
    for item in parse_goods_items(values):
        underline_values.update(v for v in item.values() if v)
    # 同步加入带逗号或固定前缀的复合填写项，确保这些黄色块也在 PDF 中带下划线。
    underline_values.add(seller_name_with_comma(values))
    underline_values.add(company_buyer_en_with_comma(values))
    underline_values.discard("")
    return underline_values


def process_docx(input_path: Path, output_path: Path, data: dict) -> None:
    values = build_values(data)
    replacements = build_replacements(data)
    fillable_values = build_fillable_values(data)
    with zipfile.ZipFile(input_path, "r") as zin, zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            # 只处理 word/document.xml，其他 XML 文件原样复制
            # 避免 ET.tostring 重写导致命名空间声明丢失
            if item.filename == "word/document.xml":
                content = process_word_xml(content, replacements, fillable_values, values)
            zout.writestr(item, content)


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

    # 使用 Word COM 复用实例转换（首次启动约 30-60 秒，后续约 3 秒）
    try:
        import pythoncom
        pythoncom.CoInitialize()
        try:
            import win32com.client
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(str(docx_path), False, True)
            doc.AcceptAllRevisions()
            doc.SaveAs(str(pdf_path), FileFormat=17)
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
        "  $word.DisplayAlerts = 0\n"
        f"  $doc = $word.Documents.OpenNoRepairDialog('{docx_literal}', $false, $true)\n"
        f"  $doc.AcceptAllRevisions()\n"
        f"  $doc.SaveAs([ref]'{pdf_literal}', [ref]17)\n"
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
        raise RuntimeError("缺少 Word 模板文件 templates/trade-contract-template.docx")
    with tempfile.TemporaryDirectory(prefix="trade_contract_pdf_") as tmp:
        tmp_dir = Path(tmp)
        docx_path = tmp_dir / "trade_contract.docx"
        process_docx(TEMPLATE_PATH, docx_path, data)
        pdf_path = convert_to_pdf(docx_path, tmp_dir)
        pdf_bytes = pdf_path.read_bytes()
    safe_no = re.sub(r"[^\w\-一-龥]+", "_", data.get("contract_no") or DEFAULTS["contract_no"])
    return pdf_bytes, f"俄英贸易合同_{safe_no}.pdf"


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
    sample = {**DEFAULTS}
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
