const importantPreviewFields = [
  "contract_no",
  "contract_date_ru",
  "contract_date_en",
  "buyer_name_ru",
  "buyer_name_en",
  "buyer_full_name_ru",
  "seller_name",
  "seller_address_en",
  "seller_bank_en",
  "seller_swift",
  "seller_account_no",
  "goods_name_ru",
  "goods_name_en",
  "goods_total",
  "form_appendix_no",
  "form_contract_no",
  "form_delivery_terms_ru",
  "form_delivery_terms_en",
  "form_payment_terms_ru",
  "form_payment_terms_en"
];

let defaults = {};
let labels = {};

const goodsFieldLabels = {
  goods_name_ru: "商品名称（俄文）",
  goods_name_en: "商品名称（英文）",
  goods_qty: "数量",
  goods_unit_price: "单价（CNY）",
  goods_total: "总价（CNY）"
};

const smartPasteAliases = [
  ["contract_no", ["合同编号", "合同号", "contract no", "contract number", "контракт №", "номер контракта"]],
  ["contract_date_ru", ["合同日期俄文", "合同日期（俄文）", "俄文日期", "дата контракта"]],
  ["contract_date_en", ["合同日期英文", "合同日期（英文）", "英文日期", "contract date"]],
  ["appendix_no", ["附件编号", "附件号", "appendix no", "приложение №"]],
  ["appendix_date_ru", ["附件日期俄文", "附件日期（俄文）"]],
  ["appendix_date_en", ["附件日期英文", "附件日期（英文）"]],
  ["buyer_name_ru", ["买方名称俄文", "买方名称（俄文）", "买方俄文", "покупатель"]],
  ["buyer_name_en", ["买方名称英文", "买方名称（英文）", "buyer", "buyer name"]],
  ["buyer_full_name_ru", ["买方完整法定名称俄文", "买方完整法定名称（俄文）", "买方法定名", "买方完整名称"]],
  ["buyer_address_ru", ["买方地址俄文", "买方地址（俄文）"]],
  ["buyer_address_en", ["买方地址英文", "买方地址（英文）", "buyer address"]],
  ["buyer_tel", ["买方电话", "buyer tel", "buyer phone"]],
  ["buyer_email", ["买方邮箱", "buyer email"]],
  ["buyer_notice_tel", ["装运通知电话", "装运通知传真", "通知电话", "通知传真"]],
  ["buyer_notice_email", ["装运通知邮箱", "通知邮箱", "通知邮件"]],
  ["buyer_representative_ru", ["买方代表俄文", "买方代表（俄文）"]],
  ["buyer_representative_en", ["买方代表英文", "买方代表（英文）"]],
  ["seller_name", ["卖方名称", "卖方", "seller", "seller name", "продавец"]],
  ["seller_address_ru", ["卖方地址俄文", "卖方地址（俄文）"]],
  ["seller_address_en", ["卖方地址英文", "卖方地址（英文）", "seller address"]],
  ["seller_tel", ["卖方电话", "seller tel", "seller phone"]],
  ["seller_email", ["卖方邮箱", "seller email"]],
  ["seller_bank_ru", ["卖方银行俄文", "卖方银行（俄文）"]],
  ["seller_bank_en", ["卖方银行英文", "卖方银行（英文）", "开户银行", "bank", "bank name"]],
  ["seller_bank_address_ru", ["卖方银行地址俄文", "卖方银行地址（俄文）"]],
  ["seller_bank_address_en", ["卖方银行地址英文", "卖方银行地址（英文）", "银行地址", "bank address"]],
  ["seller_account_no", ["账户号", "账号", "卖方账户号", "account no", "account number"]],
  ["seller_swift", ["swift", "swift代码", "swift 代码"]],
  ["seller_beneficiary", ["收款人名称", "账户名称", "beneficiary", "account name"]],
  ["buyer_corr_account", ["买方银行对应账户", "对应账户", "correspondent account"]],
  ["buyer_bic", ["买方银行bic", "bic", "бик"]],
  ["seller_representative_ru", ["卖方代表俄文", "卖方代表（俄文）"]],
  ["seller_representative_en", ["卖方代表英文", "卖方代表（英文）"]],
  ["goods_name_ru", ["商品名称俄文", "商品名称（俄文）", "货物俄文"]],
  ["goods_name_en", ["商品名称英文", "商品名称（英文）", "货物英文", "goods"]],
  ["goods_qty", ["数量", "qty", "quantity"]],
  ["goods_unit_price", ["单价", "unit price"]],
  ["goods_total", ["总价", "total"]],
  ["package_includes_ru", ["供货清单俄文", "供货清单（俄文）"]],
  ["package_includes_en", ["供货清单英文", "供货清单（英文）"]],
  ["total_price_ru", ["总价大写俄文", "总价大写（俄文）"]],
  ["total_price_en", ["总价大写英文", "总价大写（英文）"]],
  ["delivery_terms_ru", ["交付条件俄文", "交付条件（俄文）"]],
  ["delivery_terms_en", ["交付条件英文", "交付条件（英文）"]],
  ["delivery_time_ru", ["交付时间俄文", "交付时间（俄文）"]],
  ["delivery_time_en", ["交付时间英文", "交付时间（英文）"]],
  ["form_appendix_no", ["附件表单附件编号", "附件表单：附件编号", "表单附件编号", "填写区附件编号"]],
  ["form_appendix_day_ru", ["附件表单附件日俄文", "附件日俄文", "附件日（俄文）"]],
  ["form_appendix_month_ru", ["附件表单附件月俄文", "附件月俄文", "附件月（俄文）"]],
  ["form_appendix_year_ru", ["附件表单附件年俄文", "附件年俄文", "附件年（俄文）"]],
  ["form_contract_no", ["附件表单合同编号", "附件合同编号", "表单合同编号"]],
  ["form_contract_day_ru", ["附件表单合同日俄文", "合同日俄文", "合同日（俄文）"]],
  ["form_contract_month_ru", ["附件表单合同月俄文", "合同月俄文", "合同月（俄文）"]],
  ["form_contract_year_ru", ["附件表单合同年俄文", "合同年俄文", "合同年（俄文）"]],
  ["form_appendix_date_en", ["附件表单附件日期英文", "附件日期英文", "附件日期（英文）"]],
  ["form_appendix_year_en", ["附件表单附件年英文", "附件年英文", "附件年（英文）"]],
  ["form_contract_no_en", ["附件表单合同编号英文", "合同编号英文", "合同编号（英文）"]],
  ["form_contract_date_en", ["附件表单合同日期英文", "合同日期英文", "合同日期（英文）"]],
  ["form_contract_year_en", ["附件表单合同年英文", "合同年英文", "合同年（英文）"]],
  ["form_seller_name_ru", ["附件表单卖方名称俄文", "卖方公司名称俄文", "卖方公司名称（俄文）"]],
  ["form_seller_country_ru", ["附件表单卖方国家俄文", "卖方注册国家俄文", "卖方注册国家（俄文）"]],
  ["form_seller_director_ru", ["附件表单卖方法人俄文", "卖方法人代表俄文", "卖方法人代表（俄文）"]],
  ["form_seller_name_en", ["附件表单卖方名称英文", "卖方公司名称英文", "卖方公司名称（英文）"]],
  ["form_seller_country_en", ["附件表单卖方国家英文", "卖方注册国家英文", "卖方注册国家（英文）"]],
  ["form_seller_director_en", ["附件表单卖方法人英文", "卖方法人代表英文", "卖方法人代表（英文）"]],
  ["form_delivery_terms_ru", ["附件表单交付条件俄文", "4交付条件俄文", "4. 交付条件（俄文）"]],
  ["form_delivery_terms_en", ["附件表单交付条件英文", "4交付条件英文", "4. 交付条件（英文）"]],
  ["form_delivery_time_ru", ["附件表单交付时间俄文", "5交付时间俄文", "5. 交付时间（俄文）"]],
  ["form_delivery_time_en", ["附件表单交付时间英文", "5交付时间英文", "5. 交付时间（英文）"]],
  ["form_payment_terms_ru", ["附件表单付款条件俄文", "6付款条件俄文", "6. 付款条件（俄文）"]],
  ["form_payment_terms_en", ["附件表单付款条件英文", "6付款条件英文", "6. 付款条件（英文）"]],
  ["form_shipping_docs_ru", ["附件表单单据清单俄文", "7单据清单俄文", "7. 单据清单（俄文）"]],
  ["form_shipping_docs_en", ["附件表单单据清单英文", "7单据清单英文", "7. 单据清单（英文）"]],
  ["form_acceptance_ru", ["附件表单技术验收俄文", "8技术验收俄文", "8. 技术验收（俄文）"]],
  ["form_acceptance_en", ["附件表单技术验收英文", "8技术验收英文", "8. 技术验收（英文）"]],
  ["form_buyer_signer_ru", ["附件表单买方签字代表", "买方签字代表"]],
  ["form_buyer_signature_ru", ["附件表单买方签字名", "买方签字名"]],
  ["form_seller_signer_ru", ["附件表单卖方签字代表", "卖方签字代表"]],
  ["form_seller_signature_ru", ["附件表单卖方签字名", "卖方签字名"]]
];

function createGoodsItem(item = {}) {
  const holder = document.querySelector("#goodsItems");
  const index = holder.children.length + 1;
  const card = document.createElement("div");
  card.className = "goods-item";
  card.innerHTML = `
    <div class="goods-item-head">
      <div class="goods-item-title">产品 ${index}</div>
      <button type="button" class="goods-item-remove">删除</button>
    </div>
    <div class="grid">
      <label class="wide">商品名称（俄文）<textarea name="goods_name_ru" rows="2"></textarea></label>
      <label class="wide">商品名称（英文）<textarea name="goods_name_en" rows="2"></textarea></label>
      <label>数量<input name="goods_qty"></label>
      <label>单价（CNY）<input name="goods_unit_price"></label>
      <label>总价（CNY）<input name="goods_total"></label>
    </div>
  `;
  Object.entries(goodsFieldLabels).forEach(([name]) => {
    const el = card.querySelector(`[name="${name}"]`);
    if (el) {
      el.value = item[name] || "";
      el.addEventListener("input", () => {
        syncGoodsItems();
        refreshPreview();
      });
    }
  });
  card.querySelector(".goods-item-remove").addEventListener("click", () => {
    if (holder.children.length <= 1) {
      Object.keys(goodsFieldLabels).forEach((name) => {
        const el = card.querySelector(`[name="${name}"]`);
        if (el) el.value = "";
      });
    } else {
      card.remove();
    }
    renumberGoodsItems();
    syncGoodsItems();
    refreshPreview();
  });
  holder.appendChild(card);
  renumberGoodsItems();
  syncGoodsItems();
}

function renumberGoodsItems() {
  document.querySelectorAll(".goods-item").forEach((card, idx) => {
    const title = card.querySelector(".goods-item-title");
    if (title) title.textContent = `产品 ${idx + 1}`;
  });
}

function getGoodsItems() {
  return Array.from(document.querySelectorAll(".goods-item")).map((card) => {
    const item = {};
    Object.keys(goodsFieldLabels).forEach((name) => {
      item[name] = card.querySelector(`[name="${name}"]`)?.value?.trim() || "";
    });
    return item;
  }).filter((item) => Object.values(item).some(Boolean));
}

function syncGoodsItems() {
  const hidden = document.querySelector("#goodsItemsJson");
  if (hidden) hidden.value = JSON.stringify(getGoodsItems());
}

function resetGoodsItems() {
  const holder = document.querySelector("#goodsItems");
  holder.innerHTML = "";
  createGoodsItem({
    goods_name_ru: defaults.goods_name_ru || "",
    goods_name_en: defaults.goods_name_en || "",
    goods_qty: defaults.goods_qty || "",
    goods_unit_price: defaults.goods_unit_price || "",
    goods_total: defaults.goods_total || ""
  });
}

function fillDefaults() {
  resetGoodsItems();
  document.querySelectorAll("[name]").forEach((el) => {
    if (Object.prototype.hasOwnProperty.call(goodsFieldLabels, el.name)) return;
    el.value = defaults[el.name] ?? "";
  });
  syncGoodsItems();
  refreshPreview();
}

function refreshPreview() {
  const holder = document.querySelector("#previewList");
  holder.innerHTML = "";
  importantPreviewFields.forEach((name) => {
    const input = document.querySelector(`[name="${name}"]`);
    const value = input?.value?.trim() || "未填写";
    const item = document.createElement("div");
    item.className = "preview-item";
    item.innerHTML = `<span>${labels[name] || name}</span><strong></strong>`;
    item.querySelector("strong").textContent = value;
    holder.appendChild(item);
  });
}

function normalizeSmartKey(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/[：:＝=№#]/g, "")
    .replace(/[（）()【】\[\]\/\\_\-\s]/g, "")
    .trim();
}

function findSmartField(rawKey) {
  const key = normalizeSmartKey(rawKey);
  if (!key) return "";
  for (const [field, aliases] of smartPasteAliases) {
    if (normalizeSmartKey(labels[field] || "") === key) return field;
    if (aliases.some((alias) => normalizeSmartKey(alias) === key)) return field;
  }
  for (const [field, aliases] of smartPasteAliases) {
    if (aliases.some((alias) => key.includes(normalizeSmartKey(alias)))) return field;
  }
  return "";
}

function splitSmartLine(line) {
  const trimmed = line.trim();
  if (!trimmed) return null;
  const exact = trimmed.match(/^([^:：=＝]{1,36})\s*[:：=＝]\s*(.+)$/);
  if (exact) return [exact[1].trim(), exact[2].trim()];
  const loose = trimmed.match(/^([\u4e00-\u9fa5A-Za-zА-Яа-я\s\/（）()]{2,28})\s+(.{2,})$/);
  if (loose) return [loose[1].trim(), loose[2].trim()];
  return null;
}

function parseSmartPasteText(text) {
  const result = {};
  String(text || "")
    .replace(/\r/g, "")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .forEach((line) => {
      const pair = splitSmartLine(line);
      if (!pair) return;
      const [rawKey, rawValue] = pair;
      const field = findSmartField(rawKey);
      if (!field) return;
      result[field] = rawValue.trim().replace(/^["“”']|["“”']$/g, "");
    });
  return result;
}

function applySmartPaste() {
  const input = document.querySelector("#smartPasteInput");
  const status = document.querySelector("#smartPasteStatus");
  const parsed = parseSmartPasteText(input.value);
  const entries = Object.entries(parsed).filter(([name]) => document.querySelector(`[name="${name}"]`));
  if (!input.value.trim()) {
    status.textContent = "请先把需要识别的信息粘贴到上方文本框。";
    return;
  }
  if (!entries.length) {
    status.textContent = "没有识别到可填字段，请尽量使用“字段名：值”的格式。";
    return;
  }
  entries.forEach(([name, value]) => {
    const el = document.querySelector(`[name="${name}"]`);
    el.value = value;
    el.dispatchEvent(new Event("input", { bubbles: true }));
  });
  syncGoodsItems();
  refreshPreview();
  status.textContent = `已识别并填入 ${entries.length} 个字段：${entries.map(([name]) => labels[name] || name).join("、")}`;
}

async function init() {
  const res = await fetch("/defaults");
  const data = await res.json();
  defaults = data.defaults || {};
  labels = data.labels || {};
  fillDefaults();

  document.querySelector("#resetBtn").addEventListener("click", fillDefaults);
  document.querySelector("#addGoodsBtn").addEventListener("click", () => {
    createGoodsItem();
    refreshPreview();
  });
  document.querySelector("#smartParseBtn").addEventListener("click", applySmartPaste);
  document.querySelectorAll("input, textarea").forEach((el) => {
    el.addEventListener("input", refreshPreview);
  });

  document.querySelector("#contractForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    syncGoodsItems();
    const btn = document.querySelector("#submitBtn");
    const oldText = btn.textContent;
    btn.disabled = true;
    btn.textContent = "正在生成 PDF...";
    try {
      const response = await fetch("/generate", {
        method: "POST",
        body: new URLSearchParams(new FormData(event.currentTarget))
      });
      if (!response.ok) {
        const message = await response.text();
        throw new Error(message || "生成失败");
      }
      const blob = await response.blob();
      const disposition = response.headers.get("Content-Disposition") || "";
      const match = disposition.match(/filename\*=UTF-8''([^;]+)/);
      const filename = match ? decodeURIComponent(match[1]) : "俄英贸易合同.pdf";
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      alert(err.message || String(err));
    } finally {
      btn.disabled = false;
      btn.textContent = oldText;
    }
  });
}

init().catch((err) => {
  alert(`网站初始化失败：${err.message || err}`);
});
