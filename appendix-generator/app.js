const wideTextFields = new Set([
  "buyer_full_name_ru", "buyer_address_ru", "buyer_address_en", "buyer_postal_ru", "buyer_postal_en",
  "seller_address_ru", "seller_address_en", "seller_bank_address_ru", "seller_bank_address_en",
  "buyer_bank_address_ru", "buyer_bank_address_en", "package_includes_ru", "package_includes_en",
  "total_price_ru", "total_price_en", "delivery_terms_ru", "delivery_terms_en",
  "delivery_time_ru", "delivery_time_en", "form_shipping_docs_ru", "form_shipping_docs_en",
  "form_acceptance_ru", "form_acceptance_en", "payment_advance_amount_ru", "payment_advance_amount_en",
  "payment_balance_amount_ru", "payment_balance_amount_en"
]);

const goodsFields = new Set(["goods_name_ru", "goods_name_en", "goods_qty", "goods_unit_price", "goods_total", "goods_items_json"]);

const baseFieldKeys = [
  "appendix_no", "appendix_date_ru", "appendix_date_en",
  "contract_no", "contract_date_ru", "contract_date_en",
  "buyer_name_ru", "buyer_name_en", "buyer_full_name_ru", "buyer_address_ru", "buyer_address_en",
  "buyer_tel", "buyer_email", "buyer_notice_tel", "buyer_notice_email",
  "buyer_representative_ru", "buyer_representative_en", "buyer_signature_ru", "buyer_signature_en",
  "seller_name", "seller_address_ru", "seller_address_en", "seller_tel", "seller_email",
  "seller_representative_ru", "seller_representative_en"
];

const termsFieldKeys = [
  "package_includes_ru", "package_includes_en",
  "total_price_ru", "total_price_en",
  "delivery_terms_ru", "delivery_terms_en",
  "delivery_time_ru", "delivery_time_en",
  "payment_advance_percent", "payment_advance_amount", "payment_advance_amount_ru", "payment_advance_amount_en",
  "payment_balance_percent", "payment_balance_amount", "payment_balance_amount_ru", "payment_balance_amount_en",
  "payment_bank_days", "payment_bank_days_ru", "payment_bank_days_en",
  "delivery_notice_days", "delivery_notice_days_ru", "delivery_notice_days_en",
  "form_shipping_docs_ru", "form_shipping_docs_en",
  "form_acceptance_ru", "form_acceptance_en"
];

const signFieldKeys = [
  "form_buyer_signer_ru", "form_buyer_signature_ru",
  "form_seller_signer_ru", "form_seller_signature_ru"
];

const previewFields = [
  "appendix_no", "contract_no", "buyer_name_en", "seller_name",
  "delivery_terms_en", "payment_advance_amount", "payment_balance_amount", "total_price_en"
];

let defaults = {};
let labels = {};

function labelFor(key) {
  return labels[key] || key;
}

function createField(key) {
  const label = document.createElement("label");
  label.textContent = labelFor(key);
  if (wideTextFields.has(key)) label.classList.add("wide");

  const input = document.createElement(wideTextFields.has(key) ? "textarea" : "input");
  input.name = key;
  if (wideTextFields.has(key)) input.rows = 2;
  input.value = defaults[key] || "";
  input.addEventListener("input", refreshPreview);
  label.appendChild(input);
  return label;
}

function renderFieldGroup(holderId, keys) {
  const holder = document.querySelector(holderId);
  holder.innerHTML = "";
  keys.forEach((key) => {
    if (key in defaults) holder.appendChild(createField(key));
  });
}

function createGoodsItem(item = {}) {
  const holder = document.querySelector("#goodsItems");
  const card = document.createElement("div");
  card.className = "goods-item";
  card.innerHTML = `
    <div class="goods-item-head">
      <div class="goods-item-title">产品 ${holder.children.length + 1}</div>
      <button type="button" class="goods-item-remove">删除</button>
    </div>
    <div class="generated-fields">
      <label class="wide">商品名称（俄文）<textarea name="goods_name_ru_item" rows="2"></textarea></label>
      <label class="wide">商品名称（英文）<textarea name="goods_name_en_item" rows="2"></textarea></label>
      <label>数量<input name="goods_qty_item"></label>
      <label>单价（CNY）<input name="goods_unit_price_item"></label>
      <label>总价（CNY）<input name="goods_total_item"></label>
    </div>
  `;
  const map = {
    goods_name_ru_item: "goods_name_ru",
    goods_name_en_item: "goods_name_en",
    goods_qty_item: "goods_qty",
    goods_unit_price_item: "goods_unit_price",
    goods_total_item: "goods_total"
  };
  Object.entries(map).forEach(([domName, key]) => {
    const el = card.querySelector(`[name="${domName}"]`);
    el.value = item[key] || "";
    el.addEventListener("input", () => {
      syncGoods();
      refreshPreview();
    });
  });
  card.querySelector(".goods-item-remove").addEventListener("click", () => {
    if (holder.children.length <= 1) {
      card.querySelectorAll("input, textarea").forEach((el) => { el.value = ""; });
    } else {
      card.remove();
    }
    renumberGoods();
    syncGoods();
    refreshPreview();
  });
  holder.appendChild(card);
  renumberGoods();
  syncGoods();
}

function renumberGoods() {
  document.querySelectorAll(".goods-item").forEach((card, idx) => {
    card.querySelector(".goods-item-title").textContent = `产品 ${idx + 1}`;
  });
}

function getGoodsItems() {
  return Array.from(document.querySelectorAll("#goodsItems .goods-item")).map((card) => ({
    goods_name_ru: card.querySelector('[name="goods_name_ru_item"]')?.value?.trim() || "",
    goods_name_en: card.querySelector('[name="goods_name_en_item"]')?.value?.trim() || "",
    goods_qty: card.querySelector('[name="goods_qty_item"]')?.value?.trim() || "",
    goods_unit_price: card.querySelector('[name="goods_unit_price_item"]')?.value?.trim() || "",
    goods_total: card.querySelector('[name="goods_total_item"]')?.value?.trim() || ""
  })).filter((item) => Object.values(item).some(Boolean));
}

function syncGoods() {
  document.querySelector("#goodsItemsJson").value = JSON.stringify(getGoodsItems());
}

function renderGoods() {
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
  renderFieldGroup("#baseFields", baseFieldKeys);
  renderGoods();
  renderFieldGroup("#termsFields", termsFieldKeys);
  renderFieldGroup("#signFields", signFieldKeys);
  refreshPreview();
}

function refreshPreview() {
  syncGoods();
  const holder = document.querySelector("#previewList");
  holder.innerHTML = "";
  previewFields.forEach((key) => {
    const input = document.querySelector(`[name="${key}"]`);
    const value = input?.value?.trim() || "未填写";
    const item = document.createElement("div");
    item.className = "preview-item";
    item.innerHTML = `<span>${labelFor(key)}</span><strong></strong>`;
    item.querySelector("strong").textContent = value;
    holder.appendChild(item);
  });
  const goods = document.createElement("div");
  goods.className = "preview-item";
  goods.innerHTML = `<span>商品数量</span><strong>${getGoodsItems().length || 0} 项</strong>`;
  holder.appendChild(goods);
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

  document.querySelector("#appendixForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    syncGoods();
    const btn = document.querySelector("#submitBtn");
    const oldText = btn.textContent;
    btn.disabled = true;
    btn.textContent = "正在生成附件2 PDF...";
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
      const filename = match ? decodeURIComponent(match[1]) : "附件2.pdf";
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
