const wideTextFields = new Set([
  "party_a_address", "party_a_address_ru", "party_b_address", "party_b_address_ru",
  "party_c_address", "party_c_address_ru", "party_b_bank_address",
  "buyer_full_name_ru", "buyer_address_ru", "buyer_address_en", "buyer_postal_ru", "buyer_postal_en",
  "seller_address_ru", "seller_address_en", "seller_bank_address_ru", "seller_bank_address_en",
  "buyer_bank_address_ru", "buyer_bank_address_en", "package_includes_ru", "package_includes_en",
  "form_delivery_terms_ru", "form_delivery_terms_en", "form_delivery_time_ru", "form_delivery_time_en",
  "form_payment_terms_ru", "form_payment_terms_en", "form_shipping_docs_ru", "form_shipping_docs_en",
  "form_acceptance_ru", "form_acceptance_en"
]);

const tradeGoodsFields = new Set([
  "goods_name_ru", "goods_name_en", "goods_qty", "goods_unit_price", "goods_total"
]);

const previewFields = [
  ["common", "contract_no"],
  ["common", "buyer_name_en"],
  ["common", "seller_name"],
  ["common", "seller_account_no"],
  ["common", "seller_swift"],
  ["agency", "agreement_no"],
  ["trade", "total_price_en"],
  ["appendix", "appendix_no"]
];

const commonFieldDefs = [
  {
    key: "contract_no",
    label: "贸易合同号 / 合同编号",
    targets: [
      ["agency", "trade_contract_no"],
      ["agency", "trade_contract_no_with_symbol", (value) => value ? (value.trim().startsWith("№") ? value.trim() : `№ ${value.trim()}`) : ""],
      ["trade", "contract_no"],
      ["trade", "form_contract_no"],
      ["trade", "form_contract_no_en"]
    ]
  },
  { key: "contract_date_en", label: "贸易合同日期（英文）", targets: [["agency", "trade_contract_date_en"], ["trade", "contract_date_en"]] },
  { key: "contract_date_ru", label: "贸易合同日期（俄文）", targets: [["agency", "trade_contract_date_ru"], ["trade", "contract_date_ru"]] },
  { key: "buyer_name_en", label: "甲方 / 买方名称（英文）", targets: [["agency", "party_a_name_en"], ["trade", "buyer_name_en"]] },
  { key: "buyer_name_ru", label: "甲方 / 买方名称（俄文）", targets: [["agency", "party_a_name_ru"], ["trade", "buyer_name_ru"]] },
  { key: "buyer_address_en", label: "甲方 / 买方地址（英文）", wide: true, targets: [["agency", "party_a_address"], ["trade", "buyer_address_en"]] },
  { key: "buyer_address_ru", label: "甲方 / 买方地址（俄文）", wide: true, targets: [["agency", "party_a_address_ru"], ["trade", "buyer_address_ru"]] },
  { key: "buyer_representative_en", label: "甲方 / 买方代表（英文）", targets: [["agency", "party_a_representative"], ["trade", "buyer_representative_en"]] },
  { key: "buyer_representative_ru", label: "甲方 / 买方代表（俄文）", targets: [["agency", "party_a_representative_ru"], ["trade", "buyer_representative_ru"]] },
  { key: "seller_name", label: "乙方 / 卖方名称", targets: [["agency", "party_b_name_en"], ["agency", "party_b_name_ru"], ["trade", "seller_name"], ["trade", "form_seller_name_ru"], ["trade", "form_seller_name_en"]] },
  { key: "seller_address_en", label: "乙方 / 卖方地址（英文）", wide: true, targets: [["agency", "party_b_address"], ["trade", "seller_address_en"]] },
  { key: "seller_address_ru", label: "乙方 / 卖方地址（俄文）", wide: true, targets: [["agency", "party_b_address_ru"], ["trade", "seller_address_ru"]] },
  { key: "seller_representative_en", label: "乙方 / 卖方代表（英文）", targets: [["agency", "party_b_representative"], ["trade", "seller_representative_en"], ["trade", "form_seller_director_en"]] },
  { key: "seller_representative_ru", label: "乙方 / 卖方代表（俄文）", targets: [["agency", "party_b_representative_ru"], ["trade", "seller_representative_ru"], ["trade", "form_seller_director_ru"]] },
  { key: "seller_bank_en", label: "乙方 / 卖方开户银行（英文）", targets: [["agency", "party_b_bank"], ["trade", "seller_bank_en"]] },
  { key: "seller_bank_ru", label: "乙方 / 卖方开户银行（俄文）", targets: [["agency", "party_b_bank_ru"], ["trade", "seller_bank_ru"]] },
  { key: "seller_bank_address_en", label: "乙方 / 卖方银行地址（英文）", wide: true, targets: [["agency", "party_b_bank_address"], ["trade", "seller_bank_address_en"]] },
  { key: "seller_bank_address_ru", label: "乙方 / 卖方银行地址（俄文）", wide: true, targets: [["trade", "seller_bank_address_ru"]] },
  { key: "seller_swift", label: "乙方 / 卖方 SWIFT", targets: [["agency", "party_b_swift"], ["trade", "seller_swift"]] },
  { key: "seller_account_no", label: "乙方 / 卖方账户号", targets: [["agency", "party_b_account_no"], ["trade", "seller_account_no"]] },
  { key: "seller_beneficiary", label: "乙方 / 卖方收款人名称", targets: [["agency", "party_b_account_name"], ["trade", "seller_beneficiary"]] }
];

const appendixSharedTargets = {
  contract_no: [["appendix", "contract_no"], ["appendix", "form_contract_no"], ["appendix", "form_contract_no_en"]],
  contract_date_en: [["appendix", "contract_date_en"], ["appendix", "form_contract_date_en"]],
  contract_date_ru: [["appendix", "contract_date_ru"]],
  buyer_name_en: [["appendix", "buyer_name_en"]],
  buyer_name_ru: [["appendix", "buyer_name_ru"]],
  buyer_address_en: [["appendix", "buyer_address_en"]],
  buyer_address_ru: [["appendix", "buyer_address_ru"]],
  buyer_representative_en: [["appendix", "buyer_representative_en"]],
  buyer_representative_ru: [["appendix", "buyer_representative_ru"]],
  seller_name: [["appendix", "seller_name"], ["appendix", "form_seller_name_ru"], ["appendix", "form_seller_name_en"]],
  seller_address_en: [["appendix", "seller_address_en"]],
  seller_address_ru: [["appendix", "seller_address_ru"]],
  seller_representative_en: [["appendix", "seller_representative_en"], ["appendix", "form_seller_director_en"]],
  seller_representative_ru: [["appendix", "seller_representative_ru"], ["appendix", "form_seller_director_ru"]],
  seller_bank_en: [["appendix", "seller_bank_en"]],
  seller_bank_ru: [["appendix", "seller_bank_ru"]],
  seller_bank_address_en: [["appendix", "seller_bank_address_en"]],
  seller_bank_address_ru: [["appendix", "seller_bank_address_ru"]],
  seller_swift: [["appendix", "seller_swift"]],
  seller_account_no: [["appendix", "seller_account_no"]],
  seller_beneficiary: [["appendix", "seller_beneficiary"]]
};

commonFieldDefs.forEach((def) => {
  if (appendixSharedTargets[def.key]) {
    def.targets.push(...appendixSharedTargets[def.key]);
  }
});

commonFieldDefs.push(
  { key: "buyer_notice_tel", label: "装运通知电话/传真", targets: [["trade", "buyer_notice_tel"], ["appendix", "buyer_notice_tel"]] },
  { key: "buyer_notice_email", label: "装运通知邮箱", targets: [["trade", "buyer_notice_email"], ["appendix", "buyer_notice_email"]] },
  { key: "package_includes_ru", label: "供货清单/包装说明（俄文）", wide: true, targets: [["trade", "package_includes_ru"], ["appendix", "package_includes_ru"]] },
  { key: "package_includes_en", label: "供货清单/包装说明（英文）", wide: true, targets: [["trade", "package_includes_en"], ["appendix", "package_includes_en"]] },
  { key: "total_price_ru", label: "总价大写（俄文）", wide: true, targets: [["trade", "total_price_ru"], ["appendix", "total_price_ru"]] },
  { key: "total_price_en", label: "总价大写（英文）", wide: true, targets: [["trade", "total_price_en"], ["appendix", "total_price_en"]] },
  { key: "delivery_terms_ru", label: "交付条件（俄文）", wide: true, targets: [["trade", "delivery_terms_ru"], ["trade", "form_delivery_terms_ru"], ["appendix", "delivery_terms_ru"], ["appendix", "form_delivery_terms_ru"]] },
  { key: "delivery_terms_en", label: "交付条件（英文）", wide: true, targets: [["trade", "delivery_terms_en"], ["trade", "form_delivery_terms_en"], ["appendix", "delivery_terms_en"], ["appendix", "form_delivery_terms_en"]] },
  { key: "delivery_time_ru", label: "交付时间（俄文）", wide: true, targets: [["trade", "delivery_time_ru"], ["trade", "form_delivery_time_ru"], ["appendix", "delivery_time_ru"], ["appendix", "form_delivery_time_ru"]] },
  { key: "delivery_time_en", label: "交付时间（英文）", wide: true, targets: [["trade", "delivery_time_en"], ["trade", "form_delivery_time_en"], ["appendix", "delivery_time_en"], ["appendix", "form_delivery_time_en"]] },
  { key: "payment_advance_percent", label: "预付款比例", targets: [["trade", "payment_advance_percent"], ["appendix", "payment_advance_percent"]] },
  { key: "payment_advance_amount", label: "预付款金额（数字）", targets: [["trade", "payment_advance_amount"], ["appendix", "payment_advance_amount"]] },
  { key: "payment_advance_amount_ru", label: "预付款金额大写（俄文）", wide: true, targets: [["trade", "payment_advance_amount_ru"], ["appendix", "payment_advance_amount_ru"]] },
  { key: "payment_advance_amount_en", label: "预付款金额大写（英文）", wide: true, targets: [["trade", "payment_advance_amount_en"], ["appendix", "payment_advance_amount_en"]] },
  { key: "payment_balance_percent", label: "尾款比例", targets: [["trade", "payment_balance_percent"], ["appendix", "payment_balance_percent"]] },
  { key: "payment_balance_amount", label: "尾款金额（数字）", targets: [["trade", "payment_balance_amount"], ["appendix", "payment_balance_amount"]] },
  { key: "payment_balance_amount_ru", label: "尾款金额大写（俄文）", wide: true, targets: [["trade", "payment_balance_amount_ru"], ["appendix", "payment_balance_amount_ru"]] },
  { key: "payment_balance_amount_en", label: "尾款金额大写（英文）", wide: true, targets: [["trade", "payment_balance_amount_en"], ["appendix", "payment_balance_amount_en"]] },
  { key: "payment_bank_days", label: "预付款支付期限（数字）", targets: [["trade", "payment_bank_days"], ["appendix", "payment_bank_days"]] },
  { key: "payment_bank_days_ru", label: "预付款支付期限大写（俄文）", targets: [["trade", "payment_bank_days_ru"], ["appendix", "payment_bank_days_ru"]] },
  { key: "payment_bank_days_en", label: "预付款支付期限大写（英文）", targets: [["trade", "payment_bank_days_en"], ["appendix", "payment_bank_days_en"]] },
  { key: "delivery_notice_days", label: "发货前通知天数（数字）", targets: [["trade", "delivery_notice_days"], ["appendix", "delivery_notice_days"]] },
  { key: "delivery_notice_days_ru", label: "发货前通知天数大写（俄文）", targets: [["trade", "delivery_notice_days_ru"], ["appendix", "delivery_notice_days_ru"]] },
  { key: "delivery_notice_days_en", label: "发货前通知天数大写（英文）", targets: [["trade", "delivery_notice_days_en"], ["appendix", "delivery_notice_days_en"]] },
  { key: "form_acceptance_ru", label: "技术验收条款（俄文）", wide: true, targets: [["trade", "form_acceptance_ru"], ["appendix", "form_acceptance_ru"]] },
  { key: "form_acceptance_en", label: "技术验收条款（英文）", wide: true, targets: [["trade", "form_acceptance_en"], ["appendix", "form_acceptance_en"]] }
);

const syncedTargetKeys = new Set(
  commonFieldDefs.flatMap((def) => def.targets.map(([prefix, field]) => fullName(prefix, field)))
);

let agencyDefaults = {};
let tradeDefaults = {};
let appendixDefaults = {};
let agencyLabels = {};
let tradeLabels = {};
let appendixLabels = {};

function fullName(prefix, key) {
  return `${prefix}__${key}`;
}

function labelFor(prefix, key) {
  if (prefix === "common") {
    return commonFieldDefs.find((def) => def.key === key)?.label || key;
  }
  const labels = prefix === "agency" ? agencyLabels : (prefix === "trade" ? tradeLabels : appendixLabels);
  return labels[key] || key;
}

function defaultsFor(prefix) {
  if (prefix === "agency") return agencyDefaults;
  if (prefix === "trade") return tradeDefaults;
  return appendixDefaults;
}

function targetDefault([prefix, key]) {
  return defaultsFor(prefix)[key] || "";
}

function defaultForCommon(def) {
  for (const target of def.targets) {
    const value = targetDefault(target);
    if (value) return value.replace(/^№\s*/i, "");
  }
  return "";
}

function createCommonField(def) {
  const label = document.createElement("label");
  label.textContent = def.label;
  if (def.wide) label.classList.add("wide");
  const input = document.createElement(def.wide ? "textarea" : "input");
  input.name = fullName("common", def.key);
  if (def.wide) input.rows = 2;
  input.value = defaultForCommon(def);
  input.addEventListener("input", () => {
    syncCommonFields();
    refreshPreview();
  });
  label.appendChild(input);
  return label;
}

function renderCommonFields() {
  const holder = document.querySelector("#commonFields");
  const hidden = document.querySelector("#syncedHiddenFields");
  holder.innerHTML = "";
  hidden.innerHTML = "";
  const createdHidden = new Set();
  commonFieldDefs.forEach((def) => {
    holder.appendChild(createCommonField(def));
    def.targets.forEach(([prefix, key]) => {
      const name = fullName(prefix, key);
      if (createdHidden.has(name)) return;
      createdHidden.add(name);
      const input = document.createElement("input");
      input.type = "hidden";
      input.name = name;
      hidden.appendChild(input);
    });
  });
  syncCommonFields();
}

function syncCommonFields() {
  commonFieldDefs.forEach((def) => {
    const source = document.querySelector(`[name="${fullName("common", def.key)}"]`);
    const rawValue = source?.value || "";
    def.targets.forEach(([prefix, key, transform]) => {
      const target = document.querySelector(`[name="${fullName(prefix, key)}"]`);
      if (target) target.value = transform ? transform(rawValue) : rawValue;
    });
  });
}

function createField(prefix, key) {
  const label = document.createElement("label");
  const suffix = prefix === "agency" ? "（委托代理）" : (prefix === "trade" ? "（贸易合同）" : "（附件2）");
  const fieldLabel = `${labelFor(prefix, key)}${suffix}`;
  label.textContent = fieldLabel;
  if (wideTextFields.has(key)) label.classList.add("wide");

  const isTextArea = wideTextFields.has(key);
  const input = document.createElement(isTextArea ? "textarea" : "input");
  input.name = fullName(prefix, key);
  if (isTextArea) input.rows = 2;
  input.value = defaultsFor(prefix)[key] || "";
  input.addEventListener("input", refreshPreview);
  label.appendChild(input);
  return label;
}

function renderFields(prefix, holderId) {
  const holder = document.querySelector(holderId);
  holder.innerHTML = "";
  const defaults = defaultsFor(prefix);
  Object.keys(defaults).forEach((key) => {
    if (syncedTargetKeys.has(fullName(prefix, key))) return;
    if ((prefix === "trade" || prefix === "appendix") && tradeGoodsFields.has(key)) return;
    if (prefix === "appendix" && key === "goods_items_json") return;
    holder.appendChild(createField(prefix, key));
  });
  if (prefix === "trade") {
    holder.appendChild(createTradeGoodsBlock());
  }
}

function createTradeGoodsBlock() {
  const wrap = document.createElement("div");
  wrap.className = "wide goods-wrapper";
  wrap.innerHTML = `
    <input type="hidden" name="trade__goods_items_json" id="tradeGoodsItemsJson">
    <input type="hidden" name="appendix__goods_items_json" id="appendixGoodsItemsJson">
    <div class="section-title-row goods-title-row">
      <div>
        <h3>俄英贸易合同：商品明细</h3>
        <span>可添加多项商品，导出时每项商品生成一行。</span>
      </div>
      <button type="button" class="ghost" id="addTradeGoodsBtn">添加新产品</button>
    </div>
    <div id="tradeGoodsItems" class="goods-items"></div>
  `;
  setTimeout(() => {
    document.querySelector("#addTradeGoodsBtn")?.addEventListener("click", () => {
      createTradeGoodsItem();
      refreshPreview();
    });
    createTradeGoodsItem({
      goods_name_ru: tradeDefaults.goods_name_ru || "",
      goods_name_en: tradeDefaults.goods_name_en || "",
      goods_qty: tradeDefaults.goods_qty || "",
      goods_unit_price: tradeDefaults.goods_unit_price || "",
      goods_total: tradeDefaults.goods_total || ""
    });
  });
  return wrap;
}

function createTradeGoodsItem(item = {}) {
  const holder = document.querySelector("#tradeGoodsItems");
  if (!holder) return;
  const card = document.createElement("div");
  card.className = "goods-item";
  card.innerHTML = `
    <div class="goods-item-head">
      <div class="goods-item-title">产品 ${holder.children.length + 1}</div>
      <button type="button" class="goods-item-remove">删除</button>
    </div>
    <div class="generated-fields">
      <label class="wide">商品名称（俄文）<textarea name="trade_goods_name_ru" rows="2"></textarea></label>
      <label class="wide">商品名称（英文）<textarea name="trade_goods_name_en" rows="2"></textarea></label>
      <label>数量<input name="trade_goods_qty"></label>
      <label>单价（CNY）<input name="trade_goods_unit_price"></label>
      <label>总价（CNY）<input name="trade_goods_total"></label>
    </div>
  `;
  const map = {
    trade_goods_name_ru: "goods_name_ru",
    trade_goods_name_en: "goods_name_en",
    trade_goods_qty: "goods_qty",
    trade_goods_unit_price: "goods_unit_price",
    trade_goods_total: "goods_total"
  };
  Object.entries(map).forEach(([domName, key]) => {
    const el = card.querySelector(`[name="${domName}"]`);
    el.value = item[key] || "";
    el.addEventListener("input", () => {
      syncTradeGoods();
      refreshPreview();
    });
  });
  card.querySelector(".goods-item-remove").addEventListener("click", () => {
    if (holder.children.length <= 1) {
      card.querySelectorAll("input, textarea").forEach((el) => { el.value = ""; });
    } else {
      card.remove();
    }
    renumberTradeGoods();
    syncTradeGoods();
    refreshPreview();
  });
  holder.appendChild(card);
  renumberTradeGoods();
  syncTradeGoods();
}

function renumberTradeGoods() {
  document.querySelectorAll(".goods-item").forEach((card, idx) => {
    card.querySelector(".goods-item-title").textContent = `产品 ${idx + 1}`;
  });
}

function getTradeGoodsItems() {
  return Array.from(document.querySelectorAll("#tradeGoodsItems .goods-item")).map((card) => ({
    goods_name_ru: card.querySelector('[name="trade_goods_name_ru"]')?.value?.trim() || "",
    goods_name_en: card.querySelector('[name="trade_goods_name_en"]')?.value?.trim() || "",
    goods_qty: card.querySelector('[name="trade_goods_qty"]')?.value?.trim() || "",
    goods_unit_price: card.querySelector('[name="trade_goods_unit_price"]')?.value?.trim() || "",
    goods_total: card.querySelector('[name="trade_goods_total"]')?.value?.trim() || ""
  })).filter((item) => Object.values(item).some(Boolean));
}

function syncTradeGoods() {
  const json = JSON.stringify(getTradeGoodsItems());
  const tradeHidden = document.querySelector("#tradeGoodsItemsJson");
  const appendixHidden = document.querySelector("#appendixGoodsItemsJson");
  if (tradeHidden) tradeHidden.value = json;
  if (appendixHidden) appendixHidden.value = json;
}

function fillDefaults() {
  renderCommonFields();
  renderFields("agency", "#agencyFields");
  renderFields("trade", "#tradeFields");
  renderFields("appendix", "#appendixFields");
  setTimeout(() => {
    syncCommonFields();
    refreshPreview();
  });
}

function refreshPreview() {
  syncCommonFields();
  syncTradeGoods();
  const holder = document.querySelector("#previewList");
  holder.innerHTML = "";
  previewFields.forEach(([prefix, key]) => {
    const input = document.querySelector(`[name="${fullName(prefix, key)}"]`);
    const value = input?.value?.trim() || "未填写";
    const item = document.createElement("div");
    item.className = "preview-item";
    item.innerHTML = `<span>${labelFor(prefix, key)}</span><strong></strong>`;
    item.querySelector("strong").textContent = value;
    holder.appendChild(item);
  });
  const goods = getTradeGoodsItems();
  const goodsItem = document.createElement("div");
  goodsItem.className = "preview-item";
  goodsItem.innerHTML = `<span>贸易合同商品数量</span><strong>${goods.length || 0} 项</strong>`;
  holder.appendChild(goodsItem);
}

function normalizeSmartKey(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/[：:＝=№#]/g, "")
    .replace(/[（）()【】\[\]\/\\_\-\s]/g, "")
    .trim();
}

const aliasMap = [
  [["协议号", "协议编号"], [["agency", "agreement_no"]]],
  [["签订日期"], [["agency", "sign_date"]]],
  [["签订地点"], [["agency", "sign_place"]]],
  [["贸易合同号", "贸易合同编号", "合同编号", "合同号"], [["common", "contract_no"]]],
  [["合同日期俄文", "俄文日期"], [["common", "contract_date_ru"]]],
  [["合同日期英文", "英文日期", "贸易合同日期"], [["common", "contract_date_en"]]],
  [["卖方名称", "乙方名称"], [["common", "seller_name"]]],
  [["买方名称", "甲方名称"], [["common", "buyer_name_en"]]],
  [["买方俄文名称", "甲方俄文名称"], [["common", "buyer_name_ru"]]],
  [["账户号", "账号"], [["common", "seller_account_no"]]],
  [["swift"], [["common", "seller_swift"]]],
  [["开户银行", "银行名称"], [["common", "seller_bank_en"]]],
  [["开户银行俄文", "银行名称俄文"], [["common", "seller_bank_ru"]]],
  [["收款人名称", "账户名称"], [["common", "seller_beneficiary"]]],
  [["商品名称俄文", "货物俄文"], [["trade_goods", "goods_name_ru"]]],
  [["商品名称英文", "货物英文", "goods"], [["trade_goods", "goods_name_en"]]],
  [["数量", "qty"], [["trade_goods", "goods_qty"]]],
  [["单价", "unitprice"], [["trade_goods", "goods_unit_price"]]],
  [["总价", "total"], [["trade_goods", "goods_total"]]]
];

function findSmartTargets(rawKey) {
  const key = normalizeSmartKey(rawKey);
  if (!key) return [];
  for (const [aliases, targets] of aliasMap) {
    if (aliases.some((alias) => normalizeSmartKey(alias) === key)) return targets;
  }
  for (const [aliases, targets] of aliasMap) {
    if (aliases.some((alias) => key.includes(normalizeSmartKey(alias)))) return targets;
  }
  const matches = [];
  commonFieldDefs.forEach((def) => {
    if (normalizeSmartKey(def.label) === key) matches.push(["common", def.key]);
  });
  for (const [prefix, labels] of [["agency", agencyLabels], ["trade", tradeLabels], ["appendix", appendixLabels]]) {
    Object.entries(labels).forEach(([field, label]) => {
      if (normalizeSmartKey(label) === key) matches.push([prefix, field]);
    });
  }
  return matches;
}

function splitSmartLine(line) {
  const trimmed = line.trim();
  if (!trimmed) return null;
  const exact = trimmed.match(/^([^:：=＝]{1,40})\s*[:：=＝]\s*(.+)$/);
  if (exact) return [exact[1].trim(), exact[2].trim()];
  return null;
}

function applySmartPaste() {
  const input = document.querySelector("#smartPasteInput");
  const status = document.querySelector("#smartPasteStatus");
  if (!input.value.trim()) {
    status.textContent = "请先把需要识别的信息粘贴到上方文本框。";
    return;
  }
  let count = 0;
  const pendingGoods = {};
  input.value.replace(/\r/g, "").split("\n").forEach((line) => {
    const pair = splitSmartLine(line);
    if (!pair) return;
    const [rawKey, rawValue] = pair;
    const value = rawValue.trim().replace(/^["“”']|["“”']$/g, "");
    const targets = findSmartTargets(rawKey);
    targets.forEach(([prefix, field]) => {
      if (prefix === "trade_goods") {
        pendingGoods[field] = value;
        count += 1;
        return;
      }
      const el = document.querySelector(`[name="${fullName(prefix, field)}"]`);
      if (el) {
        el.value = value;
        if (prefix === "common") syncCommonFields();
        count += 1;
      }
    });
  });
  if (Object.keys(pendingGoods).length) {
    const first = document.querySelector("#tradeGoodsItems .goods-item");
    if (!first) createTradeGoodsItem();
    const target = document.querySelector("#tradeGoodsItems .goods-item");
    const map = {
      goods_name_ru: "trade_goods_name_ru",
      goods_name_en: "trade_goods_name_en",
      goods_qty: "trade_goods_qty",
      goods_unit_price: "trade_goods_unit_price",
      goods_total: "trade_goods_total"
    };
    Object.entries(pendingGoods).forEach(([key, value]) => {
      const el = target.querySelector(`[name="${map[key]}"]`);
      if (el) el.value = value;
    });
  }
  syncCommonFields();
  syncTradeGoods();
  refreshPreview();
  status.textContent = count ? `已识别并填入 ${count} 个位置。` : "没有识别到可填字段，请尽量使用“字段名：值”的格式。";
}

async function init() {
  const res = await fetch("/defaults");
  const data = await res.json();
  agencyDefaults = data.agency || {};
  tradeDefaults = data.trade || {};
  appendixDefaults = data.appendix || {};
  agencyLabels = data.agency_labels || {};
  tradeLabels = data.trade_labels || {};
  appendixLabels = data.appendix_labels || {};
  fillDefaults();

  document.querySelector("#resetBtn").addEventListener("click", fillDefaults);
  document.querySelector("#smartParseBtn").addEventListener("click", applySmartPaste);

  document.querySelector("#combinedForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    syncCommonFields();
    syncTradeGoods();
    const btn = document.querySelector("#submitBtn");
    const oldText = btn.textContent;
    btn.disabled = true;
    btn.textContent = "正在生成三份 PDF...";
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
      const filename = match ? decodeURIComponent(match[1]) : "三份合同PDF.zip";
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
