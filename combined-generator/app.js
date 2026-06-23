/**
 * combined-generator app.js
 * 合并文档生成页面 - 用户交互处理
 *
 * 功能：
 * 1. 页面加载时从 /defaults 获取默认值和字段标签
 * 2. 用默认值填充所有表单字段
 * 3. 处理表单提交：以 URL-encoded 格式 POST 数据到 /generate
 * 4. 接收 ZIP 响应（Content-Type: application/zip）并触发下载
 * 5. 支持取消/中止正在进行的提交请求（再次点击或超时触发）
 * 6. 生成按钮显示加载状态
 * 7. 实时预览关键字段（importantPreviewFields 模式）
 * 8. 5 分钟超时
 * 9. 优雅的错误处理
 */

// ============================================================
// 共享字段映射表 (Shared Field Map)
// 每个组的第一个字段是"主字段"(master)，用户在 Step 0 填写。
// 当主字段值变化时，自动同步到同组所有其他字段。
// ============================================================
const SHARED_FIELD_MAP = {
  "shared_seller_name_en": [
    ["doc8088", "party_b_name_en"],
    ["doc8089", "seller_name"],
    ["doc8091", "seller_name"],
    ["doc8094", "seller_name"],
  ],
  "shared_seller_name_ru": [
    ["doc8089", "seller_name"],
    ["doc8091", "seller_name"],
    ["doc8094", "seller_name"],
  ],
  "shared_seller_address_en": [
    ["doc8088", "party_b_address"],
    ["doc8089", "seller_address_en"],
    ["doc8091", "seller_address_en"],
  ],
  "shared_seller_address_ru": [
    ["doc8088", "party_b_address_ru"],
    ["doc8089", "seller_address_ru"],
    ["doc8091", "seller_address_ru"],
  ],
  "shared_seller_representative_en": [
    ["doc8088", "party_b_representative"],
    ["doc8089", "seller_representative_en"],
    ["doc8091", "seller_representative_en"],
  ],
  "shared_seller_representative_ru": [
    ["doc8088", "party_b_representative_ru"],
    ["doc8089", "seller_representative_ru"],
    ["doc8091", "seller_representative_ru"],
  ],
  "shared_seller_tel": [
    ["doc8089", "seller_tel"],
    ["doc8091", "seller_tel"],
  ],
  "shared_seller_email": [
    ["doc8089", "seller_email"],
    ["doc8091", "seller_email"],
  ],
  "shared_buyer_name_en": [
    ["doc8088", "party_a_name_en"],
    ["doc8089", "buyer_name_en"],
    ["doc8091", "buyer_name_en"],
    ["doc8093", "partyB_name_zh"],
    ["doc8094", "buyer_name_en"],
  ],
  "shared_buyer_name_ru": [
    ["doc8088", "party_a_name_ru"],
    ["doc8089", "buyer_name_ru"],
    ["doc8091", "buyer_name_ru"],
    ["doc8093", "partyB_name_ru"],
    ["doc8094", "buyer_name_ru"],
  ],
  "shared_buyer_full_name_ru": [
    ["doc8088", "party_a_name_ru"],
    ["doc8089", "buyer_full_name_ru"],
    ["doc8091", "buyer_full_name_ru"],
  ],
  "shared_buyer_address_en": [
    ["doc8088", "party_a_address"],
    ["doc8089", "buyer_address_en"],
    ["doc8091", "buyer_address_en"],
    ["doc8093", "partyB_address_zh"],
  ],
  "shared_buyer_address_ru": [
    ["doc8088", "party_a_address_ru"],
    ["doc8089", "buyer_address_ru"],
    ["doc8091", "buyer_address_ru"],
    ["doc8093", "partyB_address_ru"],
  ],
  "shared_buyer_representative_en": [
    ["doc8088", "party_a_representative"],
    ["doc8089", "buyer_representative_en"],
    ["doc8091", "buyer_representative_en"],
    ["doc8093", "partyB_contact_zh"],
  ],
  "shared_buyer_representative_ru": [
    ["doc8088", "party_a_representative_ru"],
    ["doc8089", "buyer_representative_ru"],
    ["doc8091", "buyer_representative_ru"],
    ["doc8093", "partyB_contact_ru"],
  ],
  "shared_buyer_tel": [
    ["doc8089", "buyer_tel"],
    ["doc8091", "buyer_tel"],
    ["doc8093", "partyB_tel"],
  ],
  "shared_buyer_email": [
    ["doc8089", "buyer_email"],
    ["doc8091", "buyer_email"],
    ["doc8093", "partyB_email"],
  ],
  "shared_buyer_notice_tel": [
    ["doc8089", "buyer_notice_tel"],
    ["doc8091", "buyer_notice_tel"],
  ],
  "shared_buyer_notice_email": [
    ["doc8089", "buyer_notice_email"],
    ["doc8091", "buyer_notice_email"],
  ],
  "shared_agent_name_cn": [
    ["doc8088", "party_c_name_cn"],
    ["doc8092", "partyA_name_zh"],
    ["doc8093", "partyA_name_zh"],
    ["doc8094", "recipient_zh"],
  ],
  "shared_agent_name_ru": [
    ["doc8088", "party_c_name_ru"],
    ["doc8092", "partyA_name_ru"],
    ["doc8093", "partyA_name_ru"],
    ["doc8094", "recipient_ru"],
  ],
  "shared_agent_name_en": [
    ["doc8095", "beneficiary_name"],
  ],
  "shared_agent_address_cn": [
    ["doc8088", "party_c_address"],
    ["doc8093", "partyA_address_zh"],
  ],
  "shared_agent_address_ru": [
    ["doc8088", "party_c_address_ru"],
    ["doc8093", "partyA_address_ru"],
  ],
  "shared_agent_representative_cn": [
    ["doc8088", "party_c_representative"],
    ["doc8093", "partyA_contact_zh"],
    ["doc8093", "partyA_signer"],
  ],
  "shared_agent_representative_ru": [
    ["doc8088", "party_c_representative_ru"],
    ["doc8093", "partyA_contact_ru"],
    ["doc8093", "partyA_signer_ru"],
  ],
  "shared_agent_tel": [
    ["doc8093", "partyA_tel"],
  ],
  "shared_agent_email": [
    ["doc8093", "partyA_email"],
  ],
  "shared_seller_bank_en": [
    ["doc8088", "party_b_bank"],
    ["doc8089", "seller_bank_en"],
    ["doc8091", "seller_bank_en"],
    ["doc8094", "supplier_bank_zh"],
  ],
  "shared_seller_bank_ru": [
    ["doc8088", "party_b_bank_ru"],
    ["doc8089", "seller_bank_ru"],
    ["doc8091", "seller_bank_ru"],
    ["doc8094", "supplier_bank_ru"],
  ],
  "shared_seller_bank_address_en": [
    ["doc8088", "party_b_bank_address"],
    ["doc8089", "seller_bank_address_en"],
    ["doc8091", "seller_bank_address_en"],
  ],
  "shared_seller_bank_address_ru": [
    ["doc8089", "seller_bank_address_ru"],
    ["doc8091", "seller_bank_address_ru"],
  ],
  "shared_seller_swift": [
    ["doc8088", "party_b_swift"],
    ["doc8089", "seller_swift"],
    ["doc8091", "seller_swift"],
  ],
  "shared_seller_account_name": [
    ["doc8088", "party_b_account_name"],
    ["doc8089", "seller_beneficiary"],
    ["doc8091", "seller_beneficiary"],
    ["doc8094", "supplier_payee_name"],
  ],
  "shared_seller_account_no": [
    ["doc8088", "party_b_account_no"],
    ["doc8089", "seller_account_no"],
    ["doc8091", "seller_account_no"],
    ["doc8094", "supplier_account_no"],
  ],
  "shared_buyer_bank_ru": [
    ["doc8089", "buyer_bank_ru"],
    ["doc8091", "buyer_bank_ru"],
  ],
  "shared_buyer_bank_en": [
    ["doc8089", "buyer_bank_en"],
    ["doc8091", "buyer_bank_en"],
  ],
  "shared_buyer_bank_address_ru": [
    ["doc8089", "buyer_bank_address_ru"],
    ["doc8091", "buyer_bank_address_ru"],
  ],
  "shared_buyer_bank_address_en": [
    ["doc8089", "buyer_bank_address_en"],
    ["doc8091", "buyer_bank_address_en"],
  ],
  "shared_buyer_account_no": [
    ["doc8089", "buyer_account_no"],
    ["doc8091", "buyer_account_no"],
  ],
  "shared_buyer_swift": [
    ["doc8089", "buyer_swift"],
    ["doc8091", "buyer_swift"],
  ],
  "shared_buyer_corr_account": [
    ["doc8089", "buyer_corr_account"],
    ["doc8091", "buyer_corr_account"],
  ],
  "shared_buyer_bic": [
    ["doc8089", "buyer_bic"],
    ["doc8091", "buyer_bic"],
  ],
  "shared_trade_contract_no": [
    ["doc8088", "trade_contract_no"],
    ["doc8089", "contract_no"],
    ["doc8091", "contract_no"],
    ["doc8094", "trade_contract_no"],
  ],
  "shared_trade_contract_date_ru": [
    ["doc8088", "trade_contract_date_ru"],
    ["doc8089", "contract_date_ru"],
    ["doc8091", "contract_date_ru"],
  ],
  "shared_trade_contract_date_en": [
    ["doc8088", "trade_contract_date_en"],
    ["doc8089", "contract_date_en"],
    ["doc8091", "contract_date_en"],
  ],
  "shared_appendix_no": [
    ["doc8088", "appendix_no"],
    ["doc8089", "appendix_no"],
    ["doc8091", "appendix_no"],
  ],
  "shared_sign_date_zh": [
    ["doc8092", "sign_date_zh"],
    ["doc8093", "sign_date_zh"],
  ],
  "shared_sign_date_ru": [
    ["doc8092", "sign_date_ru"],
    ["doc8093", "sign_date_ru"],
  ],
  "shared_sign_place_zh": [
    ["doc8088", "sign_place"],
    ["doc8092", "sign_place_zh"],
    ["doc8093", "sign_place_zh"],
  ],
  "shared_sign_place_ru": [
    ["doc8092", "sign_place_ru"],
    ["doc8093", "sign_place_ru"],
  ],
  "shared_sign_year": [
    ["doc8092", "sign_year"],
    ["doc8093", "sign_year"],
    ["doc8094", "sign_year"],
  ],
  "shared_sign_month": [
    ["doc8092", "sign_month"],
    ["doc8093", "sign_month"],
    ["doc8094", "sign_month"],
  ],
  "shared_sign_day": [
    ["doc8092", "sign_day"],
    ["doc8093", "sign_day"],
    ["doc8094", "sign_day"],
  ],
  "shared_delivery_terms_ru": [
    ["doc8089", "delivery_terms_ru"],
    ["doc8089", "form_delivery_terms_ru"],
    ["doc8091", "delivery_terms_ru"],
    ["doc8091", "form_delivery_terms_ru"],
  ],
  "shared_delivery_terms_en": [
    ["doc8089", "delivery_terms_en"],
    ["doc8089", "form_delivery_terms_en"],
    ["doc8091", "delivery_terms_en"],
    ["doc8091", "form_delivery_terms_en"],
  ],
  "shared_delivery_time_ru": [
    ["doc8089", "delivery_time_ru"],
    ["doc8089", "form_delivery_time_ru"],
    ["doc8091", "delivery_time_ru"],
    ["doc8091", "form_delivery_time_ru"],
  ],
  "shared_delivery_time_en": [
    ["doc8089", "delivery_time_en"],
    ["doc8089", "form_delivery_time_en"],
    ["doc8091", "delivery_time_en"],
    ["doc8091", "form_delivery_time_en"],
  ],
  "shared_payment_advance_percent": [
    ["doc8089", "payment_advance_percent"],
    ["doc8091", "payment_advance_percent"],
  ],
  "shared_payment_advance_amount": [
    ["doc8089", "payment_advance_amount"],
    ["doc8091", "payment_advance_amount"],
  ],
  "shared_payment_advance_amount_ru": [
    ["doc8089", "payment_advance_amount_ru"],
    ["doc8091", "payment_advance_amount_ru"],
  ],
  "shared_payment_advance_amount_en": [
    ["doc8089", "payment_advance_amount_en"],
    ["doc8091", "payment_advance_amount_en"],
  ],
  "shared_payment_balance_percent": [
    ["doc8089", "payment_balance_percent"],
    ["doc8091", "payment_balance_percent"],
  ],
  "shared_payment_balance_amount": [
    ["doc8089", "payment_balance_amount"],
    ["doc8091", "payment_balance_amount"],
  ],
  "shared_payment_balance_amount_ru": [
    ["doc8089", "payment_balance_amount_ru"],
    ["doc8091", "payment_balance_amount_ru"],
  ],
  "shared_payment_balance_amount_en": [
    ["doc8089", "payment_balance_amount_en"],
    ["doc8091", "payment_balance_amount_en"],
  ],
  "shared_payment_bank_days": [
    ["doc8089", "payment_bank_days"],
    ["doc8091", "payment_bank_days"],
  ],
  "shared_payment_bank_days_ru": [
    ["doc8089", "payment_bank_days_ru"],
    ["doc8091", "payment_bank_days_ru"],
  ],
  "shared_payment_bank_days_en": [
    ["doc8089", "payment_bank_days_en"],
    ["doc8091", "payment_bank_days_en"],
  ],
  "shared_delivery_notice_days": [
    ["doc8089", "delivery_notice_days"],
    ["doc8091", "delivery_notice_days"],
  ],
  "shared_delivery_notice_days_ru": [
    ["doc8089", "delivery_notice_days_ru"],
    ["doc8091", "delivery_notice_days_ru"],
  ],
  "shared_delivery_notice_days_en": [
    ["doc8089", "delivery_notice_days_en"],
    ["doc8091", "delivery_notice_days_en"],
  ],
  "shared_fee_percent": [
    ["doc8088", "agent_fee_percent"],
    ["doc8092", "fee_percent"],
  ],
  "shared_fee_annual_percent": [
    ["doc8093", "fee_annual_percent"],
    ["doc8094", "fee_annual_percent"],
  ],
  "shared_fee_overdue_percent": [
    ["doc8093", "fee_overdue_percent"],
    ["doc8094", "fee_overdue_percent"],
  ],
  "shared_trade_seller_name": [
    ["doc8088", "party_b_name_en"],
    ["doc8089", "seller_name"],
    ["doc8093", "trade_seller_name"],
  ],
  "shared_form_appendix_no": [
    ["doc8089", "form_appendix_no"],
    ["doc8091", "form_appendix_no"],
  ],
  "shared_form_contract_no": [
    ["doc8089", "form_contract_no"],
    ["doc8091", "form_contract_no"],
  ],
  "shared_form_seller_name_ru": [
    ["doc8089", "form_seller_name_ru"],
    ["doc8091", "form_seller_name_ru"],
  ],
  "shared_form_seller_name_en": [
    ["doc8089", "form_seller_name_en"],
    ["doc8091", "form_seller_name_en"],
  ],
  "shared_form_seller_country_ru": [
    ["doc8089", "form_seller_country_ru"],
    ["doc8091", "form_seller_country_ru"],
  ],
  "shared_form_seller_country_en": [
    ["doc8089", "form_seller_country_en"],
    ["doc8091", "form_seller_country_en"],
  ],
  "shared_form_seller_director_ru": [
    ["doc8089", "form_seller_director_ru"],
    ["doc8091", "form_seller_director_ru"],
  ],
  "shared_form_seller_director_en": [
    ["doc8089", "form_seller_director_en"],
    ["doc8091", "form_seller_director_en"],
  ],
  "shared_form_buyer_signer_ru": [
    ["doc8089", "form_buyer_signer_ru"],
    ["doc8091", "form_buyer_signer_ru"],
  ],
  "shared_form_seller_signer_ru": [
    ["doc8089", "form_seller_signer_ru"],
    ["doc8091", "form_seller_signer_ru"],
  ],
  "shared_goods_name_ru": [
    ["doc8088", "goods_ru"],
    ["doc8089", "goods_name_ru"],
    ["doc8091", "goods_name_ru"],
  ],
  "shared_goods_name_en": [
    ["doc8088", "goods_cn"],
    ["doc8089", "goods_name_en"],
    ["doc8091", "goods_name_en"],
  ],
  "shared_total_price_ru": [
    ["doc8089", "total_price_ru"],
    ["doc8091", "total_price_ru"],
  ],
  "shared_total_price_en": [
    ["doc8089", "total_price_en"],
    ["doc8091", "total_price_en"],
  ],
  "shared_package_includes_ru": [
    ["doc8089", "package_includes_ru"],
    ["doc8091", "package_includes_ru"],
  ],
  "shared_package_includes_en": [
    ["doc8089", "package_includes_en"],
    ["doc8091", "package_includes_en"],
  ],
  "shared_guarantee_no": [
    ["doc8095", "guarantee_no"],
  ],
  "shared_guarantee_issue_date": [
    ["doc8095", "issue_date"],
  ],
  "shared_buyer_signature_ru": [
    ["doc8089", "buyer_signature_ru"],
    ["doc8091", "buyer_signature_ru"],
  ],
  "shared_buyer_signature_en": [
    ["doc8089", "buyer_signature_en"],
    ["doc8091", "buyer_signature_en"],
  ],
  "shared_account_company": [
    ["doc8092", "account_company"],
  ],
  "shared_account_bank": [
    ["doc8092", "account_bank"],
  ],
  "shared_account_no": [
    ["doc8092", "account_no"],
  ],
};

// ============================================================
// 共享字段客户端同步逻辑
// ============================================================

/**
 * 获取共享组的主字段全名 (prefix_field 格式)
 */
function getMasterFieldName(sharedGroupName) {
  const group = SHARED_FIELD_MAP[sharedGroupName];
  if (!group || group.length === 0) return null;
  return group[0][0] + "_" + group[0][1];
}

/**
 * 将主字段的值同步到同组所有其他字段
 */
function syncSharedGroup(sharedGroupName) {
  const group = SHARED_FIELD_MAP[sharedGroupName];
  if (!group || group.length === 0) return;

  const masterFullName = group[0][0] + "_" + group[0][1];
  const masterInput = document.querySelector(`[name="${CSS.escape(masterFullName)}"]`);
  if (!masterInput) return;

  const value = masterInput.value;

  // 同步到同组所有其他字段（跳过主字段自身）
  for (let i = 1; i < group.length; i++) {
    const [prefix, fieldName] = group[i];
    const fullName = prefix + "_" + fieldName;
    const targetInput = document.querySelector(`[name="${CSS.escape(fullName)}"]`);
    if (targetInput) {
      targetInput.value = value;
    }
  }
}

/**
 * 对所有共享组执行一次同步（页面加载后调用）
 */
function syncAllSharedFields() {
  for (const groupName of Object.keys(SHARED_FIELD_MAP)) {
    syncSharedGroup(groupName);
  }
}

/**
 * 为所有主字段绑定 input 事件监听器
 */
function bindSharedFieldListeners() {
  for (const groupName of Object.keys(SHARED_FIELD_MAP)) {
    const masterFullName = getMasterFieldName(groupName);
    if (!masterFullName) continue;
    const masterInput = document.querySelector(`[name="${CSS.escape(masterFullName)}"]`);
    if (masterInput) {
      masterInput.addEventListener("input", () => {
        syncSharedGroup(groupName);
        refreshPreview();
      });
    }
  }
}

/**
 * 隐藏步骤 1-7 中的从属共享字段（slave fields）
 * 规则：
 * 1. 对于 SHARED_FIELD_MAP 中的每个组，如果主字段在 Step 0 中，
 *    则隐藏该组在所有步骤 1-7 中的从属字段。
 * 2. 此外，如果 Step 1-7 中有任何字段名与 Step 0 中的字段名完全相同
 *    （即同一字段在 Step 0 和后续步骤中重复出现），也隐藏后续步骤中的该字段。
 * 仍保留在 DOM 中，确保表单提交时数据完整。
 */
function hideSharedSlaveFields() {
  // 先收集 Step 0 中所有字段名
  const step0 = document.querySelector('[data-wizard-step="0"]');
  const step0FieldNames = new Set();
  if (step0) {
    step0.querySelectorAll('input, textarea').forEach(el => {
      if (el.name) step0FieldNames.add(el.name);
    });
  }

  // ---- 规则 1：按 SHARED_FIELD_MAP 隐藏跨文档的从属字段 ----
  for (const groupName of Object.keys(SHARED_FIELD_MAP)) {
    const group = SHARED_FIELD_MAP[groupName];
    if (!group || group.length <= 1) continue;

    // 检查主字段是否在 Step 0 中
    const masterFullName = group[0][0] + "_" + group[0][1];
    if (!step0FieldNames.has(masterFullName)) continue;

    // 从索引 1 开始，跳过主字段（master）
    for (let i = 1; i < group.length; i++) {
      const [prefix, fieldName] = group[i];
      const fullName = prefix + "_" + fieldName;
      const targetInput = document.querySelector(`[name="${CSS.escape(fullName)}"]`);
      if (targetInput) {
        const label = targetInput.closest("label");
        if (label) {
          label.classList.add("hidden-shared");
        }
      }
    }
  }

  // ---- 规则 2：隐藏步骤 1-7 中与 Step 0 同名的重复字段 ----
  for (let stepIdx = 1; stepIdx <= 7; stepIdx++) {
    const step = document.querySelector(`[data-wizard-step="${stepIdx}"]`);
    if (!step) continue;
    step.querySelectorAll('input, textarea').forEach(el => {
      if (el.name && step0FieldNames.has(el.name)) {
        const label = el.closest("label");
        if (label) {
          label.classList.add("hidden-shared");
        }
      }
    });
  }
}

// ============================================================
// 重要预览字段列表（importantPreviewFields 模式）
// 沿用现有 app.js 的命名约定：prefix__field_name
// ============================================================
const importantPreviewFields = [
  // ---- 委托代理协议关键字段 ----
  "agency__agreement_no",
  "agency__sign_date",
  "agency__party_a_name_en",
  "agency__party_a_address",
  "agency__party_a_representative",
  "agency__party_b_name_en",
  "agency__party_b_address",
  "agency__party_b_representative",
  "agency__trade_contract_no",
  "agency__trade_contract_date_en",
  "agency__party_b_bank",
  "agency__party_b_swift",
  "agency__party_b_account_no",

  // ---- 俄英贸易合同关键字段 ----
  "trade__contract_no",
  "trade__contract_date_en",
  "trade__contract_date_ru",
  "trade__buyer_name_en",
  "trade__buyer_name_ru",
  "trade__seller_name",
  "trade__seller_address_en",
  "trade__total_price_en",
  "trade__total_price_ru",
  "trade__seller_bank_en",
  "trade__seller_swift",
  "trade__seller_account_no",
  "trade__delivery_terms_en",
  "trade__delivery_terms_ru",

  // ---- 附件2 关键字段 ----
  "appendix__appendix_no",
  "appendix__contract_no",
  "appendix__buyer_name_en",
  "appendix__seller_name",
  "appendix__total_price_en",
  "appendix__delivery_terms_en",

  // ---- 其他可选文档预留 ----
  "settlement__contract_no",
  "payable__contract_no",
  "guarantee__guarantee_no"
];

let defaults = {};
let labels = {};
let currentAbortController = null;

// ============================================================
// 根据字段全名查询可读标签
// ============================================================
function labelFor(name) {
  // 优先匹配扁平 labels（如 labels["agency__agreement_no"]）
  if (labels[name]) return labels[name];
  // 其次拆解 prefix__key，查 labels[prefix][key]
  const parts = name.split("__");
  if (parts.length === 2) {
    const prefix = parts[0];
    const key = parts[1];
    if (labels[prefix] && labels[prefix][key]) return labels[prefix][key];
  }
  // 后备：直接返回字段名
  return name;
}

// ============================================================
// 用默认值填充所有表单字段
// ============================================================
function fillDefaults() {
  document.querySelectorAll("[name]").forEach((el) => {
    const value = defaults[el.name] ?? "";
    el.value = value;
  });
  refreshPreview();
}

// ============================================================
// 刷新预览面板 — 展示 importantPreviewFields 中每个字段的当前值
// ============================================================
function refreshPreview() {
  const holder = document.querySelector("#previewList");
  if (!holder) return;
  holder.innerHTML = "";

  importantPreviewFields.forEach((name) => {
    const input = document.querySelector(`[name="${CSS.escape(name)}"]`);
    const value = input?.value?.trim() || "";
    const item = document.createElement("div");
    item.className = "preview-item";
    const labelSpan = document.createElement("span");
    labelSpan.textContent = labelFor(name);
    const valueStrong = document.createElement("strong");
    valueStrong.textContent = value || "未填写";
    if (!value) valueStrong.style.opacity = "0.45";
    item.appendChild(labelSpan);
    item.appendChild(valueStrong);
    holder.appendChild(item);
  });
}

// ============================================================
// 获取可读的错误消息
// ============================================================
function friendlyError(err) {
  if (err.name === "AbortError") return "请求已被取消";
  if (err.message) return err.message;
  return String(err);
}

// ============================================================
// 初始化
// ============================================================
async function init() {
  let data;

  // 1. 加载默认值
  try {
    const res = await fetch("/defaults");
    if (!res.ok) {
      const text = await res.text().catch(() => "");
      throw new Error(text || `HTTP ${res.status} — 无法加载表单默认值`);
    }
    data = await res.json();
  } catch (err) {
    throw new Error(`加载默认值失败：${friendlyError(err)}`);
  }

  // 2. 解析数据结构
  //    支持两种格式：
  //      a) 扁平格式：{ defaults: { field: val, ... }, labels: { field: label, ... } }
  //      b) 分文档格式：{ agency: { ... }, trade: { ... }, appendix: { ... },
  //                      agency_labels: { ... }, trade_labels: { ... }, appendix_labels: { ... } }
  defaults = data.defaults || {};
  labels = data.labels || {};

  // 若存在 per-document 数据，合并展开为 prefix__key 格式
  const knownPrefixes = ["agency", "trade", "appendix", "settlement", "payable", "guarantee"];
  const subDefaults = {};
  const subLabels = {};

  knownPrefixes.forEach((prefix) => {
    if (data[prefix] && typeof data[prefix] === "object" && !Array.isArray(data[prefix])) {
      Object.entries(data[prefix]).forEach(([key, value]) => {
        subDefaults[`${prefix}__${key}`] = value;
      });
    }
    const labelKey = `${prefix}_labels`;
    if (data[labelKey] && typeof data[labelKey] === "object" && !Array.isArray(data[labelKey])) {
      Object.entries(data[labelKey]).forEach(([key, value]) => {
        subLabels[`${prefix}__${key}`] = value;
      });
    }
  });

  Object.assign(defaults, subDefaults);
  Object.assign(labels, subLabels);

  // 3. 填充表单
  fillDefaults();

  // 3.5 初始同步：将 Step 0 主字段的默认值同步到所有从字段
  syncAllSharedFields();
  bindSharedFieldListeners();

  // 3.6 隐藏步骤 1-7 中的从属共享字段（已在 Step 0 统一填写）
  hideSharedSlaveFields();

  // 4. 绑定实时预览响应
  document.querySelectorAll("input, textarea, select").forEach((el) => {
    el.addEventListener("input", refreshPreview);
    el.addEventListener("change", refreshPreview);
  });

  // 5. 重置按钮
  const resetBtn = document.querySelector("#resetBtn");
  if (resetBtn) {
    resetBtn.addEventListener("click", () => fillDefaults());
  }

  // 6. 表单提交
  const form = document.querySelector("#combinedForm");
  if (!form) throw new Error("未找到 #combinedForm 表单元素");

  form.addEventListener("submit", async (event) => {
    event.preventDefault();

    // 如果已有进行中的请求，先中止它
    if (currentAbortController) {
      currentAbortController.abort();
    }

    currentAbortController = new AbortController();
    const signal = currentAbortController.signal;

    const btn = document.querySelector("#generateAllBtn");
    if (!btn) {
      currentAbortController = null;
      return;
    }

    const oldText = btn.textContent;
    const oldDisabled = btn.disabled;
    btn.disabled = true;
    btn.textContent = "正在生成并打包 ZIP ...";

    // 5 分钟超时（300,000 ms）
    const TIMEOUT_MS = 300000;
    const timeoutId = setTimeout(() => {
      if (currentAbortController) {
        currentAbortController.abort();
        currentAbortController = null;
      }
      btn.disabled = false;
      btn.textContent = "生成超时（5 分钟），请重试";
      setTimeout(() => { btn.textContent = oldText; }, 3000);
    }, TIMEOUT_MS);

    try {
      const response = await fetch("/generate", {
        method: "POST",
        signal,
        body: new URLSearchParams(new FormData(form))
      });

      clearTimeout(timeoutId);

      // 检查 HTTP 状态
      if (!response.ok) {
        const message = await response.text().catch(() => "");
        throw new Error(message || `服务器返回 ${response.status}，生成失败`);
      }

      // 验证响应类型
      const contentType = response.headers.get("Content-Type") || "";
      if (!contentType.includes("zip") && !contentType.includes("octet-stream")) {
        // 不是 ZIP 响应，尝试读取文本作为错误信息
        const text = await response.text().catch(() => "");
        throw new Error(text || "服务器未返回 ZIP 文件，响应格式异常");
      }

      // 获取文件名
      const disposition = response.headers.get("Content-Disposition") || "";
      const match = disposition.match(/filename\*=UTF-8''([^;]+)/);
      const filename = match
        ? decodeURIComponent(match[1])
        : "合并文档.zip";

      // 触发下载
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);

      // 成功反馈
      btn.textContent = "生成完成！";
      setTimeout(() => { btn.textContent = oldText; }, 2000);
    } catch (err) {
      clearTimeout(timeoutId);

      if (err.name === "AbortError") {
        // 用户主动取消 或 超时取消
        btn.textContent = "已取消生成";
      } else {
        // 其他错误
        btn.textContent = "生成失败";
        const errMsg = friendlyError(err);
        console.error("生成请求失败:", err);
        alert(`生成失败：${errMsg}${errMsg.length > 80 ? "\n\n详细错误已记录到控制台。" : ""}`);
      }
      setTimeout(() => { btn.textContent = oldText; }, 3000);
    } finally {
      btn.disabled = false;
      currentAbortController = null;
    }
  });
}

// ============================================================
// 启动应用
// ============================================================
init().catch((err) => {
  console.error("页面初始化失败:", err);
  const msg = friendlyError(err);
  alert(`页面初始化失败：${msg}`);
});