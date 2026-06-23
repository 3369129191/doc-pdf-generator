const importantPreviewFields = [
  "agreement_no",
  "sign_date",
  "sign_place",
  "party_a_name_en",
  "party_a_address",
  "party_a_representative",
  "party_b_name_en",
  "party_b_address",
  "party_b_representative",
  "trade_contract_no",
  "trade_contract_date_en",
  "goods_cn",
  "party_b_bank",
  "party_b_swift",
  "party_b_account_no"
];

let defaults = {};
let labels = {};

const smartPasteAliases = [
  ["agreement_no", ["协议号", "协议编号", "合同编号", "agreement no", "agreement number", "договор №", "номер договора"]],
  ["sign_date", ["签订日期", "签署日期", "日期", "date", "дата подписания"]],
  ["sign_place", ["签订地点", "签署地点", "地点", "place", "место подписания"]],
  ["party_a_name_en", ["甲方英文名称", "甲方名称", "甲方", "付款方", "委托方", "party a", "payer", "principal"]],
  ["party_a_name_ru", ["甲方俄文名称", "甲方俄语名称", "сторона а", "сторона a"]],
  ["party_a_address", ["甲方法律地址", "甲方地址", "付款方地址", "委托方地址", "party a address"]],
  ["party_a_address_ru", ["甲方俄文地址", "甲方俄语地址", "адрес стороны а", "адрес стороны a"]],
  ["party_a_representative", ["甲方代表", "甲方授权代表", "party a representative"]],
  ["party_a_representative_ru", ["甲方俄文代表", "甲方俄语代表", "представитель стороны а", "представитель стороны a"]],
  ["party_b_name_en", ["乙方英文名称", "乙方名称", "乙方", "收款方", "合同相对方", "party b", "recipient", "counterparty"]],
  ["party_b_name_ru", ["乙方俄文名称", "乙方俄语名称", "сторона в", "сторона b"]],
  ["party_b_address", ["乙方法律地址", "乙方地址", "收款方地址", "party b address"]],
  ["party_b_address_ru", ["乙方俄文地址", "乙方俄语地址", "адрес стороны в", "адрес стороны b"]],
  ["party_b_representative", ["乙方代表", "乙方授权代表", "party b representative"]],
  ["party_b_representative_ru", ["乙方俄文代表", "乙方俄语代表", "представитель стороны в", "представитель стороны b"]],
  ["party_c_name_cn", ["丙方中文名称", "丙方名称", "丙方", "代理方"]],
  ["party_c_name_ru", ["丙方俄文名称", "丙方俄语名称", "сторона с", "сторона c", "агент"]],
  ["party_c_address", ["丙方法律地址", "丙方地址", "代理方地址"]],
  ["party_c_address_ru", ["丙方俄文地址", "丙方俄语地址", "адрес стороны с", "адрес стороны c"]],
  ["party_c_representative", ["丙方代表", "丙方授权代表"]],
  ["party_c_representative_ru", ["丙方俄文代表", "丙方俄语代表", "представитель стороны с", "представитель стороны c"]],
  ["trade_contract_no", ["贸易合同号", "贸易合同编号", "贸易合同", "contract no", "контракт №", "номер контракта"]],
  ["trade_contract_no_with_symbol", ["带№的贸易合同号", "带 № 的贸易合同号", "№贸易合同号"]],
  ["trade_contract_date_en", ["贸易合同日期", "合同日期", "contract date"]],
  ["trade_contract_date_ru", ["贸易合同俄文日期", "贸易合同俄语日期", "дата контракта"]],
  ["appendix_no", ["附件编号", "附件号", "appendix", "приложение"]],
  ["goods_cn", ["支付用途", "货物中文", "支付用途/货物中文", "用途", "商品", "货物"]],
  ["goods_ru", ["支付用途俄文", "货物俄文", "支付用途/货物俄文", "товар", "назначение платежа"]],
  ["agent_fee_percent", ["代理费比例", "代理费", "佣金比例", "费率", "agent fee", "комиссия"]],
  ["party_b_bank", ["开户银行", "乙方开户银行", "收款银行", "银行名称", "bank", "банк получателя"]],
  ["party_b_bank_ru", ["开户银行俄文", "开户银行俄语", "乙方开户银行俄文", "название банка"]],
  ["party_b_swift", ["swift", "swift代码", "swift 代码", "银行代码"]],
  ["party_b_bank_address", ["银行地址", "乙方银行地址", "开户行地址", "bank address", "адрес банка"]],
  ["party_b_account_name", ["收款人名称", "账户名称", "户名", "account name", "beneficiary", "наименование получателя"]],
  ["party_b_account_no", ["账户号", "账号", "银行账户", "收款账号", "account no", "account number", "номер счета"]]
];

function fillDefaults() {
  document.querySelectorAll("[name]").forEach((el) => {
    const value = defaults[el.name] ?? "";
    el.value = value;
  });
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
  const exact = trimmed.match(/^([^:：=＝]{1,32})\s*[:：=＝]\s*(.+)$/);
  if (exact) return [exact[1].trim(), exact[2].trim()];
  const loose = trimmed.match(/^([\u4e00-\u9fa5A-Za-zА-Яа-я\s\/（）()]{2,24})\s+(.{2,})$/);
  if (loose) return [loose[1].trim(), loose[2].trim()];
  return null;
}

function parseSmartPasteText(text) {
  const result = {};
  const lines = String(text || "")
    .replace(/\r/g, "")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);

  for (const line of lines) {
    const pair = splitSmartLine(line);
    if (!pair) continue;
    const [rawKey, rawValue] = pair;
    const field = findSmartField(rawKey);
    if (!field) continue;
    let value = rawValue.trim().replace(/^["“”']|["“”']$/g, "");
    if (field === "agent_fee_percent") value = value.replace(/[％%]/g, "").trim();
    if (field === "trade_contract_no") value = value.replace(/^№\s*/i, "").trim();
    if (field === "trade_contract_no_with_symbol" && !/^№/.test(value)) value = `№ ${value}`;
    result[field] = value;
  }

  if (result.trade_contract_no && !result.trade_contract_no_with_symbol) {
    result.trade_contract_no_with_symbol = `№ ${result.trade_contract_no}`;
  }
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
    status.textContent = "没有识别到可填字段，请尽量使用“字段名：值”的格式，例如“协议号：THOS-SMA”。";
    return;
  }
  entries.forEach(([name, value]) => {
    const el = document.querySelector(`[name="${name}"]`);
    el.value = value;
    el.dispatchEvent(new Event("input", { bubbles: true }));
  });
  refreshPreview();
  const names = entries.map(([name]) => labels[name] || name).join("、");
  status.textContent = `已识别并填入 ${entries.length} 个字段：${names}`;
}

async function init() {
  const res = await fetch("/defaults");
  const data = await res.json();
  defaults = data.defaults || {};
  labels = data.labels || {};
  fillDefaults();

  document.querySelector("#resetBtn").addEventListener("click", fillDefaults);
  document.querySelector("#smartParseBtn").addEventListener("click", applySmartPaste);
  document.querySelectorAll("input, textarea").forEach((el) => {
    el.addEventListener("input", refreshPreview);
  });

  document.querySelector("#contractForm").addEventListener("submit", async (event) => {
    event.preventDefault();
    const btn = document.querySelector("#submitBtn");
    const oldText = btn.textContent;
    btn.disabled = true;
    btn.textContent = "正在生成 PDF...";
    try {
      const form = event.currentTarget;
      const response = await fetch("/generate", {
        method: "POST",
        body: new URLSearchParams(new FormData(form))
      });
      if (!response.ok) {
        const message = await response.text();
        throw new Error(message || "生成失败");
      }
      const blob = await response.blob();
      const disposition = response.headers.get("Content-Disposition") || "";
      const match = disposition.match(/filename\*=UTF-8''([^;]+)/);
      const filename = match ? decodeURIComponent(match[1]) : "委托代理协议.pdf";
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
