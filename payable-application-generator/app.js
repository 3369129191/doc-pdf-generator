(() => {
  const form = document.getElementById('applicationForm');
  const statusEl = document.getElementById('status');
  const generateBtn = document.getElementById('generate-btn');
  const resetBtn = document.getElementById('resetBtn');
  const previewList = document.getElementById('previewList');

  let defaults = {};

  // 恢复默认值按钮
  if (resetBtn) {
    resetBtn.addEventListener('click', () => {
      form.reset();
      Object.entries(defaults).forEach(([k, v]) => {
        const el = form.elements[k];
        if (el) el.value = v;
      });
      updatePreview();
    });
  }

  fetch('/defaults')
    .then(r => r.json())
    .then(data => {
      defaults = data.defaults || {};
      Object.entries(defaults).forEach(([k, v]) => {
        const el = form.elements[k];
        if (el && !el.value) el.value = v;
      });
      updatePreview();
    })
    .catch(() => {});

  const previewMap = {
    buyer_name_en: '买方公司名',
    seller_name: '卖方公司名',
    advance_amount: '申请垫付金额',
    advance_term_days: '垫付期限（天）',
    trade_contract_no: '贸易合同号',
    trade_contract_amount: '贸易合同金额',
    invoice_amount: '商业发票金额',
    buyer_confirmed_amount: '买方确认应付金额',
    interest_start_date: '原应付到期日',
    advance_due_date: '垫付到期日',
    fee_annual_percent: '年化费率（%）',
    claim_days: '通知后偿付天数',
    applicant_company_name: '盖章公司',
    legal_rep_name: '法定代表人',
    sign_year: '申请年份',
  };

  function updatePreview() {
    if (!previewList) return;
    previewList.innerHTML = '';
    Object.entries(previewMap).forEach(([key, label]) => {
      const el = form.elements[key];
      const val = el ? el.value.trim() : '';
      if (!val) return;
      const item = document.createElement('div');
      item.className = 'preview-item';
      item.innerHTML = `<span>${label}</span><strong>${val}</strong>`;
      previewList.appendChild(item);
    });
  }

  form.addEventListener('input', updatePreview);

  // 当用户修改"卖方公司名"时，自动同步"供应商收款人"
  // 当用户修改"买方公司名"时，自动同步"申请方公司名（盖章）"
  const sellerInput = form.elements['seller_name'];
  const supplierPayeeInput = form.elements['supplier_payee_name'];
  const buyerInput = form.elements['buyer_name_en'];
  const applicantInput = form.elements['applicant_company_name'];

  if (sellerInput) {
    sellerInput.addEventListener('input', () => {
      if (supplierPayeeInput && (!supplierPayeeInput.dataset.userEdited)) {
        supplierPayeeInput.value = sellerInput.value;
      }
    });
  }
  if (buyerInput) {
    buyerInput.addEventListener('input', () => {
      if (applicantInput && (!applicantInput.dataset.userEdited)) {
        applicantInput.value = buyerInput.value;
      }
    });
  }
  if (supplierPayeeInput) {
    supplierPayeeInput.addEventListener('input', () => {
      supplierPayeeInput.dataset.userEdited = 'true';
    });
  }
  if (applicantInput) {
    applicantInput.addEventListener('input', () => {
      applicantInput.dataset.userEdited = 'true';
    });
  }

  function setStatus(text, kind) {
    if (statusEl) {
      statusEl.textContent = text;
      statusEl.className = kind || '';
    }
  }

  // 绑定 wizard 按钮事件
  const prevBtn = document.getElementById('prevBtn');
  const nextBtn = document.getElementById('nextBtn');

  if (prevBtn) {
    prevBtn.addEventListener('click', () => {
      if (window.Wizard && window.Wizard.prev) window.Wizard.prev();
    });
  }
  if (nextBtn) {
    nextBtn.addEventListener('click', () => {
      if (window.Wizard && window.Wizard.next) window.Wizard.next();
    });
  }

  // 表单提交：生成 PDF
  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    setStatus('正在生成 PDF，请稍候（约 10-30 秒）…');
    if (generateBtn) generateBtn.disabled = true;

    // 在提交前，把 supplier_payee_name 同步为 seller_name（用户没改的话）
    if (sellerInput && supplierPayeeInput && !supplierPayeeInput.dataset.userEdited) {
      supplierPayeeInput.value = sellerInput.value;
    }
    if (buyerInput && applicantInput && !applicantInput.dataset.userEdited) {
      applicantInput.value = buyerInput.value;
    }

    const formData = new FormData(form);
    const params = new URLSearchParams();
    formData.forEach((value, key) => params.append(key, value));

    try {
      const resp = await fetch('/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
        body: params.toString(),
      });
      if (!resp.ok) {
        const text = await resp.text();
        throw new Error(text || `HTTP ${resp.status}`);
      }
      const blob = await resp.blob();
      const disposition = resp.headers.get('Content-Disposition') || '';
      const match = /filename\*=UTF-8''([^;]+)/i.exec(disposition);
      const filename = match ? decodeURIComponent(match[1]) : '应付账款垫付申请书.pdf';
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      setStatus('已生成并开始下载：' + filename, 'success');
    } catch (err) {
      setStatus('生成失败：' + (err.message || err), 'error');
    } finally {
      if (generateBtn) generateBtn.disabled = false;
    }
  });
})();
