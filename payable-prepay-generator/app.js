(() => {
  const form = document.getElementById('payableForm');
  const statusEl = document.getElementById('status');
  const generateBtn = document.getElementById('generate-btn');
  const resetBtn = document.getElementById('resetBtn');
  const previewList = document.getElementById('previewList');
  const serviceItems = document.getElementById('serviceItems');
  const addServiceBtn = document.getElementById('addServiceBtn');
  const serviceItemsJson = document.getElementById('serviceItemsJson');

  let defaults = {};

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

  resetBtn.addEventListener('click', () => {
    form.reset();
    Object.entries(defaults).forEach(([k, v]) => {
      const el = form.elements[k];
      if (el) el.value = v;
    });
    serviceItems.innerHTML = '';
    updatePreview();
  });

  const previewMap = {
    contract_no: '协议编号',
    sign_date_zh: '签订日期（中文）',
    sign_place_zh: '签订地点（中文）',
    partyA_name_zh: '甲方名称',
    partyB_name_zh: '乙方名称',
    trade_seller_name: '贸易出口商',
    fee_annual_percent: '年化费率（%）',
    fee_overdue_percent: '逾期年化（%）',
    contact_china_name: '中国联络人',
    total_cny: '合计金额',
    sign_year: '签署年份',
  };

  function updatePreview() {
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

  // ---------- 服务费用明细行 ----------
  let serviceSeq = 0;

  function buildServiceRow(initial) {
    const row = document.createElement('div');
    row.className = 'good-row';
    serviceSeq += 1;
    row.innerHTML = `
      <div class="good-no">${serviceSeq}</div>
      <label>服务费用类别<input data-key="category" placeholder="例：资金占用费" /></label>
      <label>服务费用金额<input data-key="amount" placeholder="例：5000.00" /></label>
      <button type="button" class="remove" title="删除该行">×</button>
    `;
    if (initial) {
      Object.entries(initial).forEach(([k, v]) => {
        const input = row.querySelector(`input[data-key="${k}"]`);
        if (input) input.value = v;
      });
    }
    row.querySelector('button.remove').addEventListener('click', () => {
      row.remove();
      renumberService();
    });
    return row;
  }

  function renumberService() {
    serviceItems.querySelectorAll('.good-row').forEach((row, idx) => {
      const no = row.querySelector('.good-no');
      if (no) no.textContent = idx + 1;
    });
    serviceSeq = serviceItems.querySelectorAll('.good-row').length;
  }

  addServiceBtn.addEventListener('click', () => {
    serviceItems.appendChild(buildServiceRow());
  });

  function collectServiceItems() {
    const items = [];
    serviceItems.querySelectorAll('.good-row').forEach(row => {
      const item = {};
      row.querySelectorAll('input[data-key]').forEach(input => {
        item[input.dataset.key] = input.value.trim();
      });
      if (Object.values(item).some(v => v)) items.push(item);
    });
    return items;
  }

  function setStatus(text, kind) {
    statusEl.textContent = text;
    statusEl.className = kind || '';
  }

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    setStatus('正在生成 PDF，请稍候（约 10-30 秒）…');
    generateBtn.disabled = true;

    const formData = new FormData(form);
    formData.append('service_items_json', JSON.stringify(collectServiceItems()));
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
      const filename = match ? decodeURIComponent(match[1]) : '应付账款垫付协议.pdf';
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
      generateBtn.disabled = false;
    }
  });
})();
