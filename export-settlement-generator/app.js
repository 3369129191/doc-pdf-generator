(() => {
  const form = document.getElementById('settlementForm');
  const statusEl = document.getElementById('status');
  const goodsList = document.getElementById('goods-list');
  const addGoodBtn = document.getElementById('add-good');
  const generateBtn = document.getElementById('generate-btn');
  const resetBtn = document.getElementById('resetBtn');
  const previewList = document.getElementById('previewList');

  let defaults = {};

  // ---------- 默认值 ----------
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
    // 清空商品行，重建一行
    goodsList.innerHTML = '';
    goodsList.appendChild(buildGoodRow());
    updatePreview();
  });

  // ---------- 实时预览 ----------
  const previewMap = {
    contract_no: '协议编号',
    sign_date_zh: '签订日期（中文）',
    sign_place_zh: '签订地点（中文）',
    partyA_name_zh: '甲方名称',
    partyB_name_zh: '乙方名称',
    fee_percent: '代理费率',
    account_company: '收款公司',
    partyB_controller: '实际控制人',
    sign_year: '签署日期',
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

  // ---------- 商品行 ----------
  let goodSeq = 0;

  function buildGoodRow(initial) {
    const row = document.createElement('div');
    row.className = 'good-row';
    goodSeq += 1;
    row.innerHTML = `
      <div class="good-no">${countRows() + 1}</div>
      <label>海关编码<input data-key="customs_code" placeholder="例：8517.13.00" /></label>
      <label>中文品名<input data-key="name_zh" placeholder="例：某某型号手机" /></label>
      <label>俄文品名<input data-key="name_ru" placeholder="例：смартфон" /></label>
      <label>数量<input data-key="quantity" placeholder="例：100" /></label>
      <label>金额<input data-key="amount" placeholder="例：USD 12,000" /></label>
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
      renumber();
    });
    return row;
  }

  function countRows() {
    return goodsList.querySelectorAll('.good-row').length;
  }

  function renumber() {
    goodsList.querySelectorAll('.good-row').forEach((row, idx) => {
      const no = row.querySelector('.good-no');
      if (no) no.textContent = idx + 1;
    });
  }

  addGoodBtn.addEventListener('click', () => {
    goodsList.appendChild(buildGoodRow());
  });

  // ---------- 提交 ----------
  function collectGoods() {
    const items = [];
    goodsList.querySelectorAll('.good-row').forEach(row => {
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
    formData.append('table_items_json', JSON.stringify(collectGoods()));

    const params = new URLSearchParams();
    formData.forEach((value, key) => {
      params.append(key, value);
    });

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
      const filename = match ? decodeURIComponent(match[1]) : '代理出口结算协议.pdf';
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

  // 默认放一行空白
  goodsList.appendChild(buildGoodRow());
})();
