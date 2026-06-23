(() => {
  const form = document.getElementById('guaranteeForm');
  const statusEl = document.getElementById('status');
  const generateBtn = document.getElementById('generate-btn');
  const resetBtn = document.getElementById('resetBtn');
  const previewList = document.getElementById('previewList');
  let defaults = {};

  const previewMap = {
    guarantee_no: '保函编号',
    issue_date: '出具日期',
    beneficiary_name: '受益人',
    applicant_name: '申请人',
    presentation_bank: '担保/提交银行',
    presentation_swift: 'SWIFT',
    finance_agreement_no: '协议编号',
    guarantee_amount: '担保金额',
    expiry_days: '到期天数',
    sign_bank_name: '签署银行',
  };

  function setStatus(text, kind) {
    statusEl.textContent = text;
    statusEl.className = kind || '';
  }

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

  fetch('/defaults')
    .then(r => r.json())
    .then(data => {
      defaults = data.defaults || {};
      Object.entries(defaults).forEach(([key, value]) => {
        const el = form.elements[key];
        if (el && !el.value) el.value = value;
      });
      updatePreview();
    })
    .catch(() => {});

  resetBtn.addEventListener('click', () => {
    form.reset();
    Object.entries(defaults).forEach(([key, value]) => {
      const el = form.elements[key];
      if (el) el.value = value;
    });
    updatePreview();
  });

  form.addEventListener('input', updatePreview);

  form.addEventListener('submit', async (event) => {
    event.preventDefault();
    setStatus('正在生成 PDF，请稍候…');
    generateBtn.disabled = true;

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
        throw new Error(await resp.text());
      }
      const blob = await resp.blob();
      const disposition = resp.headers.get('Content-Disposition') || '';
      const match = /filename\*=UTF-8''([^;]+)/i.exec(disposition);
      const filename = match ? decodeURIComponent(match[1]) : '付款保函.pdf';
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
