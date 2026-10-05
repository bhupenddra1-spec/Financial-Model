(() => {
  const $ = (s, el = document) => el.querySelector(s);
  const busy = $('#busy');
  let results = [];

  const esc = (v) => String(v ?? '').replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const badge = (s) => s ? `<span class="badge b-${esc(s).replace(/[^A-Za-z]+/g, '-')}">${esc(s)}</span>` : '';

  async function call(url, opts) {
    busy.hidden = false;
    try {
      const res = await fetch(url, opts);
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data.error || `Request failed (${res.status})`);
      return data;
    } finally {
      busy.hidden = true;
    }
  }
  const postJSON = (url, body) => call(url, {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body),
  });

  document.querySelectorAll('.tab').forEach((btn) => btn.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach((b) => b.classList.toggle('active', b === btn));
    document.querySelectorAll('.panel').forEach((p) => { p.hidden = p.id !== `tab-${btn.dataset.tab}`; });
  }));

  // Bulk
  function render(data) {
    results = data.results;
    $('#so-summary').innerHTML = Object.entries(data.summary)
      .map(([k, v]) => `<span class="stat"><b>${v}</b>${esc(k)}</span>`).join('');
    $('#so-filter').value = '';
    draw(results);
    $('#so-output').hidden = false;
  }

  function draw(rows) {
    $('#so-results tbody').innerHTML = rows.map((r) => {
      const p = r.records[0] || {};
      const more = r.double_status ? `<br><small>+${r.records.length - 1} more CIN</small>` : '';
      return `<tr class="risk-${esc(r.risk)}">
        <td>${esc(r.sno)}</td><td class="wrap-cell">${esc(r.name)}</td><td>${esc(r.gst)}</td><td>${esc(r.pan)}</td>
        <td>${badge(r.outcome)}</td><td>${esc(p.cin)}${more}</td><td class="wrap-cell">${esc(p.company_name)}</td>
        <td>${esc(r.status)}</td><td>${esc(p.date_of_last_agm)}</td>
        <td class="wrap-cell">${esc(r.remarks.join('; '))}</td></tr>`;
    }).join('');
  }

  $('#so-filter').addEventListener('input', (e) => {
    const q = e.target.value.trim().toLowerCase();
    draw(q ? results.filter((r) => JSON.stringify(r).toLowerCase().includes(q)) : results);
  });

  async function run(promise) {
    try {
      render(await promise);
    } catch (err) {
      $('#so-output').hidden = true;
      alert(err.message);
    }
  }

  $('#so-file').addEventListener('submit', (e) => {
    e.preventDefault();
    run(call('/api/struck-off/bulk', { method: 'POST', body: new FormData(e.target) }));
  });

  $('#so-paste').addEventListener('submit', (e) => {
    e.preventDefault();
    const suppliers = e.target.elements.ids.value.split('\n').map((line) => line.trim()).filter(Boolean)
      .map((line, i) => {
        const [id, ...rest] = line.split(/\t|,/).map((s) => s.trim());
        const v = id.toUpperCase().replace(/\s+/g, '');
        const key = /^\d{2}[A-Z]{5}\d{4}[A-Z]/.test(v) ? 'gst' : /^[LU]\d{5}/.test(v) ? 'cin' : 'pan';
        return { sno: i + 1, name: rest.join(', '), [key]: v };
      });
    run(postJSON('/api/struck-off/bulk', { suppliers }));
  });

  $('#so-export').addEventListener('click', async () => {
    busy.hidden = false;
    try {
      const res = await fetch('/api/struck-off/report', {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ results }),
      });
      if (!res.ok) throw new Error('Report download failed');
      const blob = await res.blob();
      const name = (res.headers.get('Content-Disposition') || '').match(/filename="?([^";]+)/)?.[1] || 'Struck_Off_Report.xlsx';
      const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: name });
      a.click();
      URL.revokeObjectURL(a.href);
    } catch (err) {
      alert(err.message);
    } finally {
      busy.hidden = true;
    }
  });

  // Single
  const LABELS = [['cin', 'CIN'], ['company_name', 'Company Name As Per CIN'], ['status', 'Company Status (for efiling)'],
    ['roc_code', 'ROC Code'], ['registration_number', 'Registration Number'], ['category', 'Company Category'],
    ['subcategory', 'Company SubCategory'], ['class_of_company', 'Class of Company'], ['paid_up_capital', 'Paid up Capital (Rs)'],
    ['date_of_incorporation', 'Date of Incorporation'], ['registered_address', 'Registered Address'], ['email', 'Email Id'],
    ['listed', 'Whether Listed or not'], ['active_compliance', 'ACTIVE Compliance'],
    ['date_of_last_agm', 'Date of last AGM'], ['date_of_balance_sheet', 'Date of Balance Sheet']];

  $('#so-single').addEventListener('submit', async (e) => {
    e.preventDefault();
    const out = $('#so-single-result');
    try {
      const r = await postJSON('/api/struck-off/check', { id: e.target.elements.id.value });
      const recs = r.records.map((rec, i) => `
        ${r.records.length > 1 ? `<h4>MCA record ${i + 1} of ${r.records.length}</h4>` : ''}
        <dl class="details">${LABELS.map(([k, l]) => `<div><dt>${l}</dt><dd>${esc(rec[k]) || '—'}</dd></div>`).join('')}</dl>`).join('');
      out.innerHTML = `<article class="result">
        <div class="result-head"><h3>${esc(r.pan || r.gst || r.cin_input)}</h3>${badge(r.outcome)}</div>
        ${recs ? `<div class="records">${recs}</div>` : ''}
        ${r.status && !r.records.length ? `<p class="note">${esc(r.status)}</p>` : ''}
        ${r.remarks.length ? `<p class="note">${esc(r.remarks.join('; '))}</p>` : ''}
      </article>`;
    } catch (err) {
      out.innerHTML = `<div class="error">${esc(err.message)}</div>`;
    }
  });
})();
