(() => {
  const $ = (s, el = document) => el.querySelector(s);
  const busy = $('#busy');
  let bulkResults = [];

  const esc = (v) => String(v ?? '').replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const badge = (s) => `<span class="badge b-${esc(s).replace(/ /g, '-')}">${esc(s)}</span>`;
  const fmtDate = (d) => {
    if (!d || !/^\d{4}-\d{2}-\d{2}/.test(d)) return d || '';
    const [y, m, day] = d.slice(0, 10).split('-');
    return `${day}/${m}/${y}`;
  };

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

  // Tabs
  document.querySelectorAll('.tab').forEach((btn) => btn.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach((b) => b.classList.toggle('active', b === btn));
    document.querySelectorAll('.panel').forEach((p) => { p.hidden = p.id !== `tab-${btn.dataset.tab}`; });
  }));

  // Single search
  function renderSingle(r) {
    const fields = [
      ['MSME Registration Status', badge(r.status)],
      ['Name of Enterprise', esc(r.enterprise_name) || '—'],
      ['Type of Enterprise', esc(r.enterprise_type) || '—'],
      ['Major Activity of Enterprise', esc(r.major_activity) || '—'],
      ['Type of Organisation', esc(r.organisation_type) || '—'],
      ['Date of Incorporation', esc(fmtDate(r.date_of_incorporation)) || '—'],
      ['Udyam Registration Number', esc(r.udyam_number) || '—'],
      ['PAN', esc(r.pan) || '—'],
      ['Date of Udyam Registration', esc(fmtDate(r.date_of_udyam_registration)) || '—'],
      ['State / District', esc([r.state, r.district].filter(Boolean).join(' / ')) || '—'],
      ['45-day rule / Sec 43B(h)', esc(r.msme_payment_rule) || '—'],
    ];
    return `<article class="result">
      <div class="result-head"><h3>${esc(r.query)}</h3>${badge(r.status)}</div>
      <dl class="details">${fields.map(([k, v]) => `<div><dt>${k}</dt><dd>${v}</dd></div>`).join('')}</dl>
      ${r.remarks ? `<p class="note">${esc(r.remarks)}</p>` : ''}
    </article>`;
  }

  document.querySelectorAll('form.search').forEach((form) => form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const out = $('#single-result');
    try {
      const r = await call('/api/verify', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: form.elements.id.value }),
      });
      out.innerHTML = renderSingle(r);
    } catch (err) {
      out.innerHTML = `<div class="error">${esc(err.message)}</div>`;
    }
  }));

  // Bulk
  function renderBulk(data) {
    bulkResults = data.results;
    const s = data.summary;
    const stats = [['total', 'Total'], ['Registered', 'Registered'], ['Not Registered', 'Not registered'],
      ['Invalid Format', 'Invalid'], ['Error', 'Errors'], ['Micro', 'Micro'], ['Small', 'Small'], ['Medium', 'Medium']];
    $('#summary').innerHTML = stats.filter(([k]) => s[k])
      .map(([k, label]) => `<span class="stat"><b>${s[k]}</b>${label}</span>`).join('');
    $('#filter').value = '';
    drawRows(bulkResults);
    $('#bulk-output').hidden = false;
  }

  function drawRows(rows) {
    $('#results tbody').innerHTML = rows.map((r) => `<tr>
      <td>${esc(r.query)}</td><td>${badge(r.status)}</td><td>${esc(r.udyam_number)}</td><td>${esc(r.pan)}</td>
      <td class="wrap-cell">${esc(r.enterprise_name)}</td><td>${esc(r.enterprise_type)}</td><td>${esc(r.major_activity)}</td>
      <td>${esc(r.organisation_type)}</td><td>${esc(fmtDate(r.date_of_incorporation))}</td>
      <td>${esc(r.msme_payment_rule)}</td><td class="wrap-cell">${esc(r.remarks)}</td></tr>`).join('');
  }

  $('#filter').addEventListener('input', (e) => {
    const q = e.target.value.trim().toLowerCase();
    drawRows(q ? bulkResults.filter((r) => Object.values(r).join(' ').toLowerCase().includes(q)) : bulkResults);
  });

  async function runBulk(opts) {
    try {
      renderBulk(await call('/api/bulk', opts));
    } catch (err) {
      $('#bulk-output').hidden = true;
      alert(err.message);
    }
  }

  $('#bulk-file').addEventListener('submit', (e) => {
    e.preventDefault();
    runBulk({ method: 'POST', body: new FormData(e.target) });
  });

  $('#bulk-paste').addEventListener('submit', (e) => {
    e.preventDefault();
    const ids = e.target.elements.ids.value.split(/[\s,;]+/).filter(Boolean);
    runBulk({ method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ids }) });
  });

  $('#export').addEventListener('click', async () => {
    busy.hidden = false;
    try {
      const res = await fetch('/api/export', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ results: bulkResults }),
      });
      if (!res.ok) throw new Error('Export failed');
      const blob = await res.blob();
      const name = (res.headers.get('Content-Disposition') || '').match(/filename="?([^";]+)/)?.[1] || 'MSME_Verification.xlsx';
      const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: name });
      a.click();
      URL.revokeObjectURL(a.href);
    } catch (err) {
      alert(err.message);
    } finally {
      busy.hidden = true;
    }
  });
})();
