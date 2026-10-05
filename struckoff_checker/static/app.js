const $ = (s) => document.querySelector(s);
let current = [];

const esc = (v) => String(v ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const cls = (flag) => "f-" + flag.replace(/\s+/g, "-");

function show(msg, error) {
  const el = $("#msg");
  el.hidden = !msg;
  el.textContent = msg || "";
  el.className = "msg" + (error ? " error" : "");
}

async function post(url, options) {
  const resp = await fetch(url, options);
  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) throw new Error(data.error || `Request failed (${resp.status})`);
  return data;
}

function render() {
  const onlyFlagged = $("#only-flagged").checked;
  const flagged = ["Struck-Off", "Under Strike-Off Process", "Double Status"];
  const rows = [];
  current.forEach((r, i) => {
    if (onlyFlagged && !flagged.includes(r.flag)) return;
    (r.records.length ? r.records : [null]).forEach((rec) => {
      rows.push(`<tr>
        <td>${i + 1}</td><td>${esc(r.input.name)}</td>
        <td>${esc(r.input.pan || rec?.pan)}</td><td>${esc(rec?.cin || r.input.cin)}</td>
        <td>${esc(rec?.company_name)}</td><td>${esc(rec?.status)}</td>
        <td class="flag ${cls(r.flag)}">${esc(r.flag)}</td>
        <td>${esc(rec?.date_of_incorporation)}</td><td>${esc(rec?.roc_code)}</td><td>${esc(r.remarks)}</td></tr>`);
    });
  });
  $("#rows").innerHTML = rows.join("") || `<tr><td colspan="10">Nothing to show.</td></tr>`;
  $("#results").hidden = current.length === 0;
  const counts = {};
  current.forEach((r) => (counts[r.flag] = (counts[r.flag] || 0) + 1));
  $("#summary").innerHTML = `<span class="chip">Total ${current.length}</span>` +
    Object.entries(counts).map(([k, v]) => `<span class="chip ${cls(k)}">${esc(k)}: ${v}</span>`).join("");
}

async function run(button, fn) {
  button.disabled = true;
  show("Checking…");
  try {
    current = await fn();
    show("");
    render();
  } catch (e) {
    show(e.message, true);
  } finally {
    button.disabled = false;
  }
}

document.querySelectorAll(".tab").forEach((t) => t.addEventListener("click", () => {
  document.querySelectorAll(".tab").forEach((x) => x.classList.toggle("active", x === t));
  ["single", "bulk"].forEach((id) => ($("#" + id).hidden = id !== t.dataset.tab));
}));

$("#single-form").addEventListener("submit", (e) => {
  e.preventDefault();
  const body = Object.fromEntries(new FormData(e.target));
  run(e.target.querySelector("button"), async () =>
    [await post("/api/check", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) })]);
});

$("#bulk-file-btn").addEventListener("click", (e) => {
  const file = $("#bulk-file").files[0];
  if (!file) return show("Choose a file first", true);
  const fd = new FormData();
  fd.append("file", file);
  run(e.target, async () => (await post("/api/bulk", { method: "POST", body: fd })).results);
});

$("#bulk-text-btn").addEventListener("click", (e) => {
  const rows = $("#bulk-text").value.split("\n").map((l) => l.trim()).filter(Boolean).map((l) => {
    const [name = "", gstin = "", pan = "", cin = ""] = l.split(",").map((x) => x.trim());
    return { name, gstin, pan, cin };
  });
  if (!rows.length) return show("Paste at least one line", true);
  run(e.target, async () =>
    (await post("/api/bulk", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ rows }) })).results);
});

$("#only-flagged").addEventListener("change", render);

$("#download").addEventListener("click", async () => {
  try {
    const resp = await fetch("/api/export", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ results: current }) });
    if (!resp.ok) throw new Error("Export failed");
    const url = URL.createObjectURL(await resp.blob());
    const a = Object.assign(document.createElement("a"), { href: url, download: "Struck_Off_Companies.xlsx" });
    a.click();
    URL.revokeObjectURL(url);
  } catch (e) {
    show(e.message, true);
  }
});
