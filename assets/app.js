const DEFAULT_PUBHTML = window.__SHEET_PUBHTML__;

const els = {
  subtitle: document.getElementById("subtitle"),
  table: document.getElementById("dataTable"),
  search: document.getElementById("searchInput"),
  filterCol: document.getElementById("filterColumnSelect"),
  multiBtn: document.getElementById("multiBtn"),
  multiBtnText: document.getElementById("multiBtnText"),
  multiMenu: document.getElementById("multiMenu"),
  chips: document.getElementById("activeChips"),
  kpiTotal: document.getElementById("kpiTotalConvites"),
  kpiHint: document.getElementById("kpiConvitesHint"),
  kpiRows: document.getElementById("kpiRows"),
  btnClear: document.getElementById("btnClear"),
  btnExport: document.getElementById("btnExport"),
  btnRefresh: document.getElementById("btnRefresh"),
  btnTheme: document.getElementById("btnTheme"),
  pagination: document.getElementById("pagination"),
};

let raw = { headers: [], rows: [] };
let view = [];
let state = {
  query: "",
  filter: { col: null, values: new Set() },
  sort: { key: null, dir: "asc" },
  page: 1,
  pageSize: 25,
  conviteKey: null,
};

function inferCsvUrl(pubhtmlUrl) {
  const base = pubhtmlUrl.replace(/\/pubhtml.*$/i, "");
  return [base + "/pub?output=csv", base + "/gviz/tq?tqx=out:csv"];
}

async function fetchTextWithFallback(urls) {
  let lastErr = null;
  for (const u of urls) {
    try {
      const res = await fetch(u, { redirect: "follow" });
      if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
      const text = await res.text();
      if (text.trim().startsWith("<")) throw new Error("Retornou HTML, não CSV.");
      return text;
    } catch (e) {
      lastErr = e;
    }
  }
  throw lastErr ?? new Error("Falha ao baixar dados.");
}

function parseCsv(csvText) {
  const rows = [];
  let cur = [];
  let field = "";
  let inQuotes = false;

  for (let i = 0; i < csvText.length; i++) {
    const ch = csvText[i];
    const next = csvText[i + 1];

    if (inQuotes) {
      if (ch === '"' && next === '"') { field += '"'; i++; continue; }
      if (ch === '"') { inQuotes = false; continue; }
      field += ch;
      continue;
    }
    if (ch === '"') { inQuotes = true; continue; }
    if (ch === ",") { cur.push(field); field = ""; continue; }
    if (ch === "\n") { cur.push(field); rows.push(cur); cur = []; field = ""; continue; }
    if (ch === "\r") { continue; }
    field += ch;
  }
  if (field.length || cur.length) { cur.push(field); rows.push(cur); }

  const cleaned = rows.filter(r => r.some(v => String(v).trim() !== ""));
  if (!cleaned.length) return { headers: [], data: [] };

  const headers = cleaned[0].map(h => String(h).trim()).filter(Boolean);
  const data = cleaned.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, idx) => obj[h] = (r[idx] ?? "").trim());
    return obj;
  });
  return { headers, data };
}

function toNumberMaybe(v) {
  if (v == null) return null;
  const s = String(v).trim();
  if (!s) return null;
  const cleaned = s.replace(/R\$\s?/g, "").replace(/\./g, "").replace(/,/g, ".").replace(/\s/g, "");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function inferConviteColumn(headers, data) {
  const byName = headers.filter(h => /(convite|convites|qtd|quantidade|convidad)/i.test(h));
  const tryKeys = byName.length ? byName : headers;

  let best = null;
  for (const key of tryKeys) {
    let numericCount = 0, sum = 0;
    for (const row of data) {
      const n = toNumberMaybe(row[key]);
      if (n != null) { numericCount++; sum += n; }
    }
    if (numericCount >= Math.max(3, Math.floor(data.length * 0.3))) {
      const nameBonus = /(convite|convites)/i.test(key) ? 1000000 : 0;
      const score = nameBonus + sum;
      if (!best || score > best.score) best = { key, score };
    }
  }
  return best?.key ?? null;
}

function formatBR(n) {
  try { return new Intl.NumberFormat("pt-BR").format(n); } catch { return String(n); }
}

function buildFilters() {
  els.filterCol.innerHTML = "";
  raw.headers.forEach(h => {
    const opt = document.createElement("option");
    opt.value = h;
    opt.textContent = h;
    els.filterCol.appendChild(opt);
  });
  state.filter.col = raw.headers[0] ?? null;
  els.filterCol.value = state.filter.col ?? "";

  els.filterCol.addEventListener("change", () => {
    state.filter.col = els.filterCol.value;
    state.filter.values = new Set();
    rebuildMultiMenu();
    render();
  });

  rebuildMultiMenu();
}

function rebuildMultiMenu() {
  const col = state.filter.col;
  if (!col) return;

  const values = new Map();
  for (const r of raw.rows) {
    const v = (r[col] ?? "").trim();
    if (!v) continue;
    values.set(v, (values.get(v) ?? 0) + 1);
  }
  const items = Array.from(values.entries()).sort((a,b) => b[1]-a[1] || a[0].localeCompare(b[0]));

  els.multiMenu.innerHTML = "";
  const header = document.createElement("div");
  header.className = "px-2 py-2 text-xs text-slate-300/80 flex items-center justify-between gap-2";
  header.innerHTML = `<span>${items.length} valores</span><span class="opacity-70">multi-seleção</span>`;
  els.multiMenu.appendChild(header);

  const searchBox = document.createElement("input");
  searchBox.className = "input !rounded-xl !py-2 !text-sm";
  searchBox.placeholder = "Filtrar valores…";
  searchBox.addEventListener("input", () => {
    const q = searchBox.value.trim().toLowerCase();
    for (const node of Array.from(els.multiMenu.querySelectorAll("[data-val]"))) {
      const v = node.getAttribute("data-val") || "";
      node.classList.toggle("hidden", q && !v.toLowerCase().includes(q));
    }
  });
  const wrap = document.createElement("div");
  wrap.className = "px-2 pb-2";
  wrap.appendChild(searchBox);
  els.multiMenu.appendChild(wrap);

  for (const [val, count] of items) {
    const div = document.createElement("div");
    div.className = "menu-item";
    div.setAttribute("data-val", val);
    const checked = state.filter.values.has(val);
    div.innerHTML = `
      <input type="checkbox" ${checked ? "checked" : ""} />
      <div class="flex-1 min-w-0">
        <div class="text-sm truncate">${escapeHtml(val)}</div>
        <div class="text-[11px] text-slate-300/70">${count} ocorrência(s)</div>
      </div>
    `;
    div.addEventListener("click", (e) => {
      e.preventDefault();
      const isOn = state.filter.values.has(val);
      if (isOn) state.filter.values.delete(val);
      else state.filter.values.add(val);
      rebuildMultiMenu();
      updateMultiBtnText();
      render();
    });
    els.multiMenu.appendChild(div);
  }

  updateMultiBtnText();
}

function updateMultiBtnText() {
  const n = state.filter.values.size;
  els.multiBtnText.textContent = n ? `${n} selecionado(s)` : "Selecione valores…";
}

function applyFilters() {
  const q = state.query.trim().toLowerCase();
  const col = state.filter.col;
  const values = state.filter.values;

  const filtered = raw.rows.filter(r => {
    if (q) {
      const hit = raw.headers.some(h => String(r[h] ?? "").toLowerCase().includes(q));
      if (!hit) return false;
    }
    if (col && values.size) {
      const v = String(r[col] ?? "").trim();
      if (!values.has(v)) return false;
    }
    return true;
  });

  const { key, dir } = state.sort;
  if (key) {
    filtered.sort((a,b) => {
      const av = a[key] ?? "";
      const bv = b[key] ?? "";
      const an = toNumberMaybe(av);
      const bn = toNumberMaybe(bv);
      let cmp;
      if (an != null && bn != null) cmp = an - bn;
      else cmp = String(av).localeCompare(String(bv), "pt-BR", { numeric: true, sensitivity: "base" });
      return dir === "asc" ? cmp : -cmp;
    });
  }
  return filtered;
}

function computeKPIs(rows) {
  let totalConvites = null;
  if (state.conviteKey) {
    totalConvites = rows.reduce((acc, r) => acc + (toNumberMaybe(r[state.conviteKey]) ?? 0), 0);
  }
  const rowsCount = rows.length;

  if (totalConvites == null || totalConvites === 0) {
    els.kpiTotal.textContent = formatBR(rowsCount);
    els.kpiHint.textContent = state.conviteKey
      ? `Coluna '${state.conviteKey}' parece vazia; usando contagem`
      : "Nenhuma coluna numérica de convites detectada; usando contagem";
  } else {
    els.kpiTotal.textContent = formatBR(totalConvites);
    els.kpiHint.textContent = `Somando a coluna '${state.conviteKey}'`;
  }
  els.kpiRows.textContent = formatBR(rowsCount);
}

function renderChips() {
  els.chips.innerHTML = "";
  const chips = [];

  if (state.query.trim()) chips.push({ label: `Busca: ${state.query.trim()}`, onRemove: () => { state.query=""; els.search.value=""; render(); }});
  if (state.filter.col && state.filter.values.size) {
    chips.push({ label: `${state.filter.col}: ${state.filter.values.size} valor(es)`, onRemove: () => { state.filter.values=new Set(); rebuildMultiMenu(); render(); }});
  }

  for (const c of chips) {
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.innerHTML = `<span>${escapeHtml(c.label)}</span><button type="button" aria-label="Remover">✕</button>`;
    chip.querySelector("button").addEventListener("click", c.onRemove);
    els.chips.appendChild(chip);
  }
}

function renderTable(rows) {
  const headers = raw.headers;
  const pageCount = Math.max(1, Math.ceil(rows.length / state.pageSize));
  state.page = Math.min(state.page, pageCount);

  const start = (state.page - 1) * state.pageSize;
  const pageRows = rows.slice(start, start + state.pageSize);

  const thead = document.createElement("thead");
  const trh = document.createElement("tr");
  headers.forEach(h => {
    const th = document.createElement("th");
    const isSort = state.sort.key === h;
    const arrow = isSort ? (state.sort.dir === "asc" ? " ▲" : " ▼") : "";
    th.innerHTML = `<span>${escapeHtml(h)}${arrow}</span>`;
    th.addEventListener("click", () => {
      if (state.sort.key === h) state.sort.dir = state.sort.dir === "asc" ? "desc" : "asc";
      else { state.sort.key = h; state.sort.dir = "asc"; }
      render();
    });
    trh.appendChild(th);
  });
  thead.appendChild(trh);

  const tbody = document.createElement("tbody");
  for (const r of pageRows) {
    const tr = document.createElement("tr");
    headers.forEach(h => {
      const td = document.createElement("td");
      const v = r[h] ?? "";
      const num = toNumberMaybe(v);
      td.className = num != null ? "mono" : "";
      td.textContent = num != null ? formatBR(num) : String(v);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }

  els.table.innerHTML = "";
  els.table.appendChild(thead);
  els.table.appendChild(tbody);

  renderPagination(pageCount, rows.length);
}

function renderPagination(pageCount, totalRows) {
  els.pagination.innerHTML = "";
  const left = document.createElement("div");
  left.className = "text-xs text-slate-300/80";
  left.textContent = `Mostrando ${Math.min(totalRows, (state.page-1)*state.pageSize+1)}–${Math.min(totalRows, state.page*state.pageSize)} de ${totalRows}`;

  const right = document.createElement("div");
  right.className = "flex items-center gap-2";

  const mkBtn = (label, onClick, disabled=false) => {
    const b = document.createElement("button");
    b.className = "pager-btn";
    b.textContent = label;
    b.disabled = disabled;
    b.style.opacity = disabled ? ".5" : "1";
    b.addEventListener("click", onClick);
    return b;
  };

  right.appendChild(mkBtn("‹", () => { state.page = Math.max(1, state.page-1); render(); }, state.page===1));
  right.appendChild(mkBtn("›", () => { state.page = Math.min(pageCount, state.page+1); render(); }, state.page===pageCount));

  const sizeSel = document.createElement("select");
  sizeSel.className = "select !w-auto";
  [10,25,50,100].forEach(n => {
    const o=document.createElement("option");
    o.value=String(n); o.textContent=`${n}/página`;
    if (state.pageSize===n) o.selected=true;
    sizeSel.appendChild(o);
  });
  sizeSel.addEventListener("change", () => { state.pageSize = Number(sizeSel.value); state.page=1; render(); });
  right.appendChild(sizeSel);

  els.pagination.appendChild(left);
  els.pagination.appendChild(right);
}

function exportCsv(rows) {
  const headers = raw.headers;
  const lines = [headers.map(csvEscape).join(",")];
  for (const r of rows) lines.push(headers.map(h => csvEscape(r[h] ?? "")).join(","));
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "dados_filtrados.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function csvEscape(v) {
  const s = String(v ?? "");
  return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g,'""') + '"' : s;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({ "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;" }[c]));
}

function wireUI() {
  els.search.addEventListener("input", () => { state.query = els.search.value; state.page=1; render(); });

  els.btnClear.addEventListener("click", () => {
    state.query = "";
    state.filter.values = new Set();
    state.sort = { key: null, dir: "asc" };
    state.page = 1;
    els.search.value = "";
    rebuildMultiMenu();
    render();
  });

  els.btnExport.addEventListener("click", () => exportCsv(view));
  els.btnRefresh.addEventListener("click", () => init(true));

  els.multiBtn.addEventListener("click", () => els.multiMenu.classList.toggle("hidden"));
  document.addEventListener("click", (e) => {
    if (!els.multiMenu.contains(e.target) && !els.multiBtn.contains(e.target)) els.multiMenu.classList.add("hidden");
  });

  const key = "invites_theme";
  const saved = localStorage.getItem(key);
  if (saved === "light") document.documentElement.setAttribute("data-theme","light");
  els.btnTheme.addEventListener("click", () => {
    const cur = document.documentElement.getAttribute("data-theme");
    if (cur === "light") {
      document.documentElement.removeAttribute("data-theme");
      localStorage.setItem(key, "dark");
    } else {
      document.documentElement.setAttribute("data-theme","light");
      localStorage.setItem(key, "light");
    }
  });
}

function render() {
  view = applyFilters();
  computeKPIs(view);
  renderChips();
  renderTable(view);
  const key = state.conviteKey ? ` • coluna convites: ${state.conviteKey}` : "";
  els.subtitle.textContent = `${view.length} linha(s) • filtros aplicados${key}`;
}

async function init(force=false) {
  try {
    els.subtitle.textContent = "Carregando…";
    if (force) state.page = 1;

    const urls = inferCsvUrl(DEFAULT_PUBHTML);
    const csv = await fetchTextWithFallback(urls);
    const parsed = parseCsv(csv);

    raw.headers = parsed.headers;
    raw.rows = parsed.data;

    state.conviteKey = inferConviteColumn(raw.headers, raw.rows);

    buildFilters();
    render();
  } catch (e) {
    console.error(e);
    els.subtitle.textContent = "Erro ao carregar dados. Verifique se a planilha está publicada.";
    els.kpiTotal.textContent = "—";
    els.kpiRows.textContent = "—";
    els.kpiHint.textContent = String(e?.message || e);
  }
}

wireUI();
init();
