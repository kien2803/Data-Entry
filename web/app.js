// DocLift — upload -> extract -> review/edit -> download

const $ = (id) => document.getElementById(id);
const fmt = new Intl.NumberFormat("vi-VN");

const state = {
  file: null, invoice: null, warnings: [],
  mode: "auto",        // auto | text | vision
  layout: "template",  // template | flat
  sample: false, effectiveMode: null, durationMs: null,
};

const KV_FIELDS = [
  ["Số hóa đơn", "so_hoa_don", "mono"],
  ["Ký hiệu", "ky_hieu", "mono"],
  ["Ngày lập", "ngay_lap", "mono"],
  ["HT thanh toán", "hinh_thuc_thanh_toan"],
  ["Đơn vị bán", "ten_don_vi_ban"],
  ["MST bên bán", "ma_so_thue_ban", "mono"],
  ["Đơn vị mua", "ten_don_vi_mua"],
  ["MST bên mua", "ma_so_thue_mua", "mono"],
];
const SEV_PILL = { critical: "neg", check: "warn", info: "" };
const MODE_LABEL = { text: "Đọc text", vision: "Đọc ảnh", sample: "Dữ liệu mẫu" };
const vnd = (n) => (n == null ? "—" : fmt.format(n));
const esc = (s) => String(s ?? "").replace(/[&<>"]/g,
  c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const num = (v) => v == null || v === "" ? null
  : parseFloat(String(v).replace(/[^\d.,-]/g, "").replace(/\./g, "").replace(",", "."));

// ---------- render ----------

function renderResult() {
  const { invoice: inv, warnings } = state;
  const body = $("resultBody"), pills = $("statusPills");

  const nCrit = warnings.filter(w => w.severity === "critical").length;
  const nCheck = warnings.filter(w => w.severity === "check").length;
  pills.innerHTML =
    `<span class="pill">${MODE_LABEL[state.effectiveMode] || "—"}</span>` +
    `<span class="pill">${inv.items.length} dòng hàng</span>` +
    (nCrit ? `<span class="pill neg">${nCrit} lỗi quan trọng</span>` : "") +
    (nCheck ? `<span class="pill warn">${nCheck} nên kiểm tra</span>` : "") +
    (!nCrit && !nCheck
      ? `<span class="pill pos">Tổng khớp · MST hợp lệ</span>` : "");

  const grouped = { critical: [], check: [], info: [] };
  warnings.forEach(w => (grouped[w.severity] || grouped.check).push(w));
  let stripHtml = "";
  if (grouped.critical.length)
    stripHtml += `<div class="warn-strip crit">${grouped.critical.map(
      w => `<div>– ${esc(w.message)}</div>`).join("")}</div>`;
  if (grouped.check.length)
    stripHtml += `<div class="warn-strip">${grouped.check.map(
      w => `<div>– ${esc(w.message)}</div>`).join("")}</div>`;
  if (grouped.info.length)
    stripHtml += `<div class="warn-strip info">${grouped.info.map(
      w => `<div>– ${esc(w.message)}</div>`).join("")}</div>`;
  if (!warnings.length)
    stripHtml = `<div class="ok-strip">Số liệu nhất quán: cộng hàng + thuế = tổng thanh toán.</div>`;

  const unsure = new Set(inv.truong_can_kiem_tra || []);
  const kv = KV_FIELDS.map(([label, key, cls]) => {
    const flag = unsure.has(key) ? " unsure" : "";
    const title = flag ? ` title="AI không chắc — hãy đối chiếu"` : "";
    return `<div class="k">${label}</div>` +
      `<div class="v ed ${cls || ""}${flag}" contenteditable${title}` +
      ` data-field="${key}">${esc(inv[key] ?? "")}</div>`;
  }).join("");

  const rows = inv.items.map((it, i) => `
    <tr>
      <td class="c">${i + 1}</td>
      <td class="ed" contenteditable data-idx="${i}" data-key="ten_hang">${esc(it.ten_hang)}</td>
      <td class="ed" contenteditable data-idx="${i}" data-key="don_vi_tinh">${esc(it.don_vi_tinh)}</td>
      <td class="r ed" contenteditable data-idx="${i}" data-key="so_luong">${vnd(it.so_luong)}</td>
      <td class="r ed" contenteditable data-idx="${i}" data-key="don_gia">${vnd(it.don_gia)}</td>
      <td class="r ed" contenteditable data-idx="${i}" data-key="thanh_tien">${vnd(it.thanh_tien)}</td>
      <td class="r ed" contenteditable data-idx="${i}" data-key="thue_suat">${it.thue_suat ?? ""}</td>
    </tr>`).join("");

  body.innerHTML = `
    ${stripHtml}
    <div class="kv">${kv}</div>
    <div class="table-wrap">
      <table class="items">
        <thead><tr>
          <th class="c" style="width:36px">STT</th><th>Tên hàng hóa, dịch vụ</th>
          <th style="width:70px">ĐVT</th><th class="r" style="width:60px">SL</th>
          <th class="r" style="width:110px">Đơn giá</th><th class="r" style="width:120px">Thành tiền</th>
          <th class="r" style="width:60px">Thuế%</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    <div class="totals">
      <div class="trow"><span class="l">Cộng tiền hàng</span><span class="n">${vnd(inv.tong_tien_hang)}</span></div>
      <div class="trow"><span class="l">Thuế GTGT ${inv.thue_suat != null ? `(${inv.thue_suat}%)` : ""}</span><span class="n">${vnd(inv.tien_thue)}</span></div>
      <div class="trow grand"><span class="l">Tổng thanh toán</span><span class="n">${vnd(inv.tong_thanh_toan)} ₫</span></div>
    </div>
    ${inv.so_tien_bang_chu ? `<div class="inwords">${esc(inv.so_tien_bang_chu)}</div>` : ""}
    <div class="edit-hint">Bấm vào ô bất kỳ để sửa trước khi tải Excel</div>`;

  const sev = nCrit ? `${nCrit} lỗi cần kiểm tra` :
    nCheck ? `${nCheck} điểm nên đối chiếu` : "không có cảnh báo";
  const dur = state.durationMs ? ` · ${(state.durationMs / 1000).toFixed(1)}s` : "";
  $("finding").textContent =
    `Hóa đơn ${inv.so_hoa_don || "—"} · ${inv.items.length} dòng hàng · ${sev}${dur}.`;
  $("downloadBtn").disabled = false;
}

// ---------- inline editing (delegated) ----------

const NUMERIC_ITEM_KEYS = new Set(["so_luong", "don_gia", "thanh_tien", "thue_suat"]);
$("resultBody").addEventListener("input", (e) => {
  const el = e.target;
  if (!state.invoice) return;
  const text = el.textContent.trim();
  if (el.dataset.field) {
    const key = el.dataset.field;
    state.invoice[key] = key.startsWith("ma_so_thue") || key.endsWith("_hoa_don")
      || key === "ky_hieu" || key === "ngay_lap" ? text : text;
  } else if (el.dataset.idx != null && el.dataset.key) {
    const it = state.invoice.items[+el.dataset.idx];
    if (!it) return;
    const k = el.dataset.key;
    it[k] = NUMERIC_ITEM_KEYS.has(k) ? num(text) : text;
  }
});

// ---------- document preview ----------

function showPreview(url, name, mime) {
  $("dropzone").hidden = true;
  $("docWrap").hidden = false;
  $("fileName").textContent = name;
  $("docMeta").textContent = "";
  $("docview").innerHTML = mime === "application/pdf"
    ? `<div class="empty" style="border:0">Đang tạo bản xem trước…</div>`
    : `<img src="${url}" alt="${esc(name)}">`;
}

function showSampleDoc() {
  $("dropzone").hidden = true;
  $("docWrap").hidden = false;
  $("fileName").textContent = "hoa_don_mau.pdf";
  $("docMeta").textContent = "File mẫu";
  $("docview").innerHTML = `<img src="/sample/preview.png" alt="Hóa đơn mẫu">`;
}

function showExtractedPreview(b64) {
  if (b64) $("docview").innerHTML =
    `<img src="data:image/png;base64,${b64}" alt="Trang 1 chứng từ">`;
}

// ---------- states ----------

function setLoading(msg) {
  $("resultBody").innerHTML =
    `<div class="loading"><div class="bar"></div><div>${msg}</div></div>`;
  $("statusPills").innerHTML = "";
  $("downloadBtn").disabled = true;
}

function setError(title, detail) {
  $("resultBody").innerHTML =
    `<div class="error-box"><b>${esc(title)}</b>${esc(detail)}</div>`;
  $("statusPills").innerHTML = `<span class="pill neg">Lỗi</span>`;
  $("finding").textContent = "Trích xuất chưa thành công. Xem chi tiết bên dưới.";
  $("downloadBtn").disabled = true;
}

// ---------- flows ----------

async function handleFile(file) {
  state.file = file; state.sample = false;
  const url = URL.createObjectURL(file);
  showPreview(url, file.name, file.type);
  const modeHint = state.mode === "vision" || file.type !== "application/pdf"
    ? "Đang đọc ảnh bằng AI — có thể mất 10–30 giây…" : "Đang đọc chứng từ…";
  setLoading(modeHint);
  $("finding").textContent = `Đang trích xuất ${file.name}…`;

  const fd = new FormData();
  fd.append("file", file);
  fd.append("mode", state.mode);
  try {
    const res = await fetch("/extract/json", { method: "POST", body: fd });
    if (!res.ok) {
      const j = await res.json().catch(() => ({}));
      throw new Error(j.detail || `Lỗi máy chủ (${res.status})`);
    }
    const j = await res.json();
    state.invoice = j.invoice; state.warnings = j.warnings;
    state.effectiveMode = j.mode; state.durationMs = j.duration_ms;
    showExtractedPreview(j.preview_b64);
    renderResult();
  } catch (e) {
    setError("Không trích xuất được", e.message);
  }
}

async function loadSample() {
  state.sample = true; state.file = null;
  showSampleDoc();
  setLoading("Đang tải hóa đơn mẫu…");
  try {
    const res = await fetch("/sample");
    const j = await res.json();
    state.invoice = j.invoice; state.warnings = j.warnings;
    state.effectiveMode = "sample"; state.durationMs = null;
    renderResult();
  } catch (e) {
    setError("Không tải được mẫu", e.message);
  }
}

async function download() {
  if (!state.invoice) return;
  setLoading("Đang tạo file Excel…");
  try {
    const res = await fetch("/export/xlsx", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ invoice: state.invoice, layout: state.layout }),
    });
    if (!res.ok) throw new Error(
      (await res.json().catch(() => ({}))).detail || "Lỗi tạo Excel");
    const blob = await res.blob();
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = state.layout === "flat" ? "du_lieu_phang.xlsx" : "ket_qua.xlsx";
    a.click();
    renderResult();
  } catch (e) {
    setError("Không tạo được Excel", e.message);
  }
}

// ---------- views ----------

const VIEWS = ["extract", "history", "templates", "settings"];
const loaded = { history: false, templates: false, settings: false };

function switchView(name) {
  VIEWS.forEach(v => { $("view-" + v).hidden = v !== name; });
  document.querySelectorAll(".nav-item[data-view]").forEach(a =>
    a.setAttribute("aria-current", a.dataset.view === name ? "page" : "false"));
  $("sidebar").classList.remove("open");
  if (!loaded[name]) {
    loaded[name] = true;
    if (name === "history") loadHistory();
    if (name === "templates") renderTemplates();
    if (name === "settings") loadSettings();
  } else if (name === "history") loadHistory();
}

// ---------- history ----------

async function loadHistory() {
  const body = $("historyBody");
  try {
    const res = await fetch("/history");
    const j = await res.json();
    if (!j.items.length) {
      body.innerHTML = `<div class="empty">Chưa có chứng từ nào được xử lý.</div>`;
      return;
    }
    const rows = j.items.map(r => `
      <tr>
        <td class="c">#${r.id}</td>
        <td class="truncate" title="${esc(r.file_name)}">${esc(r.file_name)}</td>
        <td>${r.mode || "—"}</td>
        <td>${r.ok ? '<span class="pill pos">OK</span>' : '<span class="pill neg">Lỗi</span>'}</td>
        <td class="r">${r.n_warnings || 0}</td>
        <td>${(r.created_at || "").replace("T", " ").slice(0, 16)}</td>
        <td class="r">
          <button class="btn sm" data-act="open" data-id="${r.id}">Mở</button>
          <button class="btn sm" data-act="xlsx" data-id="${r.id}">xlsx</button>
          <button class="btn sm" data-act="retry" data-id="${r.id}">Thử lại</button>
          <button class="btn sm danger" data-act="del" data-id="${r.id}">Xóa</button>
        </td>
      </tr>`).join("");
    body.innerHTML = `<div class="table-wrap"><table class="items">
      <thead><tr><th class="c">ID</th><th>File</th><th>Mode</th><th>Trạng thái</th>
      <th class="r">Cảnh báo</th><th>Thời gian</th><th class="r">Hành động</th></tr></thead>
      <tbody>${rows}</tbody></table></div>`;
  } catch (e) {
    body.innerHTML = `<div class="error-box"><b>Không tải được lịch sử</b>${esc(e.message)}</div>`;
  }
}

$("historyBody").addEventListener("click", async (e) => {
  const b = e.target.closest("button[data-act]");
  if (!b) return;
  const id = b.dataset.id, act = b.dataset.act;
  if (act === "xlsx") { location.href = `/history/${id}/xlsx`; return; }
  if (act === "del") {
    await fetch(`/history/${id}`, { method: "DELETE" });
    loadHistory(); return;
  }
  if (act === "retry") {
    b.disabled = true; b.textContent = "…";
    await fetch(`/history/${id}/retry`, { method: "POST" });
    loadHistory(); return;
  }
  if (act === "open") {
    const res = await fetch(`/history/${id}`);
    const d = await res.json();
    if (!d.invoice) { switchView("history"); return; }
    state.invoice = d.invoice; state.warnings = d.warnings;
    state.effectiveMode = d.mode; state.sample = true; state.file = null;
    $("dropzone").hidden = true; $("docWrap").hidden = false;
    $("fileName").textContent = d.file_name;
    $("docMeta").textContent = `#${d.id} · ${d.created_at.slice(0, 16)}`;
    $("docview").innerHTML = `<img src="/history/${d.id}/preview.png" alt="">`;
    renderResult();
    $("finding").textContent =
      `Đang xem lại extraction #${d.id} — dữ liệu đã lưu, không gọi AI.`;
    switchView("extract");
  }
});

// ---------- templates ----------

function renderTemplates() {
  const tpl = [
    { name: "Hóa đơn GTGT (VAT)", status: "active",
      desc: "Hóa đơn giá trị gia tăng điện tử — mẫu phổ biến nhất tại VN",
      fields: ["Số hóa đơn", "Ký hiệu", "Ngày lập", "MST hai bên",
               "Bảng hàng hóa", "Thuế GTGT", "Tổng thanh toán"] },
    { name: "Phiếu thu / chi", status: "soon",
      desc: "Chứng từ thu-chi nội bộ, quỹ tiền mặt", fields: [] },
    { name: "Sao kê ngân hàng", status: "soon",
      desc: "Bảng kê giao dịch tài khoản ngân hàng", fields: [] },
    { name: "Đơn hàng / báo giá", status: "soon",
      desc: "Đơn đặt hàng, báo giá, hợp đồng nguyên tắc", fields: [] },
  ];
  $("templatesBody").innerHTML = tpl.map(t => `
    <section class="panel span6 tpl">
      <header><h2>${esc(t.name)}</h2>
        ${t.status === "active" ? '<span class="pill pos">Đang hỗ trợ</span>'
                               : '<span class="pill">Sắp có</span>'}</header>
      <p class="tpl-desc">${esc(t.desc)}</p>
      ${t.fields.length ? `<div class="fields">${t.fields.map(
        f => `<span class="pill">${esc(f)}</span>`).join("")}</div>` : ""}
    </section>`).join("");
}

// ---------- settings ----------

async function loadSettings() {
  try {
    const res = await fetch("/stats");
    const s = await res.json();
    $("statsMeta").textContent = s.last ? `Lần cuối: ${s.last.ts}` : "";
    const modes = Object.entries(s.by_mode || {})
      .map(([k, v]) => `${k}: ${v}`).join(" · ") || "—";
    $("statsBody").innerHTML = `
      <div class="kv">
        <div class="k">Tổng lượt trích xuất</div><div class="v mono">${s.total}</div>
        <div class="k">Thành công</div><div class="v mono">${s.ok}</div>
        <div class="k">Thất bại</div><div class="v mono">${s.failed}</div>
        <div class="k">Tỷ lệ thành công</div><div class="v mono">${s.success_rate ?? "—"}%</div>
        <div class="k">Thời gian xử lý TB</div><div class="v mono">${s.avg_ms ? (s.avg_ms/1000).toFixed(1) + "s" : "—"}</div>
        <div class="k">Theo chế độ</div><div class="v mono">${modes}</div>
      </div>`;
  } catch (e) {
    $("statsBody").innerHTML =
      `<div class="error-box"><b>Không tải được thống kê</b>${esc(e.message)}</div>`;
  }
}

// defaults persisted locally — applied to next uploads
const saved = JSON.parse(localStorage.getItem("doclift-settings") || "{}");
if (saved.mode) { state.mode = saved.mode;
  document.querySelectorAll("[data-mode]").forEach(x =>
    x.setAttribute("aria-pressed", x.dataset.mode === saved.mode ? "true" : "false"));
}
if (saved.layout) { state.layout = saved.layout;
  document.querySelectorAll("[data-layout]").forEach(x =>
    x.setAttribute("aria-pressed", x.dataset.layout === saved.layout ? "true" : "false"));
}
["smode", "slayout"].forEach(attr =>
  document.querySelectorAll(`[data-${attr}]`).forEach(b =>
    b.addEventListener("click", () => {
      const group = attr === "smode" ? "setMode" : "setLayout";
      const val = b.dataset[attr];
      if (attr === "smode") {
        state.mode = val;
        document.querySelectorAll("[data-mode]").forEach(x =>
          x.setAttribute("aria-pressed", x.dataset.mode === val ? "true" : "false"));
      } else {
        state.layout = val;
        document.querySelectorAll("[data-layout]").forEach(x =>
          x.setAttribute("aria-pressed", x.dataset.layout === val ? "true" : "false"));
      }
      document.querySelectorAll(`#${group} button`).forEach(x =>
        x.setAttribute("aria-pressed", x === b ? "true" : "false"));
      localStorage.setItem("doclift-settings",
        JSON.stringify({ mode: state.mode, layout: state.layout }));
    })));

// nav + collapse
document.querySelectorAll(".nav-item[data-view]").forEach(a =>
  a.addEventListener("click", () => switchView(a.dataset.view)));
$("collapseBtn").addEventListener("click", () => {
  const s = $("sidebar");
  s.classList.toggle("collapsed");
  localStorage.setItem("doclift-sidebar",
    s.classList.contains("collapsed") ? "1" : "0");
});
if (localStorage.getItem("doclift-sidebar") === "1")
  $("sidebar").classList.add("collapsed");

// ---------- wiring ----------

const dz = $("dropzone"), fi = $("fileInput");
dz.addEventListener("click", () => fi.click());
dz.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") fi.click(); });
fi.addEventListener("change", () => fi.files[0] && handleFile(fi.files[0]));
dz.addEventListener("dragover", e => { e.preventDefault(); dz.classList.add("drag"); });
dz.addEventListener("dragleave", () => dz.classList.remove("drag"));
dz.addEventListener("drop", e => {
  e.preventDefault(); dz.classList.remove("drag");
  const f = e.dataTransfer.files[0];
  if (f) handleFile(f);
});
$("changeBtn").addEventListener("click", () => fi.click());
$("sampleBtn").addEventListener("click", loadSample);
$("downloadBtn").addEventListener("click", download);
$("menuToggle").addEventListener("click", () => $("sidebar").classList.toggle("open"));

document.querySelectorAll("[data-mode]").forEach(b =>
  b.addEventListener("click", () => {
    state.mode = b.dataset.mode;
    document.querySelectorAll("[data-mode]").forEach(x =>
      x.setAttribute("aria-pressed", x === b ? "true" : "false"));
  }));
document.querySelectorAll("[data-layout]").forEach(b =>
  b.addEventListener("click", () => {
    state.layout = b.dataset.layout;
    document.querySelectorAll("[data-layout]").forEach(x =>
      x.setAttribute("aria-pressed", x === b ? "true" : "false"));
  }));

// First screen shows the product working, not an empty box.
loadSample();
