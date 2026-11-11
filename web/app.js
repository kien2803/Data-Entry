// DocLift — upload -> extract -> review -> download

const $ = (id) => document.getElementById(id);
const fmt = new Intl.NumberFormat("vi-VN");

const state = { file: null, invoice: null, warnings: [], mode: null, sample: false };

const VND_FIELDS = new Set(["tong_tien_hang", "tien_thue", "tong_thanh_toan",
  "don_gia", "thanh_tien"]);
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
const vnd = (n) => (n == null ? "—" : fmt.format(n));

// ---------- render ----------

function renderResult() {
  const { invoice: inv, warnings, mode } = state;
  const body = $("resultBody"), pills = $("statusPills");

  const nWarn = warnings.length;
  pills.innerHTML =
    `<span class="pill">${mode === "vision" ? "Đọc ảnh" : mode === "sample" ? "Dữ liệu mẫu" : "Đọc text"}</span>` +
    `<span class="pill">${inv.items.length} dòng hàng</span>` +
    (nWarn ? `<span class="pill warn">${nWarn} cảnh báo</span>`
           : `<span class="pill pos">Tổng khớp · MST hợp lệ</span>`);

  const warnHtml = nWarn
    ? `<div class="warn-strip">${warnings.map(w => `<div>– ${esc(w)}</div>`).join("")}</div>`
    : `<div class="ok-strip">Số liệu nhất quán: cộng hàng + thuế = tổng thanh toán.</div>`;

  const kv = KV_FIELDS.map(([label, key, cls]) =>
    `<div class="k">${label}</div><div class="v ${cls || ""}">${esc(fmtField(key, inv[key]))}</div>`
  ).join("");

  const rows = inv.items.map((it, i) => `
    <tr>
      <td class="c">${i + 1}</td>
      <td>${esc(it.ten_hang)}</td>
      <td>${esc(it.don_vi_tinh)}</td>
      <td class="r">${vnd(it.so_luong)}</td>
      <td class="r">${vnd(it.don_gia)}</td>
      <td class="r">${vnd(it.thanh_tien)}</td>
    </tr>`).join("");

  body.innerHTML = `
    ${warnHtml}
    <div class="kv">${kv}</div>
    <div class="table-wrap">
      <table class="items">
        <thead><tr>
          <th class="c" style="width:36px">STT</th><th>Tên hàng hóa, dịch vụ</th>
          <th style="width:70px">ĐVT</th><th class="r" style="width:70px">SL</th>
          <th class="r" style="width:110px">Đơn giá</th><th class="r" style="width:120px">Thành tiền</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    <div class="totals">
      <div class="trow"><span class="l">Cộng tiền hàng</span><span class="n">${vnd(inv.tong_tien_hang)}</span></div>
      <div class="trow"><span class="l">Thuế GTGT ${inv.thue_suat != null ? `(${inv.thue_suat}%)` : ""}</span><span class="n">${vnd(inv.tien_thue)}</span></div>
      <div class="trow grand"><span class="l">Tổng thanh toán</span><span class="n">${vnd(inv.tong_thanh_toan)} ₫</span></div>
    </div>
    ${inv.so_tien_bang_chu ? `<div class="inwords">${esc(inv.so_tien_bang_chu)}</div>` : ""}`;

  const issues = nWarn ? `${nWarn} cảnh báo cần kiểm tra` : "không có cảnh báo";
  $("finding").textContent =
    `Hóa đơn ${inv.so_hoa_don || "—"} · ${inv.items.length} dòng hàng · ${issues} · sẵn sàng tải Excel.`;
  $("downloadBtn").disabled = false;
}

function fmtField(key, v) {
  if (v == null || v === "") return "";
  if (VND_FIELDS.has(key)) return fmt.format(v) + " ₫";
  return String(v);
}

function esc(s) {
  return String(s ?? "").replace(/[&<>"]/g,
    c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}

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
  $("docview").innerHTML =
    `<img src="/sample/preview.png" alt="Hóa đơn mẫu">`;
}

function showExtractedPreview(b64) {
  if (b64) $("docview").innerHTML =
    `<img src="data:image/png;base64,${b64}" alt="Trang 1 chứng từ">`;
}

// ---------- flows ----------

async function handleFile(file) {
  state.file = file; state.sample = false;
  const url = URL.createObjectURL(file);
  showPreview(url, file.name, file.type);
  setLoading("Đang đọc chứng từ…");
  $("finding").textContent = `Đang trích xuất ${file.name}…`;

  const fd = new FormData();
  fd.append("file", file);
  try {
    const res = await fetch("/extract/json", { method: "POST", body: fd });
    if (!res.ok) {
      const j = await res.json().catch(() => ({}));
      throw new Error(j.detail || `Lỗi máy chủ (${res.status})`);
    }
    const j = await res.json();
    state.invoice = j.invoice; state.warnings = j.warnings; state.mode = j.mode;
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
    state.invoice = j.invoice; state.warnings = j.warnings; state.mode = "sample";
    renderResult();
  } catch (e) {
    setError("Không tải được mẫu", e.message);
  }
}

function download() {
  if (state.sample) {
    const a = document.createElement("a");
    a.href = "/sample/xlsx"; a.download = "mau.xlsx"; a.click();
    return;
  }
  if (!state.file) return;
  setLoading("Đang tạo file Excel…");
  const fd = new FormData();
  fd.append("file", state.file);
  fetch("/extract", { method: "POST", body: fd })
    .then(async res => {
      if (!res.ok) throw new Error((await res.json().catch(() => ({}))).detail || "Lỗi tạo Excel");
      return res.blob();
    })
    .then(blob => {
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "ket_qua.xlsx"; a.click();
      renderResult();
    })
    .catch(e => setError("Không tạo được Excel", e.message));
}

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

// First screen shows the product working, not an empty box.
loadSample();
