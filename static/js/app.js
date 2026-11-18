function formatNumber(value) {
  if (value === null || value === undefined || value === "") return "-";
  const num = Number(value);
  if (Number.isNaN(num)) return value;
  return num.toLocaleString("fa-IR", { maximumFractionDigits: 3 });
}

function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function csrfToken() {
  const el = document.querySelector('meta[name="csrf-token"]');
  return el ? el.content : "";
}

function fillSelect(select, items) {
  const current = select.value;
  select.innerHTML = '<option value="">همه</option>' + (items || []).map(v => {
    const safe = escapeHtml(v);
    return `<option value="${safe}">${safe}</option>`;
  }).join("");
  select.value = current;
}

function formToPayload(form) {
  const data = new FormData(form);
  const payload = {};
  data.forEach((value, key) => {
    payload[key] = value;
  });
  if (form.allow_over_declaration) {
    payload.allow_over_declaration = form.allow_over_declaration.checked;
  }
  payload.csrf_token = csrfToken();
  return payload;
}

async function api(url, options = {}) {
  const headers = Object.assign({
    "Content-Type": "application/json",
    "X-CSRF-Token": csrfToken()
  }, options.headers || {});
  const opts = Object.assign({}, options, { headers: headers });
  const res = await fetch(url, opts);
  const data = await res.json().catch(() => ({ ok: false, error: "پاسخ نامعتبر از سرور" }));
  if (!res.ok || data.ok === false) {
    alert(data.error || "خطا در انجام عملیات");
  }
  return data;
}

document.addEventListener("DOMContentLoaded", function () {
  const toggle = document.getElementById("menuToggle");
  const sidebar = document.getElementById("sidebar");
  if (toggle && sidebar) {
    toggle.addEventListener("click", function () {
      sidebar.classList.toggle("open");
    });
  }
});
