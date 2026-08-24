const form = document.querySelector("#uploadForm");
const runUpload = document.querySelector("#runUpload");
const targetDate = document.querySelector("#targetDate");
const statusList = document.querySelector("#statusList");
const previewBody = document.querySelector("#previewBody");
const downloadLink = document.querySelector("#downloadLink");
const healthBadge = document.querySelector("#healthBadge");
const tabs = document.querySelector("#tabs");

let activeKey = "";
let latestPreview = [];

const marketOrder = ["省内优先", "留存", "省内市场", "省间外送"];

function formatNumber(value, digits = 2) {
  if (value === null || value === undefined || Number.isNaN(Number(value))) {
    return "-";
  }
  return Number(value).toLocaleString("zh-CN", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}

function setBusy(isBusy) {
  runUpload.disabled = isBusy;
  runUpload.textContent = isBusy ? "计算中..." : "上传并计算";
}

function renderStatus(checks = []) {
  if (!checks.length) {
    statusList.innerHTML = '<div class="empty">暂无处理状态</div>';
    return;
  }
  statusList.innerHTML = checks
    .map((item) => {
      const text = item.status === "ok" ? "通过" : item.status === "warn" ? "提示" : "异常";
      const dates = item.datesSeen ? `<span>日期样例：${item.datesSeen.join("、")}</span>` : "";
      return `
        <div class="status-item">
          <div class="status-mark ${item.status}">${text}</div>
          <div>
            <b>${item.title}</b>
            <span>${item.detail}</span>
            ${dates}
          </div>
        </div>
      `;
    })
    .join("");
}

function renderPreview() {
  renderTabs();
  const rows = latestPreview.filter((row) => `${row.sheet}-${row.unit}` === activeKey);
  if (!rows.length) {
    previewBody.innerHTML = '<tr><td colspan="9" class="empty-cell">暂无预览数据</td></tr>';
    return;
  }

  previewBody.innerHTML = rows
    .map((row) => {
      const cells = marketOrder.flatMap((market) => {
        const item = row.markets[market] || {};
        return [
          `<td>${formatNumber(item.quarterQuantity, 2)}</td>`,
          `<td>${formatNumber(item.price, 2)}</td>`,
        ];
      });
      return `<tr><td>${row.slot}</td>${cells.join("")}</tr>`;
    })
    .join("");
}

function renderTabs() {
  const seen = new Map();
  latestPreview.forEach((row) => {
    const key = `${row.sheet}-${row.unit}`;
    if (!seen.has(key)) seen.set(key, `${row.sheet} · ${row.unit}`);
  });
  if (!activeKey || !seen.has(activeKey)) {
    activeKey = seen.keys().next().value || "";
  }
  tabs.innerHTML = Array.from(seen.entries())
    .map(([key, label]) => {
      const active = key === activeKey ? "active" : "";
      return `<button class="tab ${active}" data-key="${key}">${label}</button>`;
    })
    .join("");
}

function applyResult(result) {
  latestPreview = result.preview || [];
  renderStatus(result.checks || []);
  renderPreview();
  if (result.output?.downloadUrl) {
    downloadLink.href = result.output.downloadUrl;
    downloadLink.classList.remove("disabled");
  }
}

function showError(message) {
  renderStatus([{ status: "error", title: "处理失败", detail: message }]);
  latestPreview = [];
  renderPreview();
  downloadLink.classList.add("disabled");
}

function syncNoDataToggles() {
  document.querySelectorAll(".no-data-toggle input").forEach((checkbox) => {
    const row = checkbox.closest(".file-row");
    const toggle = checkbox.closest(".no-data-toggle");
    const fileInput = row?.querySelector('input[type="file"]');
    if (!row || !fileInput) return;
    fileInput.disabled = checkbox.checked;
    row.classList.toggle("no-data-selected", checkbox.checked);
    toggle?.classList.toggle("checked", checkbox.checked);
  });
}

function syncFileRows() {
  document.querySelectorAll(".file-row").forEach((row) => {
    const fileInput = row.querySelector('input[type="file"]');
    if (!fileInput) return;
    const hasFile = fileInput.files && fileInput.files.length > 0;
    row.classList.toggle("has-file", hasFile);
    row.dataset.fileName = hasFile ? fileInput.files[0].name : "";
  });
}

function setupDropUploads() {
  document.querySelectorAll(".file-row").forEach((row) => {
    const fileInput = row.querySelector('input[type="file"]');
    if (!fileInput) return;

    row.addEventListener("dragover", (event) => {
      event.preventDefault();
      if (!fileInput.disabled) row.classList.add("drag-over");
    });

    row.addEventListener("dragleave", () => {
      row.classList.remove("drag-over");
    });

    row.addEventListener("drop", (event) => {
      event.preventDefault();
      row.classList.remove("drag-over");
      if (fileInput.disabled) return;

      const file = event.dataTransfer.files?.[0];
      if (!file) return;
      if (!file.name.toLowerCase().endsWith(".xlsx")) {
        showError("请拖入 .xlsx 格式的 Excel 文件");
        return;
      }

      const transfer = new DataTransfer();
      transfer.items.add(file);
      fileInput.files = transfer.files;

      const checkbox = row.querySelector(".no-data-toggle input");
      if (checkbox) checkbox.checked = false;
      syncNoDataToggles();
      syncFileRows();
    });
  });
}

async function runWithUploads() {
  const data = new FormData(form);
  data.append("target_date", targetDate.value);
  setBusy(true);
  try {
    const response = await fetch("/api/process", { method: "POST", body: data });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "计算失败");
    applyResult(result);
  } catch (error) {
    showError(error.message);
  } finally {
    setBusy(false);
  }
}

async function checkHealth() {
  try {
    const response = await fetch("/api/health");
    if (!response.ok) throw new Error("服务未响应");
    healthBadge.textContent = "服务正常";
    healthBadge.classList.add("ok");
  } catch {
    healthBadge.textContent = "服务异常";
  }
}

tabs.addEventListener("click", (event) => {
  const tab = event.target.closest(".tab");
  if (tab) {
    activeKey = tab.dataset.key;
    renderPreview();
  }
});

runUpload.addEventListener("click", runWithUploads);
document.addEventListener("change", (event) => {
  if (event.target.matches(".no-data-toggle input")) {
    syncNoDataToggles();
  }
  if (event.target.matches('input[type="file"]')) {
    const row = event.target.closest(".file-row");
    const checkbox = row?.querySelector(".no-data-toggle input");
    if (checkbox && event.target.files.length) checkbox.checked = false;
    syncNoDataToggles();
    syncFileRows();
  }
});
document.addEventListener(
  "click",
  (event) => {
    const toggle = event.target.closest(".no-data-toggle");
    if (!toggle) return;
    event.preventDefault();
    event.stopPropagation();
    const checkbox = toggle.querySelector("input");
    checkbox.checked = !checkbox.checked;
    checkbox.dispatchEvent(new Event("change", { bubbles: true }));
  },
  true
);
checkHealth();
syncNoDataToggles();
syncFileRows();
setupDropUploads();
