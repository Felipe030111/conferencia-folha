const form = document.querySelector("#uploadForm");
const submitButton = form.querySelector('button[type="submit"]');
const statusValue = document.querySelector("#statusValue");
const statusLabel = document.querySelector("#statusLabel");
const periodInput = form.querySelector('input[name="period"]');
const periodKpi = document.querySelector("#periodKpi");
const totalKpi = document.querySelector("#totalKpi");
const okKpi = document.querySelector("#okKpi");
const warningKpi = document.querySelector("#warningKpi");
const errorKpi = document.querySelector("#errorKpi");
const statusList = document.querySelector("#statusList");
const launchRows = document.querySelector("#launchRows");
const exportBtn = document.querySelector("#exportBtn");
const filterButtons = Array.from(document.querySelectorAll(".filter"));

let currentRows = [];
let currentFilter = "all";

const statusText = {
  ok: "Conciliado",
  warning: "Alerta",
  error: "Divergencia",
};

periodInput.addEventListener("change", () => {
  periodKpi.textContent = formatPeriod(periodInput.value);
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  submitButton.disabled = true;
  statusValue.textContent = "Lendo";
  statusLabel.textContent = "Processando arquivos";

  try {
    const response = await fetch("/api/conferencia", {
      method: "POST",
      body: new FormData(form),
    });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "Nao foi possivel conferir os arquivos.");
    }
    renderReport(payload);
    statusValue.textContent = "Concluido";
    statusLabel.textContent = `${payload.company} - ${formatPeriod(payload.period)}`;
  } catch (error) {
    statusValue.textContent = "Erro";
    statusLabel.textContent = error.message;
  } finally {
    submitButton.disabled = false;
  }
});

filterButtons.forEach((button) => {
  button.addEventListener("click", () => {
    filterButtons.forEach((item) => item.classList.remove("active"));
    button.classList.add("active");
    currentFilter = button.dataset.filter;
    renderRows();
  });
});

exportBtn.addEventListener("click", () => {
  if (currentRows.length === 0) return;
  const csvRows = [
    ["codigo", "colaborador", "evento", "valor", "status", "observacao"],
    ...currentRows.map((row) => [row.code, row.name, row.label, row.amount, statusText[row.status], row.message]),
  ];
  const csv = csvRows.map((row) => row.map(csvCell).join(";")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "conferencia-lancamentos.csv";
  link.click();
  URL.revokeObjectURL(link.href);
});

function renderReport(report) {
  currentRows = report.launches || [];
  const totals = summarize(currentRows);
  periodKpi.textContent = formatPeriod(report.period);
  totalKpi.textContent = currentRows.length;
  okKpi.textContent = totals.ok;
  warningKpi.textContent = totals.warning;
  errorKpi.textContent = totals.error;
  statusList.innerHTML = `
    <div class="status-item">
      <div>
        <strong>Folha de Pagamento</strong>
        <span>${escapeHtml(report.file_company || report.company || "Arquivo carregado")}</span>
      </div>
      <div class="pill pill-ok">Referencia OK</div>
    </div>
    <div class="status-item">
      <div>
        <strong>Planilha de Lancamentos</strong>
        <span>${currentRows.length} eventos analisados</span>
      </div>
      <div class="pill ${totals.error ? "pill-erro" : totals.warning ? "pill-alerta" : "pill-ok"}">
        ${totals.error ? `${totals.error} divergencias` : totals.warning ? `${totals.warning} alertas` : "Conferida"}
      </div>
    </div>
  `;
  renderRows();
}

function renderRows() {
  const rows = currentFilter === "all" ? currentRows : currentRows.filter((row) => row.status === currentFilter);
  if (rows.length === 0) {
    launchRows.innerHTML = `<tr><td colspan="5" class="empty">Nenhum lancamento para este filtro.</td></tr>`;
    return;
  }
  launchRows.innerHTML = rows
    .map(
      (row) => `
        <tr>
          <td>
            <span class="employee">${escapeHtml(row.name)}</span>
            <span class="sub">Codigo ${escapeHtml(row.code)}</span>
          </td>
          <td>${escapeHtml(row.label)}</td>
          <td>${escapeHtml(row.amount)}</td>
          <td><span class="diff ${statusClass(row.status)}">${statusText[row.status]}</span></td>
          <td>${escapeHtml(row.message)}</td>
        </tr>
      `,
    )
    .join("");
}

function summarize(rows) {
  return rows.reduce(
    (acc, row) => {
      acc[row.status] = (acc[row.status] || 0) + 1;
      return acc;
    },
    { ok: 0, warning: 0, error: 0 },
  );
}

function statusClass(status) {
  if (status === "ok") return "green";
  if (status === "warning") return "orange";
  return "red";
}

function formatPeriod(value) {
  if (!value) return "-";
  const match = String(value).match(/^(\d{4})-(\d{2})$/);
  if (match) return `${match[2]}/${match[1]}`;
  return value;
}

function csvCell(value) {
  return `"${String(value ?? "").replaceAll('"', '""')}"`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}
