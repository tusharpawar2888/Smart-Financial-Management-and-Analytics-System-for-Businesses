const state = { modifiedAt: null, isLoading: false };

const $ = (selector) => document.querySelector(selector);
const moneyFormatter = new Intl.NumberFormat("en-IN", {
    style: "currency", currency: "INR", maximumFractionDigits: 0
});

function formatMoney(value) {
    return moneyFormatter.format(Number(value || 0));
}

function escapeHtml(value) {
    return String(value ?? "").replace(/[&<>'"]/g, (char) => ({
        "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;"
    }[char]));
}

function showMessage(message, type = "success") {
    const node = document.createElement("div");
    node.className = `message ${type === "error" ? "error" : ""}`;
    node.textContent = message;
    $("#message-area").appendChild(node);
    window.setTimeout(() => node.remove(), 4500);
}

function setConnection(text, status) {
    const node = $("#connection-status");
    node.className = `connection-status ${status}`;
    node.querySelector("span:last-child").textContent = text;
}

function parameters() {
    const params = new URLSearchParams();
    const values = {
        start: $("#start-date").value,
        end: $("#end-date").value,
        type: $("#type-filter").value,
        category: $("#category-filter").value
    };
    Object.entries(values).forEach(([key, value]) => { if (value) params.set(key, value); });
    return params.toString();
}

async function refreshDashboard(options = {}) {
    if (state.isLoading) return;
    state.isLoading = true;
    setConnection("Reading Excel workbook…", "");
    try {
        const response = await fetch(`/api/dashboard?${parameters()}`, { cache: "no-store" });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "Could not load the dashboard.");

        const changedExternally = state.modifiedAt && state.modifiedAt !== data.modified_at;
        state.modifiedAt = data.modified_at;
        updateCategories(data.categories);
        updateDashboard(data);
        setConnection(`Excel workbook connected · updated ${new Date(data.modified_at).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}`, "connected");
        if (changedExternally && options.background) showMessage("Workbook update detected. Dashboard refreshed.");
    } catch (error) {
        setConnection("Workbook needs attention", "error");
        if (!options.background) showMessage(error.message, "error");
    } finally {
        state.isLoading = false;
    }
}

function updateCategories(categories) {
    const categoryFilter = $("#category-filter");
    const currentCategory = categoryFilter.value;
    categoryFilter.innerHTML = `<option value="">All categories</option>${categories.map((category) => `<option value="${escapeHtml(category)}">${escapeHtml(category)}</option>`).join("")}`;
    if (categories.includes(currentCategory)) categoryFilter.value = currentCategory;
    $("#category-options").innerHTML = categories.map((category) => `<option value="${escapeHtml(category)}"></option>`).join("");
}

function updateDashboard(data) {
    const kpis = data.analysis.kpis;
    $("#kpi-income").textContent = formatMoney(kpis.income);
    $("#kpi-expenses").textContent = formatMoney(kpis.expenses);
    $("#kpi-profit").textContent = formatMoney(kpis.profit);
    $("#kpi-margin").textContent = `${Number(kpis.profit_margin).toFixed(1)}% profit margin`;
    $("#kpi-outstanding").textContent = formatMoney(kpis.outstanding);
    $("#kpi-pending-count").textContent = `${kpis.pending_payment_count} invoice${kpis.pending_payment_count === 1 ? "" : "s"} awaiting payment`;
    $("#kpi-due-soon").textContent = formatMoney(kpis.due_soon);
    $("#kpi-overdue").textContent = `${formatMoney(kpis.overdue)} overdue`;
    $("#kpi-company-payable").textContent = formatMoney(kpis.company_payable);
    $("#kpi-company-overdue").textContent = `${formatMoney(kpis.company_overdue)} overdue`;
    $("#kpi-forecast").textContent = formatMoney(kpis.forecast_30_days);
    $("#workbook-label").textContent = data.workbook;
    $("#transaction-count").textContent = `${kpis.transaction_count} record${kpis.transaction_count === 1 ? "" : "s"} shown`;
    renderMonthlyChart(data.analysis.monthly_trend);
    renderExpenseCategories(data.analysis.expense_categories);
    renderRevenue(data.analysis.top_revenue);
    renderInsights(data.analysis.insights);
    renderPaymentFollowups(data.analysis.payment_followups);
    renderCompanyPaymentFollowups(data.analysis.company_payment_followups);
    renderTransactions(data.transactions);
}

function renderMonthlyChart(rows) {
    const target = $("#monthly-chart");
    if (!rows.length) {
        target.innerHTML = '<div class="empty-chart">No transactions are available for this period.</div>';
        return;
    }
    const width = Math.max(rows.length * 86, 540);
    const height = 218;
    const left = 44, top = 14, bottom = 39, plotHeight = height - top - bottom;
    const maxValue = Math.max(...rows.flatMap((row) => [Number(row.income), Number(row.expense)]), 1);
    const roundedMax = Math.ceil(maxValue / 1000) * 1000 || 1;
    const slot = (width - left - 14) / rows.length;
    const barWidth = Math.min(19, slot * .25);
    let svg = `<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Monthly income and expenses chart">`;
    [0, .5, 1].forEach((step) => {
        const y = top + plotHeight - plotHeight * step;
        svg += `<line x1="${left}" y1="${y}" x2="${width - 9}" y2="${y}" stroke="#e5edeb" stroke-width="1"/>`;
        svg += `<text x="${left - 7}" y="${y + 4}" text-anchor="end" fill="#849390" font-size="10">${escapeHtml(compactNumber(roundedMax * step))}</text>`;
    });
    rows.forEach((row, index) => {
        const xCenter = left + slot * index + slot / 2;
        const incomeHeight = Number(row.income) / roundedMax * plotHeight;
        const expenseHeight = Number(row.expense) / roundedMax * plotHeight;
        svg += `<rect x="${xCenter - barWidth - 2}" y="${top + plotHeight - incomeHeight}" width="${barWidth}" height="${incomeHeight}" rx="3" fill="#0f766e"><title>${row.month} income: ${formatMoney(row.income)}</title></rect>`;
        svg += `<rect x="${xCenter + 2}" y="${top + plotHeight - expenseHeight}" width="${barWidth}" height="${expenseHeight}" rx="3" fill="#c65d45"><title>${row.month} expenses: ${formatMoney(row.expense)}</title></rect>`;
        svg += `<text x="${xCenter}" y="${height - 14}" text-anchor="middle" fill="#71817e" font-size="10">${escapeHtml(monthLabel(row.month))}</text>`;
    });
    svg += "</svg>";
    target.innerHTML = svg;
}

function compactNumber(value) {
    if (value >= 100000) return `₹${(value / 100000).toFixed(1)}L`;
    if (value >= 1000) return `₹${(value / 1000).toFixed(0)}k`;
    return `₹${Math.round(value)}`;
}

function monthLabel(value) {
    const [year, month] = value.split("-");
    return new Date(Number(year), Number(month) - 1).toLocaleString("en-IN", { month: "short", year: "2-digit" });
}

function renderExpenseCategories(rows) {
    const target = $("#expense-categories");
    if (!rows.length) { target.innerHTML = '<div class="empty-list">No expenses in this period.</div>'; return; }
    const topRows = rows.slice(0, 6);
    const largest = Math.max(...topRows.map((row) => row.amount), 1);
    target.innerHTML = topRows.map((row) => `<div class="bar-row">
        <div class="bar-meta"><span title="${escapeHtml(row.category)}">${escapeHtml(row.category)}</span><strong>${formatMoney(row.amount)}</strong></div>
        <div class="bar-track"><div class="bar-fill" style="width: ${Math.max(row.amount / largest * 100, 2)}%"></div></div>
    </div>`).join("");
}

function renderRevenue(rows) {
    const target = $("#revenue-list");
    if (!rows.length) { target.innerHTML = '<div class="empty-list">No income in this period.</div>'; return; }
    target.innerHTML = rows.map((row, index) => `<div class="rank-row">
        <span title="${escapeHtml(row.description)}"><i class="rank-number">${index + 1}</i>${escapeHtml(row.description)}</span><strong>${formatMoney(row.amount)}</strong>
    </div>`).join("");
}

function renderInsights(insights) {
    const icons = { info: "i", success: "✓", warning: "!" };
    $("#insights-list").innerHTML = insights.map((insight) => `<div class="insight ${escapeHtml(insight.level)}"><span class="insight-icon">${icons[insight.level] || "i"}</span><span>${escapeHtml(insight.text)}</span></div>`).join("");
}

function renderPaymentFollowups(payments) {
    const target = $("#payment-followups-body");
    if (!payments.length) {
        target.innerHTML = '<tr><td colspan="5" class="empty-row">No pending customer payments for the selected period.</td></tr>';
        return;
    }
    target.innerHTML = payments.map((payment) => {
        const statusClass = payment.status === "Overdue" ? "status-overdue" : "status-pending";
        return `<tr>
            <td>${escapeHtml(payment.party)}</td>
            <td>${escapeHtml(payment.description)}</td>
            <td>${escapeHtml(payment.due_date || "—")} <small class="due-note">${escapeHtml(payment.due_message)}</small></td>
            <td><span class="status-tag ${statusClass}">${escapeHtml(payment.status)}</span></td>
            <td class="amount">${formatMoney(payment.amount)}</td>
        </tr>`;
    }).join("");
}

function renderCompanyPaymentFollowups(payments) {
    const target = $("#company-payment-followups-body");
    if (!payments.length) {
        target.innerHTML = '<tr><td colspan="5" class="empty-row">No pending company payments for the selected period.</td></tr>';
        return;
    }
    target.innerHTML = payments.map((payment) => {
        const statusClass = payment.status === "Overdue" ? "status-overdue" : "status-pending";
        return `<tr>
            <td>${escapeHtml(payment.party)}</td>
            <td>${escapeHtml(payment.description)}</td>
            <td>${escapeHtml(payment.due_date || "—")} <small class="due-note">${escapeHtml(payment.due_message)}</small></td>
            <td><span class="status-tag ${statusClass}">${escapeHtml(payment.status)}</span></td>
            <td class="amount">${formatMoney(payment.amount)}</td>
        </tr>`;
    }).join("");
}

function renderTransactions(rows) {
    const target = $("#transactions-body");
    if (!rows.length) {
        target.innerHTML = '<tr><td colspan="8" class="empty-row">No transactions match the selected filters.</td></tr>';
        return;
    }
    target.innerHTML = rows.map((row) => {
        const typeClass = row["Transaction Type"] === "Income" ? "type-income" : "type-expense";
        const status = row["Payment Status"] || "Paid";
        const statusClass = `status-${status.toLowerCase()}`;
        return `<tr>
            <td>${escapeHtml(row.Date)}</td><td><span class="type-tag ${typeClass}">${escapeHtml(row["Transaction Type"])}</span></td>
            <td>${escapeHtml(row.Category)}</td><td>${escapeHtml(row.Description)}</td><td>${escapeHtml(row.Party || "—")}</td>
            <td>${escapeHtml(row["Due Date"] || "—")}</td><td><span class="status-tag ${statusClass}">${escapeHtml(status)}</span></td>
            <td class="amount">${formatMoney(row.Amount)}</td>
        </tr>`;
    }).join("");
}

async function submitTransaction(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const button = form.querySelector("button");
    const payload = Object.fromEntries(new FormData(form));
    button.disabled = true;
    try {
        const response = await fetch("/api/transactions", {
            method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload)
        });
        const result = await response.json();
        if (!response.ok) throw new Error(result.error || "The transaction could not be saved.");
        form.reset();
        $("#transaction-date").value = new Date().toISOString().slice(0, 10);
        showMessage(result.message);
        await refreshDashboard();
    } catch (error) { showMessage(error.message, "error"); }
    finally { button.disabled = false; }
}

async function submitUpload(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const button = form.querySelector("button");
    button.disabled = true;
    try {
        const response = await fetch("/api/upload", { method: "POST", body: new FormData(form) });
        const result = await response.json();
        if (!response.ok) throw new Error(result.error || "The Excel workbook could not be imported.");
        form.reset();
        showMessage(result.message);
        await refreshDashboard();
    } catch (error) { showMessage(error.message, "error"); }
    finally { button.disabled = false; }
}

function setToday() { $("#transaction-date").value = new Date().toISOString().slice(0, 10); }

document.addEventListener("DOMContentLoaded", () => {
    setToday();
    $("#transaction-form").addEventListener("submit", submitTransaction);
    $("#upload-form").addEventListener("submit", submitUpload);
    $("#refresh-button").addEventListener("click", () => refreshDashboard());
    ["#start-date", "#end-date", "#type-filter", "#category-filter"].forEach((selector) => {
        $(selector).addEventListener("change", () => refreshDashboard());
    });
    $("#clear-filters").addEventListener("click", () => {
        ["#start-date", "#end-date", "#type-filter", "#category-filter"].forEach((selector) => { $(selector).value = ""; });
        refreshDashboard();
    });
    refreshDashboard();
    window.setInterval(() => refreshDashboard({ background: true }), 10000);
});
