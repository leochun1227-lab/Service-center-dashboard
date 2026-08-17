const dashboardData = window.serviceCentreData;

const state = {
  page: "overview",
  month: dashboardData.meta.months[0],
  yard: dashboardData.meta.yards[0],
  invoice: dashboardData.meta.invoiceScopes[0],
  ticketType: dashboardData.meta.ticketTypes?.[0] || "All Ticket Types",
  workflowDealer: "Perth",
};

const pageConfig = {
  overview: { title: "Dealers Workflow", subtitle: "" },
  backlog: { title: "Open Backlog", subtitle: "Month-end open tickets, ageing and blocking reasons." },
  hours: { title: "Hours & Payroll", subtitle: "Technician list, payroll hours, assigned hours and invoiced hours." },
  adoption: { title: "Dealer Adoption", subtitle: "Placeholder view for phase 2 detail." },
  quality: { title: "Data Quality", subtitle: "Placeholder view for phase 2 detail." },
  reports: { title: "Reports", subtitle: "Placeholder view for exports and scheduled reporting." }
};

const WORKFLOW_DEALERS = ["Perth", "Traralgon", "Launceston", "Geelong", "Frankston"];

function initFilters() {
  renderSelect("monthFilter", dashboardData.meta.months, state.month);
  if (document.getElementById("yardFilter")) {
    renderSelect("yardFilter", dashboardData.meta.yards, state.yard);
  }
  if (document.getElementById("ticketTypeFilter")) {
    renderSelect("ticketTypeFilter", dashboardData.meta.ticketTypes || ["All Ticket Types"], state.ticketType);
  }
  if (document.getElementById("invoiceFilter")) {
    renderSelect("invoiceFilter", dashboardData.meta.invoiceScopes, state.invoice);
  }
  document.getElementById("lastUpdated").textContent = `Last updated: ${dashboardData.meta.lastUpdated}`;
  renderPeriodBanner();
}

function renderSelect(id, options, selected) {
  const select = document.getElementById(id);
  if (!select) return;
  select.innerHTML = options.map((option) => `<option ${option === selected ? "selected" : ""}>${option}</option>`).join("");
  select.onchange = () => {
    const key = id === "monthFilter" ? "month" : id === "yardFilter" ? "yard" : id === "ticketTypeFilter" ? "ticketType" : "invoice";
    state[key] = select.value;
    renderAll();
  };
}

function getSelectedDealerRows() {
  const rows = dashboardData.pages.overview.monthlyDealerActivityByType?.[state.month]?.[state.ticketType]
    || dashboardData.pages.overview.monthlyDealerActivity?.[state.month]
    || dashboardData.pages.overview.yardActivity
    || [];
  if (state.page === "overview") {
    return rows.filter((row) => row.yard === state.workflowDealer);
  }
  if (!state.yard || state.yard === "All Dealers") return rows;
  return rows.filter((row) => row.yard === state.yard);
}

function getSelectedPeriodRows() {
  return dashboardData.pages.overview.monthlyDealerActivityByType?.[state.month]?.[state.ticketType]
    || dashboardData.pages.overview.monthlyDealerActivity?.[state.month]
    || dashboardData.pages.overview.yardActivity
    || [];
}

function getPeriodRowsByType(type) {
  return dashboardData.pages.overview.monthlyDealerActivityByType?.[state.month]?.[type]
    || dashboardData.pages.overview.monthlyDealerActivity?.[state.month]
    || dashboardData.pages.overview.yardActivity
    || [];
}

function getSelectedPeriodRowsByType(type) {
  return getPeriodRowsByType(type).filter((row) => row.yard === state.workflowDealer);
}

function combineRowsByDealer(repairRows, pdiRows) {
  const pdiByDealer = new Map(pdiRows.map((row) => [row.yard, row]));
  const repairByDealer = new Map(repairRows.map((row) => [row.yard, row]));
  const dealers = [...new Set([...repairByDealer.keys(), ...pdiByDealer.keys()])];
  return dealers.map((yard) => ({
    yard,
    repair: repairByDealer.get(yard) || { yard },
    pdi: pdiByDealer.get(yard) || { yard },
  }));
}

function getWorkflowPeriodLabels() {
  return dashboardData.meta.months
    .filter((period) => !/^\d{4}$/.test(period))
    .slice()
    .reverse();
}

function isYearPeriod(period) {
  return /^\d{4}$/.test(period);
}

function getWorkflowPeriod() {
  return state.month;
}

function getWorkflowRows() {
  return getWorkflowPeriodLabels().map((period) => {
    const rows = dashboardData.pages.overview.monthlyDealerActivityByType?.[period]?.[state.ticketType]
      || dashboardData.pages.overview.monthlyDealerActivity?.[period]
      || [];
    const row = rows.find((item) => item.yard === state.workflowDealer) || {};
    return { period, ...row };
  });
}

function getWorkflowTicketType() {
  return state.ticketType === "PDI ticket" ? "PDI ticket" : "Repair ticket";
}

function getWorkflowCurrentRow(typeOverride) {
  const period = getWorkflowPeriod();
  const type = typeOverride || getWorkflowTicketType();
  const rows = dashboardData.pages.overview.monthlyDealerActivityByType?.[period]?.[type]
    || dashboardData.pages.overview.monthlyDealerActivity?.[period]
    || [];
  return rows.find((item) => item.yard === state.workflowDealer) || {};
}

function getWorkflowDailyData(typeOverride) {
  const type = typeOverride || getWorkflowTicketType();
  const period = getWorkflowPeriod();
  return dashboardData.pages.overview.workflowDaily?.[period]?.[type]?.[state.workflowDealer]
    || { daily: [], pipeline: [], totals: { created: 0, completed: 0, open: 0 } };
}

function monthSortValue(label) {
  const date = new Date(`1 ${label}`);
  return Number.isNaN(date.getTime()) ? 0 : date.getTime();
}

function getWorkflowMonthlyLabels(year) {
  return getWorkflowPeriodLabels()
    .filter((period) => period.endsWith(` ${year}`))
    .sort((a, b) => monthSortValue(a) - monthSortValue(b));
}

function getWorkflowSeriesData(typeOverride) {
  const type = typeOverride || getWorkflowTicketType();
  const period = getWorkflowPeriod();
  if (!isYearPeriod(period)) return getWorkflowDailyData(type);

  const months = getWorkflowMonthlyLabels(period);
  const daily = months.map((month) => {
    const monthData = dashboardData.pages.overview.workflowDaily?.[month]?.[type]?.[state.workflowDealer] || {};
    const monthRow = (dashboardData.pages.overview.monthlyDealerActivityByType?.[month]?.[type] || [])
      .find((item) => item.yard === state.workflowDealer) || {};
    const completedAmount = (monthData.daily || []).reduce((sum, row) => sum + (row.completedAmount || 0), 0);
    return {
      date: month,
      label: month.replace(` ${period}`, ""),
      created: monthRow.newTickets || monthData.totals?.created || 0,
      createdAmount: monthRow.newQuoteAmount || 0,
      createdAmountLabel: monthRow.newQuoteAmountLabel || "$0",
      completed: monthData.totals?.completed || 0,
      completedAmount,
      completedAmountLabel: formatMoney(completedAmount),
      openBalance: monthRow.openTickets || monthData.totals?.open || 0,
      openAmount: monthRow.openQuoteAmount || 0,
      openAmountLabel: monthRow.openQuoteAmountLabel || "$0",
    };
  });
  const yearRow = getWorkflowCurrentRow(type);
  const createdAmount = daily.reduce((sum, row) => sum + (row.createdAmount || 0), 0);
  const completedAmount = daily.reduce((sum, row) => sum + (row.completedAmount || 0), 0);
  return {
    dealer: state.workflowDealer,
    ticketType: type,
    period,
    daily,
    totals: {
      created: daily.reduce((sum, row) => sum + (row.created || 0), 0),
      completed: daily.reduce((sum, row) => sum + (row.completed || 0), 0),
      open: yearRow.openTickets || 0,
      createdAmountLabel: formatMoney(createdAmount),
      completedAmountLabel: formatMoney(completedAmount),
      openAmountLabel: yearRow.openQuoteAmountLabel || "$0",
    },
  };
}

function getWorkflowPipelineData(typeOverride) {
  const row = getWorkflowCurrentRow(typeOverride);
  return {
    totals: { open: row.openTickets || 0 },
    pipeline: (row.openStatusMix?.segments || []).map((segment) => ({
      status: segment.name,
      qty: segment.qty || 0,
      quoteAmount: segment.quoteAmount || 0,
      quoteAmountLabel: segment.quoteAmountLabel || formatMoney(segment.quoteAmount),
      rawStatuses: segment.rawStatuses || [],
      aging: segment.aging || [],
      color: segment.color,
    })),
  };
}

function renderWorkflowDealerTabs() {
  const target = document.getElementById("workflowDealerTabs");
  if (!target) return;
  target.innerHTML = WORKFLOW_DEALERS.map((dealer) => `
    <button class="dealer-tab ${dealer === state.workflowDealer ? "is-active" : ""}" data-workflow-dealer="${dealer}">${dealer}</button>
  `).join("");
  target.querySelectorAll("[data-workflow-dealer]").forEach((button) => {
    button.addEventListener("click", () => {
      state.workflowDealer = button.dataset.workflowDealer;
      renderAll();
    });
  });
}

function getInvoiceMetrics(row) {
  if (state.invoice === "Internal") {
    return {
      qty: row.internalInvoicedTickets || 0,
      amount: row.internalInvoicedAmount || 0,
      label: row.internalInvoicedAmountLabel || "$0",
    };
  }
  if (state.invoice === "External") {
    return {
      qty: row.externalInvoicedTickets || 0,
      amount: row.externalInvoicedAmount || 0,
      label: row.externalInvoicedAmountLabel || "$0",
    };
  }
  return {
    qty: row.invoicedTickets || 0,
    amount: row.invoicedAmount || 0,
    label: row.invoicedAmountLabel || "$0",
  };
}

function buildSelectedKpis() {
  const rows = getSelectedDealerRows();
  const periodEndLabel = /^\d{4}$/.test(state.month) ? "year-end" : "month-end";
  const totalQty = rows.reduce((sum, row) => sum + row.newTickets, 0);
  const totalAmount = rows.reduce((sum, row) => sum + row.newAmount, 0);
  const invoicedQty = rows.reduce((sum, row) => sum + getInvoiceMetrics(row).qty, 0);
  const invoicedAmount = rows.reduce((sum, row) => sum + getInvoiceMetrics(row).amount, 0);
  const openQty = rows.reduce((sum, row) => sum + (row.openTickets || 0), 0);
  const openQuoteAmount = rows.reduce((sum, row) => sum + (row.openQuoteAmount || 0), 0);
  return [
    {
      title: "New Tickets Created",
      value: totalQty.toLocaleString("en-US"),
      detail: `${state.month} total across five target dealers plus Other`,
      icon: "NT",
      tone: "blue",
      badge: state.month,
      badgeTone: "up",
    },
    {
      title: "New Ticket Quote Amount",
      value: formatMoney(totalAmount),
      detail: "AmountIncludingTax from exported C4C ticket result",
      icon: "AM",
      tone: "teal",
      badge: "CreatedOn",
      badgeTone: "up",
    },
    {
      title: "Tickets Invoiced",
      value: invoicedQty.toLocaleString("en-US"),
      detail: `${state.month} ${state.invoice}: Billing date plus Create invoice placeholders`,
      icon: "TI",
      tone: "green",
      badge: "Billing date",
      badgeTone: "up",
    },
    {
      title: "Invoiced Amount",
      value: formatMoney(invoicedAmount),
      detail: "ERPInvoiceNumberPrice; missing invoice placeholders are $0",
      icon: "IA",
      tone: "orange",
      badge: "SAP",
      badgeTone: "up",
    },
    {
      title: "Open Tickets",
      value: openQty.toLocaleString("en-US"),
      detail: `${state.month} ${periodEndLabel} open snapshot by CreatedOn`,
      icon: "OP",
      tone: "blue",
      badge: "Month end",
      badgeTone: "warn",
    },
    {
      title: "Open Tickets Quote Amount",
      value: formatMoney(openQuoteAmount),
      detail: "AmountIncludingTax for tickets not finished by status rule",
      icon: "OQ",
      tone: "teal",
      badge: "Quote",
      badgeTone: "warn",
    },
  ];
}

function buildWorkflowKpis() {
  const data = getWorkflowSeriesData();
  const totals = data.totals || {};
  const suffix = isYearPeriod(getWorkflowPeriod()) ? "This Year" : "This Month";
  return [
    { title: `Created ${suffix}`, value: (totals.created || 0).toLocaleString("en-US") },
    { title: "Created Quote Amount", value: totals.createdAmountLabel || "$0" },
    { title: `Completed ${suffix}`, value: (totals.completed || 0).toLocaleString("en-US") },
    { title: "Completed Quote Amount", value: totals.completedAmountLabel || "$0" },
    { title: "Still Open", value: (totals.open || 0).toLocaleString("en-US") },
    { title: "Open Quote Amount", value: totals.openAmountLabel || "$0" },
  ];
}

function formatMoney(value) {
  const amount = Number(value) || 0;
  if (Math.abs(amount) >= 1000000) return `$${(amount / 1000000).toFixed(2)}M`;
  if (Math.abs(amount) >= 1000) return `$${(amount / 1000).toFixed(1)}K`;
  return `$${Math.round(amount)}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function tooltipAttr(lines) {
  return `data-tooltip="${escapeHtml(lines.filter(Boolean).join("\n"))}"`;
}

function formatTooltipContent(value) {
  const lines = String(value || "").split("\n").filter(Boolean);
  if (!lines.length) return "";
  const [title, ...details] = lines;
  return `
    <strong>${escapeHtml(title)}</strong>
    ${details.map((line) => `<span>${escapeHtml(line)}</span>`).join("")}
  `;
}

function renderPageKpis(targetId, kpis) {
  document.getElementById(targetId).innerHTML = kpis.map((kpi) => `
    <article class="kpi-card">
      <div>
        <span class="kpi-title">${kpi.title}</span>
        <strong>${kpi.value}</strong>
      </div>
    </article>
  `).join("");
}

function renderMetricChart(targetId, config) {
  const rows = getSelectedDealerRows();
  const chartRows = rows.map((row) => config.mapRow(row));
  const totalQty = chartRows.reduce((sum, row) => sum + row.qty, 0);
  const totalAmount = chartRows.reduce((sum, row) => sum + row.amount, 0);
  const maxQtyShare = Math.max(0.01, ...chartRows.map((row) => share(row.qty, totalQty)));
  const maxAmountShare = Math.max(0.01, ...chartRows.map((row) => share(row.amount, totalAmount)));
  document.getElementById(targetId).innerHTML = chartRows.map((row) => `
    <div class="yard-group">
      <div class="bar-stack">
        ${renderBarPair(
          share(row.qty, totalQty),
          maxQtyShare,
          config.qtyClass,
          row.qty.toLocaleString("en-US"),
          config.qtyTooltip(row),
          share(row.amount, totalAmount),
          maxAmountShare,
          config.amountClass,
          row.amountLabel,
          config.amountTooltip(row),
        )}
      </div>
      <div class="yard-label">${row.yard}</div>
    </div>
  `).join("");
}

function renderTypeSplitMetricChart(targetId, config) {
  const repairRows = getSelectedPeriodRowsByType("Repair ticket");
  const pdiRows = getSelectedPeriodRowsByType("PDI ticket");
  const chartRows = combineRowsByDealer(repairRows, pdiRows).map(({ yard, repair, pdi }) => ({
    yard,
    repair: config.mapRow(repair || {}),
    pdi: config.mapRow(pdi || {}),
  }));
  const totalRepair = chartRows.reduce((sum, row) => sum + row.repair.qty, 0);
  const totalPdi = chartRows.reduce((sum, row) => sum + row.pdi.qty, 0);
  const maxRepairShare = Math.max(0.01, ...chartRows.map((row) => share(row.repair.qty, totalRepair)));
  const maxPdiShare = Math.max(0.01, ...chartRows.map((row) => share(row.pdi.qty, totalPdi)));
  document.getElementById(targetId).innerHTML = chartRows.map((row) => `
    <div class="yard-group">
      <div class="bar-stack">
        ${renderBarPair(
          share(row.repair.qty, totalRepair),
          maxRepairShare,
          "repair",
          row.repair.qty.toLocaleString("en-US"),
          [row.yard, state.month, "Repair ticket", `${config.qtyName}: ${row.repair.qty.toLocaleString("en-US")}`, row.repair.amountLabel ? `${config.amountName}: ${row.repair.amountLabel}` : ""],
          share(row.pdi.qty, totalPdi),
          maxPdiShare,
          "pdi",
          row.pdi.qty.toLocaleString("en-US"),
          [row.yard, state.month, "PDI ticket", `${config.qtyName}: ${row.pdi.qty.toLocaleString("en-US")}`, row.pdi.amountLabel ? `${config.amountName}: ${row.pdi.amountLabel}` : ""],
        )}
      </div>
      <div class="yard-label">${row.yard}</div>
    </div>
  `).join("");
}

function share(value, total) {
  return total > 0 ? value / total : 0;
}

function scaledShareHeight(valueShare, maxShare) {
  if (valueShare <= 0) return 18;
  const scaled = Math.log1p(valueShare * 100) / Math.log1p(maxShare * 100);
  return Math.round(18 + scaled * 132);
}

function renderBar(valueShare, maxShare, type) {
  const height = scaledShareHeight(valueShare, maxShare);
  return `<div class="bar ${type}" style="height:${height}px"></div>`;
}

function renderBarPair(leftShare, leftMax, leftClass, leftLabel, leftTooltip, rightShare, rightMax, rightClass, rightLabel, rightTooltip) {
  const leftHeight = scaledShareHeight(leftShare, leftMax);
  const rightHeight = scaledShareHeight(rightShare, rightMax);
  const close = Math.abs(leftHeight - rightHeight) < 18;
  const leftLift = close && leftHeight <= rightHeight ? " lift-label" : "";
  const rightLift = close && rightHeight < leftHeight ? " lift-label" : "";
  return `
    <div class="bar ${leftClass}" style="height:${leftHeight}px" ${tooltipAttr(leftTooltip || [leftLabel])}><span class="bar-value${leftLift}">${leftLabel}</span></div>
    <div class="bar ${rightClass}" style="height:${rightHeight}px" ${tooltipAttr(rightTooltip || [rightLabel])}><span class="bar-value${rightLift}">${rightLabel}</span></div>
  `;
}

function renderGroupedBars() {
  renderTypeSplitMetricChart("newTicketsChart", {
    qtyName: "New tickets",
    amountName: "New ticket quote amount",
    mapRow: (row) => ({
      qty: row.newTickets || 0,
      amountLabel: row.newQuoteAmountLabel || row.newAmountLabel || "$0",
    }),
  });

  renderTypeSplitMetricChart("invoicedTicketsChart", {
    qtyName: "Tickets invoiced",
    amountName: "Invoiced amount",
    mapRow: (row) => {
      const invoice = getInvoiceMetrics(row);
      return {
        qty: invoice.qty,
        amountLabel: invoice.label,
      };
    },
  });

  renderTypeSplitMetricChart("openTicketsChart", {
    qtyName: "Open tickets",
    amountName: "Open tickets quote amount",
    mapRow: (row) => ({
      qty: row.openTickets || 0,
      amountLabel: row.openQuoteAmountLabel || "$0",
    }),
  });

  renderInvoiceScopeChart();
}

function renderInvoiceScopeChart() {
  const rows = getSelectedDealerRows();
  const totalInternal = rows.reduce((sum, row) => sum + (row.internalInvoicedAmount || 0), 0);
  const totalExternal = rows.reduce((sum, row) => sum + (row.externalInvoicedAmount || 0), 0);
  const maxInternalShare = Math.max(0.01, ...rows.map((row) => share(row.internalInvoicedAmount || 0, totalInternal)));
  const maxExternalShare = Math.max(0.01, ...rows.map((row) => share(row.externalInvoicedAmount || 0, totalExternal)));
  document.getElementById("invoiceScopeChart").innerHTML = rows.map((row) => `
    <div class="yard-group">
      <div class="bar-stack">
        ${renderBarPair(
          share(row.internalInvoicedAmount || 0, totalInternal),
          maxInternalShare,
          "invoice-amount",
          row.internalInvoicedAmountLabel || "$0",
          [row.yard, state.month, `Internal invoice amount: ${row.internalInvoicedAmountLabel || "$0"}`, `Internal tickets: ${(row.internalInvoicedTickets || 0).toLocaleString("en-US")}`],
          share(row.externalInvoicedAmount || 0, totalExternal),
          maxExternalShare,
          "invoiced",
          row.externalInvoicedAmountLabel || "$0",
          [row.yard, state.month, `External invoice amount: ${row.externalInvoicedAmountLabel || "$0"}`, `External tickets: ${(row.externalInvoicedTickets || 0).toLocaleString("en-US")}`],
        )}
      </div>
      <div class="yard-label">${row.yard}</div>
    </div>
  `).join("");
}

function renderInvoiceMix() {
  const mix = buildInvoiceMixFromRows(getSelectedDealerRows());
  renderInvoiceMixSummary("invoiceMixSummary", mix);
}

function renderInvoiceMixSummary(targetId, mix) {
  const target = document.getElementById(targetId);
  if (!target) return;
  const segments = mix?.segments || [];
  const total = segments.reduce((sum, segment) => sum + (segment.qty || 0), 0);
  target.innerHTML = segments.map((segment) => {
    const percentLabel = segment.percentLabel || `${segment.percent || 0}%`;
    return `
      <div class="invoice-mix-pill" ${tooltipAttr([segment.name, state.month, `Amount: ${segment.amount}`, `Tickets: ${(segment.qty || 0).toLocaleString("en-US")}`, `Share: ${percentLabel}`])}>
        <span><i style="background:${segment.color}"></i>${segment.name}</span>
        <strong>${percentLabel}</strong>
        <em>${segment.amount}</em>
      </div>
    `;
  }).join("") + `
    <div class="invoice-mix-total" ${tooltipAttr(["Invoice Mix", state.month, `Total invoiced amount: ${mix.total}`, `Total tickets: ${total.toLocaleString("en-US")}`])}>
      <span>Total invoiced</span>
      <strong>${mix.total}</strong>
    </div>
  `;
}

function renderDonutMixTo(chartId, copyId, mix, title, centerLabel) {
  const chart = document.getElementById(chartId);
  const copy = document.getElementById(copyId);
  if (!chart || !copy) return;
  if (!mix || !Array.isArray(mix.segments) || mix.segments.length < 2) {
    chart.style.background = "conic-gradient(#d7e2ee 0 100%)";
    chart.innerHTML = `<div class="donut-center"><div><strong>${mix?.total || "0"}</strong><span>${centerLabel}</span></div></div>`;
    copy.innerHTML = `<div class="mix-row"><strong>No data</strong><span>0</span></div>`;
    return;
  }
  chart.style.background = "transparent";
  chart.setAttribute("data-tooltip", `${title}\n${state.month}\nTotal: ${mix.total}`);
  chart.innerHTML = `${renderDonutSvg(mix.segments)}<div class="donut-center"><div><strong>${mix.total}</strong><span>${centerLabel}</span></div></div>`;
  copy.innerHTML = mix.segments.map((segment) => `
    <div class="mix-row" ${tooltipAttr([segment.name, state.month, `Amount: ${segment.amount}`, `Tickets: ${(segment.qty || 0).toLocaleString("en-US")}`, `Share: ${segment.percentLabel || `${segment.percent}%`}`])}>
      <strong><i style="background:${segment.color}"></i>${segment.name} ${segment.percentLabel || `${segment.percent}%`}</strong>
      <span>${segment.amount}</span>
    </div>
  `).join("");
}

function buildInvoiceMixFromRows(rows) {
  const internalAmount = rows.reduce((sum, row) => sum + (row.internalInvoicedAmount || 0), 0);
  const externalAmount = rows.reduce((sum, row) => sum + (row.externalInvoicedAmount || 0), 0);
  const internalQty = rows.reduce((sum, row) => sum + (row.internalInvoicedTickets || 0), 0);
  const externalQty = rows.reduce((sum, row) => sum + (row.externalInvoicedTickets || 0), 0);
  const totalAmount = internalAmount + externalAmount;
  const externalPercent = totalAmount ? Math.round((externalAmount / totalAmount) * 100) : 0;
  const internalPercent = totalAmount ? 100 - externalPercent : 0;
  return {
    total: formatMoney(totalAmount),
    segments: [
      { name: "External", percent: externalPercent, amount: formatMoney(externalAmount), qty: externalQty, color: "#1f6feb" },
      { name: "Internal", percent: internalPercent, amount: formatMoney(internalAmount), qty: internalQty, color: "#f58b1f" },
    ],
  };
}

function renderOpenStatusMix() {
  const repairMix = buildOpenStatusMixFromRows(getSelectedPeriodRowsByType("Repair ticket"));
  const pdiMix = buildOpenStatusMixFromRows(getSelectedPeriodRowsByType("PDI ticket"));
  renderOpenStatusMixTo("repairOpenStatusMixChart", "repairOpenStatusMixCopy", repairMix);
  renderOpenStatusMixTo("pdiOpenStatusMixChart", "pdiOpenStatusMixCopy", pdiMix);
}

function renderOpenStatusMixTo(chartId, copyId, mix) {
  const chart = document.getElementById(chartId);
  const copy = document.getElementById(copyId);
  if (!chart || !copy) return;
  if (!mix || !Array.isArray(mix.segments) || mix.segments.length === 0) {
    chart.style.background = "conic-gradient(#d7e2ee 0 100%)";
    chart.innerHTML = `<div class="donut-center"><div><strong>0</strong><span>Open tickets</span></div></div>`;
    copy.innerHTML = `<div class="mix-row"><strong>No open tickets</strong><span>0 tickets</span></div>`;
    return;
  }

  chart.style.background = "transparent";
  chart.setAttribute("data-tooltip", `Open Ticket Status Mix\n${state.month}\nTotal open tickets: ${mix.total}`);
  chart.innerHTML = `${renderDonutSvg(mix.segments)}<div class="donut-center"><div><strong>${mix.total}</strong><span>Open tickets</span></div></div>`;
  copy.innerHTML = mix.segments.map((segment) => `
    <div class="mix-row" ${tooltipAttr([segment.name, state.month, `Qty: ${(segment.qty || 0).toLocaleString("en-US")} tickets`, `Share: ${segment.percentLabel || `${segment.percent}%`}`, segment.rawStatuses?.length ? `Raw C4C status: ${segment.rawStatuses.join(", ")}` : ""])}>
      <strong><i style="background:${segment.color}"></i>${segment.name} ${segment.percentLabel || `${segment.percent}%`}</strong>
      <span>${(segment.qty || 0).toLocaleString("en-US")} tickets</span>
    </div>
  `).join("");
}

function renderDonutSvg(segments) {
  const visible = (segments || []).filter((segment) => (segment.share ?? segment.percent) > 0);
  if (!visible.length) return "";
  let cursor = 0;
  const paths = visible.map((segment, index) => {
    const size = index === visible.length - 1 ? 100 - cursor : (segment.share ?? segment.percent);
    const start = cursor;
    const end = cursor + size;
    cursor = end;
    const overlap = 0.08;
    return donutArcPath(
      index === 0 ? start : Math.max(0, start - overlap),
      index === visible.length - 1 ? 100 : Math.min(100, end + overlap),
      segment.color,
    );
  }).join("");
  return `<svg class="donut-svg" viewBox="0 0 200 200" aria-hidden="true">${paths}</svg>`;
}

function renderWorkflowLineChart(targetId, config) {
  const rows = getWorkflowRows();
  const width = 640;
  const height = 260;
  const pad = { top: 28, right: 64, bottom: 42, left: 58 };
  const points = rows.map((row, index) => {
    const x = rows.length <= 1 ? width / 2 : pad.left + index * ((width - pad.left - pad.right) / (rows.length - 1));
    return {
      label: row.period.replace(" 20", " '"),
      x,
      qty: config.qty(row),
      amount: config.amount(row),
      amountLabel: config.amountLabel(row),
    };
  });
  const maxQty = Math.max(1, ...points.map((point) => point.qty));
  const maxAmount = Math.max(1, ...points.map((point) => point.amount));
  const niceQtyMax = niceMax(maxQty);
  const niceAmountMax = niceMax(maxAmount);
  const yFor = (value, max) => {
    const ratio = max ? value / max : 0;
    return pad.top + (1 - ratio) * (height - pad.top - pad.bottom);
  };
  const qtyPoints = points.map((point) => ({ ...point, y: yFor(point.qty, niceQtyMax) }));
  const amountPoints = points.map((point) => ({ ...point, y: yFor(point.amount, niceAmountMax) }));
  const path = (series) => smoothPath(series);
  const gridTicks = [0, 0.25, 0.5, 0.75, 1];
  const grid = gridTicks.map((tick) => {
    const y = pad.top + (1 - tick) * (height - pad.top - pad.bottom);
    const qtyLabel = Math.round(niceQtyMax * tick).toLocaleString("en-US");
    const amountLabel = formatMoney(niceAmountMax * tick);
    return `
      <line class="line-grid" x1="${pad.left}" y1="${y}" x2="${width - pad.right}" y2="${y}"></line>
      <text class="line-y-label qty-axis-label" x="${pad.left - 10}" y="${y + 4}" text-anchor="end">${qtyLabel}</text>
      <text class="line-y-label amount-axis-label" x="${width - pad.right + 10}" y="${y + 4}" text-anchor="start">${amountLabel}</text>
    `;
  }).join("");
  const axisLabels = points.map((point, index) => {
    if (points.length > 8 && index % 2 === 1) return "";
    return `<text x="${point.x}" y="${height - 12}" text-anchor="middle">${point.label}</text>`;
  }).join("");
  const circles = (series, type, labelFor) => series.map((point) => `
    <circle class="line-dot ${type}" cx="${point.x}" cy="${point.y}" r="5" ${tooltipAttr([
      state.workflowDealer,
      point.period,
      labelFor(point),
    ])}></circle>
  `).join("");
  const hoverBands = points.map((point) => `
    <rect class="line-hover-band" x="${point.x - 14}" y="${pad.top}" width="28" height="${height - pad.top - pad.bottom}" ${tooltipAttr([
      state.workflowDealer,
      point.period,
      `Qty: ${point.qty.toLocaleString("en-US")}`,
      `${config.amountName}: ${point.amountLabel}`,
    ])}></rect>
  `).join("");

  document.getElementById(targetId).innerHTML = `
    <svg class="line-svg" viewBox="0 0 ${width} ${height}" preserveAspectRatio="xMidYMid meet">
      <rect class="line-panel" x="${pad.left}" y="${pad.top}" width="${width - pad.left - pad.right}" height="${height - pad.top - pad.bottom}"></rect>
      <rect class="line-axis-bg left-axis-bg" x="0" y="${pad.top - 10}" width="${pad.left - 5}" height="${height - pad.top - pad.bottom + 20}"></rect>
      <rect class="line-axis-bg right-axis-bg" x="${width - pad.right + 5}" y="${pad.top - 10}" width="${pad.right - 5}" height="${height - pad.top - pad.bottom + 20}"></rect>
      ${grid}
      <line class="line-axis" x1="${pad.left}" y1="${height - pad.bottom}" x2="${width - pad.right}" y2="${height - pad.bottom}"></line>
      <path class="line-path ${config.qtyLineClass || "qty-line"}" d="${path(qtyPoints)}"></path>
      <path class="line-path ${config.amountLineClass || "amount-line"}" d="${path(amountPoints)}"></path>
      ${circles(qtyPoints, config.qtyDotClass || "qty-dot", (point) => `Qty: ${point.qty.toLocaleString("en-US")}`)}
      ${circles(amountPoints, config.amountDotClass || "amount-dot", (point) => `${config.amountName}: ${point.amountLabel}`)}
      ${hoverBands}
      <g class="line-labels">${axisLabels}</g>
    </svg>
  `;
}

function renderDailyWorkflowChart() {
  const data = getWorkflowSeriesData();
  const rows = data.daily || [];
  const card = document.getElementById("workflowDailyLine")?.closest(".card");
  if (card) {
    const title = card.querySelector("h2");
    const note = card.querySelector(".card-note");
    if (title) title.textContent = isYearPeriod(getWorkflowPeriod()) ? "Monthly Workflow" : "Daily Workflow";
    if (note) note.textContent = isYearPeriod(getWorkflowPeriod()) ? "Monthly workflow trend" : "Created, completed and open balance";
  }
  const width = 760;
  const height = 300;
  const pad = { top: 28, right: 62, bottom: 44, left: 54 };
  const plotBottom = height - pad.bottom;
  const plotWidth = width - pad.left - pad.right;
  const bandWidth = rows.length ? plotWidth / rows.length : plotWidth;
  const barWidth = Math.max(5, Math.min(12, bandWidth * 0.22));
  const points = rows.map((row, index) => {
    const x = rows.length <= 1 ? width / 2 : pad.left + index * ((width - pad.left - pad.right) / (rows.length - 1));
    return { ...row, x };
  });
  const maxDaily = niceMax(Math.max(1, ...points.map((point) => Math.max(point.created || 0, point.completed || 0))));
  const maxOpen = niceMax(Math.max(1, ...points.map((point) => point.openBalance || 0)));
  const yFor = (value, max) => pad.top + (1 - (max ? value / max : 0)) * (height - pad.top - pad.bottom);
  const created = points.map((point) => ({ ...point, y: yFor(point.created || 0, maxDaily) }));
  const completed = points.map((point) => ({ ...point, y: yFor(point.completed || 0, maxDaily) }));
  const open = points.map((point) => ({ ...point, y: yFor(point.openBalance || 0, maxOpen) }));
  const gridTicks = [0, 0.5, 1];
  const grid = gridTicks.map((tick) => {
    const y = pad.top + (1 - tick) * (height - pad.top - pad.bottom);
    return `
      <line class="line-grid workflow-grid" x1="${pad.left}" y1="${y}" x2="${width - pad.right}" y2="${y}"></line>
      <text class="line-y-label qty-axis-label" x="${pad.left - 10}" y="${y + 4}" text-anchor="end">${Math.round(maxDaily * tick).toLocaleString("en-US")}</text>
      <text class="line-y-label amount-axis-label" x="${width - pad.right + 10}" y="${y + 4}" text-anchor="start">${Math.round(maxOpen * tick).toLocaleString("en-US")}</text>
    `;
  }).join("");
  const axisLabels = points.map((point, index) => {
    if (points.length > 16 && index % 2 === 1) return "";
    return `<text x="${point.x}" y="${height - 12}" text-anchor="middle">${isYearPeriod(getWorkflowPeriod()) ? point.label : point.label.replace(" Aug", "")}</text>`;
  }).join("");
  const bars = points.map((point) => {
    const createdHeight = Math.max(0, plotBottom - yFor(point.created || 0, maxDaily));
    const completedHeight = Math.max(0, plotBottom - yFor(point.completed || 0, maxDaily));
    return `
      <rect class="workflow-bar created-bar" x="${(point.x - barWidth - 2).toFixed(1)}" y="${(plotBottom - createdHeight).toFixed(1)}" width="${barWidth}" height="${createdHeight.toFixed(1)}" rx="4" ${tooltipAttr([
        state.workflowDealer,
        getWorkflowTicketType(),
        getWorkflowPeriod(),
        point.label,
        `Created: ${(point.created || 0).toLocaleString("en-US")}`,
      ])}></rect>
      <rect class="workflow-bar completed-bar" x="${(point.x + 2).toFixed(1)}" y="${(plotBottom - completedHeight).toFixed(1)}" width="${barWidth}" height="${completedHeight.toFixed(1)}" rx="4" ${tooltipAttr([
        state.workflowDealer,
        getWorkflowTicketType(),
        getWorkflowPeriod(),
        point.label,
        `Completed: ${(point.completed || 0).toLocaleString("en-US")}`,
      ])}></rect>
    `;
  }).join("");
  const openDots = open.map((point, index) => {
    const isEndpoint = index === 0 || index === open.length - 1;
    return `
      <circle class="open-balance-dot ${isEndpoint ? "is-endpoint" : ""}" cx="${point.x}" cy="${point.y}" r="${isEndpoint ? 3.8 : 2.4}" ${tooltipAttr([
      state.workflowDealer,
      getWorkflowTicketType(),
      getWorkflowPeriod(),
      point.label,
      `Open balance: ${(point.openBalance || 0).toLocaleString("en-US")}`,
    ])}></circle>
    `;
  }).join("");
  const hoverBands = points.map((point) => `
    <rect class="line-hover-band" x="${point.x - Math.max(10, bandWidth / 2)}" y="${pad.top}" width="${Math.max(20, bandWidth)}" height="${height - pad.top - pad.bottom}" ${tooltipAttr([
      state.workflowDealer,
      getWorkflowTicketType(),
      getWorkflowPeriod(),
      point.label,
      `Created: ${(point.created || 0).toLocaleString("en-US")}`,
      `Completed: ${(point.completed || 0).toLocaleString("en-US")}`,
      `Open balance: ${(point.openBalance || 0).toLocaleString("en-US")}`,
    ])}></rect>
  `).join("");
  document.getElementById("workflowDailyLine").innerHTML = `
    <svg class="line-svg daily-line-svg" viewBox="0 0 ${width} ${height}" preserveAspectRatio="xMidYMid meet">
      <rect class="line-panel" x="${pad.left}" y="${pad.top}" width="${width - pad.left - pad.right}" height="${height - pad.top - pad.bottom}"></rect>
      ${grid}
      <line class="line-axis workflow-axis" x1="${pad.left}" y1="${plotBottom}" x2="${width - pad.right}" y2="${plotBottom}"></line>
      ${bars}
      <path class="open-balance-path" d="${smoothPath(open)}"></path>
      ${openDots}
      ${hoverBands}
      <g class="line-labels">${axisLabels}</g>
    </svg>
  `;
}

function renderWorkflowPipeline() {
  const types = ["Repair ticket", "PDI ticket"];
  document.getElementById("workflowPipeline").innerHTML = types.map((type) => {
    const data = getWorkflowPipelineData(type);
    const rows = data.pipeline || [];
    const maxAgeQty = Math.max(1, ...rows.flatMap((row) => (row.aging || []).map((bucket) => bucket.qty || 0)));
    const body = rows.length ? rows.map((row) => `
      <div class="pipeline-row stage-card ${row.qty ? "" : "is-empty"}" ${tooltipAttr([row.status, type, state.workflowDealer, `Qty: ${(row.qty || 0).toLocaleString("en-US")}`, row.rawStatuses?.length ? `Raw C4C status: ${row.rawStatuses.join(", ")}` : ""])}>
        <div class="stage-card-head">
          <span style="--stage-color:${row.color || "#64748b"}">${row.status}</span>
          <strong>${(row.qty || 0).toLocaleString("en-US")}</strong>
        </div>
        <div class="stage-aging">
          ${(row.aging || []).map((bucket) => `
            <div class="stage-aging-row">
              <span>${bucket.label}</span>
              <div class="stage-aging-track"><i style="width:${Math.round(((bucket.qty || 0) / maxAgeQty) * 100)}%; background:${row.color || "#64748b"}"></i></div>
              <strong>${(bucket.qty || 0).toLocaleString("en-US")}</strong>
            </div>
          `).join("")}
        </div>
      </div>
    `).join("") : `<div class="pipeline-empty">No open ${type.toLowerCase()} workload.</div>`;
    return `
      <div class="pipeline-section">
        <div class="pipeline-section-head">
          <strong>${type === "Repair ticket" ? "Repair" : "PDI"}</strong>
          <span>${(data.totals?.open || 0).toLocaleString("en-US")} open</span>
        </div>
        <div class="pipeline-stage-grid">${body}</div>
      </div>
    `;
  }).join("");
}

function niceMax(value) {
  if (value <= 0) return 1;
  const exponent = Math.floor(Math.log10(value));
  const base = 10 ** exponent;
  const fraction = value / base;
  const niceFraction = fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 5 ? 5 : 10;
  return niceFraction * base;
}

function smoothPath(points) {
  if (points.length === 0) return "";
  if (points.length === 1) return `M ${points[0].x} ${points[0].y}`;
  let d = `M ${points[0].x.toFixed(1)} ${points[0].y.toFixed(1)}`;
  for (let i = 1; i < points.length; i += 1) {
    const prev = points[i - 1];
    const current = points[i];
    const midX = (prev.x + current.x) / 2;
    d += ` C ${midX.toFixed(1)} ${prev.y.toFixed(1)}, ${midX.toFixed(1)} ${current.y.toFixed(1)}, ${current.x.toFixed(1)} ${current.y.toFixed(1)}`;
  }
  return d;
}

function donutArcPath(startPercent, endPercent, color) {
  const startAngle = (startPercent / 100) * 360 - 90;
  const endAngle = (endPercent / 100) * 360 - 90;
  const largeArc = endPercent - startPercent > 50 ? 1 : 0;
  const start = polarPoint(100, 100, 72, startAngle);
  const end = polarPoint(100, 100, 72, endAngle);
  if (endPercent - startPercent >= 99.999) {
    return `<circle cx="100" cy="100" r="72" fill="none" stroke="${color}" stroke-width="56"></circle>`;
  }
  return `<path d="M ${start.x} ${start.y} A 72 72 0 ${largeArc} 1 ${end.x} ${end.y}" fill="none" stroke="${color}" stroke-width="56" stroke-linecap="butt"></path>`;
}

function polarPoint(cx, cy, radius, angleDegrees) {
  const radians = (angleDegrees * Math.PI) / 180;
  return {
    x: (cx + radius * Math.cos(radians)).toFixed(3),
    y: (cy + radius * Math.sin(radians)).toFixed(3),
  };
}

function buildOpenStatusMixFromRows(rows) {
  const totals = new Map();
  rows.forEach((row) => {
    (row.openStatusMix?.segments || []).forEach((segment) => {
      const current = totals.get(segment.name) || { name: segment.name, qty: 0, quoteAmount: 0, color: segment.color };
      current.qty += segment.qty || 0;
      current.quoteAmount += segment.quoteAmount || 0;
      totals.set(segment.name, current);
    });
  });
  const total = [...totals.values()].reduce((sum, segment) => sum + segment.qty, 0);
  const segments = [...totals.values()]
    .sort((a, b) => b.qty - a.qty || a.name.localeCompare(b.name))
    .map((segment) => ({
      ...segment,
      percent: total ? Math.round((segment.qty / total) * 100) : 0,
      percentLabel: total && segment.qty > 0 && (segment.qty / total) * 100 < 1 ? "<1%" : `${total ? Math.round((segment.qty / total) * 100) : 0}%`,
      amount: `${segment.qty.toLocaleString("en-US")} tickets`,
    }));
  const allocated = segments.reduce((sum, segment) => sum + segment.percent, 0);
  if (segments.length) segments[segments.length - 1].percent += 100 - allocated;
  return { total: total.toLocaleString("en-US"), segments };
}

function renderYardSummary() {
  const rows = getSelectedDealerRows().map((row) => ({
    ...row,
    invoiceMetrics: getInvoiceMetrics(row),
    status: row.yard === "Other" ? "Other" : "Target",
    statusTone: row.yard === "Other" ? "watch" : "good",
  }));
  document.getElementById("yardSummaryRows").innerHTML = rows.map((row) => `
    <tr>
      <td><span class="yard-name"><i style="background:${row.color}"></i>${row.yard}</span></td>
      <td class="num">${row.newTickets}</td>
      <td class="num">${row.newAmountLabel}</td>
      <td class="num">${row.invoiceMetrics.qty.toLocaleString("en-US")}</td>
      <td class="num">${row.invoiceMetrics.label}</td>
      <td class="num">${(row.openTickets || 0).toLocaleString("en-US")}</td>
      <td class="num">${row.openQuoteAmountLabel || "$0"}</td>
      <td><span class="status-pill ${row.statusTone}">${row.status}</span></td>
    </tr>
  `).join("");
}

function renderUtilisation(targetId, rows) {
  if (!Array.isArray(rows) || rows.length === 0) {
    document.getElementById(targetId).innerHTML = `<div class="util-row"><div class="util-head"><strong>Pending data source</strong><span>TBC</span></div><div class="track"><div class="fill watch" style="width:0%"></div></div><div class="util-head"><span>This KPI will be connected later.</span></div></div>`;
    return;
  }
  document.getElementById(targetId).innerHTML = rows.map((row) => `
    <div class="util-row">
      <div class="util-head">
        <strong>${row.yard}</strong>
        <span>${row.ratio}%</span>
      </div>
      <div class="track"><div class="fill ${row.tone}" style="width:${row.ratio}%"></div></div>
      <div class="util-head"><span>${row.note}</span></div>
    </div>
  `).join("");
}

function renderAgeing() {
  const rows = dashboardData.pages.overview.ageing;
  if (!Array.isArray(rows) || rows.length === 0) {
    document.getElementById("ageingChart").innerHTML = `<div class="stack-group"><div class="yard-label">Pending<br><small>Open ticket KPI next</small></div></div>`;
    return;
  }
  const max = Math.max(...rows.map((row) => row.fresh + row.warm + row.stale));
  document.getElementById("ageingChart").innerHTML = rows.map((row) => {
    const total = row.fresh + row.warm + row.stale;
    const fresh = Math.round((row.fresh / max) * 150);
    const warm = Math.round((row.warm / max) * 150);
    const stale = Math.round((row.stale / max) * 150);
    return `
      <div class="stack-group">
        <div class="stack-columns">
          <div class="stack-column">
            <div class="bar stale" style="height:${Math.max(12, stale)}px"></div>
            <div class="bar warm" style="height:${Math.max(12, warm)}px"></div>
            <div class="bar fresh" style="height:${Math.max(12, fresh)}px"></div>
          </div>
        </div>
        <div class="yard-label">${row.yard}<br><small>${total} open</small></div>
      </div>
    `;
  }).join("");
}

function renderLifecycle() {
  const rows = dashboardData.pages.overview.lifecycle;
  if (!Array.isArray(rows) || rows.length === 0) {
    document.getElementById("lifecycleGrid").innerHTML = `
      <article class="life-card">
        <span>Next KPI</span>
        <strong>TBC</strong>
        <small>Invoiced/open ticket logic will be added after new tickets.</small>
      </article>
    `;
    return;
  }
  document.getElementById("lifecycleGrid").innerHTML = rows.map((row) => `
    <article class="life-card">
      <span>${row.label}</span>
      <strong>${row.value}</strong>
      <small>${row.detail}</small>
    </article>
  `).join("");
}

function renderTrend() {
  if (!document.getElementById("monthlyTrend")) return;
  const rows = dashboardData.pages.activity.trend;
  const max = Math.max(1, ...rows.map((row) => row.newTickets));
  document.getElementById("monthlyTrend").innerHTML = rows.map((row) => `
    <div class="trend-row">
      <strong>${row.monthFull || row.month}</strong>
      <div class="trend-bars">
        <div class="track"><div class="fill good" style="width:${Math.round((row.newTickets / max) * 100)}%"></div></div>
      </div>
      <span>${row.newTickets}</span>
    </div>
  `).join("");
}

function renderBreakdown() {
  const rows = getSelectedDealerRows().map((row) => ({
    title: row.yard,
    amount: getInvoiceMetrics(row).label,
    count: `${getInvoiceMetrics(row).qty.toLocaleString("en-US")} invoiced tickets`,
    note: `${state.month} ${state.invoice}`,
  }));
  document.getElementById("invoiceBreakdown").innerHTML = rows.map((item) => `
    <div class="breakdown-item">
      <div class="breakdown-head">
        <strong>${item.title}</strong>
        <strong>${item.amount}</strong>
      </div>
      <div class="metric-pair"><span>${item.count}</span><span>${item.note}</span></div>
    </div>
  `).join("");
}

function renderWorkflow() {
  renderWorkflowDealerTabs();
  renderWorkflowPipeline();
}

function renderWorkflowInvoiceScope(row) {
  const rows = [row || {}];
  const totalInternal = rows.reduce((sum, item) => sum + (item.internalInvoicedAmount || 0), 0);
  const totalExternal = rows.reduce((sum, item) => sum + (item.externalInvoicedAmount || 0), 0);
  const maxInternalShare = Math.max(0.01, ...rows.map((item) => share(item.internalInvoicedAmount || 0, totalInternal)));
  const maxExternalShare = Math.max(0.01, ...rows.map((item) => share(item.externalInvoicedAmount || 0, totalExternal)));
  document.getElementById("workflowInvoiceScopeChart").innerHTML = rows.map((item) => `
    <div class="yard-group">
      <div class="bar-stack">
        ${renderBarPair(
          share(item.internalInvoicedAmount || 0, totalInternal),
          maxInternalShare,
          "invoice-amount",
          item.internalInvoicedAmountLabel || "$0",
          [state.workflowDealer, state.month, `Internal invoice amount: ${item.internalInvoicedAmountLabel || "$0"}`, `Internal tickets: ${(item.internalInvoicedTickets || 0).toLocaleString("en-US")}`],
          share(item.externalInvoicedAmount || 0, totalExternal),
          maxExternalShare,
          "invoiced",
          item.externalInvoicedAmountLabel || "$0",
          [state.workflowDealer, state.month, `External invoice amount: ${item.externalInvoicedAmountLabel || "$0"}`, `External tickets: ${(item.externalInvoicedTickets || 0).toLocaleString("en-US")}`],
        )}
      </div>
      <div class="yard-label">${state.workflowDealer}</div>
    </div>
  `).join("");
}

function renderLegacyActivity() {
  if (!document.getElementById("invoiceBreakdown")) return;
  renderBreakdown();
  renderInsights();
}

function renderInsights() {
  document.getElementById("activityInsights").innerHTML = dashboardData.pages.activity.insights.map((item) => `
    <div class="insight-item">
      <strong>${item.title}</strong>
      <p>${item.text}</p>
    </div>
  `).join("");
}

function renderBacklogRows() {
  document.getElementById("backlogRows").innerHTML = dashboardData.pages.backlog.rows.map((row) => `
    <tr>
      <td>${row.yard}</td>
      <td class="num">${row.open}</td>
      <td class="num">${row.stale}</td>
      <td class="num">${row.value}</td>
      <td><span class="status-pill ${row.tone}">${row.risk}</span></td>
    </tr>
  `).join("");
}

function renderBacklogIssues() {
  document.getElementById("backlogIssues").innerHTML = dashboardData.pages.backlog.issues.map((item) => `
    <div class="issue-item">
      <strong>${item.title}</strong>
      <p>${item.text}</p>
    </div>
  `).join("");
}

function renderHoursRows() {
  document.getElementById("hoursRows").innerHTML = dashboardData.pages.hours.rows.map((row) => `
    <tr>
      <td>${row.yard}</td>
      <td class="num">${row.techs}</td>
      <td class="num">${row.payroll.toLocaleString("en-US")}</td>
      <td class="num">${row.assigned.toLocaleString("en-US")}</td>
      <td class="num">${row.invoiced.toLocaleString("en-US")}</td>
      <td class="num">${row.coverage}</td>
    </tr>
  `).join("");
}

function switchPage(page) {
  state.page = pageConfig[page] ? page : "overview";
  const appShell = document.getElementById("appShell");
  if (appShell) {
    appShell.classList.toggle("is-workflow-page", state.page === "overview");
  }
  document.querySelectorAll(".page-view").forEach((view) => {
    view.classList.toggle("is-active", view.dataset.page === state.page);
  });
  document.querySelectorAll(".nav-item").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.view === state.page);
  });
  const next = pageConfig[state.page];
  document.getElementById("pageTitle").textContent = next.title;
}

function bindNav() {
  document.querySelectorAll(".nav-item").forEach((button) => {
    button.addEventListener("click", () => {
      const page = button.dataset.view;
      if (pageConfig[page]) {
        window.location.hash = page;
        switchPage(page);
        renderAll();
      }
    });
  });

  document.getElementById("sidebarToggle").addEventListener("click", () => {
    document.getElementById("appShell").classList.toggle("sidebar-collapsed");
  });

  document.getElementById("resetBtn").addEventListener("click", () => {
    state.month = dashboardData.meta.months[0];
    state.yard = dashboardData.meta.yards[0];
    state.invoice = dashboardData.meta.invoiceScopes[0];
    state.ticketType = dashboardData.meta.ticketTypes?.[0] || "All Ticket Types";
    initFilters();
    renderAll();
  });

  document.getElementById("exportBtn").addEventListener("click", () => {
    window.location.href = dashboardData.meta.abnormalExportFile || "../abnormal_tickets.xlsx";
  });
}

function getInitialPage() {
  const hashPage = window.location.hash.replace("#", "");
  return pageConfig[hashPage] ? hashPage : "overview";
}

function bindChartTooltips() {
  const tooltip = document.getElementById("chartTooltip");
  if (!tooltip) return;

  document.addEventListener("pointerover", (event) => {
    const target = event.target.closest("[data-tooltip]");
    if (!target) return;
    tooltip.innerHTML = formatTooltipContent(target.dataset.tooltip);
    tooltip.classList.add("is-visible");
  });

  document.addEventListener("pointermove", (event) => {
    if (!tooltip.classList.contains("is-visible")) return;
    const offset = 16;
    const maxLeft = window.innerWidth - tooltip.offsetWidth - 12;
    const maxTop = window.innerHeight - tooltip.offsetHeight - 12;
    tooltip.style.left = `${Math.max(12, Math.min(event.clientX + offset, maxLeft))}px`;
    tooltip.style.top = `${Math.max(12, Math.min(event.clientY + offset, maxTop))}px`;
  });

  document.addEventListener("pointerout", (event) => {
    if (!event.target.closest("[data-tooltip]")) return;
    tooltip.classList.remove("is-visible");
  });
}

function renderAll() {
  renderPeriodBanner();
  renderCardPeriodNotes();
  const selectedKpis = buildSelectedKpis();
  renderPageKpis("overviewKpis", selectedKpis);
  renderWorkflow();
  renderPageKpis("backlogKpis", dashboardData.pages.backlog.kpis);
  renderPageKpis("hoursKpis", dashboardData.pages.hours.kpis);
  renderGroupedBars();
  renderInvoiceMix();
  renderOpenStatusMix();
  renderTrend();
  renderLegacyActivity();
  renderBacklogRows();
  renderBacklogIssues();
  renderHoursRows();
  renderUtilisation("hoursBalance", dashboardData.pages.hours.balance);
}

function renderPeriodBanner() {
  const periodType = /^\d{4}$/.test(state.month) ? "year" : "month";
  document.getElementById("stateBanner").textContent =
    `Current ${periodType}: ${state.month}. Dealer: ${state.workflowDealer}. Ticket type: Repair + PDI. Abnormal tickets: ${(dashboardData.meta.abnormalTickets || 0).toLocaleString("en-US")}.`;
}

function renderCardPeriodNotes() {
  document.querySelectorAll(".metric-chart-card .card-note").forEach((note) => {
    note.textContent = state.month;
  });
}

initFilters();
bindNav();
bindChartTooltips();
switchPage(getInitialPage());
window.addEventListener("hashchange", () => {
  switchPage(getInitialPage());
  renderAll();
});
renderAll();
