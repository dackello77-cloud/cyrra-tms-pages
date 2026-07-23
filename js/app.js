import {
  checkExistingTables,
  checkTable,
  checkTmsTables,
  createCarrier,
  createCarrierTender,
  createCarrierLoadBid,
  carrierPortalUpdateLoadStatus,
  createCustomer,
  createCustomerContact,
  createCustomerShipmentRequest,
  createCustomerPortalMessage,
  createDriver,
  createFacility,
  createLoad,
  createCustomerRateAgreement,
  createFreightQuote,
  createRecurringLoadTemplate,
  createLoadMessage,
  createDriverProblemReport,
  createPublicTrackingLink,
  createLoadGeofenceEvent,
  createOperationalException,
  createLoadStop,
  createMaintenance,
  createPayment,
  createInvoiceWithLineItems,
  createInvoiceBatch,
  createInvoiceCollectionNote,
  createInvoiceCreditMemo,
  createLoadSettlement,
  createSettlementPayBatch,
  createTrailer,
  createTruck,
  deleteTmsDocument,
  deleteLoadStop,
  duplicateLoad,
  driverUpdateLoadStatus,
  getLoadFinancials,
  getCurrentDriverProfile,
  getPublicTrackingSnapshot,
  getInvoiceDetails,
  getCompanyBillingSettings,
  getCurrentCustomerPortalAccess,
  getCurrentCarrierPortalAccess,
  getCurrentTmsAccess,
  getLoad,
  getSession,
  confirmLoadGeofenceEvent,
  generateMockTelematics,
  createNotification,
  listIntegrationConnections,
  listIntegrationExternalIds,
  listIntegrationRateLimits,
  listIntegrationWebhookEvents,
  listDriverProblemReports,
  listPublicTrackingAccessLogs,
  listPublicTrackingLinks,
  listBranches,
  listCompanyTenants,
  listNotificationDeliveryLogs,
  listNotificationPreferences,
  listNotifications,
  listLoadGeofenceEvents,
  listOperationalExceptions,
  listTelematicsHosSnapshots,
  listTelematicsPositions,
  listCarriers,
  listCarrierTenders,
  listCarrierLoadBids,
  listCarrierPortalTenders,
  listCustomers,
  listCustomerContacts,
  listCustomerShipmentRequests,
  listCustomerPortalMessages,
  listDocuments,
  listDispatchDocumentStates,
  listDrivers,
  listFacilities,
  listLoadDocuments,
  listLoadActivity,
  listLoadMessages,
  listLoads,
  listLoadsPage,
  listCustomerRateAgreements,
  listFreightQuotes,
  listRecurringLoadTemplates,
  listLoadStops,
  listLoadProfitability,
  listSavedReportViews,
  listScheduledReports,
  listReportDeliveryLogs,
  listMaintenance,
  listInvoices,
  listInvoicesPage,
  listInvoiceBatches,
  listInvoiceCollectionNotes,
  listSettlementCandidates,
  listSettlementPayBatches,
  listSettlements,
  manageTmsUsers,
  listTrailers,
  listTrucks,
  moveLoadStop,
  onAuthStateChange,
  saveLoadFinancials,
  saveNotificationPreference,
  saveReportView,
  deleteSavedReportView,
  listTableSavedViews,
  saveTableSavedView,
  deleteTableSavedView,
  saveScheduledReport,
  deleteScheduledReport,
  recordReportDelivery,
  updateScheduledReportRun,
  signIn,
  signOut,
  signUp,
  sendTmsInvoice,
  sendTmsInvoiceBatch,
  updateInvoiceStatus,
  updateInvoiceLineItemApproval,
  attachInvoiceLineItemReceipt,
  bulkUpdateInvoiceStatus,
  bulkUpdateLoadStatus,
  updateInvoiceCreditMemoStatus,
  updateOperationalException,
  updateIntegrationWebhookStatus,
  updateSettlementStatus,
  updateSettlementPayBatchStatus,
  updateCarrier,
  updateCustomer,
  updateCustomerShipmentRequest,
  updateDriver,
  updateFacility,
  updateLoad,
  updateFreightQuoteStatus,
  convertFreightQuoteToLoad,
  generateLoadFromRecurringTemplate,
  updateLoadAssignment,
  updateLoadBillingStatus,
  updateLoadStatus,
  updateLoadStop,
  updateMaintenance,
  updateMaintenanceStatus,
  markNotificationsRead,
  revokePublicTrackingLink,
  respondCarrierTender,
  submitCarrierLoadBid,
  selectCarrierLoadBid,
  updateTrailer,
  updateTruck,
  updateTmsDocumentStatus,
  updateCompanyBillingSettings,
  updateCompanyTenantProfile,
  uploadTmsDocument,
} from "./api.js";
import { appConfig } from "./config.js";
import { buildXlsxBlob } from "./xlsx.js";
import {
  calculateLoadProfitability,
  calculateQuotePricing,
  applyEtaSignalFreshness,
  carrierStatusHelp,
  collectionFollowUpState,
  customerPortalSummary,
  carrierPortalSummary,
  carrierBidComparison,
  differenceInCalendarDays,
  dispatchAssignmentConflicts,
  getComplianceDueState,
  getDispatchComplianceState,
  getMaintenanceDueState,
  hasTmsPermission,
  hosRiskState,
  invoiceBatchCandidates,
  invoiceFinancialState,
  invoiceLineItemsTotal,
  canManuallySetInvoiceStatus,
  integrationHealthSummary,
  notificationDeliverySummary,
  notificationEventTypes,
  notificationSummary,
  billingLoadStatus,
  driverStatusOptions,
  estimateStopArrival,
  exceptionTypeForPrediction,
  nextOperationalStop,
  loadDocumentStatus,
  loadAuditChanges,
  loadTotalMiles,
  latestTelematicsRecords,
  matchesDispatchSavedView,
  operationalLoadStatus,
  settlementStatusOptions,
  settlementPayBatchCandidates,
  telematicsFreshness,
  validateMessageBody,
  validateDuplicateLoadInput,
  validateCustomerBillingCycle,
  validateCompanyTenantProfile,
  validateDriverProblemReport,
  validateFacilityInput,
  validateLoadOperationalDetails,
  validateOperationalExceptionInput,
  validateInvoiceLineItems,
  validateInvoiceNumberingSettings,
  validatePublicTrackingLinkInput,
  validateFreightQuoteInput,
  validateCarrierBidInput,
  validateRecurringLoadTemplate,
  validateSavedReportView,
  validateTableSavedView,
  applyReportView,
  operationalKpiSummary,
  nextScheduledReportRun,
  validateScheduledReport,
  profitabilityPivot,
  rankGlobalSearchRecords,
  reportRecordValue,
  validateProfileFinancialTerms,
  validateStopWindows,
  manualInvoiceStatusOptions,
} from "./domain.js";

const authScreen = document.querySelector("#auth-screen");
const appShell = document.querySelector("#app-shell");
const loginForm = document.querySelector("#login-form");
const loginMessage = document.querySelector("#login-message");
const signupButton = document.querySelector("#signup-button");
const logoutButton = document.querySelector("#logout-button");
const addOrderButton = document.querySelector("#add-order-button");
const pageRoot = document.querySelector("#page-root");
const navLinks = Array.from(document.querySelectorAll(".nav-list a"));
const refreshButton = document.querySelector("#refresh-data");
const environmentBadge = document.querySelector("#environment-badge");
const profileBadge = document.querySelector("#profile-badge");
const globalSearch = document.querySelector("#global-search");
const globalSearchResults = document.querySelector("#global-search-results");
const isProduction = appConfig.name === "production";
const mapboxToken = appConfig.mapboxToken;

let currentSession = null;
let currentAccess = [];
let currentPortalAccess = [];
let currentCarrierPortalAccess = [];
let currentRoles = [];
let accessLoadError = "";
let authViewRequest = 0;
let fleetMapInstance = null;
let publicTrackingMapInstance = null;
let fleetTelemetryData = { positions: [], hos: [], unavailableMessage: "" };
let fleetNotice = "";
let settingsNotice = "";
let integrationNotice = "";
let settingsActiveTab = "Users";
let notificationNotice = "";
let publicTrackingNotice = "";
let reportsActiveTab = "overview";
let globalSearchCache = { records: [], loadedAt: 0 };
let globalSearchTimer = null;
let globalSearchActiveIndex = -1;
const loadListState = {
  loads: [],
  search: "",
  status: "",
  customerId: "",
  driverId: "",
  dateFrom: "",
  dateTo: "",
  view: "all",
  page: 1,
  pageSize: 10,
  total: 0,
  sortBy: "created_at",
  sortDirection: "desc",
  columns: ["status", "customer", "route", "pickup"],
  selectedIds: new Set(),
};
const invoiceListState = { records: [], search: "", status: "", page: 1, pageSize: 5, total: 0, sortBy: "created_at", sortDirection: "desc", columns: ["customer", "load", "amount", "due", "status"], selectedIds: new Set() };
let loadFilterTimer = null;
let invoiceFilterTimer = null;
const tableColumnOptions = {
  loads: [["status", "Status"], ["customer", "Customer"], ["route", "Carrier / Route"], ["pickup", "Pickup"]],
  invoices: [["customer", "Customer"], ["load", "Load"], ["amount", "Amount"], ["due", "Due"], ["status", "Status"]],
};

const tableSavedViewState = { loads: { loaded: false, records: [], error: "" }, invoices: { loaded: false, records: [], error: "" } };

function tableViewStorageKey() {
  return `cyrra-table-views-${currentSession?.user?.id || "anonymous"}`;
}

function savedTableViews() {
  try {
    return JSON.parse(localStorage.getItem(tableViewStorageKey()) || "{}") || {};
  } catch {
    return {};
  }
}

async function ensureTableSavedViews(type) {
  if (tableSavedViewState[type]?.loaded) return;
  try {
    const records = await listTableSavedViews(type);
    tableSavedViewState[type] = { loaded: true, records, error: "" };
  } catch (error) {
    tableSavedViewState[type] = { loaded: true, records: [], error: error.message };
  }
}

async function currentTableBranchId(type) {
  const accessBranch = currentAccess.find((access) => access.branch_id)?.branch_id;
  if (accessBranch) return accessBranch;
  const stateBranch = type === "loads"
    ? loadListState.loads.find((load) => load.branch_id)?.branch_id
    : invoiceListState.records.find((invoice) => invoice.branch_id)?.branch_id;
  if (stateBranch) return stateBranch;
  const branches = await listBranches();
  return branches[0]?.id || "";
}

function tableSavedViewsFor(type) {
  const remote = tableSavedViewState[type]?.records || [];
  if (remote.length || !tableSavedViewState[type]?.loaded) return remote;
  return (savedTableViews()[type] || []).map((view, index) => ({
    id: `local-${index}`,
    name: view.name,
    settings: view.settings,
    visibility: "private",
    isLocalFallback: true,
  }));
}

function renderTableViewControls(type) {
  const views = tableSavedViewsFor(type);
  const state = type === "loads" ? loadListState : invoiceListState;
  const canBulk = type === "loads" ? can("manage_operations") : can("manage_finance");
  const bulkOptions = type === "loads" ? [["booked", "Book selected"], ["new", "Return to New"]] : [["sent", "Mark Sent"], ["overdue", "Mark Overdue"]];
  return `<div class="table-view-controls" data-table-view-controls="${type}">
    ${canBulk ? `<select data-bulk-status hidden aria-label="Bulk status action">${bulkOptions.map(([value, label]) => `<option value="${value}">${label}</option>`).join("")}</select><button type="button" data-apply-bulk hidden>Apply to 0</button>` : ""}
    <details class="column-chooser"><summary>Columns</summary><div>${tableColumnOptions[type].map(([key, label]) => `<label><input type="checkbox" value="${key}" ${state.columns.includes(key) ? "checked" : ""}>${label}</label>`).join("")}<button type="button" data-reset-columns>Reset columns</button></div></details>
    <select data-saved-view aria-label="Saved ${type} view"><option value="">Saved views</option>${views.map((view, index) => `<option value="${index}">${escapeHtml(view.name)}${view.visibility === "shared" ? " · Shared" : ""}${view.isLocalFallback ? " · Local" : ""}</option>`).join("")}</select>
    <button type="button" data-save-view>Save view</button>
    <button type="button" data-delete-view disabled>Delete view</button>
    <button type="button" data-export-table>Export CSV</button>
    ${tableSavedViewState[type]?.error ? `<span class="list-helper">Using browser-local views. ${escapeHtml(tableSavedViewState[type].error)}</span>` : ""}
  </div>`;
}

function tableGridTemplate(type, columns) {
  const widths = type === "loads"
    ? { status: "minmax(110px,.7fr)", customer: "minmax(130px,1fr)", route: "minmax(170px,1.2fr)", pickup: "minmax(110px,.7fr)" }
    : { customer: "minmax(130px,1fr)", load: "minmax(70px,.5fr)", amount: "minmax(125px,.8fr)", due: "minmax(105px,.7fr)", status: "minmax(120px,.75fr)" };
  const canBulk = type === "loads" ? can("manage_operations") : can("manage_finance");
  return [canBulk ? "34px" : null, `minmax(100px,.75fr)`, ...columns.map((column) => widths[column]), type === "loads" ? "196px" : "minmax(360px,2fr)"].filter(Boolean).join(" ");
}

function tableRowStyle(type, columns) {
  const minWidth = type === "invoices" ? 500 + columns.length * 125 : 390 + columns.length * 125;
  return `grid-template-columns:${tableGridTemplate(type, columns)};min-width:${minWidth}px`;
}

function tableViewSnapshot(type) {
  if (type === "loads") return {
    search: loadListState.search, status: loadListState.status, customerId: loadListState.customerId, driverId: loadListState.driverId,
    dateFrom: loadListState.dateFrom, dateTo: loadListState.dateTo, view: loadListState.view, pageSize: loadListState.pageSize,
    sortBy: loadListState.sortBy, sortDirection: loadListState.sortDirection, columns: [...loadListState.columns],
  };
  return { search: invoiceListState.search, status: invoiceListState.status, pageSize: invoiceListState.pageSize, sortBy: invoiceListState.sortBy, sortDirection: invoiceListState.sortDirection, columns: [...invoiceListState.columns] };
}

function persistTableViews(type, views) {
  const stored = savedTableViews();
  stored[type] = views;
  localStorage.setItem(tableViewStorageKey(), JSON.stringify(stored));
}

function updateTableSavedViewOptions(type, select, selectedId = "") {
  const views = tableSavedViewsFor(type);
  select.innerHTML = '<option value="">Saved views</option>' + views.map((view, index) => `<option value="${index}">${escapeHtml(view.name)}${view.visibility === "shared" ? " · Shared" : ""}${view.isLocalFallback ? " · Local" : ""}</option>`).join("");
  if (selectedId) {
    const index = views.findIndex((view) => view.id === selectedId);
    if (index >= 0) select.value = String(index);
  }
}

function clearSelectedTableView(type) {
  const controls = document.querySelector(`[data-table-view-controls="${type}"]`);
  if (!controls) return;
  controls.querySelector("[data-saved-view]").value = "";
  controls.querySelector("[data-delete-view]").disabled = true;
}

function bindTableViewControls(type) {
  const controls = document.querySelector(`[data-table-view-controls="${type}"]`);
  if (!controls) return;
  const state = type === "loads" ? loadListState : invoiceListState;
  const reload = type === "loads" ? loadLoads : loadInvoices;
  const rerender = type === "loads" ? renderFilteredLoads : loadInvoices;
  const select = controls.querySelector("[data-saved-view]");
  const deleteButton = controls.querySelector("[data-delete-view]");
  controls.querySelectorAll('.column-chooser input[type="checkbox"]').forEach((checkbox) => checkbox.addEventListener("change", () => {
    const selected = [...controls.querySelectorAll('.column-chooser input[type="checkbox"]:checked')].map((input) => input.value);
    state.columns = selected;
    clearSelectedTableView(type);
    rerender();
  }));
  controls.querySelector("[data-reset-columns]").addEventListener("click", () => {
    state.columns = tableColumnOptions[type].map(([key]) => key);
    clearSelectedTableView(type);
    controls.querySelectorAll('.column-chooser input[type="checkbox"]').forEach((input) => { input.checked = true; });
    rerender();
  });
  controls.querySelector("[data-save-view]").addEventListener("click", async () => {
    const name = window.prompt("View name:")?.trim();
    if (!name) return;
    const branchId = await currentTableBranchId(type);
    try {
      if (!branchId) throw new Error("A branch is required to save shared table views.");
      const visibility = window.confirm("Share this view with users in this branch?") ? "shared" : "private";
      const validated = validateTableSavedView(type, { name, visibility, settings: tableViewSnapshot(type) });
      const saved = await saveTableSavedView({ ...validated, branch_id: branchId, created_by: currentSession.user.id });
      tableSavedViewState[type].records = [saved, ...tableSavedViewState[type].records.filter((view) => view.id !== saved.id && view.name.toLowerCase() !== saved.name.toLowerCase())]
        .sort((a, b) => String(a.visibility).localeCompare(String(b.visibility)) || String(a.name).localeCompare(String(b.name)));
      updateTableSavedViewOptions(type, select, saved.id);
      deleteButton.disabled = false;
    } catch (error) {
      const views = savedTableViews()[type] || [];
      const existing = views.findIndex((view) => view.name.toLowerCase() === name.toLowerCase());
      const saved = { name, settings: tableViewSnapshot(type) };
      if (existing >= 0) views[existing] = saved; else views.push(saved);
      persistTableViews(type, views);
      tableSavedViewState[type] = { loaded: true, records: [], error: error.message };
      updateTableSavedViewOptions(type, select);
      select.value = String(existing >= 0 ? existing : views.length - 1);
      deleteButton.disabled = false;
    }
  });
  select.addEventListener("change", () => {
    const selected = select.value === "" ? null : tableSavedViewsFor(type)[Number(select.value)];
    deleteButton.disabled = !selected || (!selected.isLocalFallback && selected.created_by !== currentSession?.user?.id);
    if (select.value === "") return;
    const saved = tableSavedViewsFor(type)[Number(select.value)];
    if (!saved) return;
    Object.assign(state, saved.settings, { page: 1 });
    if (type === "loads") syncLoadFilterForm(); else syncInvoiceFilterForm();
    controls.querySelectorAll('.column-chooser input[type="checkbox"]').forEach((input) => { input.checked = state.columns.includes(input.value); });
    reload();
  });
  deleteButton.addEventListener("click", async () => {
    if (select.value === "") return;
    const views = tableSavedViewsFor(type);
    const selectedIndex = Number(select.value);
    const selectedView = views[selectedIndex];
    if (!selectedView || !window.confirm(`Delete saved view "${selectedView.name}"?`)) return;
    if (selectedView.isLocalFallback) {
      const localViews = savedTableViews()[type] || [];
      localViews.splice(selectedIndex, 1);
      persistTableViews(type, localViews);
    } else {
      try {
        await deleteTableSavedView(selectedView.id);
        tableSavedViewState[type].records = tableSavedViewState[type].records.filter((view) => view.id !== selectedView.id);
      } catch (error) {
        window.alert(error.message);
        return;
      }
    }
    updateTableSavedViewOptions(type, select);
    deleteButton.disabled = true;
  });
  controls.querySelector("[data-export-table]").addEventListener("click", (event) => handleTableCsvExport(type, event.currentTarget));
  controls.querySelector("[data-apply-bulk]")?.addEventListener("click", () => handleBulkTableAction(type));
}

function updateBulkTableControls(type) {
  const state = type === "loads" ? loadListState : invoiceListState;
  const controls = document.querySelector(`[data-table-view-controls="${type}"]`);
  const select = controls?.querySelector("[data-bulk-status]");
  const button = controls?.querySelector("[data-apply-bulk]");
  if (!select || !button) return;
  const count = state.selectedIds.size;
  select.hidden = count === 0;
  button.hidden = count === 0;
  button.disabled = count === 0;
  button.textContent = `Apply to ${count}`;
}

function bindTableRowSelection(type) {
  const state = type === "loads" ? loadListState : invoiceListState;
  const table = document.querySelector(type === "loads" ? "#loads-table" : "#invoices-table");
  const visibleRecords = type === "loads" ? loadListState.loads : invoiceListState.records;
  table?.querySelector("[data-select-page]")?.addEventListener("change", (event) => {
    visibleRecords.forEach((record) => event.currentTarget.checked ? state.selectedIds.add(record.id) : state.selectedIds.delete(record.id));
    table.querySelectorAll("[data-select-row]").forEach((input) => { input.checked = event.currentTarget.checked; });
    updateBulkTableControls(type);
  });
  table?.querySelectorAll("[data-select-row]").forEach((input) => input.addEventListener("change", () => {
    if (input.checked) state.selectedIds.add(input.value); else state.selectedIds.delete(input.value);
    const pageToggle = table.querySelector("[data-select-page]");
    if (pageToggle) pageToggle.checked = visibleRecords.length > 0 && visibleRecords.every((record) => state.selectedIds.has(record.id));
    updateBulkTableControls(type);
  }));
  updateBulkTableControls(type);
}

async function handleBulkTableAction(type) {
  const state = type === "loads" ? loadListState : invoiceListState;
  const controls = document.querySelector(`[data-table-view-controls="${type}"]`);
  const targetStatus = controls.querySelector("[data-bulk-status]").value;
  const ids = [...state.selectedIds];
  const label = formatStatus(targetStatus);
  if (!ids.length || !window.confirm(`Change ${ids.length} selected ${type} to ${label}?\n\nThe entire action will fail if any selected record is ineligible.`)) return;
  const button = controls.querySelector("[data-apply-bulk]");
  button.disabled = true;
  button.textContent = "Applying...";
  try {
    if (type === "loads") await bulkUpdateLoadStatus(ids, targetStatus); else await bulkUpdateInvoiceStatus(ids, targetStatus);
    state.selectedIds.clear();
    await (type === "loads" ? loadLoads() : loadInvoices());
  } catch (error) {
    window.alert(`Bulk action failed: ${error.message}`);
    button.disabled = false;
    updateBulkTableControls(type);
  }
}

async function allPagedRecords(type, state) {
  const records = [];
  let page = 1;
  let total = 0;
  do {
    const result = type === "loads"
      ? await listLoadsPage({ ...state, page, pageSize: 100, today: localDateIso() })
      : await listInvoicesPage({ ...state, page, pageSize: 100 });
    if (!result.records.length && records.length < result.count) throw new Error("Export pagination stopped before all matching records were loaded.");
    records.push(...result.records);
    total = result.count;
    page += 1;
  } while (records.length < total);
  return records;
}

function loadCsvRows(loads) {
  return loads.map((load) => {
    const row = { "Load #": load.load_no };
    if (loadListState.columns.includes("status")) Object.assign(row, { "Operational status": formatStatus(operationalLoadStatus(load)), "Billing status": formatStatus(billingLoadStatus(load)) });
    if (loadListState.columns.includes("customer")) row.Customer = load.customers?.name || "";
    if (loadListState.columns.includes("route")) Object.assign(row, { Carrier: load.carriers?.name || "", Origin: load.origin || "", Destination: load.destination || "" });
    if (loadListState.columns.includes("pickup")) row.Pickup = load.pickup_date || "";
    return row;
  });
}

function invoiceCsvRows(invoices) {
  return invoices.map((invoice) => {
    const financial = invoiceFinancialState(invoice, localDateIso());
    const row = { "Invoice #": invoice.invoice_no };
    if (invoiceListState.columns.includes("customer")) row.Customer = invoice.customers?.name || "";
    if (invoiceListState.columns.includes("load")) row.Load = invoice.loads?.load_no || "";
    if (invoiceListState.columns.includes("amount")) Object.assign(row, { Amount: numberOrZero(invoice.amount).toFixed(2), Balance: financial.balance.toFixed(2) });
    if (invoiceListState.columns.includes("due")) row.Due = invoice.due_date || "";
    if (invoiceListState.columns.includes("status")) row.Status = formatStatus(financial.displayStatus);
    return row;
  });
}

async function handleTableCsvExport(type, button) {
  const originalLabel = button.textContent;
  button.disabled = true;
  button.textContent = "Exporting...";
  try {
    const state = type === "loads" ? loadListState : invoiceListState;
    const records = await allPagedRecords(type, state);
    if (!records.length) {
      window.alert("No matching records to export.");
      return;
    }
    const rows = type === "loads" ? loadCsvRows(records) : invoiceCsvRows(records);
    downloadCsv(`cyrra-${type}-${localDateIso()}.csv`, rows);
  } catch (error) {
    window.alert(`CSV export failed: ${error.message}`);
  } finally {
    button.disabled = false;
    button.textContent = originalLabel;
  }
}

document.addEventListener("click", (event) => {
  document.querySelectorAll(".column-chooser[open]").forEach((chooser) => {
    if (!chooser.contains(event.target)) chooser.removeAttribute("open");
  });
});

document.addEventListener("keydown", (event) => {
  if (event.key !== "Escape") return;
  document.querySelectorAll(".column-chooser[open]").forEach((chooser) => chooser.removeAttribute("open"));
});

environmentBadge.textContent = appConfig.name;
environmentBadge.dataset.environment = appConfig.name;
signupButton.hidden = isProduction;
signupButton.disabled = isProduction;

function route() {
  const requested = window.location.hash.replace("#", "");
  if (requested.startsWith("loads/")) return "loads";
  if (requested) return requested;
  if (currentRoles.includes("customer_portal") && currentRoles.length === 1) return "customer_portal";
  if (currentRoles.includes("carrier_portal") && currentRoles.length === 1) return "carrier_portal";
  return currentRoles.includes("driver") && currentRoles.length === 1 ? "driver" : "dashboard";
}

function loadIdFromHash() {
  const requested = window.location.hash.replace("#", "");
  return requested.startsWith("loads/") ? decodeURIComponent(requested.slice("loads/".length)) : "";
}

function publicTrackingTokenFromHash() {
  const hash = window.location.hash.replace("#", "");
  if (!hash.startsWith("track/")) return "";
  return decodeURIComponent(hash.slice("track/".length));
}

async function setAuthView(session) {
  const requestId = ++authViewRequest;
  currentSession = session;
  Object.assign(tableSavedViewState, { loads: { loaded: false, records: [], error: "" }, invoices: { loaded: false, records: [], error: "" } });

  const publicTrackingToken = publicTrackingTokenFromHash();
  if (publicTrackingToken) {
    authScreen.hidden = true;
    appShell.hidden = false;
    currentAccess = [];
    currentPortalAccess = [];
    currentCarrierPortalAccess = [];
    currentRoles = [];
    applyRoleAwareShell();
    renderPublicTracking(publicTrackingToken);
    return;
  }

  authScreen.hidden = Boolean(session);
  appShell.hidden = !session;

  if (session) {
    pageRoot.innerHTML = '<div class="empty-state">Loading user access...</div>';
    try {
      [currentAccess, currentPortalAccess, currentCarrierPortalAccess] = await Promise.all([
        getCurrentTmsAccess(session.user.id),
        getCurrentCustomerPortalAccess(session.user.id),
        getCurrentCarrierPortalAccess(session.user.id),
      ]);
      accessLoadError = "";
    } catch (error) {
      currentAccess = [];
      currentPortalAccess = [];
      currentCarrierPortalAccess = [];
      accessLoadError = error.message;
    }
    if (requestId !== authViewRequest) return;
    currentRoles = [...new Set([
      ...currentAccess.map((access) => access.role),
      ...(currentPortalAccess.length ? ["customer_portal"] : []),
      ...(currentCarrierPortalAccess.length ? ["carrier_portal"] : []),
    ])];
    applyRoleAwareShell();
    refreshNotificationBadge();
    renderRoute();
  } else {
    currentAccess = [];
    currentPortalAccess = [];
    currentCarrierPortalAccess = [];
    currentRoles = [];
    accessLoadError = "";
  }
}

function can(permission) {
  return hasTmsPermission(currentRoles, permission);
}

function applyRoleAwareShell() {
  navLinks.forEach((link) => {
    link.hidden = !can(`view_${link.dataset.route}`);
  });
  addOrderButton.hidden = !can("manage_operations");
  const notificationButton = document.querySelector("#notification-button");
  notificationButton.hidden = !can("view_notifications");
  const avatar = document.querySelector(".user-avatar");
  const roleLabel = currentRoles.length ? currentRoles.map(formatStatus).join(", ") : "No active role";
  updateProfileBadge();
  avatar.title = roleLabel;
  avatar.setAttribute("aria-label", roleLabel);
}

function accessScopeLabel() {
  const branchNames = currentAccess
    .map((access) => access.branches?.name)
    .filter(Boolean);
  const customerNames = currentPortalAccess
    .map((access) => access.customers?.name)
    .filter(Boolean);
  const carrierNames = currentCarrierPortalAccess
    .map((access) => access.carriers?.name)
    .filter(Boolean);
  const scopes = [
    ...branchNames.map((name) => `Branch: ${name}`),
    ...customerNames.map((name) => `Customer: ${name}`),
    ...carrierNames.map((name) => `Carrier: ${name}`),
  ];
  return scopes.length ? [...new Set(scopes)].join(" · ") : "No active scope";
}

function userDisplayName() {
  const metadata = currentSession?.user?.user_metadata || {};
  const fullName = [metadata.first_name, metadata.last_name].filter(Boolean).join(" ").trim();
  return fullName || metadata.full_name || metadata.name || currentSession?.user?.email || "Signed out";
}

function userInitials(name) {
  const text = String(name || currentSession?.user?.email || "U").trim();
  const parts = text.includes("@") ? [text.charAt(0)] : text.split(/\s+/).filter(Boolean);
  return (parts.length > 1 ? `${parts[0][0]}${parts[1][0]}` : parts[0]?.slice(0, 2) || "U").toUpperCase();
}

function updateProfileBadge() {
  const avatar = document.querySelector(".user-avatar");
  if (!profileBadge || !avatar) return;
  if (!currentSession) {
    profileBadge.innerHTML = '<span class="profile-badge-label">Signed out</span><strong>Not connected</strong><small>No active role</small>';
    profileBadge.title = "No active session";
    avatar.textContent = "MS";
    return;
  }
  const name = userDisplayName();
  const roles = currentRoles.length ? currentRoles.map(formatStatus).join(", ") : "No active role";
  const scope = accessScopeLabel();
  const email = currentSession.user.email || "";
  profileBadge.innerHTML = `
    <span class="profile-badge-label">${escapeHtml(appConfig.name)}</span>
    <strong>${escapeHtml(name)}</strong>
    <small>${escapeHtml(roles)} · ${escapeHtml(scope)}</small>
  `;
  profileBadge.title = [email, roles, scope].filter(Boolean).join(" · ");
  avatar.textContent = userInitials(name);
}

function renderAccessDenied(activeRoute) {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Access", title: "Access Restricted" })}
    <section class="panel access-denied-panel">
      <h2>${accessLoadError ? "Access profile unavailable" : "This module is not assigned to your role"}</h2>
      <p>${escapeHtml(accessLoadError || `Your active role (${currentRoles.map(formatStatus).join(", ") || "none"}) cannot open ${formatStatus(activeRoute)}.`)}</p>
      <a class="button-link" href="#dashboard">Return to Dashboard</a>
    </section>`;
}

async function handleLogin(event) {
  event.preventDefault();
  loginMessage.textContent = "Logging in...";

  const formData = new FormData(loginForm);
  const email = formData.get("email") || document.querySelector("#login-email").value;
  const password = formData.get("password") || document.querySelector("#login-password").value;

  try {
    await signIn(email, password);
    loginForm.reset();
    loginMessage.textContent = "";
  } catch (error) {
    loginMessage.textContent = error.message;
  }
}

async function handleSignup() {
  if (isProduction) {
    loginMessage.textContent = "Account creation is disabled in production.";
    return;
  }

  loginMessage.textContent = "Creating account...";

  const formData = new FormData(loginForm);
  const email = formData.get("email") || document.querySelector("#login-email").value;
  const password = formData.get("password") || document.querySelector("#login-password").value;

  if (!email || !password) {
    loginMessage.textContent = "Enter email and password first.";
    return;
  }

  try {
    const session = await signUp(email, password);
    loginMessage.textContent = session
      ? "Account created."
      : "Account created. Check email confirmation settings in Supabase if login is not automatic.";
  } catch (error) {
    loginMessage.textContent = error.message;
  }
}

function updateActiveNav(activeRoute) {
  navLinks.forEach((link) => {
    link.classList.toggle("active", link.dataset.route === activeRoute);
  });
}

function renderRoute() {
  const publicTrackingToken = publicTrackingTokenFromHash();
  if (publicTrackingToken) {
    authScreen.hidden = true;
    appShell.hidden = false;
    renderPublicTracking(publicTrackingToken);
    return;
  }

  if (!currentSession) {
    return;
  }

  const activeRoute = route();
  updateActiveNav(activeRoute);

  if (!can(`view_${activeRoute}`)) {
    renderAccessDenied(activeRoute);
    return;
  }

  if (activeRoute === "customers") {
    renderCustomers();
    return;
  }

  if (activeRoute === "carriers") {
    renderCarriers();
    return;
  }

  if (activeRoute === "drivers") {
    renderDrivers();
    return;
  }

  if (activeRoute === "equipment") {
    renderEquipment();
    return;
  }

  if (activeRoute === "maintenance") {
    renderMaintenance();
    return;
  }

  if (activeRoute === "tracking") {
    renderFleetMap();
    return;
  }

  if (activeRoute === "driver") {
    renderDriverPortal();
    return;
  }

  if (activeRoute === "customer_portal") {
    renderCustomerPortal();
    return;
  }

  if (activeRoute === "carrier_portal") {
    renderCarrierPortal();
    return;
  }

  if (activeRoute === "documents") {
    renderDocumentCenter();
    return;
  }

  if (activeRoute === "loads") {
    renderLoads();
    return;
  }

  if (activeRoute === "quotes") {
    renderQuotes();
    return;
  }

  if (activeRoute === "dispatch") {
    renderDispatchBoard();
    return;
  }

  if (activeRoute === "billing") {
    renderBilling();
    return;
  }

  if (activeRoute === "settlements") {
    renderSettlements();
    return;
  }

  if (activeRoute === "reports") {
    renderReports();
    return;
  }

  if (activeRoute === "notifications") {
    renderNotifications();
    return;
  }

  if (activeRoute === "settings") {
    renderSettings();
    return;
  }

  renderDashboard();
}

function renderPageHeader({ eyebrow, title, status }) {
  return `
    <section class="page-header">
      <div>
        <p class="eyebrow">${eyebrow}</p>
        <h1>${title}</h1>
      </div>
      ${status ? `<span id="connection-status" class="status-pill">${status}</span>` : ""}
    </section>
  `;
}

async function renderPublicTracking(token) {
  updateActiveNav("__public_tracking");
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Public Tracking", title: "Shipment Status", status: "Loading" })}
    <section class="public-tracking-page">
      <article class="panel">
        <div class="panel-header">
          <div><h2>Tracking Link</h2><p>Limited public view controlled by the carrier.</p></div>
          <button id="public-tracking-refresh" type="button">Refresh</button>
        </div>
        <div id="public-tracking-content" class="public-tracking-content">
          <div class="empty-state">Loading public shipment status...</div>
        </div>
      </article>
    </section>
  `;
  document.querySelector("#public-tracking-refresh").addEventListener("click", () => renderPublicTracking(token));
  const content = document.querySelector("#public-tracking-content");
  try {
    const snapshot = await getPublicTrackingSnapshot(token);
    document.querySelector("#connection-status").textContent = "Live";
    content.innerHTML = renderPublicTrackingSnapshot(snapshot);
    initPublicTrackingMap(snapshot);
  } catch (error) {
    document.querySelector("#connection-status").textContent = "Unavailable";
    content.innerHTML = `<div class="empty-state">Tracking link is unavailable: ${escapeHtml(error.message)}</div>`;
  }
}

function renderPublicTrackingSnapshot(snapshot) {
  const privacy = snapshot.privacy || {};
  return `
    <section class="public-tracking-hero">
      <div>
        <span class="eyebrow">${escapeHtml(snapshot.load_no || "Shipment")}</span>
        <h2>${escapeHtml([snapshot.origin, snapshot.destination].filter(Boolean).join(" -> ") || "Shipment in progress")}</h2>
      </div>
      ${snapshot.status ? `<span class="status-pill">${escapeHtml(formatStatus(snapshot.status))}</span>` : ""}
    </section>
    <section class="public-tracking-grid">
      <article><span>Pickup</span><strong>${escapeHtml(snapshot.pickup_date || "-")}</strong></article>
      <article><span>Delivery</span><strong>${escapeHtml(snapshot.delivery_date || "-")}</strong></article>
      <article><span>Link expires</span><strong>${formatDateTime(snapshot.expires_at)}</strong></article>
      ${privacy.show_eta ? `<article><span>Next stop</span><strong>${escapeHtml(snapshot.next_stop?.facility || "No pending stop")}</strong><small>${snapshot.next_stop ? formatDateTime(snapshot.next_stop.appointment_from) : ""}</small></article>` : ""}
      ${privacy.show_driver && snapshot.driver ? `<article><span>Driver</span><strong>${escapeHtml(snapshot.driver.name || "-")}</strong><small>${escapeHtml(snapshot.driver.phone || "")}</small></article>` : ""}
      ${privacy.show_equipment && snapshot.equipment ? `<article><span>Equipment</span><strong>${escapeHtml([snapshot.equipment.truck, snapshot.equipment.trailer].filter(Boolean).join(" / ") || "-")}</strong></article>` : ""}
    </section>
    ${privacy.show_location && snapshot.location ? `<section class="public-location-box">
      <h3>Latest Location</h3>
      <div id="public-tracking-map" class="public-tracking-map" aria-label="Latest shipment location map"></div>
      <p>${escapeHtml(`${snapshot.location.latitude}, ${snapshot.location.longitude}`)}</p>
      <small>${formatDateTime(snapshot.location.event_at)}</small>
    </section>` : ""}
  `;
}

function initPublicTrackingMap(snapshot) {
  publicTrackingMapInstance?.remove();
  publicTrackingMapInstance = null;
  const location = snapshot?.location;
  const mapEl = document.querySelector("#public-tracking-map");
  if (!mapEl || !location?.latitude || !location?.longitude) return;
  if (!window.mapboxgl) {
    mapEl.innerHTML = '<div class="mapbox-message">Map is unavailable.</div>';
    return;
  }
  window.mapboxgl.accessToken = mapboxToken;
  const coordinates = [Number(location.longitude), Number(location.latitude)];
  publicTrackingMapInstance = new window.mapboxgl.Map({
    container: mapEl,
    style: "mapbox://styles/mapbox/streets-v12",
    center: coordinates,
    zoom: 9,
    attributionControl: false,
  });
  publicTrackingMapInstance.addControl(new window.mapboxgl.NavigationControl({ showCompass: false }), "top-right");
  publicTrackingMapInstance.on("load", () => {
    const markerEl = document.createElement("div");
    markerEl.className = "mapbox-live-marker";
    markerEl.dataset.state = "live";
    new window.mapboxgl.Marker(markerEl)
      .setLngLat(coordinates)
      .setPopup(new window.mapboxgl.Popup({ offset: 18 }).setHTML(`
        <strong>Latest location</strong>
        <p>${formatDateTime(location.event_at)}</p>
      `))
      .addTo(publicTrackingMapInstance);
  });
}

async function renderCustomerPortal() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Customer Portal", title: "My Shipments", status: "Loading" })}
    <section id="customer-portal-root" class="customer-portal-page">
      <article class="panel">
        <div class="empty-state">Loading customer shipments...</div>
      </article>
    </section>
  `;

  const root = document.querySelector("#customer-portal-root");
  try {
    const [loads, invoices, requests, portalMessages] = await Promise.all([listLoads(), listInvoices(), listCustomerShipmentRequests(), listCustomerPortalMessages()]);
    const visibleInvoices = invoices.filter((invoice) => invoice.status !== "void");
    const documentPairs = await Promise.all(loads.map(async (load) => [load.id, await listLoadDocuments(load.id)]));
    const documentsByLoad = new Map(documentPairs);
    const summary = customerPortalSummary(loads, visibleInvoices);
    document.querySelector("#connection-status").textContent = "Ready";
    root.innerHTML = `
      <section class="portal-summary">
        <article><span>Shipments</span><strong>${summary.shipments}</strong></article>
        <article><span>Active</span><strong>${summary.active}</strong></article>
        <article><span>Delivered</span><strong>${summary.delivered}</strong></article>
        <article><span>Open invoices</span><strong>${summary.openInvoices}</strong></article>
      </section>
      <section class="panel">
        <div class="panel-header"><div><h2>Shipment Requests</h2><p>Request a quote or a new shipment and continue the conversation with operations.</p></div></div>
        <form id="customer-request-form" class="customer-request-form">
          <label><span>Request type</span><select name="request_type"><option value="quote">Quote</option><option value="shipment">Shipment</option></select></label>
          <label><span>Pickup date</span><input name="pickup_date" type="date" required></label>
          <label><span>Delivery date</span><input name="delivery_date" type="date"></label>
          <label><span>Origin</span><input name="origin" required placeholder="Chicago, IL"></label>
          <label><span>Destination</span><input name="destination" required placeholder="Dallas, TX"></label>
          <label><span>Equipment</span><input name="equipment_type" placeholder="Dry Van"></label>
          <label><span>Commodity</span><input name="commodity" placeholder="General Freight"></label>
          <label><span>Loaded miles</span><input name="loaded_miles" type="number" min="0" step="0.1"></label>
          <label class="form-field-wide"><span>Notes</span><textarea name="notes" rows="3"></textarea></label>
          <button class="primary-button form-field-wide" type="submit">Submit Request</button><p class="form-message form-field-wide" aria-live="polite"></p>
        </form>
        <div class="portal-load-list">${requests.length?requests.map((request)=>renderCustomerRequest(request,portalMessages.filter((message)=>message.request_id===request.id))).join(""):'<div class="empty-state compact-empty">No shipment requests yet.</div>'}</div>
      </section>
      <section class="portal-layout">
        <article class="panel portal-shipments">
          <div class="panel-header">
            <div><h2>Shipments</h2><p>Status, stops, approved documents and visible invoices for your account.</p></div>
            <button id="reload-customer-portal" type="button">Refresh</button>
          </div>
          <div class="portal-account-list">
            ${currentPortalAccess.length ? currentPortalAccess.map((access) => `<span>${escapeHtml(access.customers?.name || "Customer account")}</span>`).join("") : '<span>No customer account assigned</span>'}
          </div>
          <div class="portal-load-list">
            ${loads.length ? loads.map((load) => renderCustomerPortalLoad(load, documentsByLoad.get(load.id) || [], visibleInvoices.filter((invoice) => invoice.load_id === load.id))).join("") : '<div class="empty-state compact-empty">No visible shipments yet.</div>'}
          </div>
        </article>
        <aside class="panel portal-invoices">
          <div class="panel-header"><div><h2>Invoices</h2><p>Draft invoices and internal payment notes are not shown here.</p></div></div>
          <div class="portal-invoice-list">
            ${visibleInvoices.length ? visibleInvoices.map(renderCustomerPortalInvoice).join("") : '<div class="empty-state compact-empty">No visible invoices yet.</div>'}
          </div>
        </aside>
      </section>
    `;
    document.querySelector("#reload-customer-portal").addEventListener("click", renderCustomerPortal);
    document.querySelector("#customer-request-form").addEventListener("submit",handleCustomerRequestSubmit);
    document.querySelectorAll("[data-customer-request-message]").forEach((form)=>form.addEventListener("submit",handleCustomerRequestMessage));
  } catch (error) {
    document.querySelector("#connection-status").dataset.state = "error";
    document.querySelector("#connection-status").textContent = "Unavailable";
    root.innerHTML = `<article class="panel"><div class="empty-state">Customer Portal is not ready: ${escapeHtml(error.message)}</div></article>`;
  }
}

function renderCustomerRequest(request,messages=[]){return `<article class="portal-load-card customer-request-card"><header><div><strong>${escapeHtml(request.request_no)}</strong><span>${escapeHtml(formatStatus(request.request_type))} · ${escapeHtml(request.origin)} → ${escapeHtml(request.destination)}</span></div><span class="status-pill" data-status="${escapeHtml(request.status)}">${escapeHtml(formatStatus(request.status))}</span></header><small>Pickup ${formatDate(request.pickup_date)}${request.delivery_date?` · Delivery ${formatDate(request.delivery_date)}`:""}</small>${request.notes?`<p class="customer-request-note">${escapeHtml(request.notes)}</p>`:""}<section class="portal-mini-section customer-request-thread"><h3>Messages</h3><div class="customer-request-messages">${messages.map((message)=>`<article class="customer-request-message ${message.sender_id===currentSession?.user?.id?"is-own":""}"><div><strong>${message.sender_id===currentSession?.user?.id?"You":"Operations"}</strong><small>${formatDateTime(message.created_at)}</small></div><p>${escapeHtml(message.body)}</p></article>`).join("")||'<small>No messages yet.</small>'}</div><form data-customer-request-message="${request.id}" class="customer-request-message-form"><input name="body" required maxlength="2000" placeholder="Write a message"><button type="submit">Send</button></form></section></article>`;}
async function handleCustomerRequestSubmit(event){event.preventDefault();const form=event.currentTarget;const message=form.querySelector(".form-message");try{const data=Object.fromEntries(new FormData(form));const access=currentPortalAccess[0];if(!access)throw new Error("Customer portal access is required.");if(data.delivery_date&&data.delivery_date<data.pickup_date)throw new Error("Delivery cannot be before pickup.");await createCustomerShipmentRequest({...data,customer_id:access.customer_id,branch_id:access.branch_id,loaded_miles:data.loaded_miles?Number(data.loaded_miles):null,delivery_date:data.delivery_date||null});await renderCustomerPortal();}catch(error){message.textContent=error.message;}}
async function handleCustomerRequestMessage(event){event.preventDefault();const form=event.currentTarget;try{const body=String(new FormData(form).get("body")||"").trim();if(!body)throw new Error("Message is required.");await createCustomerPortalMessage({request_id:form.dataset.customerRequestMessage,body});await renderCustomerPortal();}catch(error){window.alert(error.message);}}

async function renderCarrierPortal() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Carrier Portal", title: "Assigned Freight", status: "Loading" })}
    <section id="carrier-portal-root" class="customer-portal-page"><article class="panel"><div class="empty-state">Loading carrier freight...</div></article></section>`;
  const root = document.querySelector("#carrier-portal-root");
  try {
    const [loads, tenders, settlements, bids] = await Promise.all([listLoads(), listCarrierPortalTenders(), listSettlements(), listCarrierLoadBids()]);
    const [documentPairs, messagePairs] = await Promise.all([
      Promise.all(loads.map(async (load) => [load.id, await listLoadDocuments(load.id)])),
      Promise.all(loads.map(async (load) => [load.id, await listLoadMessages(load.id)])),
    ]);
    const documentsByLoad = new Map(documentPairs);
    const messagesByLoad = new Map(messagePairs);
    const summary = carrierPortalSummary(loads, tenders, settlements);
    document.querySelector("#connection-status").textContent = "Ready";
    root.innerHTML = `
      <section class="portal-summary">
        <article><span>Assigned loads</span><strong>${summary.loads}</strong></article>
        <article><span>Pending tenders</span><strong>${summary.pendingTenders}</strong></article>
        <article><span>Active</span><strong>${summary.active}</strong></article>
        <article><span>Approved / paid</span><strong>${formatMoney(summary.payable)}</strong></article>
      </section>
      <section class="portal-layout">
        <article class="panel portal-shipments">
          <div class="panel-header"><div><h2>Freight</h2><p>Tenders, assigned loads, stops and carrier-facing documents.</p></div><button id="reload-carrier-portal" type="button">Refresh</button></div>
          <div class="portal-account-list">${currentCarrierPortalAccess.map((access) => `<span>${escapeHtml(access.carriers?.name || "Carrier account")}</span>`).join("") || "<span>No carrier account assigned</span>"}</div>
          ${bids.filter((item) => ["invited", "submitted"].includes(item.status)).map(renderCarrierBidInvitation).join("")}
          ${tenders.filter((item) => item.status === "pending").map(renderCarrierTender).join("")}
          <div class="portal-load-list">${loads.length ? loads.map((load) => renderCarrierPortalLoad(load, documentsByLoad.get(load.id) || [], messagesByLoad.get(load.id) || [])).join("") : '<div class="empty-state compact-empty">No assigned loads yet.</div>'}</div>
        </article>
        <aside class="panel portal-invoices"><div class="panel-header"><div><h2>Statements</h2><p>Carrier settlements visible to this account.</p></div></div>
          <div class="portal-invoice-list">${settlements.length ? settlements.map(renderCarrierPortalStatement).join("") : '<div class="empty-state compact-empty">No statements yet.</div>'}</div>
        </aside>
      </section>`;
    document.querySelector("#reload-carrier-portal").addEventListener("click", renderCarrierPortal);
    root.querySelectorAll("[data-tender-response]").forEach((button) => button.addEventListener("click", async () => {
      button.disabled = true;
      try {
        await respondCarrierTender(button.dataset.tenderId, button.dataset.tenderResponse);
        await renderCarrierPortal();
      } catch (error) {
        window.alert(error.message);
        button.disabled = false;
      }
    }));
    root.querySelectorAll("[data-carrier-status-form]").forEach((form) => form.addEventListener("submit", handleCarrierPortalStatus));
    root.querySelectorAll("[data-carrier-bid-form]").forEach((form) => form.addEventListener("submit", handleCarrierBidSubmit));
    root.querySelectorAll("[data-carrier-message-form]").forEach((form) => form.addEventListener("submit", handleCarrierPortalMessage));
    root.querySelectorAll("[data-carrier-document-form]").forEach((form) => form.addEventListener("submit", handleCarrierPortalDocument));
  } catch (error) {
    document.querySelector("#connection-status").textContent = "Unavailable";
    root.innerHTML = `<article class="panel"><div class="empty-state">Carrier Portal is not ready: ${escapeHtml(error.message)}</div></article>`;
  }
}

function renderCarrierBidInvitation(bid) {
  const load = bid.loads || {};
  return `<article class="portal-load-card"><header><div><span class="eyebrow">Bid invitation</span><h2>${escapeHtml(load.load_no || "Load")}: ${escapeHtml([load.origin, load.destination].filter(Boolean).join(" → "))}</h2></div><span class="status-pill">${formatStatus(bid.status)}</span></header>
    <div class="portal-load-meta"><span><strong>Target</strong>${bid.target_amount == null ? "Open bid" : formatMoney(bid.target_amount)}</span><span><strong>Expires</strong>${formatDateTime(bid.expires_at)}</span></div>
    <form class="record-form compact-form" data-carrier-bid-form><input name="bid_id" type="hidden" value="${bid.id}"><label><span>Your amount</span><input name="amount" type="number" min="0.01" step="0.01" value="${bid.bid_amount || ""}" required></label><label><span>Transit days</span><input name="transit_days" type="number" min="1" step="1" value="${bid.transit_days || ""}"></label><label><span>Note</span><input name="note" value="${escapeAttribute(bid.carrier_note || "")}"></label><button type="submit">${bid.status === "submitted" ? "Update Bid" : "Submit Bid"}</button><p class="form-message"></p></form></article>`;
}

async function handleCarrierBidSubmit(event) {
  event.preventDefault(); const form=event.currentTarget; const message=form.querySelector(".form-message"); const values=Object.fromEntries(new FormData(form));
  try { await submitCarrierLoadBid(values.bid_id, numberOrZero(values.amount), numberOrNull(values.transit_days), values.note || null); await renderCarrierPortal(); }
  catch(error) { message.textContent=error.message; }
}

function renderCarrierTender(tender) {
  const load = tender.loads || {};
  return `<article class="portal-load-card"><header><div><span class="eyebrow">Tender</span><h2>${escapeHtml(load.load_no || "Load")}: ${escapeHtml([load.origin, load.destination].filter(Boolean).join(" → "))}</h2></div><span class="status-pill">Pending</span></header>
    <div class="portal-load-meta"><span><strong>Pickup</strong>${formatDate(load.pickup_date)}</span><span><strong>Delivery</strong>${formatDate(load.delivery_date)}</span><span><strong>Expires</strong>${formatDateTime(tender.expires_at)}</span></div>
    <div class="settlement-actions"><button type="button" class="primary-button" data-tender-id="${tender.id}" data-tender-response="accepted">Accept</button><button type="button" data-tender-id="${tender.id}" data-tender-response="rejected">Reject</button></div></article>`;
}

function renderCarrierPortalStatement(statement) {
  return `<article class="portal-statement-row">
    <div class="portal-statement-heading">
      <div><span>Settlement</span><strong>${escapeHtml(statement.settlement_no)}</strong></div>
      <span class="status-pill">${escapeHtml(formatStatus(statement.status))}</span>
    </div>
    <div class="portal-statement-meta">
      <div><span>Statement date</span><strong>${formatDate(statement.settlement_date)}</strong></div>
      <div class="portal-statement-total"><span>Total amount</span><strong>${formatMoney(statement.total_amount)}</strong></div>
    </div>
  </article>`;
}

function renderCarrierPortalLoad(load, documents = [], messages = []) {
  const statusOptions = driverStatusOptions(load);
  return `<article class="portal-load-card"><header><div><span class="eyebrow">${escapeHtml(load.load_no)}</span><h2>${escapeHtml(routeLabel(load))}</h2></div><span class="status-pill">${escapeHtml(formatStatus(operationalLoadStatus(load)))}</span></header>
    <div class="portal-load-meta"><span><strong>Pickup</strong>${formatDate(load.pickup_date)}</span><span><strong>Delivery</strong>${formatDate(load.delivery_date)}</span><span><strong>Equipment</strong>${escapeHtml(load.equipment_type || "-")}</span></div>
    <section class="portal-mini-section"><h3>Stops</h3><div class="portal-stop-list">${(load.load_stops || []).map((stop) => `<div><strong>${escapeHtml(formatStatus(stop.type))}</strong><span>${escapeHtml(stop.facility || stop.address || "Stop")}</span><small>${formatStopWindow(stop, load.pickup_date)}</small></div>`).join("") || '<div class="empty-state compact-empty">No stops visible.</div>'}</div></section>
    <section class="carrier-action-grid">
      <form class="carrier-action-card carrier-status-card" data-carrier-status-form>
        <input name="load_id" type="hidden" value="${load.id}">
        <header><div><span class="carrier-action-icon">↗</span><div><h3>Operational Status</h3><p>Share load progress with dispatch.</p></div></div></header>
        <label><span>Next status</span><select name="status" ${statusOptions.length ? "" : "disabled"}>${statusOptions.length ? statusOptions.map((status) => `<option value="${status}">${formatStatus(status)}</option>`).join("") : '<option>No action available</option>'}</select></label>
        ${statusOptions.length ? "" : `<small class="carrier-action-help">${escapeHtml(carrierStatusHelp(load))}</small>`}
        <button type="submit" ${statusOptions.length ? "" : "disabled"}>Update Status</button><span class="form-message"></span>
      </form>
      <form class="carrier-action-card carrier-upload-card" data-carrier-document-form>
        <input name="load_id" type="hidden" value="${load.id}">
        <header><div><span class="carrier-action-icon">↑</span><div><h3>Upload Document</h3><p>PDF, JPG or PNG · maximum 15 MB.</p></div></div></header>
        <div class="carrier-upload-fields"><label><span>Type</span><select name="doc_type"><option value="bol">BOL</option><option value="pod">POD</option><option value="photo">Photo</option></select></label><label class="carrier-file-field"><span>File</span><input name="file" type="file" accept="application/pdf,image/jpeg,image/png" required></label></div>
        <button type="submit">Upload for Review</button><span class="form-message"></span>
      </form>
    </section>
    <section class="carrier-documents-section"><div class="carrier-section-heading"><div><h3>Documents</h3><p>Carrier-facing files and review status.</p></div><span>${documents.length}</span></div><div class="carrier-document-list">${documents.map((document) => `<article><span class="carrier-doc-icon">${document.doc_type === "photo" ? "IMG" : "DOC"}</span><div><strong>${escapeHtml(document.file_name || "Document")}</strong><small>${escapeHtml(formatStatus(document.doc_type))} · uploaded ${formatDateTime(document.created_at)}</small></div><span class="status-pill">${escapeHtml(formatStatus(document.status))}</span></article>`).join("") || '<div class="empty-state compact-empty">No carrier-facing documents.</div>'}</div></section>
    <section class="carrier-messages-section"><div class="carrier-section-heading"><div><h3>Dispatch Messages</h3><p>Updates shared with the operations team.</p></div><span>${messages.length}</span></div><div class="driver-message-list">${renderDriverMessages(messages)}</div><form class="carrier-message-form" data-carrier-message-form><input name="load_id" type="hidden" value="${load.id}"><label><span>New message</span><textarea name="body" rows="3" maxlength="4000" placeholder="Write an update for dispatch..." required></textarea></label><div><small>Visible to dispatch for this load</small><button type="submit">Send Message</button></div><span class="form-message"></span></form></section></article>`;
}

async function handleCarrierPortalStatus(event) {
  event.preventDefault(); const form=event.currentTarget; const message=form.querySelector(".form-message"); const button=form.querySelector("button"); button.disabled=true;
  try { await carrierPortalUpdateLoadStatus(form.elements.load_id.value,form.elements.status.value); await renderCarrierPortal(); }
  catch(error){ message.textContent=error.message; button.disabled=false; }
}

async function handleCarrierPortalMessage(event) {
  event.preventDefault(); const form=event.currentTarget; const message=form.querySelector(".form-message"); const button=form.querySelector("button"); button.disabled=true;
  try { const body=validateMessageBody(form.elements.body.value); await createLoadMessage(form.elements.load_id.value,currentSession.user.id,body); await renderCarrierPortal(); }
  catch(error){ message.textContent=error.message; button.disabled=false; }
}

async function handleCarrierPortalDocument(event) {
  event.preventDefault(); const form=event.currentTarget; const message=form.querySelector(".form-message"); const button=form.querySelector("button"); button.disabled=true; message.textContent="Uploading...";
  try { await uploadTmsDocument({ entityType:"load",entityId:form.elements.load_id.value,docType:form.elements.doc_type.value,status:"pending",file:form.elements.file.files[0],uploadedBy:currentSession.user.id }); await renderCarrierPortal(); }
  catch(error){ message.textContent=error.message; button.disabled=false; }
}

function renderCustomerPortalLoad(load, documents = [], invoices = []) {
  const nextStop = nextOperationalStop(load);
  return `
    <article class="portal-load-card">
      <header>
        <div>
          <span class="eyebrow">${escapeHtml(load.load_no)}</span>
          <h2>${escapeHtml(routeLabel(load))}</h2>
        </div>
        <span class="status-pill">${escapeHtml(formatStatus(operationalLoadStatus(load)))}</span>
      </header>
      <div class="portal-load-meta">
        <span><strong>Pickup</strong>${formatDate(load.pickup_date)}</span>
        <span><strong>Delivery</strong>${formatDate(load.delivery_date)}</span>
        <span><strong>Next stop</strong>${escapeHtml(nextStop?.facility || nextStop?.address || "No pending stop")}</span>
        <span><strong>Equipment</strong>${escapeHtml([load.trucks?.unit_no, load.trailers?.unit_no].filter(Boolean).join(" / ") || "-")}</span>
      </div>
      ${load.customer_notes ? `<p class="portal-customer-note">${escapeHtml(load.customer_notes)}</p>` : ""}
      <section class="portal-mini-section">
        <h3>Stops</h3>
        <div class="portal-stop-list">
          ${(load.load_stops || []).length ? [...load.load_stops].sort((a, b) => a.stop_order - b.stop_order).map((stop) => `
            <div><strong>${escapeHtml(formatStatus(stop.type))}</strong><span>${escapeHtml(stop.facility || stop.address || "Stop")}</span><small>${formatStopWindow(stop, load.pickup_date)}</small></div>
          `).join("") : '<div class="empty-state compact-empty">No stops visible yet.</div>'}
        </div>
      </section>
      <section class="portal-mini-section">
        <h3>Documents</h3>
        <div class="portal-doc-list">
          ${documents.length ? documents.map((document) => `<a href="${escapeAttribute(document.signed_url || document.file_url || "#")}" target="_blank" rel="noopener">${escapeHtml(formatStatus(document.doc_type))}<small>${escapeHtml(document.file_name || document.status)}</small></a>`).join("") : '<div class="empty-state compact-empty">No approved customer documents yet.</div>'}
        </div>
      </section>
      ${invoices.length ? `<section class="portal-mini-section"><h3>Invoices</h3><div class="portal-doc-list">${invoices.map(renderCustomerPortalInvoice).join("")}</div></section>` : ""}
    </article>`;
}

function renderCustomerPortalInvoice(invoice) {
  const { balance, displayStatus } = invoiceFinancialState(invoice, localDateIso());
  const customerBalance = invoice.status === "paid" ? 0 : balance;
  const loadLabel = invoice.loads?.load_no ? `Load ${invoice.loads.load_no}` : "Shipment invoice";
  return `
    <article class="portal-invoice-row">
      <div class="portal-invoice-title"><strong>${escapeHtml(invoice.invoice_no)}</strong><span>${escapeHtml(loadLabel)}</span></div>
      <span class="status-pill">${escapeHtml(formatStatus(displayStatus))}</span>
      <small>${formatMoney(invoice.amount)} · balance ${formatMoney(customerBalance)} · due ${formatDate(invoice.due_date)}</small>
    </article>`;
}

function renderDashboard() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "M&M Safety operations", title: "Cyrra TMS Dashboard", status: "Checking Supabase" })}
    <section class="phoenix-layout">
      <div class="main-column">
        <section class="quick-stats" aria-label="Operational alerts">
          <article>
            <span class="stat-icon green">✓</span>
            <div><strong id="dash-active-loads">0 active loads</strong><small>In dispatch pipeline</small></div>
          </article>
          <article>
            <span class="stat-icon amber">!</span>
            <div><strong id="dash-unassigned">0 unassigned</strong><small>Need dispatch attention</small></div>
          </article>
          <article>
            <span class="stat-icon red">×</span>
            <div><strong id="dash-docs-review">0 document issues</strong><small>Pending, rejected or needs review</small></div>
          </article>
        </section>

        <article class="panel chart-panel">
          <div class="panel-header chart-header">
            <div>
              <h2>Load Status</h2>
              <p>Operational distribution across the dispatch pipeline</p>
            </div>
            <a class="panel-link" href="#dispatch" ${can("view_dispatch") ? "" : "hidden"}>Open Dispatch</a>
          </div>
          <div id="dashboard-status-bars" class="dashboard-status-bars">
            <div class="empty-state">Loading load status...</div>
          </div>
        </article>
      </div>

      <aside class="metric-grid">
        <a class="panel metric-card metric-link" href="#loads">
          <div><h3>Total orders</h3><small>All loads</small></div>
          <strong id="dash-total-loads">0</strong>
          <span class="trend-pill" id="dash-today-work">0 today</span>
        </a>
        <a class="panel metric-card metric-link" href="#tracking" ${can("view_tracking") ? "" : "hidden"}>
          <div><h3>Fleet availability</h3><small>Drivers / trucks / trailers</small></div>
          <strong id="dash-fleet-available">0</strong>
          <div id="dash-fleet-breakdown" class="mini-list"></div>
        </a>
        <a class="panel metric-card metric-link compliance-metric" href="#maintenance" ${can("view_maintenance") ? "" : "hidden"}>
          <div><h3>Safety compliance</h3><small>Expired and upcoming credentials</small></div>
          <strong id="dash-compliance-risk">0</strong>
          <div id="dash-compliance-breakdown" class="mini-list"></div>
        </a>
        <a class="panel metric-card metric-link" href="#billing" ${can("view_billing") ? "" : "hidden"}>
          <div><h3>Billing snapshot</h3><small>Invoices and ready-to-bill loads</small></div>
          <strong id="dash-unpaid">0</strong>
          <div id="dash-billing-breakdown" class="mini-list"></div>
        </a>
        <a class="panel metric-card metric-link" href="#documents">
          <div><h3>Documents</h3><small>Review queue</small></div>
          <strong id="dash-doc-count">0</strong>
          <div id="dash-doc-breakdown" class="mini-list"></div>
        </a>
      </aside>
    </section>

    <section class="dashboard-lower-grid">
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2>Action Queue</h2>
            <p>Items that need operational follow-up</p>
          </div>
          <button id="check-tables" type="button">Refresh</button>
        </div>
        <div id="dashboard-alerts" class="action-list">
          <div class="empty-state">Loading action queue...</div>
        </div>
      </article>

      <article class="panel" ${can("manage_users") ? "" : "hidden"}>
        <div class="panel-header">
          <div>
            <h2>Database Health</h2>
            <p>Supabase setup and required TMS tables</p>
          </div>
        </div>
        <div class="health-grid">
          <article>
            <span>TMS schema</span>
            <strong id="schema-count">0%</strong>
            <div id="tms-table-status" class="compact-status"></div>
          </article>
          <article>
            <span>Supabase setup</span>
            <strong id="setup-count">0%</strong>
            <div id="table-status" class="compact-status"></div>
          </article>
        </div>
      </article>
    </section>

    <section class="panel review-panel">
      <div class="panel-header">
        <div>
          <h2>Recent Activity</h2>
          <p>Latest loads, documents and invoices from Supabase</p>
        </div>
        <div class="table-tools">
          <a class="panel-link" href="#loads">Loads</a>
          <a class="panel-link" href="#documents">Documents</a>
          <a class="panel-link" href="#billing" ${can("view_billing") ? "" : "hidden"}>Billing</a>
        </div>
      </div>
      <div id="dashboard-recent" class="operations-table">
        <div class="empty-state">Loading recent activity...</div>
      </div>
    </section>
  `;

  document.querySelector("#check-tables").addEventListener("click", loadDashboardData);
  loadDashboardData();
}

async function loadDashboardData() {
  const statusEl = document.querySelector("#connection-status");
  const tableStatusEl = document.querySelector("#table-status");
  const tmsTableStatusEl = document.querySelector("#tms-table-status");
  const statusBarsEl = document.querySelector("#dashboard-status-bars");
  const alertsEl = document.querySelector("#dashboard-alerts");
  const recentEl = document.querySelector("#dashboard-recent");

  statusEl.textContent = "Checking Supabase";
  tableStatusEl.innerHTML = '<div class="empty-state">Checking database tables...</div>';
  tmsTableStatusEl.innerHTML = '<div class="empty-state">Checking TMS schema...</div>';
  statusBarsEl.innerHTML = '<div class="empty-state">Loading load status...</div>';
  alertsEl.innerHTML = '<div class="empty-state">Loading action queue...</div>';
  recentEl.innerHTML = '<div class="empty-state">Loading recent activity...</div>';

  try {
    const [results, tmsResults, drivers, trucks, trailers, carriers, loads, documents, invoices] = await Promise.all([
      checkExistingTables(),
      checkTmsTables(),
      safeDashboardList(listDrivers),
      safeDashboardList(listTrucks),
      safeDashboardList(listTrailers),
      safeDashboardList(listCarriers),
      safeDashboardList(listLoads),
      safeDashboardList(listDocuments),
      safeDashboardList(listInvoices),
    ]);
    const hasConnection = results.some((result) => result.ok);
    const tmsReady = tmsResults.every((result) => result.ok);
    const setupPercent = Math.round((results.filter((result) => result.ok).length / results.length) * 100);
    const tmsPercent = Math.round((tmsResults.filter((result) => result.ok).length / tmsResults.length) * 100);
    const activeLoads = loads.filter((load) => !["delivered", "cancelled"].includes(operationalLoadStatus(load)));
    const unassignedLoads = loads.filter((load) => !load.driver_id || !load.truck_id || !load.trailer_id);
    const today = localDateIso();
    const todayLoads = loads.filter((load) => load.pickup_date === today || load.delivery_date === today);
    const docIssues = documents.filter((document) => ["pending", "needs_review", "rejected"].includes(document.status));
    const statusCounts = countBy(loads.map((load) => ({ status: operationalLoadStatus(load) })), "status");
    const billingStatusCounts = countBy(loads.map((load) => ({ status: billingLoadStatus(load) })), "status");
    const availableDrivers = drivers.filter((driver) => driver.status === "available").length;
    const availableTrucks = trucks.filter((truck) => truck.status === "available").length;
    const availableTrailers = trailers.filter((trailer) => trailer.status === "available").length;
    const unpaidInvoices = invoices.filter((invoice) => !["paid", "void"].includes(invoice.status));
    const overdueInvoices = invoices.filter((invoice) => invoice.status === "overdue");
    const readyToBill = loads.filter((load) => billingLoadStatus(load) === "ready_to_bill").length;
    const compliance = buildComplianceItems({ drivers, carriers, trucks, trailers });
    const expiredCompliance = compliance.filter((item) => item.state === "expired");
    const dueSoonCompliance = compliance.filter((item) => item.state === "due_soon");
    const missingCompliance = compliance.filter((item) => item.state === "missing");

    tableStatusEl.innerHTML = renderRows(results);
    tmsTableStatusEl.innerHTML = renderRows(tmsResults);
    document.querySelector("#setup-count").textContent = `${setupPercent}%`;
    document.querySelector("#schema-count").textContent = `${tmsPercent}%`;
    document.querySelector("#dash-active-loads").textContent = `${activeLoads.length} active loads`;
    document.querySelector("#dash-unassigned").textContent = `${unassignedLoads.length} unassigned`;
    document.querySelector("#dash-docs-review").textContent = `${docIssues.length} document issues`;
    document.querySelector("#dash-total-loads").textContent = loads.length;
    document.querySelector("#dash-today-work").textContent = `${todayLoads.length} today`;
    document.querySelector("#dash-fleet-available").textContent = availableDrivers + availableTrucks + availableTrailers;
    document.querySelector("#dash-fleet-breakdown").innerHTML = `
      <span>Drivers <strong>${availableDrivers}/${drivers.length}</strong></span>
      <span>Trucks <strong>${availableTrucks}/${trucks.length}</strong></span>
      <span>Trailers <strong>${availableTrailers}/${trailers.length}</strong></span>
    `;
    const complianceRisk = document.querySelector("#dash-compliance-risk");
    complianceRisk.textContent = expiredCompliance.length + dueSoonCompliance.length;
    complianceRisk.dataset.state = expiredCompliance.length ? "error" : dueSoonCompliance.length ? "warning" : "ok";
    document.querySelector("#dash-compliance-breakdown").innerHTML = `
      <span>Expired <strong>${expiredCompliance.length}</strong></span>
      <span>Due soon <strong>${dueSoonCompliance.length}</strong></span>
      <span>Missing dates <strong>${missingCompliance.length}</strong></span>
    `;
    document.querySelector("#dash-unpaid").textContent = formatMoney(sumBy(unpaidInvoices, "amount"));
    document.querySelector("#dash-billing-breakdown").innerHTML = `
      <span>Ready to bill <strong>${readyToBill}</strong></span>
      <span>Overdue invoices <strong>${overdueInvoices.length}</strong></span>
      <span>Unpaid invoices <strong>${unpaidInvoices.length}</strong></span>
    `;
    document.querySelector("#dash-doc-count").textContent = documents.length;
    document.querySelector("#dash-doc-breakdown").innerHTML = `
      <span>Pending <strong>${documents.filter((document) => document.status === "pending").length}</strong></span>
      <span>Needs review <strong>${documents.filter((document) => document.status === "needs_review").length}</strong></span>
      <span>Rejected <strong>${documents.filter((document) => document.status === "rejected").length}</strong></span>
    `;
    statusBarsEl.innerHTML = dispatchStatuses
      .map((status) => renderStatusBar(status, statusCounts[status] || 0, loads.length))
      .join("");
    alertsEl.innerHTML = renderDashboardAlerts({ unassignedLoads, todayLoads, readyToBill, docIssues, overdueInvoices, expiredCompliance, dueSoonCompliance, missingCompliance });
    recentEl.innerHTML = renderDashboardRecent({ loads, documents, invoices });
    statusEl.dataset.state = hasConnection ? "ok" : "error";
    statusEl.textContent = tmsReady ? "Live TMS data" : hasConnection ? "Supabase connected" : "No readable tables";
  } catch (error) {
    tableStatusEl.innerHTML = `<div class="empty-state">Supabase check failed: ${escapeHtml(error.message)}</div>`;
    tmsTableStatusEl.innerHTML = '<div class="empty-state">TMS schema check skipped.</div>';
    statusBarsEl.innerHTML = '<div class="empty-state">Dashboard data unavailable.</div>';
    alertsEl.innerHTML = '<div class="empty-state">Action queue unavailable.</div>';
    recentEl.innerHTML = '<div class="empty-state">Recent activity unavailable.</div>';
    statusEl.dataset.state = "error";
    statusEl.textContent = "Connection failed";
  }
}

async function safeDashboardList(loader) {
  try {
    return await loader();
  } catch (_error) {
    return [];
  }
}

async function refreshNotificationBadge() {
  const navCount = document.querySelector("#nav-notification-count");
  const topbarCount = document.querySelector("#topbar-notification-count");
  if (!navCount || !topbarCount || !currentSession || !can("view_notifications")) return;
  try {
    const notifications = await listNotifications(50);
    const summary = notificationSummary(notifications);
    [navCount, topbarCount].forEach((badge) => {
      badge.textContent = String(summary.unread);
      badge.hidden = summary.unread === 0;
    });
  } catch (_error) {
    [navCount, topbarCount].forEach((badge) => { badge.hidden = true; });
  }
}

function renderDashboardAlerts({ unassignedLoads, todayLoads, readyToBill, docIssues, overdueInvoices, expiredCompliance, dueSoonCompliance, missingCompliance }) {
  const alerts = [
    { label: "Loads missing assignment", count: unassignedLoads.length, href: "#loads", tone: unassignedLoads.length ? "warn" : "ok" },
    { label: "Pickup or delivery today", count: todayLoads.length, href: "#loads", tone: todayLoads.length ? "info" : "ok" },
    can("view_billing") ? { label: "Ready to bill", count: readyToBill, href: "#billing", tone: readyToBill ? "info" : "ok" } : null,
    { label: "Documents needing attention", count: docIssues.length, href: "#documents", tone: docIssues.length ? "warn" : "ok" },
    can("view_maintenance") ? { label: "Expired safety compliance", count: expiredCompliance.length, href: "#maintenance", tone: expiredCompliance.length ? "danger" : "ok" } : null,
    can("view_maintenance") ? { label: "Compliance due within 30 days", count: dueSoonCompliance.length, href: "#maintenance", tone: dueSoonCompliance.length ? "warn" : "ok" } : null,
    can("view_maintenance") ? { label: "Missing compliance dates", count: missingCompliance.length, href: "#maintenance", tone: missingCompliance.length ? "warn" : "ok" } : null,
    can("view_billing") ? { label: "Overdue invoices", count: overdueInvoices.length, href: "#billing", tone: overdueInvoices.length ? "danger" : "ok" } : null,
  ].filter(Boolean);

  return alerts
    .map(
      (alert) => `
        <a class="action-row" data-tone="${alert.tone}" href="${alert.href}">
          <span>${alert.label}</span>
          <strong>${alert.count}</strong>
        </a>
      `
    )
    .join("");
}

function renderDashboardRecent({ loads, documents, invoices }) {
  const recent = [
    ...loads.slice(0, 5).map((load) => ({
      module: `Load ${load.load_no}`,
      owner: load.customers?.name || "No customer",
      status: formatStatus(operationalLoadStatus(load)),
      value: formatDate(load.pickup_date),
      href: "#loads",
      created_at: load.created_at,
    })),
    ...documents.slice(0, 4).map((document) => ({
      module: formatStatus(document.doc_type),
      owner: formatStatus(document.entity_type),
      status: formatStatus(document.status),
      value: formatDateTime(document.created_at),
      href: "#documents",
      created_at: document.created_at,
    })),
    ...(can("view_billing") ? invoices.slice(0, 4) : []).map((invoice) => ({
      module: `Invoice ${invoice.invoice_no}`,
      owner: invoice.customers?.name || "-",
      status: formatStatus(invoice.status),
      value: formatMoney(invoice.amount),
      href: "#billing",
      created_at: invoice.created_at,
    })),
  ]
    .sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0))
    .slice(0, 8);

  if (!recent.length) {
    return '<div class="empty-state">No activity yet. Add customers, equipment and the first load to populate this panel.</div>';
  }

  return `
    <div class="operation-row operation-head">
      <span>Item</span><span>Owner</span><span>Status</span><span>Value</span>
    </div>
    ${recent
      .map(
        (item) => `
          <a class="operation-row operation-link" href="${item.href}">
            <strong>${escapeHtml(item.module)}</strong>
            <span>${escapeHtml(item.owner)}</span>
            <em>${escapeHtml(item.status)}</em>
            <span>${escapeHtml(item.value)}</span>
          </a>
        `
      )
      .join("")}
  `;
}

function localDateIso(offsetDays = 0) {
  const date = new Date();
  date.setDate(date.getDate() + offsetDays);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

function modalizeCrudForm(formSelector, addLabel, resetForm, triggerPanelSelector = "") {
  const form = document.querySelector(formSelector);
  const panel = form?.closest("article.panel");
  const grid = panel?.parentElement;
  const listPanel = triggerPanelSelector ? document.querySelector(triggerPanelSelector) : grid?.querySelector("article.panel:first-child");
  if (!form || !panel || !grid || !listPanel) return;

  const dialog = document.createElement("dialog");
  dialog.className = "crud-dialog";
  const closeButton = document.createElement("button");
  closeButton.type = "button";
  closeButton.className = "dialog-close";
  closeButton.setAttribute("aria-label", "Close");
  closeButton.textContent = "×";
  closeButton.addEventListener("click", () => {
    resetForm();
    if (dialog.open) dialog.close();
  });
  panel.querySelector(".panel-header")?.append(closeButton);
  dialog.append(panel);
  grid.append(dialog);
  grid.classList.add("single-column-grid");

  const addButton = document.createElement("button");
  addButton.type = "button";
  addButton.className = "primary-button";
  addButton.textContent = addLabel;
  addButton.addEventListener("click", () => {
    resetForm();
    dialog.showModal();
  });
  listPanel.querySelector(".panel-header")?.append(addButton);
  dialog.addEventListener("click", (event) => {
    if (event.target === dialog) {
      resetForm();
      if (dialog.open) dialog.close();
    }
  });
}

function configureCrudFormAccess(formSelector, permission, addLabel, resetForm, triggerPanelSelector = "") {
  const form = document.querySelector(formSelector);
  const panel = form?.closest("article.panel");
  if (can(permission)) {
    modalizeCrudForm(formSelector, addLabel, resetForm, triggerPanelSelector);
    return true;
  }
  panel?.setAttribute("hidden", "");
  panel?.parentElement?.classList.add("single-column-grid");
  return false;
}

function openFormDialog(form) {
  const dialog = form.closest("dialog");
  if (dialog && !dialog.open) dialog.showModal();
}

function closeFormDialog(form) {
  const dialog = form.closest("dialog");
  if (dialog?.open) dialog.close();
}

function setupPageTabs(panels, labels, activeLabel = labels[0], onChange = null) {
  panels = panels.filter(Boolean);
  if (!panels.length) return;
  const activeIndex = Math.max(0, labels.indexOf(activeLabel));
  let previous = panels[0].previousElementSibling;
  while (previous?.classList.contains("page-tabs")) {
    const duplicate = previous;
    previous = previous.previousElementSibling;
    duplicate.remove();
  }
  const tabs = document.createElement("div");
  tabs.className = "page-tabs";
  panels[0].before(tabs);
  panels.forEach((panel, index) => {
    panel.hidden = index !== activeIndex;
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = labels[index];
    button.dataset.active = index === activeIndex ? "true" : "false";
    button.addEventListener("click", () => {
      panels.forEach((item, panelIndex) => { item.hidden = panelIndex !== index; });
      tabs.querySelectorAll("button").forEach((item, buttonIndex) => { item.dataset.active = buttonIndex === index ? "true" : "false"; });
      onChange?.(labels[index]);
    });
    tabs.append(button);
  });
}

let customerDirectoryData = [];
let facilityDirectoryData = [];

function renderBusinessProfileDialog() {
  return `<dialog id="business-profile-dialog" class="crud-dialog business-profile-dialog">
    <article class="panel">
      <div class="panel-header"><div><h2 id="business-profile-title">Profile</h2><p id="business-profile-subtitle"></p></div><button id="close-business-profile" class="dialog-close" type="button" aria-label="Close">×</button></div>
      <div id="business-profile-content" class="business-profile-content"><div class="empty-state">Loading profile...</div></div>
    </article>
  </dialog>`;
}

function renderCustomers() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Phase 1", title: "Customers" })}
    ${renderBusinessProfileDialog()}
    <section class="content-grid customer-grid">
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2>Customer List</h2>
            <p>Customers used by loads, invoices and the future portal.</p>
          </div>
          <button id="reload-customers" type="button">Refresh</button>
        </div>
        <div id="customers-table" class="data-table"></div>
      </article>
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2 id="customer-form-title">Add Customer</h2>
            <p>Creates a record in the TMS customers table.</p>
          </div>
        </div>
        <form id="customer-form" class="record-form">
          <input name="id" type="hidden" />
          <label>
            <span>Name</span>
            <input name="name" type="text" required />
          </label>
          <label>
            <span>Billing address</span>
            <input name="billing_address" type="text" />
          </label>
          <label><span>Billing email</span><input name="billing_email" type="email" /></label>
          <label><span>Billing phone</span><input name="billing_phone" type="tel" /></label>
          <label><span>Payment terms (days)</span><input name="payment_terms_days" type="number" min="0" max="365" value="30" /></label>
          <label><span>Billing cycle</span><select name="billing_cycle"><option value="per_load">Per load</option><option value="weekly">Weekly</option><option value="biweekly">Biweekly</option><option value="monthly">Monthly</option></select></label>
          <label><span>Cycle day</span><input name="billing_cycle_day" type="number" min="1" max="31" placeholder="Optional day of week/month" /></label>
          <label><span>Notes</span><textarea name="notes" rows="3"></textarea></label>
          <label>
            <span>Credit status</span>
            <select name="credit_status">
              <option value="ok">OK</option>
              <option value="watch">Watch</option>
              <option value="hold">Hold</option>
            </select>
          </label>
          <label class="checkbox-row">
            <input name="portal_enabled" type="checkbox" />
            <span>Portal enabled</span>
          </label>
          <div class="form-actions">
            <button id="customer-submit" type="submit">Save Customer</button>
            <button id="cancel-customer-edit" type="button" hidden>Cancel edit</button>
          </div>
          <p id="customer-message" class="form-message"></p>
        </form>
      </article>
    </section>
    <section class="content-grid customer-grid">
      <article class="panel">
        <div class="panel-header"><div><h2>Facility Address Book</h2><p>Reusable shipper and consignee locations for load stops.</p></div><button id="reload-facilities" type="button">Refresh</button></div>
        <div id="facilities-table" class="data-table"></div>
      </article>
      <article class="panel">
        <div class="panel-header"><div><h2 id="facility-form-title">Add Facility</h2><p>Creates a reusable pickup or delivery location.</p></div></div>
        <form id="facility-form" class="record-form">
          <input name="id" type="hidden" />
          <label><span>Name</span><input name="name" required /></label>
          <label><span>Type</span><select name="facility_type"><option value="both">Shipper & consignee</option><option value="shipper">Shipper</option><option value="consignee">Consignee</option><option value="other">Other</option></select></label>
          <label><span>Customer (optional)</span><select name="customer_id"><option value="">No customer</option></select></label>
          <label><span>Address line 1</span><input name="address_line1" required /></label>
          <label><span>Address line 2</span><input name="address_line2" /></label>
          <label><span>City</span><input name="city" required /></label>
          <label><span>State</span><input name="state" required maxlength="50" /></label>
          <label><span>ZIP / postal code</span><input name="postal_code" /></label>
          <label><span>Country</span><input name="country" value="US" required /></label>
          <label><span>Contact name</span><input name="contact_name" /></label>
          <label><span>Contact phone</span><input name="contact_phone" type="tel" /></label>
          <label><span>Contact email</span><input name="contact_email" type="email" /></label>
          <label><span>Hours</span><input name="hours" placeholder="Mon–Fri 08:00–17:00" /></label>
          <label><span>Instructions</span><textarea name="instructions" rows="3"></textarea></label>
          <label><span>Status</span><select name="status"><option value="active">Active</option><option value="inactive">Inactive</option></select></label>
          <div class="form-actions"><button id="facility-submit" type="submit">Save Facility</button><button id="cancel-facility-edit" type="button" hidden>Cancel edit</button></div>
          <p id="facility-message" class="form-message"></p>
        </form>
      </article>
    </section>
  `;

  document.querySelector("#reload-customers").addEventListener("click", loadCustomers);
  document.querySelector("#customer-form").addEventListener("submit", handleCreateCustomer);
  document.querySelector("#cancel-customer-edit").addEventListener("click", resetCustomerForm);
  document.querySelector("#reload-facilities").addEventListener("click", loadFacilities);
  document.querySelector("#facility-form").addEventListener("submit", handleFacilitySave);
  document.querySelector("#cancel-facility-edit").addEventListener("click", resetFacilityForm);
  document.querySelector("#close-business-profile").addEventListener("click", () => document.querySelector("#business-profile-dialog").close());
  configureCrudFormAccess("#customer-form", "manage_operations", "Add Customer", resetCustomerForm);
  configureCrudFormAccess("#facility-form", "manage_operations", "Add Facility", resetFacilityForm);
  setupPageTabs(Array.from(pageRoot.querySelectorAll(":scope > section.content-grid")), ["Customers", "Facilities"]);
  loadCustomers();
  loadFacilities();
}

async function loadCustomers() {
  const table = document.querySelector("#customers-table");
  table.innerHTML = '<div class="empty-state">Loading customers...</div>';

  try {
    const customers = await listCustomers();
    customerDirectoryData = customers;
    populateFacilityCustomerOptions();

    if (!customers.length) {
      table.innerHTML = '<div class="empty-state">No customers yet.</div>';
      return;
    }

    table.innerHTML = `
      <div class="data-row customer-row data-row-head">
        <span>Name</span>
        <span>Credit</span>
        <span>Portal</span>
        <span>Status / Actions</span>
        <span></span>
      </div>
      ${customers
        .map(
          (customer) => `
            <div class="data-row customer-row">
              <strong>${escapeHtml(customer.name)}</strong>
              <span>${customer.credit_status}</span>
              <span>${customer.portal_enabled ? "Enabled" : "Off"}</span>
              <span>${customer.status}</span>
              <div class="table-tools"><button class="row-action" type="button" data-open-profile="customer" data-profile-id="${customer.id}">Open</button>${can("manage_operations") ? `<button class="row-action" type="button" data-edit-customer="${customer.id}">Edit</button>` : ""}</div>
            </div>
          `
        )
        .join("")}
    `;

    table.querySelectorAll("[data-edit-customer]").forEach((button) => {
      const customer = customers.find((item) => item.id === button.dataset.editCustomer);
      button.addEventListener("click", () => fillCustomerForm(customer));
    });
    bindBusinessProfileButtons(table, customers);
  } catch (error) {
    table.innerHTML = `
      <div class="empty-state">
        Customers table is not ready: ${escapeHtml(error.message)}
      </div>
    `;
  }
}

function fillCustomerForm(customer) {
  const form = document.querySelector("#customer-form");
  form.elements.id.value = customer.id;
  form.elements.name.value = customer.name || "";
  form.elements.billing_address.value = customer.billing_address || "";
  form.elements.billing_email.value = customer.billing_email || "";
  form.elements.billing_phone.value = customer.billing_phone || "";
  form.elements.payment_terms_days.value = customer.payment_terms_days ?? 30;
  form.elements.billing_cycle.value = customer.billing_cycle || "per_load";
  form.elements.billing_cycle_day.value = customer.billing_cycle_day ?? "";
  form.elements.notes.value = customer.notes || "";
  form.elements.credit_status.value = customer.credit_status || "ok";
  form.elements.portal_enabled.checked = Boolean(customer.portal_enabled);
  document.querySelector("#customer-form-title").textContent = "Edit Customer";
  document.querySelector("#customer-submit").textContent = "Update Customer";
  document.querySelector("#cancel-customer-edit").hidden = false;
  document.querySelector("#customer-message").textContent = "";
  openFormDialog(form);
}

function resetCustomerForm() {
  const form = document.querySelector("#customer-form");
  form.reset();
  form.elements.id.value = "";
  document.querySelector("#customer-form-title").textContent = "Add Customer";
  document.querySelector("#customer-submit").textContent = "Save Customer";
  document.querySelector("#cancel-customer-edit").hidden = true;
  document.querySelector("#customer-message").textContent = "";
  closeFormDialog(form);
}

async function handleCreateCustomer(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const message = document.querySelector("#customer-message");
  const formData = new FormData(form);
  const customerId = formData.get("id");

  message.textContent = customerId ? "Updating customer..." : "Saving customer...";

  try {
    const payload = {
      name: formData.get("name"),
      billing_address: formData.get("billing_address") || null,
      billing_email: formData.get("billing_email") || null,
      billing_phone: formData.get("billing_phone") || null,
      payment_terms_days: numberOrZero(formData.get("payment_terms_days")),
      ...validateCustomerBillingCycle({
        billingCycle: formData.get("billing_cycle"),
        billingCycleDay: formData.get("billing_cycle_day"),
      }),
      notes: formData.get("notes") || null,
      credit_status: formData.get("credit_status"),
      portal_enabled: formData.get("portal_enabled") === "on",
    };
    validateProfileFinancialTerms({ paymentTermsDays: payload.payment_terms_days });

    if (customerId) {
      await updateCustomer(customerId, payload);
    } else {
      await createCustomer(payload);
    }

    resetCustomerForm();
    message.textContent = customerId ? "Customer updated." : "Customer saved.";
    await loadCustomers();
  } catch (error) {
    message.textContent = error.message;
  }
}

function populateFacilityCustomerOptions() {
  const select = document.querySelector('#facility-form [name="customer_id"]');
  if (!select) return;
  const selected = select.value;
  select.innerHTML = `<option value="">No customer</option>${customerDirectoryData.map((customer) => `<option value="${customer.id}">${escapeHtml(customer.name)}</option>`).join("")}`;
  select.value = selected;
}

async function loadFacilities() {
  const table = document.querySelector("#facilities-table");
  if (!table) return;
  table.innerHTML = '<div class="empty-state">Loading facilities...</div>';
  try {
    facilityDirectoryData = await listFacilities();
    table.innerHTML = facilityDirectoryData.length ? `
      <div class="data-row facility-row data-row-head"><span>Name</span><span>Type</span><span>Location</span><span>Customer</span><span></span></div>
      ${facilityDirectoryData.map((facility) => `<div class="data-row facility-row">
        <strong>${escapeHtml(facility.name)}</strong>
        <span>${formatStatus(facility.facility_type)}</span>
        <span>${escapeHtml(`${facility.city}, ${facility.state}`)}</span>
        <span>${escapeHtml(facility.customers?.name || "-")}</span>
        ${can("manage_operations") ? `<button class="row-action" type="button" data-edit-facility="${facility.id}">Edit</button>` : ""}
      </div>`).join("")}` : '<div class="empty-state">No facilities yet. Add the first reusable shipper or consignee.</div>';
    table.querySelectorAll("[data-edit-facility]").forEach((button) => button.addEventListener("click", () => {
      const facility = facilityDirectoryData.find((item) => item.id === button.dataset.editFacility);
      if (facility) fillFacilityForm(facility);
    }));
  } catch (error) {
    table.innerHTML = `<div class="empty-state">Facilities are not ready: ${escapeHtml(error.message)}</div>`;
  }
}

function fillFacilityForm(facility) {
  const form = document.querySelector("#facility-form");
  for (const field of ["id", "name", "facility_type", "customer_id", "address_line1", "address_line2", "city", "state", "postal_code", "country", "contact_name", "contact_phone", "contact_email", "hours", "instructions", "status"]) {
    form.elements[field].value = facility[field] || (field === "country" ? "US" : "");
  }
  document.querySelector("#facility-form-title").textContent = "Edit Facility";
  document.querySelector("#facility-submit").textContent = "Update Facility";
  document.querySelector("#cancel-facility-edit").hidden = false;
  document.querySelector("#facility-message").textContent = "";
  openFormDialog(form);
}

function resetFacilityForm() {
  const form = document.querySelector("#facility-form");
  form.reset();
  form.elements.id.value = "";
  form.elements.country.value = "US";
  document.querySelector("#facility-form-title").textContent = "Add Facility";
  document.querySelector("#facility-submit").textContent = "Save Facility";
  document.querySelector("#cancel-facility-edit").hidden = true;
  document.querySelector("#facility-message").textContent = "";
  closeFormDialog(form);
}

async function handleFacilitySave(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const data = new FormData(form);
  const id = data.get("id");
  const message = document.querySelector("#facility-message");
  message.textContent = id ? "Updating facility..." : "Saving facility...";
  try {
    const payload = validateFacilityInput({
      name: data.get("name").trim(), facility_type: data.get("facility_type"), customer_id: data.get("customer_id") || null,
      address_line1: data.get("address_line1").trim(), address_line2: data.get("address_line2") || null,
      city: data.get("city").trim(), state: data.get("state").trim(), postal_code: data.get("postal_code") || null,
      country: data.get("country").trim() || "US", contact_name: data.get("contact_name") || null,
      contact_phone: data.get("contact_phone") || null, contact_email: data.get("contact_email") || null,
      hours: data.get("hours") || null, instructions: data.get("instructions") || null, status: data.get("status"),
    });
    if (id) await updateFacility(id, payload); else await createFacility(payload);
    resetFacilityForm();
    await loadFacilities();
  } catch (error) {
    message.textContent = error.message;
  }
}

let activeBusinessProfile = null;

function bindBusinessProfileButtons(root, records) {
  root.querySelectorAll("[data-open-profile]").forEach((button) => button.addEventListener("click", () => {
    const record = records.find((item) => item.id === button.dataset.profileId);
    if (record) openBusinessProfile(button.dataset.openProfile, record);
  }));
}

function businessProfileField(label, value) {
  return `<div><span>${escapeHtml(label)}</span><strong>${escapeHtml(value === null || value === undefined || value === "" ? "-" : String(value))}</strong></div>`;
}

function businessProfileSummary(type, record) {
  if (type === "driver") return [
    ["Status", formatStatus(record.status)], ["Phone", record.phone], ["Email", record.email], ["Address", record.address],
    ["Hire date", formatDate(record.hire_date)], ["CDL", record.cdl_no], ["CDL expiry", formatDate(record.cdl_expiry)],
    ["Medical expiry", formatDate(record.medical_expiry)], ["Emergency contact", record.emergency_contact_name],
    ["Emergency phone", record.emergency_contact_phone], ["Relation", record.emergency_contact_relation],
    ["Pay setup", record.pay_type ? `${formatStatus(record.pay_type)} · ${formatMoney(record.pay_rate)}` : "Not set"],
  ];
  if (type === "customer") return [
    ["Status", formatStatus(record.status)], ["Credit", formatStatus(record.credit_status)], ["Billing address", record.billing_address],
    ["Billing email", record.billing_email], ["Billing phone", record.billing_phone], ["Payment terms", `${record.payment_terms_days ?? 30} days`],
    ["Billing cycle", `${formatStatus(record.billing_cycle || "per_load")}${record.billing_cycle_day ? ` · day ${record.billing_cycle_day}` : ""}`],
    ["Portal", record.portal_enabled ? "Enabled" : "Off"], ["Notes", record.notes],
  ];
  if (type === "carrier") return [
    ["Status", formatStatus(record.status)], ["MC", record.mc], ["DOT", record.dot], ["Safety rating", record.safety_rating],
    ["Insurance expiry", formatDate(record.insurance_expiry)], ["Contact", record.contact_name], ["Email", record.contact_email],
    ["Phone", record.contact_phone], ["W-9", formatStatus(record.w9_status)], ["Payment terms", `${record.payment_terms_days ?? 30} days`],
    ["Pay setup", record.pay_type ? `${formatStatus(record.pay_type)} · ${formatMoney(record.pay_rate)}` : "Not set"], ["Notes", record.notes],
  ];
  if (type === "truck") return [
    ["Status", formatStatus(record.status)], ["VIN", record.vin], ["Plate", record.plate], ["Make / Model", [record.make, record.model].filter(Boolean).join(" ")],
    ["Year", record.year], ["Odometer", formatNumber(record.odometer)], ["Registration", formatDate(record.registration_expiry)], ["Insurance", formatDate(record.insurance_expiry)],
  ];
  return [["Status", formatStatus(record.status)], ["VIN", record.vin], ["Plate", record.plate], ["Type", record.type], ["Registration", formatDate(record.registration_expiry)], ["Insurance", formatDate(record.insurance_expiry)]];
}

function settledValue(result) {
  return result.status === "fulfilled" ? result.value : [];
}

async function openBusinessProfile(type, record) {
  activeBusinessProfile = { type, record };
  const dialog = document.querySelector("#business-profile-dialog");
  const content = document.querySelector("#business-profile-content");
  const label = type === "truck" ? `Truck ${record.unit_no}` : type === "trailer" ? `Trailer ${record.unit_no}` : record.name;
  document.querySelector("#business-profile-title").textContent = label;
  document.querySelector("#business-profile-subtitle").textContent = `${formatStatus(type)} profile and related operational records`;
  content.innerHTML = '<div class="empty-state">Loading full profile...</div>';
  if (!dialog.open) dialog.showModal();

  const [loadsResult, documentsResult, maintenanceResult, invoicesResult, settlementsResult, contactsResult, facilitiesResult] = await Promise.allSettled([
    listLoads(), listDocuments(), listMaintenance(), listInvoices(), listSettlements(),
    type === "customer" ? listCustomerContacts(record.id) : Promise.resolve([]),
    type === "customer" ? listFacilities() : Promise.resolve([]),
  ]);
  const loads = settledValue(loadsResult).filter((load) => load[`${type}_id`] === record.id || (type === "truck" && load.truck_id === record.id) || (type === "trailer" && load.trailer_id === record.id));
  const documents = settledValue(documentsResult).filter((document) => document.entity_type === type && document.entity_id === record.id);
  const maintenance = settledValue(maintenanceResult).filter((item) => item.equipment_type === type && item.equipment_id === record.id);
  const invoices = settledValue(invoicesResult).filter((invoice) => type === "customer" && invoice.customer_id === record.id);
  const settlements = settledValue(settlementsResult).filter((item) => (type === "driver" && item.driver_id === record.id) || (type === "carrier" && item.carrier_id === record.id));
  const contacts = settledValue(contactsResult);
  const facilities = settledValue(facilitiesResult).filter((facility) => facility.customer_id === record.id);

  content.innerHTML = `
    <div class="business-profile-summary">${businessProfileSummary(type, record).map(([name, value]) => businessProfileField(name, value)).join("")}</div>
    ${type === "customer" ? renderCustomerProfileContacts(contacts) : ""}
    ${renderBusinessProfileSection("Recent Loads", loads.map((load) => `<div><strong>Load ${escapeHtml(load.load_no)}</strong><span>${escapeHtml(routeLabel(load))} · ${formatStatus(operationalLoadStatus(load))}</span></div>`), "No related loads.")}
    ${type === "customer" ? renderBusinessProfileSection("Facilities", facilities.map((facility) => `<div><strong>${escapeHtml(facility.name)}</strong><span>${escapeHtml(`${facility.address_line1}, ${facility.city}, ${facility.state}`)}</span></div>`), "No linked facilities.") : ""}
    ${renderBusinessProfileSection("Documents", documents.map((document) => `<div><strong>${formatStatus(document.doc_type)}</strong><span>${formatStatus(document.status)}${document.signed_url ? ` · <a href="${escapeHtml(document.signed_url)}" target="_blank" rel="noopener">Open</a>` : ""}</span></div>`), "No documents.")}
    ${["truck", "trailer"].includes(type) ? renderBusinessProfileSection("Maintenance History", maintenance.map((item) => `<div><strong>${escapeHtml(item.service_type)}</strong><span>${formatStatus(item.status)} · ${formatDate(item.completed_date || item.scheduled_date || item.due_date)} · ${formatMoney(item.cost)}</span></div>`), "No maintenance records.") : ""}
    ${type === "customer" ? renderBusinessProfileSection("Invoices", invoices.map((invoice) => `<div><strong>${escapeHtml(invoice.invoice_no)}</strong><span>${formatStatus(invoice.status)} · ${formatMoney(invoice.amount)}</span></div>`), "No invoices.") : ""}
    ${["driver", "carrier"].includes(type) ? renderBusinessProfileSection("Settlements", settlements.map((item) => `<div><strong>${escapeHtml(item.settlement_no || "Draft settlement")}</strong><span>${formatStatus(item.status)} · ${formatMoney(item.total_amount)}</span></div>`), "No settlements.") : ""}
  `;
  document.querySelector("#customer-contact-form")?.addEventListener("submit", handleCustomerContactSave);
}

function renderBusinessProfileSection(title, rows, emptyText) {
  return `<section class="business-profile-section"><h3>${title}</h3><div>${rows.length ? rows.join("") : `<span class="empty-state compact-empty">${emptyText}</span>`}</div></section>`;
}

function renderCustomerProfileContacts(contacts) {
  const rows = contacts.map((contact) => `<div><strong>${escapeHtml(contact.name)}</strong><span>${escapeHtml([contact.role, contact.email, contact.phone].filter(Boolean).join(" · ") || "Contact")}</span></div>`);
  const form = can("manage_operations") ? `<form id="customer-contact-form" class="compact-profile-form"><input name="name" placeholder="Contact name" required><input name="role" placeholder="Role"><input name="email" type="email" placeholder="Email"><input name="phone" type="tel" placeholder="Phone"><button type="submit">Add Contact</button><span id="customer-contact-message"></span></form>` : "";
  return `<section class="business-profile-section"><h3>Contacts</h3><div>${rows.length ? rows.join("") : '<span class="empty-state compact-empty">No contacts.</span>'}</div>${form}</section>`;
}

async function handleCustomerContactSave(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const data = new FormData(form);
  const message = document.querySelector("#customer-contact-message");
  try {
    await createCustomerContact({ customer_id: activeBusinessProfile.record.id, name: data.get("name").trim(), role: data.get("role") || null, email: data.get("email") || null, phone: data.get("phone") || null });
    await openBusinessProfile("customer", activeBusinessProfile.record);
  } catch (error) {
    message.textContent = error.message;
  }
}

function renderCarriers() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Phase 1", title: "Carriers" })}
    ${renderBusinessProfileDialog()}
    <section class="content-grid customer-grid">
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2>Carrier List</h2>
            <p>Outside carriers, brokered capacity and compliance visibility.</p>
          </div>
          <button id="reload-carriers" type="button">Refresh</button>
        </div>
        <div id="carriers-table" class="data-table"></div>
      </article>
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2 id="carrier-form-title">Add Carrier</h2>
            <p>Creates a record in the TMS carriers table.</p>
          </div>
        </div>
        <form id="carrier-form" class="record-form">
          <input name="id" type="hidden" />
          <label>
            <span>Name</span>
            <input name="name" type="text" required />
          </label>
          <label>
            <span>MC number</span>
            <input name="mc" type="text" />
          </label>
          <label>
            <span>DOT number</span>
            <input name="dot" type="text" />
          </label>
          <label>
            <span>Safety rating</span>
            <input name="safety_rating" type="text" placeholder="Satisfactory" />
          </label>
          <label>
            <span>Insurance expiry</span>
            <input name="insurance_expiry" type="date" />
          </label>
          <label><span>Contact name</span><input name="contact_name" /></label>
          <label><span>Contact email</span><input name="contact_email" type="email" /></label>
          <label><span>Contact phone</span><input name="contact_phone" type="tel" /></label>
          <label><span>Payment terms (days)</span><input name="payment_terms_days" type="number" min="0" max="365" value="30" /></label>
          <label><span>W-9 status</span><select name="w9_status"><option value="missing">Missing</option><option value="pending">Pending</option><option value="approved">Approved</option><option value="rejected">Rejected</option></select></label>
          <label><span>Pay type</span><select name="pay_type"><option value="">Not set</option><option value="flat">Flat</option><option value="per_mile">Per mile</option><option value="percentage">Percentage</option></select></label>
          <label><span>Pay rate</span><input name="pay_rate" type="number" min="0" step="0.01" /></label>
          <label><span>Notes</span><textarea name="notes" rows="3"></textarea></label>
          <label>
            <span>Status</span>
            <select name="status">
              <option value="active">Active</option>
              <option value="inactive">Inactive</option>
              <option value="blocked">Blocked</option>
            </select>
          </label>
          <div class="form-actions">
            <button id="carrier-submit" type="submit">Save Carrier</button>
            <button id="cancel-carrier-edit" type="button" hidden>Cancel edit</button>
          </div>
          <p id="carrier-message" class="form-message"></p>
        </form>
      </article>
    </section>
  `;

  document.querySelector("#reload-carriers").addEventListener("click", loadCarriers);
  document.querySelector("#carrier-form").addEventListener("submit", handleCreateCarrier);
  document.querySelector("#cancel-carrier-edit").addEventListener("click", resetCarrierForm);
  document.querySelector("#close-business-profile").addEventListener("click", () => document.querySelector("#business-profile-dialog").close());
  configureCrudFormAccess("#carrier-form", "manage_carriers", "Add Carrier", resetCarrierForm);
  loadCarriers();
}

async function loadCarriers() {
  const table = document.querySelector("#carriers-table");
  table.innerHTML = '<div class="empty-state">Loading carriers...</div>';

  try {
    const carriers = await listCarriers();

    if (!carriers.length) {
      table.innerHTML = '<div class="empty-state">No carriers yet.</div>';
      return;
    }

    table.innerHTML = `
      <div class="data-row carrier-row data-row-head">
        <span>Name</span>
        <span>Status</span>
        <span>MC / DOT</span>
        <span>Insurance</span>
        <span></span>
      </div>
      ${carriers
        .map(
          (carrier) => `
            <div class="data-row carrier-row">
              <strong>${escapeHtml(carrier.name)}</strong>
              <span>${formatStatus(carrier.status)}</span>
              <span>${escapeHtml([carrier.mc, carrier.dot].filter(Boolean).join(" / ") || "-")}</span>
              <span>${formatDate(carrier.insurance_expiry)}</span>
              <div class="table-tools"><button class="row-action" type="button" data-open-profile="carrier" data-profile-id="${carrier.id}">Open</button>${can("manage_carriers") ? `<button class="row-action" type="button" data-edit-carrier="${carrier.id}">Edit</button>` : ""}</div>
            </div>
          `
        )
        .join("")}
    `;

    table.querySelectorAll("[data-edit-carrier]").forEach((button) => {
      const carrier = carriers.find((item) => item.id === button.dataset.editCarrier);
      button.addEventListener("click", () => fillCarrierForm(carrier));
    });
    bindBusinessProfileButtons(table, carriers);
  } catch (error) {
    table.innerHTML = `
      <div class="empty-state">
        Carriers table is not ready: ${escapeHtml(error.message)}
      </div>
    `;
  }
}

function fillCarrierForm(carrier) {
  const form = document.querySelector("#carrier-form");
  form.elements.id.value = carrier.id;
  form.elements.name.value = carrier.name || "";
  form.elements.mc.value = carrier.mc || "";
  form.elements.dot.value = carrier.dot || "";
  form.elements.safety_rating.value = carrier.safety_rating || "";
  form.elements.insurance_expiry.value = carrier.insurance_expiry || "";
  for (const field of ["contact_name", "contact_email", "contact_phone", "pay_type", "notes"]) form.elements[field].value = carrier[field] || "";
  form.elements.payment_terms_days.value = carrier.payment_terms_days ?? 30;
  form.elements.w9_status.value = carrier.w9_status || "missing";
  form.elements.pay_rate.value = carrier.pay_rate ?? "";
  form.elements.status.value = carrier.status || "active";
  document.querySelector("#carrier-form-title").textContent = "Edit Carrier";
  document.querySelector("#carrier-submit").textContent = "Update Carrier";
  document.querySelector("#cancel-carrier-edit").hidden = false;
  document.querySelector("#carrier-message").textContent = "";
  openFormDialog(form);
}

function resetCarrierForm() {
  const form = document.querySelector("#carrier-form");
  form.reset();
  form.elements.id.value = "";
  document.querySelector("#carrier-form-title").textContent = "Add Carrier";
  document.querySelector("#carrier-submit").textContent = "Save Carrier";
  document.querySelector("#cancel-carrier-edit").hidden = true;
  document.querySelector("#carrier-message").textContent = "";
  closeFormDialog(form);
}

async function handleCreateCarrier(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const message = document.querySelector("#carrier-message");
  const formData = new FormData(form);
  const carrierId = formData.get("id");

  message.textContent = carrierId ? "Updating carrier..." : "Saving carrier...";

  try {
    const payload = {
      name: formData.get("name"),
      mc: formData.get("mc") || null,
      dot: formData.get("dot") || null,
      safety_rating: formData.get("safety_rating") || null,
      insurance_expiry: formData.get("insurance_expiry") || null,
      contact_name: formData.get("contact_name") || null,
      contact_email: formData.get("contact_email") || null,
      contact_phone: formData.get("contact_phone") || null,
      payment_terms_days: numberOrZero(formData.get("payment_terms_days")),
      w9_status: formData.get("w9_status"),
      pay_type: formData.get("pay_type") || null,
      pay_rate: numberOrNull(formData.get("pay_rate")),
      notes: formData.get("notes") || null,
      status: formData.get("status"),
    };
    validateProfileFinancialTerms({ paymentTermsDays: payload.payment_terms_days, payRate: payload.pay_rate });

    if (carrierId) {
      await updateCarrier(carrierId, payload);
    } else {
      await createCarrier(payload);
    }

    resetCarrierForm();
    message.textContent = carrierId ? "Carrier updated." : "Carrier saved.";
    await loadCarriers();
  } catch (error) {
    message.textContent = error.message;
  }
}

function renderDrivers() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Phase 1", title: "Drivers" })}
    ${renderBusinessProfileDialog()}
    <section class="content-grid customer-grid">
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2>Driver List</h2>
            <p>Driver availability, CDL records and compliance dates.</p>
          </div>
          <button id="reload-drivers" type="button">Refresh</button>
        </div>
        <div id="drivers-table" class="data-table"></div>
      </article>
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2 id="driver-form-title">Add Driver</h2>
            <p>Creates a record in the TMS drivers table.</p>
          </div>
        </div>
        <form id="driver-form" class="record-form">
          <input name="id" type="hidden" />
          <label>
            <span>Name</span>
            <input name="name" type="text" required />
          </label>
          <label>
            <span>Phone</span>
            <input name="phone" type="tel" />
          </label>
          <label>
            <span>Email</span>
            <input name="email" type="email" />
          </label>
          <label><span>Address</span><input name="address" /></label>
          <label><span>Hire date</span><input name="hire_date" type="date" /></label>
          <label><span>Emergency contact</span><input name="emergency_contact_name" /></label>
          <label><span>Emergency phone</span><input name="emergency_contact_phone" type="tel" /></label>
          <label><span>Emergency relation</span><input name="emergency_contact_relation" /></label>
          <label>
            <span>CDL number</span>
            <input name="cdl_no" type="text" />
          </label>
          <label>
            <span>CDL expiry</span>
            <input name="cdl_expiry" type="date" />
          </label>
          <label>
            <span>Medical expiry</span>
            <input name="medical_expiry" type="date" />
          </label>
          <label>
            <span>Status</span>
            <select name="status">
              <option value="available">Available</option>
              <option value="on_load">On load</option>
              <option value="home_time">Home time</option>
              <option value="vacation">Vacation</option>
              <option value="out_of_service">Out of service</option>
              <option value="inactive">Inactive</option>
            </select>
          </label>
          <label><span>Pay type</span><select name="pay_type"><option value="">Not set</option><option value="per_mile">Per mile</option><option value="percentage">Percentage</option><option value="hourly">Hourly</option><option value="salary">Salary</option><option value="per_load">Per load</option></select></label>
          <label><span>Pay rate</span><input name="pay_rate" type="number" min="0" step="0.01" /></label>
          ${can("manage_users") ? `
            <section class="driver-credential-box" data-driver-credential-box>
              <label class="checkbox-row">
                <input name="create_driver_login" type="checkbox" />
                <span>Create Driver App login now</span>
              </label>
              <div class="driver-credential-fields" hidden>
                <label><span>Login email</span><input name="login_email" type="email" autocomplete="off" /></label>
                <label><span>Temporary password</span><input name="login_password" type="password" minlength="8" autocomplete="new-password" /></label>
                <p class="form-help">Password is used once to create the Auth user. It is not saved in the driver profile or shown later.</p>
              </div>
            </section>
          ` : ""}
          <div class="form-actions">
            <button id="driver-submit" type="submit">Save Driver</button>
            <button id="cancel-driver-edit" type="button" hidden>Cancel edit</button>
          </div>
          <p id="driver-message" class="form-message"></p>
        </form>
      </article>
    </section>
  `;

  document.querySelector("#reload-drivers").addEventListener("click", loadDrivers);
  document.querySelector("#driver-form").addEventListener("submit", handleCreateDriver);
  document.querySelector("#cancel-driver-edit").addEventListener("click", resetDriverForm);
  document.querySelector("#close-business-profile").addEventListener("click", () => document.querySelector("#business-profile-dialog").close());
  bindDriverCredentialFields();
  configureCrudFormAccess("#driver-form", "manage_assets", "Add Driver", resetDriverForm);
  loadDrivers();
}

async function loadDrivers() {
  const table = document.querySelector("#drivers-table");
  table.innerHTML = '<div class="empty-state">Loading drivers...</div>';

  try {
    const drivers = await listDrivers();

    if (!drivers.length) {
      table.innerHTML = '<div class="empty-state">No drivers yet.</div>';
      return;
    }

    table.innerHTML = `
      <div class="data-row driver-row data-row-head">
        <span>Name</span>
        <span>Status</span>
        <span>Phone</span>
        <span>CDL / Medical</span>
        <span>Driver App</span>
        <span></span>
      </div>
      ${drivers
        .map(
          (driver) => `
            <div class="data-row driver-row">
              <strong>${escapeHtml(driver.name)}</strong>
              <span>${formatStatus(driver.status)}</span>
              <span>${escapeHtml(driver.phone || "-")}</span>
              <span>${formatDate(driver.cdl_expiry)} / ${formatDate(driver.medical_expiry)}</span>
              ${renderDriverAppAccess(driver)}
              <div class="table-tools"><button class="row-action" type="button" data-open-profile="driver" data-profile-id="${driver.id}">Open</button>${can("manage_assets") ? `<button class="row-action" type="button" data-edit-driver="${driver.id}">Edit</button>` : ""}</div>
            </div>
          `
        )
        .join("")}
    `;

    table.querySelectorAll("[data-edit-driver]").forEach((button) => {
      const driver = drivers.find((item) => item.id === button.dataset.editDriver);
      button.addEventListener("click", () => fillDriverForm(driver));
    });
    bindBusinessProfileButtons(table, drivers);
  } catch (error) {
    table.innerHTML = `
      <div class="empty-state">
        Drivers table is not ready: ${escapeHtml(error.message)}
      </div>
    `;
  }
}

function renderDriverAppAccess(driver) {
  const linked = Boolean(driver.auth_user_id);
  return `<span class="driver-app-access" data-linked="${linked ? "true" : "false"}">
    <strong>${linked ? "Linked" : "No login"}</strong>
    <small>${linked ? escapeHtml(driver.email || "Driver App enabled") : "Create from Add Driver or Settings"}</small>
  </span>`;
}

function fillDriverForm(driver) {
  const form = document.querySelector("#driver-form");
  form.elements.id.value = driver.id;
  form.elements.name.value = driver.name || "";
  form.elements.phone.value = driver.phone || "";
  form.elements.email.value = driver.email || "";
  for (const field of ["address", "hire_date", "emergency_contact_name", "emergency_contact_phone", "emergency_contact_relation", "pay_type"]) form.elements[field].value = driver[field] || "";
  form.elements.pay_rate.value = driver.pay_rate ?? "";
  form.elements.cdl_no.value = driver.cdl_no || "";
  form.elements.cdl_expiry.value = driver.cdl_expiry || "";
  form.elements.medical_expiry.value = driver.medical_expiry || "";
  form.elements.status.value = driver.status || "available";
  document.querySelector("#driver-form-title").textContent = "Edit Driver";
  document.querySelector("#driver-submit").textContent = "Update Driver";
  document.querySelector("#cancel-driver-edit").hidden = false;
  document.querySelector("#driver-message").textContent = "";
  setDriverCredentialBoxEnabled(false);
  openFormDialog(form);
}

function resetDriverForm() {
  const form = document.querySelector("#driver-form");
  form.reset();
  form.elements.id.value = "";
  document.querySelector("#driver-form-title").textContent = "Add Driver";
  document.querySelector("#driver-submit").textContent = "Save Driver";
  document.querySelector("#cancel-driver-edit").hidden = true;
  document.querySelector("#driver-message").textContent = "";
  setDriverCredentialBoxEnabled(true);
  closeFormDialog(form);
}

function bindDriverCredentialFields() {
  const form = document.querySelector("#driver-form");
  const checkbox = form?.elements.create_driver_login;
  if (!form || !checkbox) return;
  const fields = form.querySelector(".driver-credential-fields");
  const loginEmail = form.elements.login_email;
  const password = form.elements.login_password;
  const sync = () => {
    const enabled = checkbox.checked && !form.elements.id.value;
    fields.hidden = !enabled;
    loginEmail.required = enabled;
    password.required = enabled;
    if (enabled && !loginEmail.value) loginEmail.value = form.elements.email.value || "";
  };
  checkbox.addEventListener("change", sync);
  form.elements.email.addEventListener("input", () => {
    if (checkbox.checked) loginEmail.value = form.elements.email.value || "";
  });
  sync();
}

function setDriverCredentialBoxEnabled(enabled) {
  const form = document.querySelector("#driver-form");
  const box = form?.querySelector("[data-driver-credential-box]");
  const checkbox = form?.elements.create_driver_login;
  if (!box || !checkbox) return;
  box.hidden = !enabled;
  checkbox.checked = false;
  box.querySelector(".driver-credential-fields").hidden = true;
  if (form.elements.login_email) form.elements.login_email.required = false;
  if (form.elements.login_password) form.elements.login_password.required = false;
}

function driverNameParts(name) {
  const parts = String(name || "").trim().split(/\s+/).filter(Boolean);
  return {
    firstName: parts[0] || "Driver",
    lastName: parts.length > 1 ? parts.slice(1).join(" ") : "User",
  };
}

async function handleCreateDriver(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const message = document.querySelector("#driver-message");
  const formData = new FormData(form);
  const driverId = formData.get("id");

  message.textContent = driverId ? "Updating driver..." : "Saving driver...";

  try {
    const payload = {
      name: formData.get("name"),
      phone: formData.get("phone") || null,
      email: formData.get("email") || null,
      address: formData.get("address") || null,
      hire_date: formData.get("hire_date") || null,
      emergency_contact_name: formData.get("emergency_contact_name") || null,
      emergency_contact_phone: formData.get("emergency_contact_phone") || null,
      emergency_contact_relation: formData.get("emergency_contact_relation") || null,
      cdl_no: formData.get("cdl_no") || null,
      cdl_expiry: formData.get("cdl_expiry") || null,
      medical_expiry: formData.get("medical_expiry") || null,
      pay_type: formData.get("pay_type") || null,
      pay_rate: numberOrNull(formData.get("pay_rate")),
      status: formData.get("status"),
    };
    validateProfileFinancialTerms({ payRate: payload.pay_rate });

    let savedDriver;
    if (driverId) {
      savedDriver = await updateDriver(driverId, payload);
    } else {
      savedDriver = await createDriver(payload);
      if (formData.get("create_driver_login") === "on") {
        const branchId = savedDriver.branch_id || currentAccess.find((access) => access.branch_id)?.branch_id;
        if (!branchId) throw new Error("Driver was saved, but login was not created because a branch is required.");
        const { firstName, lastName } = driverNameParts(savedDriver.name);
        await manageTmsUsers("create", {
          firstName,
          lastName,
          email: formData.get("login_email") || savedDriver.email,
          password: formData.get("login_password"),
          role: "driver",
          branchId,
          driverId: savedDriver.id,
          status: "active",
        });
      }
    }

    resetDriverForm();
    message.textContent = driverId ? "Driver updated." : formData.get("create_driver_login") === "on" ? "Driver and Driver App login created." : "Driver saved.";
    await loadDrivers();
  } catch (error) {
    message.textContent = error.message;
  }
}

function renderEquipment() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Phase 1", title: "Equipment" })}
    ${renderBusinessProfileDialog()}
    <section class="content-grid customer-grid">
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2>Truck List</h2>
            <p>Power units used for dispatch assignment and maintenance tracking.</p>
          </div>
          <button id="reload-trucks" type="button">Refresh</button>
        </div>
        <div id="trucks-table" class="data-table"></div>
      </article>
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2 id="truck-form-title">Add Truck</h2>
            <p>Creates a record in the TMS trucks table.</p>
          </div>
        </div>
        <form id="truck-form" class="record-form">
          <input name="id" type="hidden" />
          <label>
            <span>Unit number</span>
            <input name="unit_no" type="text" required />
          </label>
          <label>
            <span>VIN</span>
            <input name="vin" type="text" />
          </label>
          <label>
            <span>Plate</span>
            <input name="plate" type="text" />
          </label>
          <label>
            <span>Make</span>
            <input name="make" type="text" />
          </label>
          <label>
            <span>Model</span>
            <input name="model" type="text" />
          </label>
          <label>
            <span>Year</span>
            <input name="year" type="number" min="1980" max="2100" />
          </label>
          <label>
            <span>Odometer</span>
            <input name="odometer" type="number" min="0" step="0.1" />
          </label>
          <label>
            <span>Registration expiry</span>
            <input name="registration_expiry" type="date" />
          </label>
          <label>
            <span>Insurance expiry</span>
            <input name="insurance_expiry" type="date" />
          </label>
          <label>
            <span>Status</span>
            <select name="status">
              <option value="available">Available</option>
              <option value="assigned">Assigned</option>
              <option value="maintenance">Maintenance</option>
              <option value="out_of_service">Out of service</option>
              <option value="inactive">Inactive</option>
            </select>
          </label>
          <div class="form-actions">
            <button id="truck-submit" type="submit">Save Truck</button>
            <button id="cancel-truck-edit" type="button" hidden>Cancel edit</button>
          </div>
          <p id="truck-message" class="form-message"></p>
        </form>
      </article>
    </section>
    <section class="content-grid customer-grid">
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2>Trailer List</h2>
            <p>Trailers available for assignment, maintenance and document control.</p>
          </div>
          <button id="reload-trailers" type="button">Refresh</button>
        </div>
        <div id="trailers-table" class="data-table"></div>
      </article>
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2 id="trailer-form-title">Add Trailer</h2>
            <p>Creates a record in the TMS trailers table.</p>
          </div>
        </div>
        <form id="trailer-form" class="record-form">
          <input name="id" type="hidden" />
          <label>
            <span>Unit number</span>
            <input name="unit_no" type="text" required />
          </label>
          <label>
            <span>VIN</span>
            <input name="vin" type="text" />
          </label>
          <label>
            <span>Plate</span>
            <input name="plate" type="text" />
          </label>
          <label>
            <span>Type</span>
            <input name="type" type="text" placeholder="Dry van, reefer, flatbed..." />
          </label>
          <label>
            <span>Registration expiry</span>
            <input name="registration_expiry" type="date" />
          </label>
          <label>
            <span>Insurance expiry</span>
            <input name="insurance_expiry" type="date" />
          </label>
          <label>
            <span>Status</span>
            <select name="status">
              <option value="available">Available</option>
              <option value="assigned">Assigned</option>
              <option value="maintenance">Maintenance</option>
              <option value="out_of_service">Out of service</option>
              <option value="inactive">Inactive</option>
            </select>
          </label>
          <div class="form-actions">
            <button id="trailer-submit" type="submit">Save Trailer</button>
            <button id="cancel-trailer-edit" type="button" hidden>Cancel edit</button>
          </div>
          <p id="trailer-message" class="form-message"></p>
        </form>
      </article>
    </section>
  `;

  document.querySelector("#reload-trucks").addEventListener("click", loadTrucks);
  document.querySelector("#reload-trailers").addEventListener("click", loadTrailers);
  document.querySelector("#truck-form").addEventListener("submit", handleCreateTruck);
  document.querySelector("#trailer-form").addEventListener("submit", handleCreateTrailer);
  document.querySelector("#cancel-truck-edit").addEventListener("click", resetTruckForm);
  document.querySelector("#cancel-trailer-edit").addEventListener("click", resetTrailerForm);
  document.querySelector("#close-business-profile").addEventListener("click", () => document.querySelector("#business-profile-dialog").close());
  configureCrudFormAccess("#truck-form", "manage_assets", "Add Truck", resetTruckForm);
  configureCrudFormAccess("#trailer-form", "manage_assets", "Add Trailer", resetTrailerForm);
  setupPageTabs(Array.from(pageRoot.querySelectorAll(":scope > section.content-grid")), ["Trucks", "Trailers"]);
  loadTrucks();
  loadTrailers();
}

async function loadTrucks() {
  const table = document.querySelector("#trucks-table");
  table.innerHTML = '<div class="empty-state">Loading trucks...</div>';

  try {
    const trucks = await listTrucks();

    if (!trucks.length) {
      table.innerHTML = '<div class="empty-state">No trucks yet.</div>';
      return;
    }

    table.innerHTML = `
      <div class="data-row truck-row data-row-head">
        <span>Unit</span>
        <span>Status</span>
        <span>Plate</span>
        <span>Odometer</span>
        <span></span>
      </div>
      ${trucks
        .map(
          (truck) => `
            <div class="data-row truck-row">
              <strong>${escapeHtml(truck.unit_no)}</strong>
              <span>${formatStatus(truck.status)}</span>
              <span>${escapeHtml(truck.plate || "-")}</span>
              <span>${formatNumber(truck.odometer)}</span>
              <div class="table-tools"><button class="row-action" type="button" data-open-profile="truck" data-profile-id="${truck.id}">Open</button>${can("manage_assets") ? `<button class="row-action" type="button" data-edit-truck="${truck.id}">Edit</button>` : ""}</div>
            </div>
          `
        )
        .join("")}
    `;

    table.querySelectorAll("[data-edit-truck]").forEach((button) => {
      const truck = trucks.find((item) => item.id === button.dataset.editTruck);
      button.addEventListener("click", () => fillTruckForm(truck));
    });
    bindBusinessProfileButtons(table, trucks);
  } catch (error) {
    table.innerHTML = `
      <div class="empty-state">
        Trucks table is not ready: ${escapeHtml(error.message)}
      </div>
    `;
  }
}

function fillTruckForm(truck) {
  const form = document.querySelector("#truck-form");
  form.elements.id.value = truck.id;
  form.elements.unit_no.value = truck.unit_no || "";
  form.elements.vin.value = truck.vin || "";
  form.elements.plate.value = truck.plate || "";
  form.elements.make.value = truck.make || "";
  form.elements.model.value = truck.model || "";
  form.elements.year.value = truck.year || "";
  form.elements.odometer.value = truck.odometer ?? "";
  form.elements.registration_expiry.value = truck.registration_expiry || "";
  form.elements.insurance_expiry.value = truck.insurance_expiry || "";
  form.elements.status.value = truck.status || "available";
  document.querySelector("#truck-form-title").textContent = "Edit Truck";
  document.querySelector("#truck-submit").textContent = "Update Truck";
  document.querySelector("#cancel-truck-edit").hidden = false;
  document.querySelector("#truck-message").textContent = "";
  openFormDialog(form);
}

function resetTruckForm() {
  const form = document.querySelector("#truck-form");
  form.reset();
  form.elements.id.value = "";
  document.querySelector("#truck-form-title").textContent = "Add Truck";
  document.querySelector("#truck-submit").textContent = "Save Truck";
  document.querySelector("#cancel-truck-edit").hidden = true;
  document.querySelector("#truck-message").textContent = "";
  closeFormDialog(form);
}

async function handleCreateTruck(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const message = document.querySelector("#truck-message");
  const formData = new FormData(form);
  const truckId = formData.get("id");

  message.textContent = truckId ? "Updating truck..." : "Saving truck...";

  try {
    const payload = {
      unit_no: formData.get("unit_no"),
      vin: formData.get("vin") || null,
      plate: formData.get("plate") || null,
      make: formData.get("make") || null,
      model: formData.get("model") || null,
      year: numberOrNull(formData.get("year")),
      odometer: numberOrNull(formData.get("odometer")),
      registration_expiry: formData.get("registration_expiry") || null,
      insurance_expiry: formData.get("insurance_expiry") || null,
      status: formData.get("status"),
    };

    if (truckId) {
      await updateTruck(truckId, payload);
    } else {
      await createTruck(payload);
    }

    resetTruckForm();
    message.textContent = truckId ? "Truck updated." : "Truck saved.";
    await loadTrucks();
  } catch (error) {
    message.textContent = error.message;
  }
}

async function loadTrailers() {
  const table = document.querySelector("#trailers-table");
  table.innerHTML = '<div class="empty-state">Loading trailers...</div>';

  try {
    const trailers = await listTrailers();

    if (!trailers.length) {
      table.innerHTML = '<div class="empty-state">No trailers yet.</div>';
      return;
    }

    table.innerHTML = `
      <div class="data-row truck-row data-row-head">
        <span>Unit</span>
        <span>Status</span>
        <span>Plate</span>
        <span>Type</span>
        <span></span>
      </div>
      ${trailers
        .map(
          (trailer) => `
            <div class="data-row truck-row">
              <strong>${escapeHtml(trailer.unit_no)}</strong>
              <span>${formatStatus(trailer.status)}</span>
              <span>${escapeHtml(trailer.plate || "-")}</span>
              <span>${escapeHtml(trailer.type || "-")}</span>
              <div class="table-tools"><button class="row-action" type="button" data-open-profile="trailer" data-profile-id="${trailer.id}">Open</button>${can("manage_assets") ? `<button class="row-action" type="button" data-edit-trailer="${trailer.id}">Edit</button>` : ""}</div>
            </div>
          `
        )
        .join("")}
    `;

    table.querySelectorAll("[data-edit-trailer]").forEach((button) => {
      const trailer = trailers.find((item) => item.id === button.dataset.editTrailer);
      button.addEventListener("click", () => fillTrailerForm(trailer));
    });
    bindBusinessProfileButtons(table, trailers);
  } catch (error) {
    table.innerHTML = `
      <div class="empty-state">
        Trailers table is not ready: ${escapeHtml(error.message)}
      </div>
    `;
  }
}

function fillTrailerForm(trailer) {
  const form = document.querySelector("#trailer-form");
  form.elements.id.value = trailer.id;
  form.elements.unit_no.value = trailer.unit_no || "";
  form.elements.vin.value = trailer.vin || "";
  form.elements.plate.value = trailer.plate || "";
  form.elements.type.value = trailer.type || "";
  form.elements.registration_expiry.value = trailer.registration_expiry || "";
  form.elements.insurance_expiry.value = trailer.insurance_expiry || "";
  form.elements.status.value = trailer.status || "available";
  document.querySelector("#trailer-form-title").textContent = "Edit Trailer";
  document.querySelector("#trailer-submit").textContent = "Update Trailer";
  document.querySelector("#cancel-trailer-edit").hidden = false;
  document.querySelector("#trailer-message").textContent = "";
  openFormDialog(form);
}

function resetTrailerForm() {
  const form = document.querySelector("#trailer-form");
  form.reset();
  form.elements.id.value = "";
  document.querySelector("#trailer-form-title").textContent = "Add Trailer";
  document.querySelector("#trailer-submit").textContent = "Save Trailer";
  document.querySelector("#cancel-trailer-edit").hidden = true;
  document.querySelector("#trailer-message").textContent = "";
  closeFormDialog(form);
}

async function handleCreateTrailer(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const message = document.querySelector("#trailer-message");
  const formData = new FormData(form);
  const trailerId = formData.get("id");

  message.textContent = trailerId ? "Updating trailer..." : "Saving trailer...";

  try {
    const payload = {
      unit_no: formData.get("unit_no"),
      vin: formData.get("vin") || null,
      plate: formData.get("plate") || null,
      type: formData.get("type") || null,
      registration_expiry: formData.get("registration_expiry") || null,
      insurance_expiry: formData.get("insurance_expiry") || null,
      status: formData.get("status"),
    };

    if (trailerId) {
      await updateTrailer(trailerId, payload);
    } else {
      await createTrailer(payload);
    }

    resetTrailerForm();
    message.textContent = trailerId ? "Trailer updated." : "Trailer saved.";
    await loadTrailers();
  } catch (error) {
    message.textContent = error.message;
  }
}

const maintenanceTransitions = {
  open: ["open", "scheduled", "completed", "cancelled"],
  scheduled: ["scheduled", "open", "completed", "cancelled"],
  completed: ["completed"],
  cancelled: ["cancelled", "open"],
};

let maintenanceData = { records: [], trucks: [], trailers: [], drivers: [], carriers: [], compliance: [], view: "all", complianceView: "all" };

function renderMaintenance() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Operations", title: "Maintenance / Safety" })}
    <dialog id="maintenance-dialog" class="crud-dialog maintenance-dialog">
      <article class="panel">
        <div class="panel-header"><div><h2 id="maintenance-form-title">Add Maintenance</h2><p>Schedule service and track equipment downtime.</p></div><button id="close-maintenance-dialog" class="dialog-close" type="button">×</button></div>
        <form id="maintenance-form" class="record-form maintenance-form">
          <input name="id" type="hidden">
          <label><span>Equipment type</span><select name="equipment_type"><option value="truck">Truck</option><option value="trailer">Trailer</option></select></label>
          <label><span>Equipment</span><select name="equipment_id" required></select></label>
          <label><span>Service type</span><input name="service_type" list="maintenance-service-types" placeholder="Preventive maintenance" required></label>
          <datalist id="maintenance-service-types"><option value="Preventive maintenance"><option value="Oil change"><option value="Annual inspection"><option value="Tires"><option value="Brake service"><option value="Repair"><option value="Registration renewal"></datalist>
          <label><span>Due date</span><input name="due_date" type="date"></label>
          <label><span>Due miles</span><input name="due_miles" type="number" min="0" step="0.1"></label>
          <label><span>Scheduled date</span><input name="scheduled_date" type="date"></label>
          <label><span>Vendor</span><input name="vendor" placeholder="Service shop"></label>
          <label><span>Cost</span><input name="cost" type="number" min="0" step="0.01" value="0"></label>
          <label><span>Reference</span><input name="service_reference" placeholder="Work order or invoice"></label>
          <label class="checkbox-row maintenance-hold"><input name="out_of_service" type="checkbox"><span>Place equipment in Maintenance status</span></label>
          <label class="maintenance-notes"><span>Notes</span><textarea name="notes" rows="3"></textarea></label>
          <div class="form-actions"><button id="maintenance-submit" type="submit">Save Maintenance</button><button id="cancel-maintenance-edit" type="button" hidden>Cancel edit</button></div>
          <p id="maintenance-message" class="form-message"></p>
        </form>
      </article>
    </dialog>
    <dialog id="compliance-renewal-dialog" class="crud-dialog compliance-renewal-dialog">
      <article class="panel">
        <div class="panel-header"><div><h2>Renew Compliance</h2><p id="compliance-renewal-summary"></p></div><button id="close-compliance-renewal" class="dialog-close" type="button" aria-label="Close">×</button></div>
        <form id="compliance-renewal-form" class="record-form">
          <input name="compliance_index" type="hidden">
          <label><span>New expiration date</span><input name="expiration_date" type="date" required></label>
          <label><span>Renewed document (optional)</span><input name="file" type="file" accept="application/pdf,image/jpeg,image/png"></label>
          <p class="form-help">PDF, JPG or PNG, maximum 15 MB. If attached, the document is saved as Approved.</p>
          <div class="form-actions"><button type="submit">Save Renewal</button></div>
          <p id="compliance-renewal-message" class="form-message"></p>
        </form>
      </article>
    </dialog>
    <section id="maintenance-root" class="maintenance-root"><div class="empty-state">Loading maintenance...</div></section>
  `;
  const form = document.querySelector("#maintenance-form");
  form.addEventListener("submit", handleMaintenanceSave);
  form.elements.equipment_type.addEventListener("change", populateMaintenanceEquipment);
  document.querySelector("#close-maintenance-dialog").addEventListener("click", resetMaintenanceForm);
  document.querySelector("#cancel-maintenance-edit").addEventListener("click", resetMaintenanceForm);
  document.querySelector("#close-compliance-renewal").addEventListener("click", closeComplianceRenewal);
  document.querySelector("#compliance-renewal-form").addEventListener("submit", handleComplianceRenewal);
  loadMaintenancePage();
}

async function loadMaintenancePage() {
  const root = document.querySelector("#maintenance-root");
  try {
    const [records, trucks, trailers, drivers, carriers] = await Promise.all([listMaintenance(), listTrucks(), listTrailers(), listDrivers(), listCarriers()]);
    const compliance = buildComplianceItems({ drivers, carriers, trucks, trailers });
    maintenanceData = { ...maintenanceData, records, trucks, trailers, drivers, carriers, compliance };
    const active = records.filter((record) => ["open", "scheduled"].includes(record.status));
    const overdue = active.filter((record) => maintenanceDueState(record) === "overdue");
    const dueSoon = active.filter((record) => maintenanceDueState(record) === "due_soon");
    const completedCost = sumBy(records.filter((record) => record.status === "completed"), "cost");
    root.innerHTML = `
      <section class="report-kpis maintenance-kpis">
        ${renderReportKpi("Open", active.length)}
        ${renderReportKpi("Overdue", overdue.length)}
        ${renderReportKpi("Due soon", dueSoon.length)}
        ${renderReportKpi("Out of service", active.filter((record) => record.out_of_service).length)}
        ${renderReportKpi("Completed cost", formatMoney(completedCost))}
      </section>
      <section class="panel">
        <div class="panel-header"><div><h2>Maintenance Schedule</h2><p>Service due dates, mileage, cost and equipment availability.</p></div><div class="table-tools"><button id="reload-maintenance" type="button">Refresh</button>${can("manage_safety") ? '<button id="open-maintenance-dialog" class="primary-button" type="button">Add Maintenance</button>' : ""}</div></div>
        <div class="load-saved-views maintenance-views">${[["all","All"],["overdue","Overdue"],["due_soon","Due Soon"],["scheduled","Scheduled"],["completed","Completed"]].map(([value,label]) => `<button type="button" data-maintenance-view="${value}" data-active="${maintenanceData.view === value}">${label}</button>`).join("")}</div>
        <div id="maintenance-table" class="data-table"></div>
      </section>
      <section class="panel compliance-panel">
        <div class="panel-header"><div><h2>Safety Compliance</h2><p>CDL, medical, insurance and registration expiration tracking.</p></div></div>
        <section class="report-kpis compliance-kpis">
          ${renderReportKpi("Expired", compliance.filter((item) => item.state === "expired").length)}
          ${renderReportKpi("Due soon", compliance.filter((item) => item.state === "due_soon").length)}
          ${renderReportKpi("Missing dates", compliance.filter((item) => item.state === "missing").length)}
          ${renderReportKpi("Current", compliance.filter((item) => item.state === "current").length)}
        </section>
        <div class="load-saved-views compliance-views">${[["all","All"],["expired","Expired"],["due_soon","Due Soon"],["missing","Missing"],["current","Current"]].map(([value,label]) => `<button type="button" data-compliance-view="${value}" data-active="${maintenanceData.complianceView === value}">${label}</button>`).join("")}</div>
        <div id="compliance-table" class="data-table"></div>
      </section>
    `;
    document.querySelector("#reload-maintenance").addEventListener("click", loadMaintenancePage);
    document.querySelector("#open-maintenance-dialog")?.addEventListener("click", openMaintenanceCreate);
    root.querySelectorAll("[data-maintenance-view]").forEach((button) => button.addEventListener("click", () => {
      maintenanceData.view = button.dataset.maintenanceView;
      root.querySelectorAll("[data-maintenance-view]").forEach((item) => item.dataset.active = String(item === button));
      renderMaintenanceTable();
    }));
    root.querySelectorAll("[data-compliance-view]").forEach((button) => button.addEventListener("click", () => {
      maintenanceData.complianceView = button.dataset.complianceView;
      root.querySelectorAll("[data-compliance-view]").forEach((item) => item.dataset.active = String(item === button));
      renderComplianceTable();
    }));
    renderMaintenanceTable();
    renderComplianceTable();
  } catch (error) {
    root.innerHTML = `<div class="empty-state">Maintenance is not ready: ${escapeHtml(error.message)}</div>`;
  }
}

function buildComplianceItems({ drivers, carriers, trucks, trailers }) {
  const today = localDateIso();
  const dueSoonDate = localDateIso(30);
  const items = [];
  const add = (entityType, entityId, entityLabel, requirement, expiryDate, href) => items.push({
    entityType,
    entityId,
    entityLabel,
    requirement,
    expiryDate,
    href,
    state: getComplianceDueState(expiryDate, today, dueSoonDate),
  });

  drivers.filter((driver) => driver.status !== "inactive").forEach((driver) => {
    add("driver", driver.id, driver.name, "CDL", driver.cdl_expiry, "#drivers");
    add("driver", driver.id, driver.name, "Medical certificate", driver.medical_expiry, "#drivers");
  });
  carriers.filter((carrier) => carrier.status === "active").forEach((carrier) => {
    add("carrier", carrier.id, carrier.name, "Insurance", carrier.insurance_expiry, "#carriers");
  });
  trucks.filter((truck) => truck.status !== "inactive").forEach((truck) => {
    add("truck", truck.id, `Truck ${truck.unit_no}`, "Registration", truck.registration_expiry, "#equipment");
    add("truck", truck.id, `Truck ${truck.unit_no}`, "Insurance", truck.insurance_expiry, "#equipment");
  });
  trailers.filter((trailer) => trailer.status !== "inactive").forEach((trailer) => {
    add("trailer", trailer.id, `Trailer ${trailer.unit_no}`, "Registration", trailer.registration_expiry, "#equipment");
    add("trailer", trailer.id, `Trailer ${trailer.unit_no}`, "Insurance", trailer.insurance_expiry, "#equipment");
  });

  return items.sort((a, b) => {
    const rank = { expired: 0, due_soon: 1, missing: 2, current: 3 };
    return rank[a.state] - rank[b.state] || String(a.expiryDate || "9999-12-31").localeCompare(String(b.expiryDate || "9999-12-31"));
  });
}

function renderComplianceTable() {
  const table = document.querySelector("#compliance-table");
  if (!table) return;
  const filtered = maintenanceData.compliance.filter((item) => maintenanceData.complianceView === "all" || item.state === maintenanceData.complianceView);
  if (!filtered.length) {
    table.innerHTML = '<div class="empty-state">No compliance records in this view.</div>';
    return;
  }
  table.innerHTML = `
    <div class="data-row compliance-row data-row-head"><span>Entity</span><span>Requirement</span><span>Expiration</span><span>State</span><span>Action</span></div>
    ${filtered.map((item) => {
      const complianceIndex = maintenanceData.compliance.indexOf(item);
      return `
      <div class="data-row compliance-row" data-compliance-state="${item.state}">
        <span><strong>${escapeHtml(item.entityLabel)}</strong><small>${formatStatus(item.entityType)}</small></span>
        <strong>${escapeHtml(item.requirement)}</strong>
        <span>${formatDate(item.expiryDate)}</span>
        <span class="status-pill" data-state="${item.state === "expired" || item.state === "missing" ? "error" : item.state === "current" ? "ok" : "warning"}">${formatStatus(item.state)}</span>
        ${can("manage_safety") ? `<button type="button" data-renew-compliance="${complianceIndex}">Renew</button>` : `<a class="button-link" href="${item.href}">Open</a>`}
      </div>`;
    }).join("")}
  `;
  table.querySelectorAll("[data-renew-compliance]").forEach((button) => button.addEventListener("click", () => openComplianceRenewal(Number(button.dataset.renewCompliance))));
}

const complianceRenewalFields = {
  "driver:CDL": { field: "cdl_expiry", docType: "cdl", update: updateDriver },
  "driver:Medical certificate": { field: "medical_expiry", docType: "medical", update: updateDriver },
  "carrier:Insurance": { field: "insurance_expiry", docType: "insurance", update: updateCarrier },
  "truck:Registration": { field: "registration_expiry", docType: "registration", update: updateTruck },
  "truck:Insurance": { field: "insurance_expiry", docType: "insurance", update: updateTruck },
  "trailer:Registration": { field: "registration_expiry", docType: "registration", update: updateTrailer },
  "trailer:Insurance": { field: "insurance_expiry", docType: "insurance", update: updateTrailer },
};

function openComplianceRenewal(index) {
  const item = maintenanceData.compliance[index];
  const config = item && complianceRenewalFields[`${item.entityType}:${item.requirement}`];
  if (!item || !config) return;
  const form = document.querySelector("#compliance-renewal-form");
  form.reset();
  form.elements.compliance_index.value = String(index);
  form.elements.expiration_date.value = item.expiryDate || "";
  document.querySelector("#compliance-renewal-summary").textContent = `${item.entityLabel} · ${item.requirement} · Current: ${formatDate(item.expiryDate)}`;
  document.querySelector("#compliance-renewal-message").textContent = "";
  document.querySelector("#compliance-renewal-dialog").showModal();
}

function closeComplianceRenewal() {
  const dialog = document.querySelector("#compliance-renewal-dialog");
  if (dialog?.open) dialog.close();
}

async function handleComplianceRenewal(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const index = Number(form.elements.compliance_index.value);
  const item = maintenanceData.compliance[index];
  const config = item && complianceRenewalFields[`${item.entityType}:${item.requirement}`];
  const message = document.querySelector("#compliance-renewal-message");
  const button = form.querySelector("button[type='submit']");
  if (!item || !config) {
    message.textContent = "Compliance record is no longer available. Refresh and try again.";
    return;
  }

  const expirationDate = form.elements.expiration_date.value;
  const file = form.elements.file.files[0];
  button.disabled = true;
  message.textContent = file ? "Saving date and uploading renewed document..." : "Saving renewed expiration date...";
  try {
    await config.update(item.entityId, { [config.field]: expirationDate });
    if (file) {
      try {
        await uploadTmsDocument({
          entityType: item.entityType,
          entityId: item.entityId,
          docType: config.docType,
          status: "approved",
          file,
          uploadedBy: currentSession?.user?.id || null,
        });
      } catch (uploadError) {
        item.expiryDate = expirationDate;
        item.state = getComplianceDueState(expirationDate, localDateIso(), localDateIso(30));
        renderComplianceTable();
        message.textContent = `Expiration date was updated, but the document upload failed: ${uploadError.message}`;
        return;
      }
    }
    closeComplianceRenewal();
    await loadMaintenancePage();
  } catch (error) {
    message.textContent = error.message;
  } finally {
    button.disabled = false;
  }
}

function maintenanceEquipment(record) {
  const collection = record.equipment_type === "truck" ? maintenanceData.trucks : maintenanceData.trailers;
  return collection.find((item) => item.id === record.equipment_id);
}

function maintenanceDueState(record) {
  const equipment = maintenanceEquipment(record);
  return getMaintenanceDueState(record, equipment, localDateIso(), localDateIso(30));
}

function renderMaintenanceTable() {
  const table = document.querySelector("#maintenance-table");
  if (!table) return;
  const filtered = maintenanceData.records.filter((record) => {
    if (maintenanceData.view === "overdue" || maintenanceData.view === "due_soon") return maintenanceDueState(record) === maintenanceData.view;
    if (["scheduled", "completed"].includes(maintenanceData.view)) return record.status === maintenanceData.view;
    return true;
  });
  if (!filtered.length) {
    table.innerHTML = '<div class="empty-state">No maintenance records in this view.</div>';
    return;
  }
  table.innerHTML = `
    <div class="data-row maintenance-row data-row-head"><span>Equipment</span><span>Service</span><span>Due</span><span>Schedule</span><span>Vendor / Cost</span><span>Status / Actions</span></div>
    ${filtered.map((record) => {
      const equipment = maintenanceEquipment(record);
      const dueState = maintenanceDueState(record);
      return `<div class="data-row maintenance-row ${dueState === "overdue" ? "is-overdue" : ""}">
        <span><strong>${formatStatus(record.equipment_type)} ${escapeHtml(equipment?.unit_no || "-")}</strong><small>${record.out_of_service ? "Maintenance hold" : formatStatus(equipment?.status)}</small></span>
        <span><strong>${escapeHtml(record.service_type)}</strong><small>${escapeHtml(record.service_reference || "No reference")}</small></span>
        <span><strong>${formatDate(record.due_date)}</strong><small>${record.due_miles !== null ? `${formatNumber(record.due_miles)} mi` : formatStatus(dueState)}</small></span>
        <span>${formatDate(record.scheduled_date)}</span>
        <span><strong>${escapeHtml(record.vendor || "-")}</strong><small>${formatMoney(record.cost)}</small></span>
        <div class="maintenance-actions"><select data-maintenance-status="${record.id}" ${record.status === "completed" || !can("manage_safety") ? "disabled" : ""}>${(maintenanceTransitions[record.status] || [record.status]).map((status) => `<option value="${status}">${formatStatus(status)}</option>`).join("")}</select>${can("manage_safety") ? `<button type="button" data-edit-maintenance="${record.id}">Edit</button>` : ""}</div>
      </div>`;
    }).join("")}`;
  table.querySelectorAll("[data-maintenance-status]").forEach((select) => select.addEventListener("change", handleMaintenanceStatus));
  table.querySelectorAll("[data-edit-maintenance]").forEach((button) => button.addEventListener("click", () => openMaintenanceEdit(maintenanceData.records.find((record) => record.id === button.dataset.editMaintenance))));
}

function populateMaintenanceEquipment(selectedId = "") {
  const form = document.querySelector("#maintenance-form");
  const type = form.elements.equipment_type.value;
  const records = type === "truck" ? maintenanceData.trucks : maintenanceData.trailers;
  form.elements.equipment_id.innerHTML = records.map((item) => `<option value="${item.id}" ${item.id === selectedId ? "selected" : ""}>${escapeHtml(item.unit_no)} · ${formatStatus(item.status)}</option>`).join("");
  form.elements.due_miles.disabled = type === "trailer";
  if (type === "trailer") form.elements.due_miles.value = "";
}

function openMaintenanceCreate() {
  resetMaintenanceForm(false);
  populateMaintenanceEquipment();
  document.querySelector("#maintenance-dialog").showModal();
}

function openMaintenanceEdit(record) {
  const form = document.querySelector("#maintenance-form");
  form.elements.id.value = record.id;
  form.elements.equipment_type.value = record.equipment_type;
  populateMaintenanceEquipment(record.equipment_id);
  form.elements.equipment_type.disabled = true;
  form.elements.equipment_id.disabled = true;
  form.elements.service_type.value = record.service_type || "";
  form.elements.due_date.value = record.due_date || "";
  form.elements.due_miles.value = record.due_miles ?? "";
  form.elements.scheduled_date.value = record.scheduled_date || "";
  form.elements.vendor.value = record.vendor || "";
  form.elements.cost.value = record.cost ?? 0;
  form.elements.service_reference.value = record.service_reference || "";
  form.elements.out_of_service.checked = Boolean(record.out_of_service);
  form.elements.notes.value = record.notes || "";
  document.querySelector("#maintenance-form-title").textContent = "Edit Maintenance";
  document.querySelector("#maintenance-submit").textContent = "Update Maintenance";
  document.querySelector("#cancel-maintenance-edit").hidden = false;
  document.querySelector("#maintenance-message").textContent = "";
  document.querySelector("#maintenance-dialog").showModal();
}

function resetMaintenanceForm(close = true) {
  const form = document.querySelector("#maintenance-form");
  form.reset();
  form.elements.id.value = "";
  form.elements.equipment_type.disabled = false;
  form.elements.equipment_id.disabled = false;
  document.querySelector("#maintenance-form-title").textContent = "Add Maintenance";
  document.querySelector("#maintenance-submit").textContent = "Save Maintenance";
  document.querySelector("#cancel-maintenance-edit").hidden = true;
  document.querySelector("#maintenance-message").textContent = "";
  if (close) document.querySelector("#maintenance-dialog").close();
}

async function handleMaintenanceSave(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const maintenanceId = form.elements.id.value;
  const message = document.querySelector("#maintenance-message");
  const button = form.querySelector("button[type='submit']");
  const payload = {
    service_type: form.elements.service_type.value,
    due_date: form.elements.due_date.value || null,
    due_miles: numberOrNull(form.elements.due_miles.value),
    scheduled_date: form.elements.scheduled_date.value || null,
    vendor: form.elements.vendor.value || null,
    cost: numberOrZero(form.elements.cost.value),
    service_reference: form.elements.service_reference.value || null,
    out_of_service: form.elements.out_of_service.checked,
    notes: form.elements.notes.value || null,
  };
  if (!maintenanceId) {
    payload.equipment_type = form.elements.equipment_type.value;
    payload.equipment_id = form.elements.equipment_id.value;
    payload.status = "open";
  }
  button.disabled = true;
  message.textContent = maintenanceId ? "Updating maintenance..." : "Saving maintenance...";
  try {
    if (maintenanceId) await updateMaintenance(maintenanceId, payload); else await createMaintenance(payload);
    resetMaintenanceForm();
    await loadMaintenancePage();
  } catch (error) {
    message.textContent = error.message;
  } finally {
    button.disabled = false;
  }
}

async function handleMaintenanceStatus(event) {
  const select = event.currentTarget;
  select.disabled = true;
  try {
    await updateMaintenanceStatus(select.dataset.maintenanceStatus, select.value);
    await loadMaintenancePage();
  } catch (error) {
    window.alert(`Maintenance update failed: ${error.message}`);
    await loadMaintenancePage();
  }
}

async function loadTelematicsFoundation() {
  try {
    const [positions, hos] = await Promise.all([listTelematicsPositions(), listTelematicsHosSnapshots()]);
    return {
      positions: latestTelematicsRecords(positions, ["truck_id", "external_vehicle_id", "driver_id", "load_id"]),
      hos: latestTelematicsRecords(hos, ["driver_id", "external_driver_id"]),
      unavailableMessage: "",
    };
  } catch (error) {
    return { positions: [], hos: [], unavailableMessage: error.message };
  }
}

function latestTelemetryRecord(records, field, value) {
  if (!value) return null;
  return records.find((record) => record[field] === value) || null;
}

function telemetryForLoad(load, telemetry = fleetTelemetryData) {
  const position = latestTelemetryRecord(telemetry.positions, "load_id", load.id)
    || latestTelemetryRecord(telemetry.positions, "truck_id", load.truck_id)
    || latestTelemetryRecord(telemetry.positions, "driver_id", load.driver_id);
  const hos = latestTelemetryRecord(telemetry.hos, "driver_id", load.driver_id);
  return { position, hos };
}

function telemetryStateTone(state) {
  if (["current", "ok", "on_time", "arrived", "tracking", "complete"].includes(state)) return "ok";
  if (["stale", "warning", "missing", "unknown", "at_risk"].includes(state)) return "warning";
  return "error";
}

function formatTelematicsAge(record) {
  const freshness = telematicsFreshness(record, new Date().toISOString());
  if (freshness.ageMinutes === null) return freshness.label;
  return `${freshness.label} · ${freshness.ageMinutes} min ago`;
}

function formatHosRemaining(seconds) {
  if (seconds === null || seconds === undefined || seconds === "") return "-";
  const value = Number(seconds);
  if (!Number.isFinite(value)) return "-";
  const hours = Math.floor(value / 3600);
  const minutes = Math.floor((value % 3600) / 60);
  return `${hours}h ${String(minutes).padStart(2, "0")}m`;
}

function renderFleetMap() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Operations", title: "Fleet Map" })}
    <section id="fleet-map-root" class="fleet-map-root">
      <div class="empty-state">Loading fleet map...</div>
    </section>
  `;

  loadFleetMap();
}

function predictionStop(load) {
  const stop = nextOperationalStop(load);
  if (stop) return stop;
  if (load.load_stops?.length) {
    return {
      type: "complete",
      facility: "No pending stops",
      address: null,
      schedule_complete: true,
    };
  }
  return {
    type: "delivery",
    facility: load.destination || "Delivery",
    address: load.destination,
    appointment_to: load.delivery_date ? `${load.delivery_date}T23:59:59` : null,
  };
}

async function buildOperationalPredictions(loads, telemetry) {
  const now = new Date().toISOString();
  const entries = await Promise.all(loads.map(async (load) => {
    const live = telemetryForLoad(load, telemetry);
    const stop = predictionStop(load);
    if (stop.schedule_complete) {
      return [load.id, {
        state: "complete",
        stop,
        address: null,
        live,
        hos: hosRiskState(live.hos),
        freshness: telematicsFreshness(live.position, now),
        distanceMiles: null,
        etaAt: null,
        deltaMinutes: null,
        insideGeofence: false,
        geofenceTarget: null,
      }];
    }
    const address = stop.address || (stop.type === "pickup" ? load.origin : load.destination);
    const geocoded = live.position && address ? await geocodeLocation(address) : null;
    const hos = hosRiskState(live.hos);
    const estimate = estimateStopArrival({
      position: live.position,
      destination: geocoded ? { latitude: geocoded.coordinates[1], longitude: geocoded.coordinates[0] } : null,
      stop,
      hosState: hos.state,
      now,
    });
    const freshness = telematicsFreshness(live.position, now);
    const prediction = applyEtaSignalFreshness(estimate, freshness);
    return [load.id, {
      ...prediction,
      stop,
      address,
      live,
      hos,
      freshness,
      geofenceTarget: geocoded ? { latitude: geocoded.coordinates[1], longitude: geocoded.coordinates[0] } : null,
    }];
  }));
  return Object.fromEntries(entries);
}

async function loadFleetMap() {
  const root = document.querySelector("#fleet-map-root");

  try {
    const [loads, drivers, trucks, trailers, telemetry, exceptions, geofenceEvents] = await Promise.all([
      listLoads(),
      listDrivers(),
      listTrucks(),
      listTrailers(),
      loadTelematicsFoundation(),
      listOperationalExceptions(),
      listLoadGeofenceEvents(),
    ]);
    fleetTelemetryData = telemetry;
    const activeLoads = loads.filter((load) => !["delivered", "cancelled"].includes(operationalLoadStatus(load)));
    const availableTrucks = trucks.filter((truck) => truck.status === "available").length;
    const availableTrailers = trailers.filter((trailer) => trailer.status === "available").length;
    const availableDrivers = drivers.filter((driver) => driver.status === "available").length;
    const telemetryFreshness = telemetry.positions.map((position) => telematicsFreshness(position, new Date().toISOString()));
    const currentSignals = telemetryFreshness.filter((item) => item.state === "current").length;
    const staleSignals = telemetryFreshness.filter((item) => ["stale", "offline"].includes(item.state)).length;
    const predictions = await buildOperationalPredictions(activeLoads, telemetry);
    const riskCount = Object.values(predictions).filter((item) => ["at_risk", "late"].includes(item.state)).length;
    const activeExceptions = exceptions.filter((item) => ["open", "acknowledged"].includes(item.status));
    const pendingGeofenceEvents = geofenceEvents.filter((item) => item.status === "pending");

    root.innerHTML = `
      <section class="fleet-kpis">
        ${renderReportKpi("Active loads", activeLoads.length)}
        ${renderReportKpi("Available drivers", availableDrivers)}
        ${renderReportKpi("Live signals", currentSignals)}
        ${renderReportKpi("Stale/offline", staleSignals)}
        ${renderReportKpi("ETA risks", riskCount)}
        ${renderReportKpi("Open exceptions", activeExceptions.length)}
      </section>
      <section class="fleet-map-panel panel">
        <div class="panel-header">
          <div>
            <h2>Mapbox Fleet Map</h2>
            <p>Planned routes plus provider-neutral live telemetry when an adapter has posted data.</p>
          </div>
          <button id="reload-fleet-map" type="button">Refresh</button>
        </div>
        <div class="mapbox-shell">
          <div id="fleet-map" class="mapbox-map" aria-label="Fleet map"></div>
          <div id="fleet-map-message" class="mapbox-message">Preparing Mapbox...</div>
        </div>
        <div class="fleet-map-canvas route-list">
          ${activeLoads.length ? activeLoads.slice(0, 10).map((load) => renderFleetRoute(load, telemetryForLoad(load, telemetry), predictions[load.id])).join("") : '<div class="empty-state compact-empty">No active loads to map.</div>'}
        </div>
      </section>
      <section class="panel eta-risk-panel">
        <div class="panel-header">
          <div>
            <h2>ETA & Geofence Watch</h2>
            <p>Estimated arrival to the next stop from the latest GPS snapshot. Estimates are advisory until a dispatcher confirms the event.</p>
          </div>
        </div>
        ${fleetNotice ? `<p class="form-message fleet-ops-message">${escapeHtml(fleetNotice)}</p>` : ""}
        <div class="eta-risk-list">
          ${activeLoads.length ? activeLoads.map((load) => renderEtaRiskRow(load, predictions[load.id], activeExceptions, pendingGeofenceEvents)).join("") : '<div class="empty-state compact-empty">No active loads to evaluate.</div>'}
        </div>
      </section>
      ${renderOperationalExceptionPanel(exceptions, pendingGeofenceEvents, loads)}
      <section class="panel telemetry-panel">
        <div class="panel-header">
          <div>
            <h2>Live Telematics</h2>
            <p>Latest GPS and HOS snapshots from server-side provider adapters.</p>
          </div>
        </div>
        ${renderTelematicsOverview(telemetry)}
      </section>
      <section class="fleet-columns">
        ${renderFleetAssetColumn("Drivers", drivers, "name", "driver_id", telemetry)}
        ${renderFleetAssetColumn("Trucks", trucks, "unit_no", "truck_id", telemetry)}
        ${renderFleetAssetColumn("Trailers", trailers, "unit_no", "trailer_id", telemetry)}
      </section>
      ${renderOperationalExceptionDialogs()}
    `;

    document.querySelector("#reload-fleet-map").addEventListener("click", loadFleetMap);
    bindOperationalExceptionActions();
    renderMapboxFleetMap(activeLoads, telemetry.positions);
  } catch (error) {
    root.innerHTML = `<div class="empty-state">Fleet map is not ready: ${escapeHtml(error.message)}</div>`;
  }
}

function renderTelematicsOverview(telemetry) {
  if (telemetry.unavailableMessage) {
    return `<div class="integration-unavailable">Telematics foundation is not available on this environment yet: ${escapeHtml(telemetry.unavailableMessage)}</div>`;
  }
  const currentPositions = telemetry.positions.filter((position) => telematicsFreshness(position, new Date().toISOString()).state === "current").length;
  const safeHos = telemetry.hos.filter((snapshot) => hosRiskState(snapshot).state === "ok").length;
  return `
    <div class="telemetry-grid">
      <article class="telemetry-card">
        <header class="telemetry-card-header">
          <div><span class="telemetry-card-icon" aria-hidden="true">GPS</span><div><h3>Vehicle Positions</h3><small>Latest location by truck</small></div></div>
          <strong>${currentPositions}/${telemetry.positions.length} live</strong>
        </header>
        <div class="telemetry-list">
          ${telemetry.positions.length ? telemetry.positions.slice(0, 8).map(renderTelemetryPosition).join("") : '<div class="empty-state compact-empty">No live position snapshots yet.</div>'}
        </div>
      </article>
      <article class="telemetry-card">
        <header class="telemetry-card-header">
          <div><span class="telemetry-card-icon" aria-hidden="true">HOS</span><div><h3>Driver Hours</h3><small>Latest hours-of-service snapshot</small></div></div>
          <strong>${safeHos}/${telemetry.hos.length} clear</strong>
        </header>
        <div class="telemetry-list">
          ${telemetry.hos.length ? telemetry.hos.slice(0, 8).map(renderHosSnapshot).join("") : '<div class="empty-state compact-empty">No HOS snapshots yet.</div>'}
        </div>
      </article>
    </div>`;
}

function renderTelemetryPosition(position) {
  const freshness = telematicsFreshness(position, new Date().toISOString());
  return `<div class="telemetry-row">
    <span class="telemetry-identity"><strong>${escapeHtml(position.trucks?.unit_no ? `Truck ${position.trucks.unit_no}` : position.external_vehicle_id || "Vehicle")}</strong><small>${escapeHtml(position.drivers?.name || "No assigned driver")}</small></span>
    <span class="status-pill" data-state="${telemetryStateTone(freshness.state)}">${formatStatus(freshness.state)}</span>
    <span class="telemetry-metric"><small>Last signal</small><strong>${escapeHtml(formatTelematicsAge(position).replace(`${freshness.label} · `, ""))}</strong></span>
    <span class="telemetry-metric"><small>Position</small><strong>${Number(position.latitude).toFixed(4)}, ${Number(position.longitude).toFixed(4)}</strong></span>
  </div>`;
}

function renderHosSnapshot(snapshot) {
  const risk = hosRiskState(snapshot);
  return `<div class="telemetry-row telemetry-hos-row">
    <span class="telemetry-identity"><strong>${escapeHtml(snapshot.drivers?.name || "Driver")}</strong><small>${escapeHtml(formatStatus(snapshot.duty_status))}</small></span>
    <span class="status-pill" data-state="${telemetryStateTone(risk.state)}">${formatStatus(risk.state)}</span>
    <span class="telemetry-metric"><small>Drive left</small><strong>${formatHosRemaining(snapshot.drive_seconds_remaining)}</strong></span>
    <span class="telemetry-metric"><small>Shift left</small><strong>${formatHosRemaining(snapshot.shift_seconds_remaining)}</strong></span>
  </div>`;
}

function renderFleetRoute(load, telemetry = telemetryForLoad(load), prediction = null) {
  const freshness = telematicsFreshness(telemetry.position, new Date().toISOString());
  const hos = hosRiskState(telemetry.hos);
  return `
    <article class="fleet-route-card">
      <div>
        <strong>${escapeHtml(load.load_no)}</strong>
        <span>${formatStatus(operationalLoadStatus(load))}</span>
      </div>
      <p>${escapeHtml(routeLabel(load))}</p>
      <small>
        ${escapeHtml(load.drivers?.name || "No driver")} ·
        ${escapeHtml(load.trucks?.unit_no || "No truck")} ·
        ${escapeHtml(load.trailers?.unit_no || "No trailer")}
      </small>
      <div class="telemetry-badges">
        <span data-tone="${telemetryStateTone(freshness.state)}">${escapeHtml(formatTelematicsAge(telemetry.position))}</span>
        <span data-tone="${telemetryStateTone(hos.state)}">${escapeHtml(hos.label)}</span>
        ${prediction ? `<span data-tone="${telemetryStateTone(prediction.state)}">ETA ${escapeHtml(formatStatus(prediction.state))}</span>` : ""}
      </div>
    </article>
  `;
}

function renderEtaRiskRow(load, prediction, activeExceptions = [], pendingEvents = []) {
  const scheduleComplete = Boolean(prediction?.stop?.schedule_complete);
  const stopLabel = scheduleComplete
    ? "All recorded stops completed"
    : prediction?.stop
    ? `${formatStatus(prediction.stop.type || "stop")} · ${prediction.stop.facility || prediction.address || "Next stop"}`
    : "Next stop unavailable";
  const distance = prediction?.distanceMiles === null || prediction?.distanceMiles === undefined
    ? "Distance unavailable"
    : `${prediction.distanceMiles.toFixed(1)} mi away`;
  const eta = scheduleComplete ? "No pending ETA" : prediction?.etaAt ? `ETA ${formatDateTime(prediction.etaAt)}` : "ETA unavailable";
  let variance = scheduleComplete ? "Confirmed arrival and departure recorded" : "No appointment comparison";
  if (!scheduleComplete && prediction?.state === "offline") variance = `Last GPS signal ${prediction.freshness?.ageMinutes ?? "-"} min ago`;
  else if (prediction?.state === "unknown") variance = "A live GPS signal is required";
  else if (prediction?.insideGeofence) variance = "Inside 0.25 mi geofence — confirmation required";
  else if (prediction?.deltaMinutes !== null && prediction?.deltaMinutes !== undefined) {
    variance = prediction.deltaMinutes > 0
      ? `${prediction.deltaMinutes} min after appointment`
      : `${Math.abs(prediction.deltaMinutes)} min before appointment`;
  }
  const exceptionType = exceptionTypeForPrediction(prediction);
  const existingException = exceptionType
    ? activeExceptions.find((item) => item.load_id === load.id && item.exception_type === exceptionType)
    : null;
  const pendingEvent = prediction?.stop?.id
    ? pendingEvents.find((item) => item.stop_id === prediction.stop.id)
    : null;
  const eventType = prediction?.stop?.actual_arrival ? "departure" : "arrival";
  const reason = suggestedExceptionReason(load, prediction, exceptionType);
  return `<article class="eta-risk-row" data-state="${telemetryStateTone(prediction?.state || "unknown")}">
    <div><strong>Load ${escapeHtml(load.load_no)}</strong><small>${escapeHtml(stopLabel)}</small></div>
    <div><strong>${escapeHtml(formatStatus(prediction?.state || "unknown"))}</strong><small>${escapeHtml(distance)}</small></div>
    <div><strong>${escapeHtml(eta)}</strong><small>${escapeHtml(variance)}</small></div>
    <div class="eta-risk-actions">
      ${existingException ? `<span class="status-pill" data-state="warning">${escapeHtml(formatStatus(existingException.status))} exception</span>` : exceptionType && exceptionType !== "geofence_review" ? `<button type="button" data-create-exception data-load-id="${load.id}" data-stop-id="${prediction?.stop?.id || ""}" data-exception-type="${exceptionType}" data-reason="${escapeHtml(reason)}">Create Exception</button>` : ""}
      ${prediction?.insideGeofence && prediction?.stop?.id ? (pendingEvent ? '<span class="status-pill" data-state="warning">Geofence review pending</span>' : `<button type="button" data-create-geofence data-load-id="${load.id}" data-stop-id="${prediction.stop.id}" data-event-type="${eventType}" data-position-id="${prediction.live?.position?.id || ""}" data-occurred-at="${prediction.live?.position?.event_at || ""}" data-distance="${prediction.distanceMiles ?? ""}">Review ${formatStatus(eventType)}</button>`) : ""}
      ${appConfig.name === "staging" && currentRoles.includes("owner") && load.truck_id && prediction?.stop?.id && prediction?.geofenceTarget && !prediction.insideGeofence ? `<button type="button" data-stage-geofence data-branch-id="${load.branch_id}" data-load-id="${load.id}" data-latitude="${prediction.geofenceTarget.latitude}" data-longitude="${prediction.geofenceTarget.longitude}">Stage Geofence Signal</button>` : ""}
    </div>
  </article>`;
}

function suggestedExceptionReason(load, prediction, type) {
  if (type === "offline_signal") return `GPS signal for load ${load.load_no} is ${prediction.state}${prediction.freshness?.ageMinutes !== null ? ` (${prediction.freshness.ageMinutes} minutes)` : ""}.`;
  if (type === "hos_risk") return `HOS risk for ${load.drivers?.name || `load ${load.load_no}`}: ${prediction.hos?.label || "review required"}.`;
  if (["late_pickup", "late_delivery"].includes(type)) return `Load ${load.load_no} is projected ${Math.max(0, prediction.deltaMinutes || 0)} minutes after the ${prediction.stop?.type || "stop"} appointment.`;
  return `Operational review required for load ${load.load_no}.`;
}

function renderOperationalExceptionPanel(exceptions, pendingEvents, loads) {
  const loadMap = Object.fromEntries(loads.map((load) => [load.id, load]));
  return `<section class="panel exception-panel">
    <div class="panel-header">
      <div><h2>Operational Exceptions</h2><p>Tracked issues with ownership, acknowledgement and resolution history.</p></div>
      <span class="status-pill" data-state="${exceptions.some((item) => item.status === "open") ? "warning" : "ok"}">${exceptions.filter((item) => ["open", "acknowledged"].includes(item.status)).length} active</span>
    </div>
    <div class="exception-list">
      ${exceptions.length ? exceptions.slice(0, 20).map((item) => renderOperationalExceptionRow(item, loadMap[item.load_id])).join("") : '<div class="empty-state compact-empty">No operational exceptions recorded.</div>'}
    </div>
    ${pendingEvents.length ? `<div class="geofence-review-section"><h3>Pending Geofence Reviews</h3>${pendingEvents.map((event) => `<article><span><strong>Load ${escapeHtml(loadMap[event.load_id]?.load_no || event.load_id)}</strong><small>${escapeHtml(formatStatus(event.event_type))} · ${event.distance_miles === null ? "Distance unavailable" : `${Number(event.distance_miles).toFixed(2)} mi`}</small></span><button type="button" data-review-geofence="${event.id}" data-event-label="${escapeHtml(`${formatStatus(event.event_type)} · Load ${loadMap[event.load_id]?.load_no || ""}`)}">Review</button></article>`).join("")}</div>` : ""}
  </section>`;
}

function renderOperationalExceptionRow(item, load) {
  const owner = item.owner_id === currentSession?.user?.id ? "You" : item.owner_id ? "Assigned user" : "Unassigned";
  const active = ["open", "acknowledged"].includes(item.status);
  return `<article class="exception-row" data-severity="${item.severity}">
    <div><strong>Load ${escapeHtml(load?.load_no || item.load_id)}</strong><small>${escapeHtml(formatStatus(item.exception_type))} · ${escapeHtml(formatStatus(item.severity))}</small></div>
    <div class="exception-reason"><strong>${escapeHtml(item.reason)}</strong><small>Detected ${formatDateTime(item.detected_at)}</small></div>
    <div><span class="status-pill" data-state="${item.status === "resolved" ? "ok" : item.status === "dismissed" ? "neutral" : "warning"}">${escapeHtml(formatStatus(item.status))}</span><small>${escapeHtml(owner)}</small></div>
    <div class="exception-actions">
      ${active && !item.owner_id ? `<button type="button" data-take-exception="${item.id}">Take Ownership</button>` : ""}
      ${item.status === "open" ? `<button type="button" data-ack-exception="${item.id}">Acknowledge</button>` : ""}
      ${active ? `<button type="button" data-close-exception="${item.id}" data-exception-label="${escapeHtml(`Load ${load?.load_no || ""} · ${formatStatus(item.exception_type)}`)}">Close</button>` : item.resolution ? `<small>${escapeHtml(item.resolution)}</small>` : ""}
    </div>
  </article>`;
}

function renderOperationalExceptionDialogs() {
  return `
    <dialog id="create-exception-dialog" class="crud-dialog exception-dialog"><article class="panel">
      <div class="panel-header"><div><h2>Create Exception</h2><p>Record an operational issue for follow-up.</p></div><button type="button" class="dialog-close" data-close-dialog>×</button></div>
      <form id="create-exception-form" class="record-form">
        <input name="load_id" type="hidden"><input name="stop_id" type="hidden">
        <label><span>Type</span><select name="exception_type"><option value="offline_signal">Offline signal</option><option value="late_pickup">Late pickup</option><option value="late_delivery">Late delivery</option><option value="hos_risk">HOS risk</option><option value="geofence_review">Geofence review</option><option value="other">Other</option></select></label>
        <label><span>Severity</span><select name="severity"><option value="warning">Warning</option><option value="critical">Critical</option><option value="info">Info</option></select></label>
        <label><span>Reason</span><textarea name="reason" maxlength="1000" required></textarea></label>
        <label class="checkbox-field"><input name="assign_to_me" type="checkbox" checked><span>Assign this exception to me</span></label>
        <div class="form-actions"><button type="submit">Save Exception</button></div><p class="form-message" id="create-exception-message"></p>
      </form>
    </article></dialog>
    <dialog id="close-exception-dialog" class="crud-dialog exception-dialog"><article class="panel">
      <div class="panel-header"><div><h2>Close Exception</h2><p id="close-exception-label"></p></div><button type="button" class="dialog-close" data-close-dialog>×</button></div>
      <form id="close-exception-form" class="record-form"><input name="exception_id" type="hidden">
        <label><span>Decision</span><select name="status"><option value="resolved">Resolved</option><option value="dismissed">Dismissed</option></select></label>
        <label><span>Resolution</span><textarea name="resolution" maxlength="2000" required placeholder="What was checked or corrected?"></textarea></label>
        <div class="form-actions"><button type="submit">Close Exception</button></div><p class="form-message" id="close-exception-message"></p>
      </form>
    </article></dialog>
    <dialog id="geofence-review-dialog" class="crud-dialog exception-dialog"><article class="panel">
      <div class="panel-header"><div><h2>Review Geofence Event</h2><p id="geofence-review-label"></p></div><button type="button" class="dialog-close" data-close-dialog>×</button></div>
      <form id="geofence-review-form" class="record-form"><input name="event_id" type="hidden">
        <label><span>Decision</span><select name="decision"><option value="confirmed">Confirm event</option><option value="rejected">Reject event</option></select></label>
        <label><span>Review note</span><textarea name="note" maxlength="2000" placeholder="Optional operational note"></textarea></label>
        <div class="form-actions"><button type="submit">Save Review</button></div><p class="form-message" id="geofence-review-message"></p>
      </form>
    </article></dialog>`;
}

function bindOperationalExceptionActions() {
  document.querySelectorAll("[data-close-dialog]").forEach((button) => button.addEventListener("click", () => button.closest("dialog").close()));
  document.querySelectorAll("[data-create-exception]").forEach((button) => button.addEventListener("click", () => {
    const form = document.querySelector("#create-exception-form");
    form.reset();
    form.elements.load_id.value = button.dataset.loadId;
    form.elements.stop_id.value = button.dataset.stopId || "";
    form.elements.exception_type.value = button.dataset.exceptionType;
    form.elements.reason.value = button.dataset.reason;
    form.elements.assign_to_me.checked = true;
    document.querySelector("#create-exception-message").textContent = "";
    document.querySelector("#create-exception-dialog").showModal();
  }));
  document.querySelectorAll("[data-create-geofence]").forEach((button) => button.addEventListener("click", () => handleCreateGeofenceReview(button)));
  document.querySelectorAll("[data-stage-geofence]").forEach((button) => button.addEventListener("click", () => handleStageGeofenceSignal(button)));
  document.querySelectorAll("[data-take-exception]").forEach((button) => button.addEventListener("click", () => handleQuickExceptionUpdate(button, { owner_id: currentSession.user.id })));
  document.querySelectorAll("[data-ack-exception]").forEach((button) => button.addEventListener("click", () => handleQuickExceptionUpdate(button, { status: "acknowledged" })));
  document.querySelectorAll("[data-close-exception]").forEach((button) => button.addEventListener("click", () => {
    const form = document.querySelector("#close-exception-form");
    form.reset();
    form.elements.exception_id.value = button.dataset.closeException;
    document.querySelector("#close-exception-label").textContent = button.dataset.exceptionLabel;
    document.querySelector("#close-exception-message").textContent = "";
    document.querySelector("#close-exception-dialog").showModal();
  }));
  document.querySelectorAll("[data-review-geofence]").forEach((button) => button.addEventListener("click", () => {
    const form = document.querySelector("#geofence-review-form");
    form.reset();
    form.elements.event_id.value = button.dataset.reviewGeofence;
    document.querySelector("#geofence-review-label").textContent = button.dataset.eventLabel;
    document.querySelector("#geofence-review-message").textContent = "";
    document.querySelector("#geofence-review-dialog").showModal();
  }));
  document.querySelector("#create-exception-form").addEventListener("submit", handleCreateOperationalException);
  document.querySelector("#close-exception-form").addEventListener("submit", handleCloseOperationalException);
  document.querySelector("#geofence-review-form").addEventListener("submit", handleReviewGeofenceEvent);
}

async function handleStageGeofenceSignal(button) {
  button.disabled = true;
  const originalLabel = button.textContent;
  button.textContent = "Staging...";
  try {
    await generateMockTelematics(button.dataset.branchId, {
      targetLoadId: button.dataset.loadId,
      targetLatitude: Number(button.dataset.latitude),
      targetLongitude: Number(button.dataset.longitude),
    });
    fleetNotice = "Fresh staging GPS signal placed inside the next-stop geofence. Review the arrival or departure event below.";
    await loadFleetMap();
  } catch (error) {
    fleetNotice = `Geofence signal could not be staged: ${error.message}`;
    await loadFleetMap();
  } finally {
    button.disabled = false;
    button.textContent = originalLabel;
  }
}

async function handleCreateOperationalException(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = document.querySelector("#create-exception-message");
  const button = form.querySelector("button[type='submit']");
  button.disabled = true;
  try {
    const validated = validateOperationalExceptionInput({ reason: form.elements.reason.value });
    await createOperationalException({
      load_id: form.elements.load_id.value,
      stop_id: form.elements.stop_id.value || null,
      exception_type: form.elements.exception_type.value,
      severity: form.elements.severity.value,
      source: ["offline_signal", "hos_risk"].includes(form.elements.exception_type.value) ? "telematics" : "eta",
      reason: validated.reason,
      owner_id: form.elements.assign_to_me.checked ? currentSession.user.id : null,
    });
    document.querySelector("#create-exception-dialog").close();
    await loadFleetMap();
  } catch (error) {
    message.textContent = error.message;
  } finally {
    button.disabled = false;
  }
}

async function handleQuickExceptionUpdate(button, updates) {
  button.disabled = true;
  try {
    await updateOperationalException(button.dataset.takeException || button.dataset.ackException, updates);
    await loadFleetMap();
  } catch (error) {
    window.alert(`Exception update failed: ${error.message}`);
    button.disabled = false;
  }
}

async function handleCloseOperationalException(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = document.querySelector("#close-exception-message");
  const button = form.querySelector("button[type='submit']");
  button.disabled = true;
  try {
    const validated = validateOperationalExceptionInput({ reason: "Existing exception", resolution: form.elements.resolution.value, closing: true });
    await updateOperationalException(form.elements.exception_id.value, { status: form.elements.status.value, resolution: validated.resolution });
    document.querySelector("#close-exception-dialog").close();
    await loadFleetMap();
  } catch (error) {
    message.textContent = error.message;
  } finally {
    button.disabled = false;
  }
}

async function handleCreateGeofenceReview(button) {
  button.disabled = true;
  try {
    await createLoadGeofenceEvent({
      load_id: button.dataset.loadId,
      stop_id: button.dataset.stopId,
      event_type: button.dataset.eventType,
      source_position_id: button.dataset.positionId || null,
      distance_miles: button.dataset.distance === "" ? null : Number(button.dataset.distance),
      occurred_at: button.dataset.occurredAt || new Date().toISOString(),
    });
    await loadFleetMap();
  } catch (error) {
    window.alert(`Geofence review could not be created: ${error.message}`);
    button.disabled = false;
  }
}

async function handleReviewGeofenceEvent(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = document.querySelector("#geofence-review-message");
  const button = form.querySelector("button[type='submit']");
  button.disabled = true;
  try {
    await confirmLoadGeofenceEvent(form.elements.event_id.value, form.elements.decision.value, form.elements.note.value.trim());
    document.querySelector("#geofence-review-dialog").close();
    await loadFleetMap();
  } catch (error) {
    message.textContent = error.message;
  } finally {
    button.disabled = false;
  }
}

async function renderMapboxFleetMap(activeLoads, positions = []) {
  const mapEl = document.querySelector("#fleet-map");
  const messageEl = document.querySelector("#fleet-map-message");

  if (!mapEl || !messageEl) {
    return;
  }

  if (!window.mapboxgl) {
    messageEl.textContent = "Mapbox could not load. Check internet access or CDN permissions.";
    return;
  }

  window.mapboxgl.accessToken = mapboxToken;
  fleetMapInstance?.remove();
  fleetMapInstance = new window.mapboxgl.Map({
    container: mapEl,
    style: "mapbox://styles/mapbox/streets-v12",
    center: [-98.5795, 39.8283],
    projection: "mercator",
    pitch: 0,
    bearing: 0,
    zoom: 3.2,
  });

  fleetMapInstance.addControl(new window.mapboxgl.NavigationControl({ showCompass: false }), "top-right");
  messageEl.textContent = activeLoads.length || positions.length ? "Preparing fleet map..." : "No active loads or live telemetry to show on map.";

  fleetMapInstance.on("load", async () => {
    const mappedRoutes = await buildMappedRoutes(activeLoads.slice(0, 12));
    const bounds = new window.mapboxgl.LngLatBounds();
    const routeFeatures = [];
    let hasBounds = false;

    mappedRoutes.forEach((route, index) => {
      const color = mapRouteColor(index);
      bounds.extend(route.origin.coordinates);
      bounds.extend(route.destination.coordinates);
      hasBounds = true;
      routeFeatures.push({
        type: "Feature",
        properties: {
          color,
          load_no: route.load.load_no,
        },
        geometry: {
          type: "LineString",
          coordinates: [route.origin.coordinates, route.destination.coordinates],
        },
      });
      addFleetMarker(route.origin.coordinates, "Pickup", route, color);
      addFleetMarker(route.destination.coordinates, "Delivery", route, color);
    });

    if (routeFeatures.length) {
      fleetMapInstance.addSource("fleet-routes", {
        type: "geojson",
        data: {
          type: "FeatureCollection",
          features: routeFeatures,
        },
      });
      fleetMapInstance.addLayer({
        id: "fleet-routes-line",
        type: "line",
        source: "fleet-routes",
        paint: {
          "line-color": ["get", "color"],
          "line-width": 3,
          "line-opacity": 0.82,
        },
      });
    }

    positions.slice(0, 50).forEach((position) => {
      const coordinates = [Number(position.longitude), Number(position.latitude)];
      if (!Number.isFinite(coordinates[0]) || !Number.isFinite(coordinates[1])) return;
      bounds.extend(coordinates);
      hasBounds = true;
      addLiveFleetMarker(coordinates, position);
    });

    if (hasBounds) {
      fleetMapInstance.fitBounds(bounds, { padding: 72, maxZoom: 8 });
    }
    const routeLabelText = `${mappedRoutes.length} planned route${mappedRoutes.length === 1 ? "" : "s"}`;
    const liveLabelText = `${positions.length} live signal${positions.length === 1 ? "" : "s"}`;
    messageEl.textContent = hasBounds ? `${routeLabelText} · ${liveLabelText}` : "Add origin/destination or telematics snapshots to draw the map.";
  });
}

async function buildMappedRoutes(loads) {
  const mappedRoutes = [];

  for (const load of loads) {
    if (!load.origin || !load.destination) {
      continue;
    }

    const [origin, destination] = await Promise.all([geocodeLocation(load.origin), geocodeLocation(load.destination)]);

    if (origin && destination) {
      mappedRoutes.push({ load, origin, destination });
    }
  }

  return mappedRoutes;
}

async function geocodeLocation(query) {
  const normalizedQuery = String(query || "").trim();

  if (!normalizedQuery) {
    return null;
  }

  const cacheKey = `cyrra-mapbox-geocode:${normalizedQuery.toLowerCase()}`;
  const cached = localStorage.getItem(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  const url = `https://api.mapbox.com/geocoding/v5/mapbox.places/${encodeURIComponent(normalizedQuery)}.json?country=US&limit=1&access_token=${mapboxToken}`;
  const response = await fetch(url);

  if (!response.ok) {
    return null;
  }

  const payload = await response.json();
  const feature = payload.features?.[0];

  if (!feature?.center) {
    return null;
  }

  const result = {
    label: feature.place_name || normalizedQuery,
    coordinates: feature.center,
  };

  localStorage.setItem(cacheKey, JSON.stringify(result));
  return result;
}

function addFleetMarker(coordinates, label, route, color) {
  const markerEl = document.createElement("div");
  markerEl.className = "mapbox-route-marker";
  markerEl.style.setProperty("--marker-color", color);
  markerEl.textContent = label === "Pickup" ? "P" : "D";

  const popup = new window.mapboxgl.Popup({ offset: 18 }).setHTML(`
    <strong>${escapeHtml(route.load.load_no)}</strong>
    <p>${escapeHtml(label)}: ${escapeHtml(label === "Pickup" ? route.origin.label : route.destination.label)}</p>
    <small>${escapeHtml(route.load.drivers?.name || "No driver")} · ${escapeHtml(route.load.trucks?.unit_no || "No truck")}</small>
  `);

  new window.mapboxgl.Marker(markerEl).setLngLat(coordinates).setPopup(popup).addTo(fleetMapInstance);
}

function mapRouteColor(index) {
  const colors = ["#2563eb", "#16a34a", "#f59e0b", "#dc2626", "#7c3aed", "#0891b2"];

  return colors[index % colors.length];
}

function addLiveFleetMarker(coordinates, position) {
  const freshness = telematicsFreshness(position, new Date().toISOString());
  const markerEl = document.createElement("div");
  markerEl.className = "mapbox-live-marker";
  markerEl.dataset.state = telemetryStateTone(freshness.state);
  markerEl.textContent = "T";

  const popup = new window.mapboxgl.Popup({ offset: 18 }).setHTML(`
    <strong>${escapeHtml(position.trucks?.unit_no ? `Truck ${position.trucks.unit_no}` : position.external_vehicle_id || "Vehicle")}</strong>
    <p>${escapeHtml(formatTelematicsAge(position))}</p>
    <small>${escapeHtml(position.provider)} · ${escapeHtml(position.drivers?.name || "No driver")} · ${Number(position.speed_mph || 0).toFixed(0)} mph</small>
  `);

  new window.mapboxgl.Marker(markerEl).setLngLat(coordinates).setPopup(popup).addTo(fleetMapInstance);
}

function renderFleetAssetColumn(title, rows, labelKey, telemetryKey, telemetry = fleetTelemetryData) {
  return `
    <article class="panel fleet-asset-panel">
      <div class="panel-header">
        <div>
          <h2>${title}</h2>
          <p>${rows.length} record${rows.length === 1 ? "" : "s"}</p>
        </div>
      </div>
      <div class="fleet-asset-list">
        ${
          rows.length
            ? rows
                .slice(0, 8)
                .map((row) => {
                  const position = latestTelemetryRecord(telemetry.positions, telemetryKey, row.id);
                  const freshness = telematicsFreshness(position, new Date().toISOString());
                  const hos = telemetryKey === "driver_id" ? hosRiskState(latestTelemetryRecord(telemetry.hos, "driver_id", row.id)) : null;
                  return `
                    <div>
                      <strong>${escapeHtml(row[labelKey] || "-")}</strong>
                      <span>${formatStatus(row.status)} · ${escapeHtml(freshness.label)}${hos ? ` · ${escapeHtml(hos.label)}` : ""}</span>
                    </div>
                  `;
                })
                .join("")
            : '<div class="empty-state compact-empty">No records.</div>'
        }
      </div>
    </article>
  `;
}

function renderDocumentCenter() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Back office", title: "Documents" })}
    <section class="content-grid customer-grid">
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2>Document Center</h2>
            <p>All uploaded document records across loads, equipment, customers and billing.</p>
          </div>
          <button id="reload-document-center" type="button">Refresh</button>
        </div>
        <div id="document-center-table" class="data-table"></div>
      </article>
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2>Add Document</h2>
            <p>Uploads a private PDF or image (maximum 15 MB).</p>
          </div>
        </div>
        <form id="document-center-form" class="record-form">
          <label>
            <span>Entity type</span>
            <select name="entity_type">
              <option value="load">Load</option>
              <option value="driver">Driver</option>
              <option value="truck">Truck</option>
              <option value="trailer">Trailer</option>
              <option value="customer">Customer</option>
              <option value="carrier">Carrier</option>
              <option value="invoice">Invoice</option>
              <option value="maintenance">Maintenance</option>
            </select>
          </label>
          <label>
            <span>Related record</span>
            <select name="entity_id" id="document-entity-id" required>
              <option value="">Loading records...</option>
            </select>
          </label>
          <label>
            <span>Document type</span>
            <select name="doc_type" required>
              <option value="rate_con">Rate confirmation</option>
              <option value="bol">BOL</option>
              <option value="pod">POD</option>
              <option value="fuel">Fuel</option>
              <option value="lumper">Lumper</option>
              <option value="invoice">Invoice</option>
              <option value="insurance">Insurance</option>
              <option value="registration">Registration</option>
              <option value="cdl">CDL</option>
              <option value="medical">Medical certificate</option>
              <option value="maintenance">Maintenance</option>
              <option value="photo">Photo</option>
              <option value="other">Other</option>
            </select>
          </label>
          <label>
            <span>File</span>
            <input name="file" type="file" accept="application/pdf,image/jpeg,image/png" required />
          </label>
          <label>
            <span>Status</span>
            <select name="status">
              <option value="pending">Pending</option>
              <option value="approved">Approved</option>
              <option value="needs_review">Needs review</option>
              <option value="rejected">Rejected</option>
            </select>
          </label>
          <button type="submit">Save Document</button>
          <p id="document-center-message" class="form-message"></p>
        </form>
      </article>
    </section>
  `;

  document.querySelector("#reload-document-center").addEventListener("click", loadDocumentCenter);
  document.querySelector("#document-center-form").addEventListener("submit", handleCreateCenterDocument);
  document.querySelector('#document-center-form [name="entity_type"]').addEventListener("change", loadDocumentEntityOptions);
  configureCrudFormAccess("#document-center-form", "manage_documents", "Add Document", resetCenterDocumentForm);
  loadDocumentEntityOptions();
  loadDocumentCenter();
}

async function loadDocumentEntityOptions() {
  const form = document.querySelector("#document-center-form");
  const select = document.querySelector("#document-entity-id");
  if (!form || !select) return;
  const entityType = form.elements.entity_type.value;
  select.innerHTML = '<option value="">Loading records...</option>';

  const loaders = {
    load: async () => (await listLoads()).map((item) => ({ id: item.id, label: `${item.load_no} · ${routeLabel(item)}` })),
    driver: async () => (await listDrivers()).map((item) => ({ id: item.id, label: item.name })),
    truck: async () => (await listTrucks()).map((item) => ({ id: item.id, label: item.unit_no })),
    trailer: async () => (await listTrailers()).map((item) => ({ id: item.id, label: item.unit_no })),
    customer: async () => (await listCustomers()).map((item) => ({ id: item.id, label: item.name })),
    carrier: async () => (await listCarriers()).map((item) => ({ id: item.id, label: item.name })),
    invoice: async () => (await listInvoices()).map((item) => ({ id: item.id, label: item.invoice_no })),
    maintenance: async () => (await listMaintenance()).map((item) => ({ id: item.id, label: `${item.service_type} · ${formatStatus(item.equipment_type)}` })),
  };

  try {
    const records = loaders[entityType] ? await loaders[entityType]() : [];
    select.innerHTML = records.length
      ? ['<option value="">Select a record</option>', ...records.map((record) => `<option value="${record.id}">${escapeHtml(record.label)}</option>`)].join("")
      : '<option value="">No available records</option>';
  } catch (error) {
    select.innerHTML = `<option value="">${escapeHtml(error.message)}</option>`;
  }
}

function resetCenterDocumentForm() {
  const form = document.querySelector("#document-center-form");
  form?.reset();
  loadDocumentEntityOptions();
  if (form) closeFormDialog(form);
}

async function loadDocumentCenter() {
  const table = document.querySelector("#document-center-table");
  table.innerHTML = '<div class="empty-state">Loading documents...</div>';

  try {
    const documents = await listDocuments();

    if (!documents.length) {
      table.innerHTML = '<div class="empty-state">No documents yet.</div>';
      return;
    }

    table.innerHTML = `
      <div class="data-row document-row data-row-head">
        <span>Type</span>
        <span>Entity</span>
        <span>Status</span>
        <span>Created</span>
        <span></span>
      </div>
      ${documents
        .map(
          (document) => `
            <div class="data-row document-row">
              <strong>${formatStatus(document.doc_type)}</strong>
              <span>${formatStatus(document.entity_type)}</span>
              <span>${formatStatus(document.status)}</span>
              <span>${formatDateTime(document.created_at)}</span>
              <span class="document-actions">${
                document.signed_url
                  ? `<a class="row-action" href="${escapeAttribute(document.signed_url)}" target="_blank" rel="noreferrer">Open</a>`
                  : "<span>-</span>"
              }<button class="delete-document danger-button" type="button" data-document-id="${document.id}">Delete</button></span>
            </div>
          `
        )
        .join("")}
    `;
    bindDocumentDeleteButtons(loadDocumentCenter);
  } catch (error) {
    table.innerHTML = `<div class="empty-state">Documents table is not ready: ${escapeHtml(error.message)}</div>`;
  }
}

async function handleCreateCenterDocument(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const message = document.querySelector("#document-center-message");
  const formData = new FormData(form);

  message.textContent = "Uploading document...";

  try {
    await uploadTmsDocument({
      entityType: formData.get("entity_type"),
      entityId: formData.get("entity_id"),
      docType: formData.get("doc_type"),
      status: formData.get("status"),
      file: formData.get("file"),
      uploadedBy: currentSession?.user?.id || null,
    });

    resetCenterDocumentForm();
    message.textContent = "Document uploaded securely.";
    await loadDocumentCenter();
  } catch (error) {
    message.textContent = error.message;
  }
}

const dispatchStatuses = [
  "new",
  "booked",
  "assigned",
  "dispatched",
  "en_route",
  "arrived_pickup",
  "loading",
  "picked_up",
  "in_transit",
  "arrived_delivery",
  "unloading",
  "delivered",
  "exception",
  "cancelled",
];

const loadBillingStatuses = ["not_ready", "ready_to_bill", "drafted", "invoiced", "partial", "paid", "void"];

const dispatchTransitions = {
  new: ["new", "booked", "cancelled"],
  booked: ["booked", "new", "assigned", "cancelled"],
  assigned: ["assigned", "booked", "dispatched", "cancelled"],
  dispatched: ["dispatched", "assigned", "en_route", "cancelled", "exception"],
  en_route: ["en_route", "dispatched", "arrived_pickup", "cancelled", "exception"],
  arrived_pickup: ["arrived_pickup", "en_route", "loading", "cancelled", "exception"],
  loading: ["loading", "arrived_pickup", "picked_up", "cancelled", "exception"],
  picked_up: ["picked_up", "loading", "in_transit", "cancelled", "exception"],
  in_transit: ["in_transit", "picked_up", "arrived_delivery", "cancelled", "exception"],
  arrived_delivery: ["arrived_delivery", "in_transit", "unloading", "cancelled", "exception"],
  unloading: ["unloading", "arrived_delivery", "delivered", "cancelled", "exception"],
  delivered: ["delivered", "unloading", "in_transit"],
  exception: ["exception", "dispatched", "en_route", "in_transit", "cancelled"],
  cancelled: ["cancelled", "new"],
};

let dispatchData = { loads: [], drivers: [], trucks: [], trailers: [], maintenance: [], documents: [], telemetry: { positions: [], hos: [], unavailableMessage: "" }, conflicts: {} };
const dispatchViewState = {
  mode: "board",
  calendarMode: "week",
  anchorDate: localDateIso(),
  groupBy: "load",
  savedView: "all",
  customerId: "",
  status: "",
  driverId: "",
  truckId: "",
};

const dispatchSavedViews = [
  ["all", "All Loads"],
  ["my_loads", "My Loads"],
  ["today_pickup", "Today Pickup"],
  ["missing_pod", "Missing POD"],
  ["unassigned", "Unassigned"],
];

function renderDispatchBoard() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Phase 1", title: "Dispatch Board" })}
    <section class="panel">
      <div class="panel-header">
        <div>
          <h2>Dispatch Workspace</h2>
          <p>Board and calendar share saved views, filters and operational warnings.</p>
        </div>
        <div class="table-tools">
          <div class="dispatch-view-switch" role="group" aria-label="Dispatch view">
            <button type="button" data-dispatch-mode="board">Board</button>
            <button type="button" data-dispatch-mode="calendar">Calendar</button>
          </div>
          <button id="reload-dispatch" type="button">Refresh</button>
        </div>
      </div>
      <div id="dispatch-controls"></div>
      <div id="dispatch-workspace"><div class="empty-state">Loading dispatch workspace...</div></div>
    </section>
    <dialog id="dispatch-assignment-dialog" class="crud-dialog dispatch-assignment-dialog">
      <article class="panel">
        <div class="panel-header">
          <div><h2>Assign Resources</h2><p id="dispatch-assignment-summary"></p></div>
          <button class="dialog-close" id="close-dispatch-assignment" type="button" aria-label="Close">×</button>
        </div>
        <form id="dispatch-assignment-form" class="record-form">
          <input type="hidden" name="load_id">
          <label><span>Driver</span><select name="driver_id"></select></label>
          <label><span>Truck</span><select name="truck_id"></select></label>
          <label><span>Trailer (optional)</span><select name="trailer_id"></select></label>
          <div class="form-actions"><button type="submit">Save Assignment</button></div>
          <p id="dispatch-assignment-message" class="form-message"></p>
        </form>
      </article>
    </dialog>
    <dialog id="dispatch-quick-dialog" class="crud-dialog dispatch-quick-dialog">
      <article class="panel">
        <div class="panel-header">
          <div><h2 id="dispatch-quick-title">Load</h2><p id="dispatch-quick-route"></p></div>
          <button class="dialog-close" id="close-dispatch-quick" type="button" aria-label="Close">×</button>
        </div>
        <div id="dispatch-quick-content" class="dispatch-quick-content"></div>
        <form id="dispatch-quick-message-form" class="record-form">
          <input type="hidden" name="load_id">
          <label><span>Internal dispatch note</span><textarea name="body" maxlength="4000" placeholder="Add a quick update for the team..."></textarea></label>
          <div class="form-actions"><button type="submit">Send Note</button></div>
          <p id="dispatch-quick-message" class="form-message"></p>
        </form>
      </article>
    </dialog>
  `;

  document.querySelector("#reload-dispatch").addEventListener("click", loadDispatchBoard);
  document.querySelectorAll("[data-dispatch-mode]").forEach((button) => button.addEventListener("click", () => {
    dispatchViewState.mode = button.dataset.dispatchMode;
    renderDispatchWorkspace();
  }));
  document.querySelector("#close-dispatch-assignment").addEventListener("click", () => document.querySelector("#dispatch-assignment-dialog").close());
  document.querySelector("#close-dispatch-quick").addEventListener("click", () => document.querySelector("#dispatch-quick-dialog").close());
  document.querySelector("#dispatch-assignment-form").addEventListener("submit", handleDispatchAssignment);
  document.querySelector("#dispatch-quick-message-form").addEventListener("submit", handleDispatchQuickMessage);
  ["driver_id", "truck_id", "trailer_id"].forEach((name) => {
    document.querySelector(`#dispatch-assignment-form [name="${name}"]`).addEventListener("change", updateDispatchAssignmentNotice);
  });
  loadDispatchBoard();
}

async function loadDispatchBoard() {
  const workspace = document.querySelector("#dispatch-workspace");
  workspace.innerHTML = '<div class="empty-state">Loading dispatch workspace...</div>';

  try {
    const [loads, drivers, trucks, trailers, maintenance, documents, telemetry] = await Promise.all([
      listLoads(), listDrivers(), listTrucks(), listTrailers(), listMaintenance(), listDispatchDocumentStates(), loadTelematicsFoundation(),
    ]);
    dispatchData = { loads, drivers, trucks, trailers, maintenance, documents, telemetry, conflicts: dispatchAssignmentConflicts(loads) };
    renderDispatchWorkspace();
  } catch (error) {
    workspace.innerHTML = `<div class="empty-state">Dispatch workspace is not ready: ${escapeHtml(error.message)}</div>`;
  }
}

function renderDispatchWorkspace() {
  const controls = document.querySelector("#dispatch-controls");
  const workspace = document.querySelector("#dispatch-workspace");
  if (!controls || !workspace) return;
  document.querySelectorAll("[data-dispatch-mode]").forEach((button) => button.classList.toggle("active", button.dataset.dispatchMode === dispatchViewState.mode));
  controls.innerHTML = renderDispatchControls();
  bindDispatchControls();
  const loads = filteredDispatchLoads();
  workspace.innerHTML = dispatchViewState.mode === "calendar" ? renderDispatchCalendar(loads) : renderDispatchColumns(loads);
  bindDispatchWorkspaceActions();
}

function renderDispatchControls() {
  const customers = [...new Map(dispatchData.loads.filter((load) => load.customer_id).map((load) => [load.customer_id, load.customers?.name || "Customer"])).entries()];
  return `
    <div class="dispatch-saved-views">
      ${dispatchSavedViews.map(([value, label]) => `<button type="button" data-dispatch-saved-view="${value}" class="${dispatchViewState.savedView === value ? "active" : ""}">${label}</button>`).join("")}
    </div>
    <div class="dispatch-filters">
      <label><span>Customer</span><select data-dispatch-filter="customerId"><option value="">All customers</option>${customers.map(([id, name]) => `<option value="${id}" ${id === dispatchViewState.customerId ? "selected" : ""}>${escapeHtml(name)}</option>`).join("")}</select></label>
      <label><span>Status</span><select data-dispatch-filter="status"><option value="">All statuses</option>${dispatchStatuses.map((status) => `<option value="${status}" ${status === dispatchViewState.status ? "selected" : ""}>${formatStatus(status)}</option>`).join("")}</select></label>
      <label><span>Driver</span><select data-dispatch-filter="driverId"><option value="">All drivers</option>${dispatchData.drivers.map((driver) => `<option value="${driver.id}" ${driver.id === dispatchViewState.driverId ? "selected" : ""}>${escapeHtml(driver.name)}</option>`).join("")}</select></label>
      <label><span>Truck</span><select data-dispatch-filter="truckId"><option value="">All trucks</option>${dispatchData.trucks.map((truck) => `<option value="${truck.id}" ${truck.id === dispatchViewState.truckId ? "selected" : ""}>${escapeHtml(truck.unit_no)}</option>`).join("")}</select></label>
      <button type="button" id="clear-dispatch-filters">Clear</button>
    </div>
    ${dispatchViewState.mode === "calendar" ? renderDispatchCalendarToolbar() : ""}
  `;
}

function bindDispatchControls() {
  document.querySelectorAll("[data-dispatch-saved-view]").forEach((button) => button.addEventListener("click", () => {
    dispatchViewState.savedView = button.dataset.dispatchSavedView;
    renderDispatchWorkspace();
  }));
  document.querySelectorAll("[data-dispatch-filter]").forEach((select) => select.addEventListener("change", () => {
    dispatchViewState[select.dataset.dispatchFilter] = select.value;
    renderDispatchWorkspace();
  }));
  document.querySelector("#clear-dispatch-filters")?.addEventListener("click", () => {
    Object.assign(dispatchViewState, { savedView: "all", customerId: "", status: "", driverId: "", truckId: "" });
    renderDispatchWorkspace();
  });
  document.querySelectorAll("[data-calendar-mode]").forEach((button) => button.addEventListener("click", () => {
    dispatchViewState.calendarMode = button.dataset.calendarMode;
    renderDispatchWorkspace();
  }));
  document.querySelector("#dispatch-calendar-group")?.addEventListener("change", (event) => {
    dispatchViewState.groupBy = event.currentTarget.value;
    renderDispatchWorkspace();
  });
  document.querySelectorAll("[data-calendar-move]").forEach((button) => button.addEventListener("click", () => {
    const direction = Number(button.dataset.calendarMove);
    const offset = dispatchViewState.calendarMode === "week" ? direction * 7 : direction;
    dispatchViewState.anchorDate = addDaysIso(dispatchViewState.anchorDate, offset);
    renderDispatchWorkspace();
  }));
  document.querySelector("#dispatch-calendar-today")?.addEventListener("click", () => {
    dispatchViewState.anchorDate = localDateIso();
    renderDispatchWorkspace();
  });
}

function filteredDispatchLoads() {
  return dispatchData.loads.filter((load) => matchesDispatchSavedView(load, dispatchViewState.savedView, {
    userId: currentSession?.user?.id,
    today: localDateIso(),
    documents: dispatchData.documents,
  })).filter((load) => !dispatchViewState.customerId || load.customer_id === dispatchViewState.customerId)
    .filter((load) => !dispatchViewState.status || operationalLoadStatus(load) === dispatchViewState.status)
    .filter((load) => !dispatchViewState.driverId || load.driver_id === dispatchViewState.driverId)
    .filter((load) => !dispatchViewState.truckId || load.truck_id === dispatchViewState.truckId);
}

function renderDispatchColumns(loads) {
  const grouped = groupLoadsByStatus(loads);
  const dragEnabled = can("manage_operations");
  return `${dragEnabled ? '<p class="dispatch-drag-help">Drag a load card to an allowed status column. The status menu remains available for keyboard use.</p>' : ""}
    <div id="dispatch-drag-message" class="dispatch-drag-message" role="status" aria-live="polite"></div>
    <div id="dispatch-board" class="dispatch-board">${dispatchStatuses.map((status) => `
    <section class="dispatch-column" data-dispatch-drop-status="${status}">
      <header><h3>${formatStatus(status)}</h3><span>${grouped[status].length}</span></header>
      <div class="dispatch-cards">${grouped[status].length ? grouped[status].map(renderDispatchCard).join("") : '<div class="empty-state compact-empty">No loads</div>'}</div>
    </section>`).join("")}</div>`;
}

function addDaysIso(value, offset) {
  const date = new Date(`${value}T12:00:00`);
  date.setDate(date.getDate() + offset);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function startOfDispatchWeek(value) {
  const date = new Date(`${value}T12:00:00`);
  const mondayOffset = date.getDay() === 0 ? -6 : 1 - date.getDay();
  return addDaysIso(value, mondayOffset);
}

function renderDispatchCalendarToolbar() {
  return `
    <div class="dispatch-calendar-toolbar">
      <div class="table-tools">
        <button type="button" data-calendar-move="-1">Previous</button>
        <button type="button" id="dispatch-calendar-today">Today</button>
        <button type="button" data-calendar-move="1">Next</button>
      </div>
      <strong>${dispatchCalendarRangeLabel()}</strong>
      <div class="table-tools">
        <label><span>Group by</span><select id="dispatch-calendar-group">
          ${[["load", "Load"], ["driver", "Driver"], ["truck", "Truck"], ["customer", "Customer"]].map(([value, label]) => `<option value="${value}" ${dispatchViewState.groupBy === value ? "selected" : ""}>${label}</option>`).join("")}
        </select></label>
        <div class="dispatch-view-switch" role="group" aria-label="Calendar range">
          <button type="button" data-calendar-mode="day" class="${dispatchViewState.calendarMode === "day" ? "active" : ""}">Day</button>
          <button type="button" data-calendar-mode="week" class="${dispatchViewState.calendarMode === "week" ? "active" : ""}">Week</button>
        </div>
      </div>
    </div>`;
}

function dispatchCalendarDates() {
  const start = dispatchViewState.calendarMode === "week" ? startOfDispatchWeek(dispatchViewState.anchorDate) : dispatchViewState.anchorDate;
  return Array.from({ length: dispatchViewState.calendarMode === "week" ? 7 : 1 }, (_, index) => addDaysIso(start, index));
}

function dispatchCalendarRangeLabel() {
  const dates = dispatchCalendarDates();
  return dates.length === 1 ? formatDate(dates[0]) : `${formatDate(dates[0])} – ${formatDate(dates.at(-1))}`;
}

function dispatchLoadDate(load) {
  const pickup = orderedLoadStops(load).find((stop) => stop.type === "pickup" && stop.appointment_from)
    || orderedLoadStops(load).find((stop) => stop.appointment_from);
  return pickup?.appointment_from?.slice(0, 10) || load.pickup_date || null;
}

function dispatchCalendarGroupLabel(load) {
  if (dispatchViewState.groupBy === "driver") return load.drivers?.name || "Unassigned driver";
  if (dispatchViewState.groupBy === "truck") return load.trucks?.unit_no ? `Truck ${load.trucks.unit_no}` : "Unassigned truck";
  if (dispatchViewState.groupBy === "customer") return load.customers?.name || "No customer";
  return `Load ${load.load_no}`;
}

function renderDispatchCalendar(loads) {
  const dates = dispatchCalendarDates();
  return `<div class="dispatch-calendar ${dates.length === 1 ? "dispatch-calendar-day-view" : ""}">
    ${dates.map((date) => {
      const dayLoads = loads.filter((load) => dispatchLoadDate(load) === date);
      return `<section class="dispatch-calendar-day ${date === localDateIso() ? "today" : ""}">
        <header><span>${new Intl.DateTimeFormat("en-US", { weekday: "short" }).format(new Date(`${date}T12:00:00`))}</span><strong>${formatDate(date)}</strong><em>${dayLoads.length}</em></header>
        <div class="dispatch-calendar-events">${dayLoads.length ? dayLoads.map((load) => `
          <button type="button" class="dispatch-calendar-event" data-dispatch-quick="${load.id}">
            <span>${escapeHtml(dispatchCalendarGroupLabel(load))}</span>
            <strong>${escapeHtml(load.load_no)} · ${formatStatus(operationalLoadStatus(load))}</strong>
            <small>${escapeHtml(routeLabel(load))}</small>
            ${renderDispatchWarningBadges(load)}
          </button>`).join("") : '<div class="empty-state compact-empty">No scheduled loads</div>'}</div>
      </section>`;
    }).join("")}
  </div>`;
}

function bindDispatchWorkspaceActions() {
  document.querySelectorAll("[data-dispatch-status]").forEach((select) => select.addEventListener("change", handleDispatchStatusChange));
  document.querySelectorAll("[data-dispatch-assign]").forEach((button) => button.addEventListener("click", openDispatchAssignment));
  document.querySelectorAll("[data-dispatch-quick]").forEach((button) => button.addEventListener("click", openDispatchQuickLoad));
  bindDispatchDragAndDrop();
}

function dispatchDropAllowed(sourceStatus, targetStatus) {
  return sourceStatus !== targetStatus && (dispatchTransitions[sourceStatus] || []).includes(targetStatus);
}

function clearDispatchDragState() {
  document.querySelectorAll(".dispatch-card.dragging").forEach((card) => card.classList.remove("dragging"));
  document.querySelectorAll("[data-dispatch-drop-status]").forEach((column) => column.classList.remove("drag-allowed", "drag-blocked", "drag-over"));
}

function bindDispatchDragAndDrop() {
  if (dispatchViewState.mode !== "board" || !can("manage_operations")) return;
  const cards = document.querySelectorAll(".dispatch-card[draggable='true']");
  const columns = document.querySelectorAll("[data-dispatch-drop-status]");
  cards.forEach((card) => {
    card.addEventListener("dragstart", (event) => {
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", card.dataset.dispatchDragId);
      card.classList.add("dragging");
      const sourceStatus = card.dataset.dispatchDragStatus;
      columns.forEach((column) => column.classList.add(dispatchDropAllowed(sourceStatus, column.dataset.dispatchDropStatus) ? "drag-allowed" : "drag-blocked"));
    });
    card.addEventListener("dragend", clearDispatchDragState);
  });
  columns.forEach((column) => {
    column.addEventListener("dragover", (event) => {
      const card = document.querySelector(".dispatch-card.dragging");
      if (!card || !dispatchDropAllowed(card.dataset.dispatchDragStatus, column.dataset.dispatchDropStatus)) return;
      event.preventDefault();
      event.dataTransfer.dropEffect = "move";
      column.classList.add("drag-over");
    });
    column.addEventListener("dragleave", (event) => {
      if (!column.contains(event.relatedTarget)) column.classList.remove("drag-over");
    });
    column.addEventListener("drop", handleDispatchDrop);
  });
}

async function handleDispatchDrop(event) {
  event.preventDefault();
  const column = event.currentTarget;
  const loadId = event.dataTransfer.getData("text/plain");
  const load = dispatchData.loads.find((item) => item.id === loadId);
  const targetStatus = column.dataset.dispatchDropStatus;
  const message = document.querySelector("#dispatch-drag-message");
  if (!load || !dispatchDropAllowed(operationalLoadStatus(load), targetStatus)) {
    clearDispatchDragState();
    return;
  }
  message.dataset.state = "working";
  message.textContent = `Moving load ${load.load_no} to ${formatStatus(targetStatus)}...`;
  document.querySelectorAll(".dispatch-card").forEach((card) => { card.draggable = false; });
  try {
    await updateLoadStatus(loadId, targetStatus);
    message.dataset.state = "ok";
    message.textContent = `Load ${load.load_no} moved to ${formatStatus(targetStatus)}.`;
    await loadDispatchBoard();
  } catch (error) {
    clearDispatchDragState();
    message.dataset.state = "error";
    message.textContent = `Move failed: ${error.message}`;
    document.querySelectorAll(".dispatch-card").forEach((card) => { card.draggable = true; });
  }
}

function dispatchWarnings(load) {
  const warnings = [];
  if (!load.driver_id || !load.truck_id) warnings.push({ tone: "warning", label: "Unassigned" });
  const conflicts = dispatchData.conflicts[load.id] || [];
  if (conflicts.length) warnings.push({ tone: "danger", label: `Conflict: ${conflicts.map((item) => item.loadNo).join(", ")}` });
  const driver = dispatchData.drivers.find((item) => item.id === load.driver_id);
  const truck = dispatchData.trucks.find((item) => item.id === load.truck_id);
  for (const [resource, type] of [[driver, "driver"], [truck, "truck"]]) {
    if (!resource) continue;
    const compliance = dispatchResourceCompliance(resource, type);
    if (compliance.state !== "current") warnings.push({ tone: compliance.state === "expired" ? "danger" : "warning", label: dispatchComplianceBadgeLabel(type, compliance.state) });
  }
  const maintenanceHold = dispatchData.maintenance.some((record) => record.equipment_type === "truck" && record.equipment_id === load.truck_id
    && record.out_of_service && ["open", "scheduled"].includes(record.status));
  if (maintenanceHold) warnings.push({ tone: "danger", label: "Maintenance hold" });
  const telemetry = telemetryForLoad(load, dispatchData.telemetry);
  const freshness = telematicsFreshness(telemetry.position, new Date().toISOString());
  if (["stale", "offline"].includes(freshness.state)) warnings.push({ tone: freshness.state === "offline" ? "danger" : "warning", label: freshness.label });
  const hos = hosRiskState(telemetry.hos);
  if (["violation", "warning"].includes(hos.state)) warnings.push({ tone: hos.state === "violation" ? "danger" : "warning", label: hos.label });
  if (matchesDispatchSavedView(load, "missing_pod", { documents: dispatchData.documents })) warnings.push({ tone: "danger", label: "Missing POD" });
  return warnings;
}

function renderDispatchWarningBadges(load) {
  const warnings = dispatchWarnings(load);
  if (!warnings.length) return "";
  return `<span class="dispatch-warning-badges">${warnings.map((warning) => `<span data-tone="${warning.tone}">${escapeHtml(warning.label)}</span>`).join("")}</span>`;
}

async function openDispatchQuickLoad(event) {
  const load = dispatchData.loads.find((item) => item.id === event.currentTarget.dataset.dispatchQuick);
  if (!load) return;
  const dialog = document.querySelector("#dispatch-quick-dialog");
  const form = document.querySelector("#dispatch-quick-message-form");
  form.reset();
  form.elements.load_id.value = load.id;
  document.querySelector("#dispatch-quick-title").textContent = `Load ${load.load_no}`;
  document.querySelector("#dispatch-quick-route").textContent = routeLabel(load);
  document.querySelector("#dispatch-quick-message").textContent = "";
  const telemetry = telemetryForLoad(load, dispatchData.telemetry);
  const freshness = telematicsFreshness(telemetry.position, new Date().toISOString());
  const hos = hosRiskState(telemetry.hos);
  document.querySelector("#dispatch-quick-content").innerHTML = `
    ${renderDispatchWarningBadges(load)}
    <div class="dispatch-quick-grid">
      <div><span>Operational</span><strong>${formatStatus(operationalLoadStatus(load))}</strong></div>
      <div><span>Customer</span><strong>${escapeHtml(load.customers?.name || "-")}</strong></div>
      <div><span>Driver</span><strong>${escapeHtml(load.drivers?.name || "Unassigned")}</strong></div>
      <div><span>Truck</span><strong>${escapeHtml(load.trucks?.unit_no || "Unassigned")}</strong></div>
      <div><span>Live signal</span><strong>${escapeHtml(freshness.ageMinutes === null ? freshness.label : `${freshness.label} · ${freshness.ageMinutes} min`)}</strong></div>
      <div><span>HOS</span><strong>${escapeHtml(hos.label)}${hos.minimumSeconds !== null ? ` · ${formatHosRemaining(hos.minimumSeconds)}` : ""}</strong></div>
    </div>
    <h3>Schedule</h3>
    <div class="dispatch-quick-stops">${orderedLoadStops(load).length ? orderedLoadStops(load).map((stop) => `<div><strong>${formatStatus(stop.type)} · ${escapeHtml(stop.facility || stop.address || "Stop")}</strong><span>${formatStopWindow(stop, load.pickup_date)}</span></div>`).join("") : `<div><span>${formatDate(load.pickup_date)} → ${formatDate(load.delivery_date)}</span></div>`}</div>
    <div id="dispatch-quick-messages" class="dispatch-quick-messages"><span>Loading recent notes...</span></div>`;
  dialog.showModal();
  try {
    const messages = await listLoadMessages(load.id);
    document.querySelector("#dispatch-quick-messages").innerHTML = messages.length
      ? messages.slice(-3).reverse().map((message) => `<div><strong>${message.sender_id === currentSession?.user?.id ? "You" : "Team"}</strong><span>${escapeHtml(message.body)}</span></div>`).join("")
      : "<span>No internal notes yet.</span>";
  } catch (error) {
    document.querySelector("#dispatch-quick-messages").textContent = error.message;
  }
}

async function handleDispatchQuickMessage(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = document.querySelector("#dispatch-quick-message");
  const submit = form.querySelector("button[type='submit']");
  submit.disabled = true;
  try {
    await createLoadMessage(form.elements.load_id.value, currentSession?.user?.id || null, validateMessageBody(form.elements.body.value));
    message.textContent = "Dispatch note sent.";
    form.elements.body.value = "";
    const messages = await listLoadMessages(form.elements.load_id.value);
    document.querySelector("#dispatch-quick-messages").innerHTML = messages.slice(-3).reverse().map((item) => `<div><strong>${item.sender_id === currentSession?.user?.id ? "You" : "Team"}</strong><span>${escapeHtml(item.body)}</span></div>`).join("");
  } catch (error) {
    message.textContent = error.message;
  } finally {
    submit.disabled = false;
  }
}

function groupLoadsByStatus(loads) {
  const grouped = Object.fromEntries(dispatchStatuses.map((status) => [status, []]));

  loads.forEach((load) => {
    const status = dispatchStatuses.includes(operationalLoadStatus(load)) ? operationalLoadStatus(load) : "new";
    grouped[status].push(load);
  });

  return grouped;
}

function renderDispatchCard(load) {
  const currentStatus = operationalLoadStatus(load);
  const allowedStatuses = dispatchTransitions[currentStatus] || [currentStatus];
  return `
    <article class="dispatch-card" data-dispatch-drag-id="${load.id}" data-dispatch-drag-status="${currentStatus}" ${can("manage_operations") ? 'draggable="true"' : ""}>
      <div class="dispatch-card-head">
        <strong>${escapeHtml(load.load_no)}</strong>
        <span>${loadPickupWindow(load)}</span>
      </div>
      <p>${escapeHtml(load.customers?.name || "No customer")}</p>
      <small>${escapeHtml(routeLabel(load))}</small>
      <div class="dispatch-meta">
        <span>${escapeHtml(load.carriers?.name || "Own fleet")}</span>
        <span>${escapeHtml(load.drivers?.name || "No driver")}</span>
        <span>${escapeHtml(load.trucks?.unit_no || "No truck")}</span>
      </div>
      ${renderDispatchWarningBadges(load)}
      <div class="dispatch-card-actions">
      <button type="button" data-dispatch-assign="${load.id}">Assign</button>
      <button type="button" data-dispatch-quick="${load.id}">Quick View</button>
      <select data-dispatch-status="${load.id}" aria-label="Change status for ${escapeHtml(load.load_no)}">
        ${allowedStatuses
          .map((status) => `<option value="${status}" ${status === currentStatus ? "selected" : ""}>${formatStatus(status)}</option>`)
          .join("")}
      </select>
      </div>
    </article>
  `;
}

function dispatchResourceCompliance(item, resourceType) {
  return getDispatchComplianceState(item, resourceType, localDateIso(), localDateIso(30));
}

function dispatchComplianceBadgeLabel(resourceType, complianceState) {
  return `${formatStatus(resourceType)} compliance ${formatStatus(complianceState).toLowerCase()}`;
}

function dispatchComplianceLabel(compliance) {
  if (compliance.state === "current") return "";
  const affected = compliance.details.filter((detail) => detail.state === compliance.state).map((detail) => detail.label).join(", ");
  return ` · ${formatStatus(compliance.state)}: ${affected}`;
}

function resourceOptions(items, selectedId, labelKey, activeStatuses, resourceType) {
  const options = items.filter((item) => activeStatuses.includes(item.status) || item.id === selectedId);
  return `<option value="">Unassigned</option>${options.map((item) => {
    const compliance = dispatchResourceCompliance(item, resourceType);
    return `<option value="${item.id}" ${item.id === selectedId ? "selected" : ""} ${compliance.state === "expired" ? "disabled" : ""}>${escapeHtml(item[labelKey])} · ${formatStatus(item.status)}${escapeHtml(dispatchComplianceLabel(compliance))}</option>`;
  }).join("")}`;
}

function openDispatchAssignment(event) {
  const load = dispatchData.loads.find((item) => item.id === event.currentTarget.dataset.dispatchAssign);
  if (!load) return;

  const dialog = document.querySelector("#dispatch-assignment-dialog");
  const form = document.querySelector("#dispatch-assignment-form");
  form.reset();
  form.elements.load_id.value = load.id;
  form.elements.driver_id.innerHTML = resourceOptions(dispatchData.drivers, load.driver_id, "name", ["available", "on_load"], "driver");
  form.elements.truck_id.innerHTML = resourceOptions(dispatchData.trucks, load.truck_id, "unit_no", ["available", "assigned"], "truck");
  form.elements.trailer_id.innerHTML = resourceOptions(dispatchData.trailers, load.trailer_id, "unit_no", ["available", "assigned"], "trailer");
  document.querySelector("#dispatch-assignment-summary").textContent = `${load.load_no} · ${routeLabel(load)}`;
  updateDispatchAssignmentNotice();
  dialog.showModal();
}

function updateDispatchAssignmentNotice() {
  const form = document.querySelector("#dispatch-assignment-form");
  const message = document.querySelector("#dispatch-assignment-message");
  const submit = form.querySelector("button[type='submit']");
  const selections = [
    ["driver_id", dispatchData.drivers, "driver", "name"],
    ["truck_id", dispatchData.trucks, "truck", "unit_no"],
    ["trailer_id", dispatchData.trailers, "trailer", "unit_no"],
  ].map(([field, items, resourceType, labelKey]) => {
    const item = items.find((candidate) => candidate.id === form.elements[field].value);
    return item ? { item, resourceType, label: item[labelKey], compliance: dispatchResourceCompliance(item, resourceType) } : null;
  }).filter(Boolean);
  const expired = selections.filter((selection) => selection.compliance.state === "expired");
  const warnings = selections.filter((selection) => ["due_soon", "missing"].includes(selection.compliance.state));
  if (expired.length) {
    submit.disabled = true;
    message.dataset.state = "error";
    message.textContent = `Blocked: ${expired.map((selection) => `${selection.label}${dispatchComplianceLabel(selection.compliance)}`).join("; ")}. Select a compliant replacement.`;
    return;
  }
  if (warnings.length) {
    submit.disabled = false;
    message.dataset.state = "warning";
    message.textContent = `Warning: ${warnings.map((selection) => `${selection.label}${dispatchComplianceLabel(selection.compliance)}`).join("; ")}. Assignment remains allowed.`;
    return;
  }
  submit.disabled = false;
  message.dataset.state = "ok";
  message.textContent = "Only operationally available, non-expired resources can be assigned.";
}

async function handleDispatchAssignment(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const submit = form.querySelector("button[type='submit']");
  const message = document.querySelector("#dispatch-assignment-message");
  submit.disabled = true;
  message.textContent = "Checking availability and saving...";

  try {
    await updateLoadAssignment(form.elements.load_id.value, {
      driver_id: form.elements.driver_id.value || null,
      truck_id: form.elements.truck_id.value || null,
      trailer_id: form.elements.trailer_id.value || null,
    });
    document.querySelector("#dispatch-assignment-dialog").close();
    await loadDispatchBoard();
  } catch (error) {
    message.textContent = error.message;
  } finally {
    submit.disabled = false;
  }
}

async function handleDispatchStatusChange(event) {
  const select = event.currentTarget;
  const loadId = select.dataset.dispatchStatus;
  const status = select.value;

  select.disabled = true;

  try {
    await updateLoadStatus(loadId, status);
    await loadDispatchBoard();
  } catch (error) {
    select.disabled = false;
    window.alert(`Status update failed: ${error.message}`);
  }
}

const invoiceStatuses = ["draft", "sent", "partial", "paid", "overdue", "void"];

function renderInvoiceStatusOptions(displayStatus) {
  const lockedCurrent = canManuallySetInvoiceStatus(displayStatus)
    ? ""
    : `<option value="${displayStatus}" selected disabled>${formatStatus(displayStatus)}</option>`;
  return `${lockedCurrent}${manualInvoiceStatusOptions.map((status) => `<option value="${status}" ${status === displayStatus ? "selected" : ""}>${formatStatus(status)}</option>`).join("")}`;
}

const invoiceLineItemTypeOptions = [
  ["linehaul", "Linehaul"],
  ["fsc", "Fuel surcharge"],
  ["detention", "Detention"],
  ["layover", "Layover"],
  ["lumper", "Lumper"],
  ["tonu", "TONU"],
  ["discount", "Discount"],
  ["tax", "Tax"],
  ["credit_memo", "Credit memo"],
  ["custom", "Custom"],
];
let invoiceFormLoadCache = [];

async function renderBilling() {
  await ensureTableSavedViews("invoices");
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Phase 2", title: "Billing / Invoices" })}
    ${renderPaymentDialog()}
    ${renderInvoiceBatchDialog()}
    ${renderCollectionNoteDialog()}
    ${renderAccessorialApprovalDialog()}
    ${renderCreditMemoDialog()}
    <section class="content-grid load-grid">
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2>Invoice List</h2>
            <p>Customer invoices connected to delivered or ready-to-bill loads.</p>
          </div>
          <button id="reload-invoices" type="button">Refresh</button>
        </div>
        <form id="invoice-filters" class="load-filters invoice-filters">
          <label><span>Search</span><input name="search" type="search" placeholder="Invoice, customer, load..."></label>
          <label><span>Status</span><select name="status"><option value="">All statuses</option>${invoiceStatuses.map((status) => `<option value="${status}">${formatStatus(status)}</option>`).join("")}</select></label>
          <label><span>Sort by</span><select name="sort_by"><option value="created_at">Created</option><option value="invoice_no">Invoice #</option><option value="amount">Amount</option><option value="due_date">Due date</option><option value="status">Status</option></select></label>
          <label><span>Direction</span><select name="sort_direction"><option value="desc">Descending</option><option value="asc">Ascending</option></select></label>
          <label><span>Rows</span><select name="page_size"><option value="5">5</option><option value="10">10</option><option value="25">25</option></select></label>
          <button type="reset">Clear</button>
        </form>
        ${renderTableViewControls("invoices")}
        <div id="invoices-table" class="data-table"></div>
        <div id="invoices-pagination" class="table-pagination"></div>
      </article>
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2>Create Invoice</h2>
            <p>Creates invoice line items and accessorials for accounting follow-up.</p>
          </div>
        </div>
        <form id="invoice-form" class="record-form">
          <label>
            <span>Invoice number</span>
            <input name="invoice_no" type="text" placeholder="Auto from numbering settings" />
          </label>
          <label>
            <span>Customer</span>
            <select name="customer_id" id="invoice-customer" required></select>
          </label>
          <label>
            <span>Load</span>
            <select name="load_id" id="invoice-load"></select>
          </label>
          <div class="invoice-line-editor">
            <div class="inline-section-header">
              <div><span>Line items</span><small>Linehaul, FSC, accessorials, discounts, tax or custom charges.</small></div>
              <button id="add-invoice-line-item" type="button">Add line</button>
            </div>
            <div id="invoice-line-items"></div>
            <div class="invoice-line-total"><span>Total</span><strong id="invoice-line-total">$0.00</strong></div>
          </div>
          <label>
            <span>Due date</span>
            <input name="due_date" type="date" />
          </label>
          <label>
            <span>Status</span>
            <select name="status">
              <option value="draft">Draft</option>
              <option value="sent">Sent</option>
              <option value="partial">Partial</option>
              <option value="paid">Paid</option>
              <option value="overdue">Overdue</option>
            </select>
          </label>
          <button type="submit">Save Invoice</button>
          <p id="invoice-message" class="form-message"></p>
        </form>
      </article>
    </section>
    <section id="invoice-batches-panel" class="panel billing-batches-panel">
      <div class="panel-header">
        <div>
          <h2>Billing Batches</h2>
          <p>Group open customer invoices into draft master billing statements.</p>
        </div>
        <div class="table-tools">
          <button id="reload-invoice-batches" type="button">Refresh</button>
          <button id="open-invoice-batch-dialog" class="primary-button" type="button">Create Batch</button>
        </div>
      </div>
      <div id="invoice-batches-table" class="data-table"></div>
    </section>
    <section id="collections-panel" class="panel collections-panel">
      <div class="panel-header">
        <div>
          <h2>Collections</h2>
          <p>Track open balances, reminders and collection follow-ups.</p>
        </div>
        <div class="table-tools">
          <button id="reload-collections" type="button">Refresh</button>
        </div>
      </div>
      <div id="collections-table" class="data-table"></div>
    </section>
    <section id="accounting-export-panel" class="panel accounting-export-panel">
      <div class="panel-header">
        <div>
          <h2>Accounting Export</h2>
          <p>Download finance CSV files for accounting review or external import prep.</p>
        </div>
      </div>
      <div class="accounting-export-actions">
        <button id="export-accounting-all" class="primary-button" type="button">Export All CSV</button>
        <button type="button" data-accounting-export="invoices">Invoices CSV</button>
        <button type="button" data-accounting-export="payments">Payments CSV</button>
        <button type="button" data-accounting-export="batches">Batches CSV</button>
        <button type="button" data-accounting-export="collections">Collections CSV</button>
      </div>
      <p id="accounting-export-message" class="form-message"></p>
    </section>
  `;

  document.querySelector("#reload-invoices").addEventListener("click", loadInvoices);
  document.querySelector("#invoice-filters").addEventListener("input", handleInvoiceFilters);
  document.querySelector("#invoice-filters").addEventListener("reset", () => setTimeout(() => {
    Object.assign(invoiceListState, { search: "", status: "", page: 1, pageSize: 5, sortBy: "created_at", sortDirection: "desc" });
    clearSelectedTableView("invoices");
    loadInvoices();
  }));
  bindTableViewControls("invoices");
  document.querySelector("#reload-invoice-batches").addEventListener("click", loadInvoiceBatches);
  document.querySelector("#reload-collections").addEventListener("click", loadCollections);
  document.querySelector("#open-invoice-batch-dialog").addEventListener("click", openInvoiceBatchDialog);
  document.querySelector("#close-invoice-batch-dialog").addEventListener("click", () => document.querySelector("#invoice-batch-dialog").close());
  document.querySelector("#invoice-batch-form").addEventListener("submit", handleCreateInvoiceBatch);
  document.querySelector("#invoice-batch-customer").addEventListener("change", renderInvoiceBatchCandidateList);
  document.querySelector("#close-collection-note-dialog").addEventListener("click", () => document.querySelector("#collection-note-dialog").close());
  document.querySelector("#close-accessorial-approval-dialog").addEventListener("click", () => document.querySelector("#accessorial-approval-dialog").close());
  document.querySelector("#close-credit-memo-dialog").addEventListener("click", () => document.querySelector("#credit-memo-dialog").close());
  document.querySelector("#collection-note-form").addEventListener("submit", handleCollectionNoteSubmit);
  document.querySelector("#export-accounting-all").addEventListener("click", () => handleAccountingExport("all"));
  document.querySelectorAll("[data-accounting-export]").forEach((button) => {
    button.addEventListener("click", () => handleAccountingExport(button.dataset.accountingExport));
  });
  document.querySelector("#invoice-form").addEventListener("submit", handleCreateInvoice);
  document.querySelector("#add-invoice-line-item").addEventListener("click", () => {
    addInvoiceLineItem({ item_type: "custom", description: "Custom charge", customer_description: "Custom charge", quantity: 1, unit_amount: 0, approved: true });
  });
  document.querySelector("#invoice-load").addEventListener("change", handleInvoiceLoadChange);
  document.querySelector("#payment-form").addEventListener("submit", handlePaymentSubmit);
  document.querySelector("#close-payment-dialog").addEventListener("click", () => document.querySelector("#payment-dialog").close());
  configureCrudFormAccess("#invoice-form", "manage_finance", "Create Invoice", resetInvoiceForm);
  loadInvoiceFormOptions();
  loadInvoices();
  loadInvoiceBatches();
  loadCollections();
}

function renderPaymentDialog() {
  return `<dialog id="payment-dialog" class="crud-dialog payment-dialog"><article class="panel"><div class="panel-header"><div><h2>Record Payment</h2><p id="payment-summary"></p></div><button id="close-payment-dialog" class="dialog-close" type="button">×</button></div><form id="payment-form" class="record-form"><input name="invoice_id" type="hidden"><label><span>Amount</span><input name="amount" type="number" min="0.01" step="0.01" required></label><label><span>Payment date</span><input name="date" type="date" required></label><label><span>Method</span><select name="method"><option>ACH</option><option>Check</option><option>Wire</option><option>Card</option><option>Cash</option><option>Other</option></select></label><label><span>Reference</span><input name="reference" placeholder="Transaction or check number"></label><div class="form-actions"><button type="submit">Save Payment</button></div><p id="payment-message" class="form-message"></p></form></article></dialog>`;
}

function renderInvoiceBatchDialog() {
  return `<dialog id="invoice-batch-dialog" class="crud-dialog invoice-batch-dialog"><article class="panel"><div class="panel-header"><div><h2>Create Billing Batch</h2><p>Draft master billing statement from open invoices.</p></div><button id="close-invoice-batch-dialog" class="dialog-close" type="button">×</button></div><form id="invoice-batch-form" class="record-form"><label><span>Customer</span><select id="invoice-batch-customer" name="customer_id" required></select></label><label><span>Period start</span><input name="period_start" type="date"></label><label><span>Period end</span><input name="period_end" type="date"></label><label><span>Due date</span><input name="due_date" type="date"></label><label class="company-payment-field"><span>Notes</span><textarea name="notes" rows="3" placeholder="Optional statement note"></textarea></label><section class="invoice-batch-candidates"><div class="inline-section-header"><div><span>Open invoices</span><small>Select invoices to include in this batch.</small></div><strong id="invoice-batch-total">$0.00</strong></div><div id="invoice-batch-candidate-list"></div></section><div class="form-actions"><button type="submit">Create Billing Batch</button></div><p id="invoice-batch-message" class="form-message"></p></form></article></dialog>`;
}

function renderCollectionNoteDialog() {
  return `<dialog id="collection-note-dialog" class="crud-dialog collection-note-dialog"><article class="panel"><div class="panel-header"><div><h2>Add Collection Note</h2><p id="collection-note-summary"></p></div><button id="close-collection-note-dialog" class="dialog-close" type="button">×</button></div><form id="collection-note-form" class="record-form"><input name="entity_type" type="hidden"><input name="entity_id" type="hidden"><label><span>Type</span><select name="note_type"><option value="internal">Internal note</option><option value="reminder">Reminder</option><option value="call">Call</option><option value="email">Email</option><option value="promise_to_pay">Promise to pay</option><option value="dispute">Dispute</option></select></label><label><span>Follow-up date</span><input name="follow_up_date" type="date"></label><label class="company-payment-field"><span>Note</span><textarea name="note" rows="4" required placeholder="What happened, next step, promise date or dispute reason"></textarea></label><div class="form-actions"><button type="submit">Save Note</button></div><p id="collection-note-message" class="form-message"></p></form></article></dialog>`;
}

function renderAccessorialApprovalDialog() {
  return `<dialog id="accessorial-approval-dialog" class="crud-dialog accessorial-approval-dialog"><article class="panel"><div class="panel-header"><div><h2>Accessorial Approval</h2><p id="accessorial-approval-summary"></p></div><button id="close-accessorial-approval-dialog" class="dialog-close" type="button">×</button></div><div id="accessorial-approval-list" class="accessorial-approval-list"></div></article></dialog>`;
}

function renderCreditMemoDialog() {
  return `<dialog id="credit-memo-dialog" class="crud-dialog credit-memo-dialog"><article class="panel"><div class="panel-header"><div><h2>Credit Memos</h2><p id="credit-memo-summary"></p></div><button id="close-credit-memo-dialog" class="dialog-close" type="button">×</button></div><div id="credit-memo-content" class="credit-memo-content"></div></article></dialog>`;
}

function resetInvoiceForm() {
  const form = document.querySelector("#invoice-form");
  form?.reset();
  renderInvoiceLineItems([defaultInvoiceLineItem()]);
  if (form) closeFormDialog(form);
}

async function loadInvoiceFormOptions() {
  const customerSelect = document.querySelector("#invoice-customer");
  const loadSelect = document.querySelector("#invoice-load");

  customerSelect.innerHTML = '<option value="">Loading customers...</option>';
  loadSelect.innerHTML = '<option value="">Loading loads...</option>';

  try {
    const [customers, loads] = await Promise.all([listCustomers(), listLoads()]);
    invoiceFormLoadCache = loads;

    customerSelect.innerHTML = customers.length
      ? customers.map((customer) => `<option value="${customer.id}">${escapeHtml(customer.name)}</option>`).join("")
      : '<option value="">Add a customer first</option>';
    loadSelect.innerHTML = [
      '<option value="">No load selected</option>',
      ...loads.filter((load) => ["ready_to_bill", "drafted", "invoiced"].includes(billingLoadStatus(load))).map((load) => `<option value="${load.id}">${escapeHtml(load.load_no)} - ${escapeHtml(routeLabel(load))}</option>`),
    ].join("");
    renderInvoiceLineItems([defaultInvoiceLineItem()]);
  } catch (error) {
    customerSelect.innerHTML = `<option value="">${escapeHtml(error.message)}</option>`;
    loadSelect.innerHTML = '<option value="">Unavailable</option>';
  }
}

function defaultInvoiceLineItem() {
  return { item_type: "linehaul", description: "Transportation services", customer_description: "Transportation services", quantity: 1, unit_amount: 0, approved: true, customer_visible: true };
}

async function handleInvoiceLoadChange(event) {
  const loadId = event.currentTarget.value;
  if (!loadId) {
    renderInvoiceLineItems([defaultInvoiceLineItem()]);
    return;
  }

  try {
    const financials = await getLoadFinancials(loadId);
    const selectedLoad = invoiceFormLoadCache.find((load) => load.id === loadId);
    const route = selectedLoad ? routeLabel(selectedLoad) : "Transportation services";
    const items = [
      numberOrZero(financials?.linehaul) > 0 ? { item_type: "linehaul", description: `Linehaul - ${route}`, customer_description: `Linehaul - ${route}`, quantity: 1, unit_amount: numberOrZero(financials.linehaul), approved: true, customer_visible: true } : null,
      numberOrZero(financials?.fsc) > 0 ? { item_type: "fsc", description: "Fuel surcharge", customer_description: "Fuel surcharge", quantity: 1, unit_amount: numberOrZero(financials.fsc), approved: true, customer_visible: true } : null,
      numberOrZero(financials?.accessorials) > 0 ? { item_type: "custom", description: "Accessorials", customer_description: "Accessorials", quantity: 1, unit_amount: numberOrZero(financials.accessorials), approved: false, customer_visible: true } : null,
    ].filter(Boolean);
    renderInvoiceLineItems(items.length ? items : [{ ...defaultInvoiceLineItem(), description: `Linehaul - ${route}`, customer_description: `Linehaul - ${route}` }]);
  } catch (error) {
    renderInvoiceLineItems([defaultInvoiceLineItem()]);
    const message = document.querySelector("#invoice-message");
    message.textContent = `Load financials unavailable: ${error.message}`;
  }
}

function renderInvoiceLineItems(items = []) {
  const container = document.querySelector("#invoice-line-items");
  if (!container) return;
  container.innerHTML = items.map((item, index) => renderInvoiceLineItem(item, index)).join("");
  container.querySelectorAll("[data-remove-invoice-line]").forEach((button) => {
    button.addEventListener("click", () => {
      button.closest(".invoice-line-item")?.remove();
      updateInvoiceLineItemsTotal();
    });
  });
  container.querySelectorAll("input, select").forEach((input) => input.addEventListener("input", updateInvoiceLineItemsTotal));
  container.querySelectorAll('[name="line_item_type"]').forEach((select) => {
    select.addEventListener("change", () => {
      if (accessorialItemTypes.has(select.value)) {
        select.closest(".invoice-line-item")?.querySelector('[name="line_approved"]')?.removeAttribute("checked");
        const approved = select.closest(".invoice-line-item")?.querySelector('[name="line_approved"]');
        if (approved) approved.checked = false;
      }
    });
  });
  updateInvoiceLineItemsTotal();
}

function addInvoiceLineItem(item) {
  const items = collectInvoiceLineItems({ allowInvalid: true });
  items.push(item);
  renderInvoiceLineItems(items);
}

function renderInvoiceLineItem(item, index) {
  const itemType = item.item_type || "custom";
  return `
    <article class="invoice-line-item">
      <label><span>Type</span><select name="line_item_type">${invoiceLineItemTypeOptions.map(([value, label]) => `<option value="${value}" ${value === itemType ? "selected" : ""}>${label}</option>`).join("")}</select></label>
      <label><span>Description</span><input name="line_description" value="${escapeHtml(item.description || "")}" required></label>
      <label><span>Customer description</span><input name="line_customer_description" value="${escapeHtml(item.customer_description || item.description || "")}" required></label>
      <label><span>Qty</span><input name="line_quantity" type="number" min="0.01" step="0.01" value="${item.quantity ?? 1}" required></label>
      <label><span>Unit amount</span><input name="line_unit_amount" type="number" step="0.01" value="${item.unit_amount ?? 0}" required></label>
      <div class="invoice-line-flags">
        <label><input name="line_receipt_required" type="checkbox" ${item.receipt_required ? "checked" : ""}> Receipt required</label>
        <label><input name="line_customer_visible" type="checkbox" ${item.customer_visible === false ? "" : "checked"}> Customer visible</label>
        <label><input name="line_approved" type="checkbox" ${item.approved ? "checked" : ""}> Approved</label>
      </div>
      <strong class="invoice-line-amount">${formatMoney(numberOrZero(item.quantity || 1) * numberOrZero(item.unit_amount))}</strong>
      <button type="button" data-remove-invoice-line="${index}" aria-label="Remove invoice line">Remove</button>
    </article>
  `;
}

function collectInvoiceLineItems({ allowInvalid = false } = {}) {
  const rows = Array.from(document.querySelectorAll("#invoice-line-items .invoice-line-item"));
  const items = rows.map((row, index) => ({
    item_type: row.querySelector('[name="line_item_type"]').value,
    description: row.querySelector('[name="line_description"]').value,
    customer_description: row.querySelector('[name="line_customer_description"]').value,
    quantity: row.querySelector('[name="line_quantity"]').value,
    unit_amount: row.querySelector('[name="line_unit_amount"]').value,
    receipt_required: row.querySelector('[name="line_receipt_required"]').checked,
    customer_visible: row.querySelector('[name="line_customer_visible"]').checked,
    approved: row.querySelector('[name="line_approved"]').checked,
    sort_order: index + 1,
  }));
  return allowInvalid ? items : validateInvoiceLineItems(items);
}

function updateInvoiceLineItemsTotal() {
  const rows = Array.from(document.querySelectorAll("#invoice-line-items .invoice-line-item"));
  const items = rows.map((row) => ({
    quantity: row.querySelector('[name="line_quantity"]')?.value,
    unit_amount: row.querySelector('[name="line_unit_amount"]')?.value,
  }));
  rows.forEach((row) => {
    const amount = numberOrZero(row.querySelector('[name="line_quantity"]').value || 1) * numberOrZero(row.querySelector('[name="line_unit_amount"]').value);
    row.querySelector(".invoice-line-amount").textContent = formatMoney(amount);
  });
  const total = document.querySelector("#invoice-line-total");
  if (total) total.textContent = formatMoney(invoiceLineItemsTotal(items));
}

let invoiceBatchState = { customers: [], invoices: [], batches: [] };
let collectionState = { rows: [], notes: [] };

async function loadInvoiceBatches() {
  const table = document.querySelector("#invoice-batches-table");
  if (!table) return;
  table.innerHTML = '<div class="empty-state">Loading billing batches...</div>';
  try {
    const [customers, invoices, batches] = await Promise.all([listCustomers(), listInvoices(), listInvoiceBatches()]);
    invoiceBatchState = { customers, invoices, batches };
    table.innerHTML = renderInvoiceBatchRows(batches);
    table.querySelectorAll("[data-preview-invoice-batch]").forEach((button) => {
      button.addEventListener("click", () => renderInvoiceBatchPreview(button.dataset.previewInvoiceBatch));
    });
    table.querySelectorAll("[data-email-invoice-batch]").forEach((button) => {
      button.addEventListener("click", () => handleEmailInvoiceBatch(button));
    });
  } catch (error) {
    table.innerHTML = `<div class="empty-state">Billing batches are not ready: ${escapeHtml(error.message)}</div>`;
  }
}

function renderInvoiceBatchRows(batches) {
  if (!batches.length) return '<div class="empty-state">No billing batches yet.</div>';
  return `
    <div class="data-row invoice-batch-row data-row-head"><span>Batch</span><span>Customer</span><span>Period</span><span>Invoices</span><span>Amount</span><span>Status</span><span>Actions</span></div>
    ${batches.map((batch) => `
      <div class="data-row invoice-batch-row">
        <strong>${escapeHtml(batch.batch_no)}</strong>
        <span>${escapeHtml(batch.customers?.name || "-")}</span>
        <span>${formatDate(batch.period_start)} - ${formatDate(batch.period_end)}</span>
        <span>${batch.invoice_count} invoice${batch.invoice_count === 1 ? "" : "s"}</span>
        <strong>${formatMoney(batch.total_amount)}</strong>
        <span class="status-pill" data-state="${batch.status === "void" ? "error" : "ok"}">${formatStatus(batch.status)}</span>
        <div class="invoice-batch-actions">
          <button type="button" data-preview-invoice-batch="${batch.id}">Preview</button>
          <button type="button" data-email-invoice-batch="${batch.id}" data-recipient="${escapeAttribute(batch.customers?.billing_email || "")}">Email</button>
        </div>
      </div>
    `).join("")}
  `;
}

async function handleEmailInvoiceBatch(button) {
  const recipient = window.prompt("Recipient email address:", button.dataset.recipient || "");
  if (!recipient) return;
  const original = button.textContent;
  button.disabled = true;
  button.textContent = "Sending...";
  try {
    await sendTmsInvoiceBatch(button.dataset.emailInvoiceBatch, recipient.trim());
    await loadInvoiceBatches();
  } catch (error) {
    window.alert(error.message);
    button.disabled = false;
    button.textContent = original || "Email";
  }
}

async function openInvoiceBatchDialog() {
  if (!invoiceBatchState.customers.length) {
    await loadInvoiceBatches();
  }
  const dialog = document.querySelector("#invoice-batch-dialog");
  const form = document.querySelector("#invoice-batch-form");
  form.reset();
  document.querySelector("#invoice-batch-message").textContent = "";
  document.querySelector("#invoice-batch-customer").innerHTML = invoiceBatchState.customers
    .map((customer) => `<option value="${customer.id}">${escapeHtml(customer.name)} · ${formatStatus(customer.billing_cycle || "per_load")}</option>`)
    .join("");
  renderInvoiceBatchCandidateList();
  dialog.showModal();
}

function renderInvoiceBatchCandidateList() {
  const customerId = document.querySelector("#invoice-batch-customer")?.value;
  const list = document.querySelector("#invoice-batch-candidate-list");
  if (!list) return;
  const batchedInvoiceIds = new Set(invoiceBatchState.batches.flatMap((batch) => (batch.invoice_batch_items || []).map((item) => item.invoice_id)));
  const candidates = invoiceBatchCandidates(invoiceBatchState.invoices, localDateIso())
    .filter((invoice) => invoice.customer_id === customerId && !batchedInvoiceIds.has(invoice.id));
  list.innerHTML = candidates.length
    ? candidates.map((invoice) => `<label class="invoice-batch-candidate"><input type="checkbox" name="invoice_ids" value="${invoice.id}" data-balance="${invoice.financialState.balance}" checked><span><strong>${escapeHtml(invoice.invoice_no)}</strong><small>${formatStatus(invoice.status)} · due ${formatDate(invoice.due_date)}</small></span><strong>${formatMoney(invoice.financialState.balance)}</strong></label>`).join("")
    : '<div class="empty-state compact-empty">No open unbatched invoices for this customer.</div>';
  list.querySelectorAll("input").forEach((input) => input.addEventListener("change", updateInvoiceBatchTotal));
  updateInvoiceBatchTotal();
}

function updateInvoiceBatchTotal() {
  const total = Array.from(document.querySelectorAll('#invoice-batch-candidate-list input[name="invoice_ids"]:checked'))
    .reduce((sum, input) => sum + numberOrZero(input.dataset.balance), 0);
  const totalEl = document.querySelector("#invoice-batch-total");
  if (totalEl) totalEl.textContent = formatMoney(total);
}

async function handleCreateInvoiceBatch(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = document.querySelector("#invoice-batch-message");
  const button = form.querySelector('button[type="submit"]');
  const invoiceIds = Array.from(form.querySelectorAll('input[name="invoice_ids"]:checked')).map((input) => input.value);
  if (!invoiceIds.length) {
    message.textContent = "Select at least one invoice.";
    return;
  }
  button.disabled = true;
  message.textContent = "Creating billing batch...";
  try {
    const formData = new FormData(form);
    await createInvoiceBatch({
      customer_id: formData.get("customer_id"),
      period_start: formData.get("period_start") || null,
      period_end: formData.get("period_end") || null,
      due_date: formData.get("due_date") || null,
      notes: formData.get("notes") || null,
    }, invoiceIds);
    document.querySelector("#invoice-batch-dialog").close();
    await loadInvoiceBatches();
  } catch (error) {
    message.textContent = error.message;
  } finally {
    button.disabled = false;
  }
}

async function loadCollections() {
  const table = document.querySelector("#collections-table");
  if (!table) return;
  table.innerHTML = '<div class="empty-state">Loading collections...</div>';
  try {
    const [invoices, batches, notes] = await Promise.all([listInvoices(), listInvoiceBatches(), listInvoiceCollectionNotes()]);
    const rows = buildCollectionRows(invoices, batches, notes);
    collectionState = { rows, notes };
    table.innerHTML = renderCollectionRows(rows);
    table.querySelectorAll("[data-add-collection-note]").forEach((button) => {
      button.addEventListener("click", () => openCollectionNoteDialog(button.dataset.entityType, button.dataset.entityId));
    });
  } catch (error) {
    table.innerHTML = `<div class="empty-state">Collections are not ready: ${escapeHtml(error.message)}</div>`;
  }
}

function buildCollectionRows(invoices, batches, notes) {
  const today = localDateIso();
  const latestNoteByEntity = new Map();
  notes.forEach((note) => {
    const key = `${note.invoice_id ? "invoice" : "batch"}:${note.invoice_id || note.invoice_batch_id}`;
    if (!latestNoteByEntity.has(key)) latestNoteByEntity.set(key, note);
  });
  const invoiceRows = invoices
    .map((invoice) => {
      const financialState = invoiceFinancialState(invoice, today);
      const note = latestNoteByEntity.get(`invoice:${invoice.id}`);
      return {
        entityType: "invoice",
        entityId: invoice.id,
        label: invoice.invoice_no,
        customer: invoice.customers?.name || "-",
        status: financialState.displayStatus,
        dueDate: invoice.due_date,
        amount: financialState.balance,
        daysPastDue: financialState.daysPastDue,
        latestNote: note,
        follow_up_date: note?.follow_up_date || null,
      };
    })
    .filter((row) => row.amount > 0 && !["paid", "void"].includes(row.status));
  const batchRows = batches
    .map((batch) => {
      const note = latestNoteByEntity.get(`batch:${batch.id}`);
      const daysPastDue = batch.due_date && batch.due_date < today ? differenceInCalendarDays(today, batch.due_date) : 0;
      return {
        entityType: "batch",
        entityId: batch.id,
        label: batch.batch_no,
        customer: batch.customers?.name || "-",
        status: daysPastDue > 0 ? "overdue" : batch.status,
        dueDate: batch.due_date,
        amount: numberOrZero(batch.total_amount),
        daysPastDue,
        latestNote: note,
        follow_up_date: note?.follow_up_date || null,
      };
    })
    .filter((row) => row.amount > 0 && !["paid", "void"].includes(row.status));
  return [...invoiceRows, ...batchRows].sort((a, b) => b.daysPastDue - a.daysPastDue || String(a.dueDate || "9999").localeCompare(String(b.dueDate || "9999")));
}

function renderCollectionRows(rows) {
  if (!rows.length) return '<div class="empty-state">No open balances to collect.</div>';
  return `
    <div class="data-row collection-row data-row-head"><span>Account</span><span>Type</span><span>Balance</span><span>Due</span><span>Follow-up</span><span>Last note</span><span>Actions</span></div>
    ${rows.map((row) => {
      const followUpState = collectionFollowUpState(row, localDateIso());
      return `
        <div class="data-row collection-row" data-collection-severity="${followUpState.severity}">
          <div><strong>${escapeHtml(row.customer)}</strong><small>${escapeHtml(row.label)} · ${formatStatus(row.status)}</small></div>
          <span>${formatStatus(row.entityType)}</span>
          <strong>${formatMoney(row.amount)}</strong>
          <span>${formatDate(row.dueDate)}${row.daysPastDue ? `<small>${row.daysPastDue} day(s) past due</small>` : ""}</span>
          <span>${row.follow_up_date ? formatDate(row.follow_up_date) : "-"}</span>
          <span>${row.latestNote ? `${formatStatus(row.latestNote.note_type)} · ${escapeHtml(row.latestNote.note)}` : "-"}</span>
          <button type="button" data-add-collection-note data-entity-type="${row.entityType}" data-entity-id="${row.entityId}">Add Note</button>
        </div>
      `;
    }).join("")}
  `;
}

function openCollectionNoteDialog(entityType, entityId) {
  const row = collectionState.rows.find((item) => item.entityType === entityType && item.entityId === entityId);
  const dialog = document.querySelector("#collection-note-dialog");
  const form = document.querySelector("#collection-note-form");
  form.reset();
  form.elements.entity_type.value = entityType;
  form.elements.entity_id.value = entityId;
  document.querySelector("#collection-note-summary").textContent = row
    ? `${row.customer} · ${row.label} · Balance ${formatMoney(row.amount)}`
    : "Collection follow-up";
  document.querySelector("#collection-note-message").textContent = "";
  dialog.showModal();
}

async function handleCollectionNoteSubmit(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const button = form.querySelector('button[type="submit"]');
  const message = document.querySelector("#collection-note-message");
  const formData = new FormData(form);
  const entityType = formData.get("entity_type");
  button.disabled = true;
  message.textContent = "Saving collection note...";
  try {
    await createInvoiceCollectionNote({
      invoice_id: entityType === "invoice" ? formData.get("entity_id") : null,
      invoice_batch_id: entityType === "batch" ? formData.get("entity_id") : null,
      note_type: formData.get("note_type"),
      note: String(formData.get("note") || "").trim(),
      follow_up_date: formData.get("follow_up_date") || null,
    });
    document.querySelector("#collection-note-dialog").close();
    await loadCollections();
  } catch (error) {
    message.textContent = error.message;
  } finally {
    button.disabled = false;
  }
}

async function handleAccountingExport(type) {
  const message = document.querySelector("#accounting-export-message");
  const buttons = Array.from(document.querySelectorAll("#accounting-export-panel button"));
  buttons.forEach((button) => { button.disabled = true; });
  message.textContent = "Preparing accounting export...";
  try {
    const files = await buildAccountingExportFiles();
    const selected = type === "all" ? files : files.filter((file) => file.type === type);
    selected.forEach((file) => downloadCsv(file.filename, file.rows));
    message.textContent = `Exported ${selected.length} CSV file${selected.length === 1 ? "" : "s"}.`;
  } catch (error) {
    message.textContent = error.message;
  } finally {
    buttons.forEach((button) => { button.disabled = false; });
  }
}

async function buildAccountingExportFiles() {
  const [invoices, batches, notes] = await Promise.all([listInvoices(), listInvoiceBatches(), listInvoiceCollectionNotes()]);
  const stamp = localDateIso();
  const invoiceRows = invoices.map((invoice) => {
    const state = invoiceFinancialState(invoice, stamp);
    return {
      invoice_no: invoice.invoice_no,
      customer: invoice.customers?.name || "",
      load_no: invoice.loads?.load_no || "",
      status: state.displayStatus,
      invoice_date: invoice.created_at ? String(invoice.created_at).slice(0, 10) : "",
      due_date: invoice.due_date || "",
      amount: numberOrZero(invoice.amount).toFixed(2),
      paid: state.paid.toFixed(2),
      balance: state.balance.toFixed(2),
      days_past_due: state.daysPastDue,
      sent_at: invoice.sent_at || "",
    };
  });
  const paymentRows = invoices.flatMap((invoice) => (invoice.payments || []).map((payment) => ({
    invoice_no: invoice.invoice_no,
    customer: invoice.customers?.name || "",
    payment_date: payment.date || "",
    amount: numberOrZero(payment.amount).toFixed(2),
    method: payment.method || "",
    reference: payment.reference || "",
  })));
  const batchRows = batches.map((batch) => ({
    batch_no: batch.batch_no,
    customer: batch.customers?.name || "",
    status: batch.status,
    billing_cycle: batch.billing_cycle,
    period_start: batch.period_start || "",
    period_end: batch.period_end || "",
    due_date: batch.due_date || "",
    invoice_count: batch.invoice_count,
    total_amount: numberOrZero(batch.total_amount).toFixed(2),
    sent_at: batch.sent_at || "",
  }));
  const collectionRows = notes.map((note) => {
    const invoice = invoices.find((item) => item.id === note.invoice_id);
    const batch = batches.find((item) => item.id === note.invoice_batch_id);
    return {
      entity_type: note.invoice_id ? "invoice" : "batch",
      entity_no: invoice?.invoice_no || batch?.batch_no || "",
      customer: invoice?.customers?.name || batch?.customers?.name || "",
      note_type: note.note_type,
      status: note.status,
      follow_up_date: note.follow_up_date || "",
      note: note.note,
      created_at: note.created_at || "",
    };
  });
  return [
    { type: "invoices", filename: `cyrra-accounting-invoices-${stamp}.csv`, rows: invoiceRows },
    { type: "payments", filename: `cyrra-accounting-payments-${stamp}.csv`, rows: paymentRows },
    { type: "batches", filename: `cyrra-accounting-batches-${stamp}.csv`, rows: batchRows },
    { type: "collections", filename: `cyrra-accounting-collections-${stamp}.csv`, rows: collectionRows },
  ];
}

function downloadCsv(filename, rows) {
  const csv = toCsv(rows);
  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" });
  downloadBlob(filename,blob);
}

function downloadBlob(filename,blob) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function toCsv(rows) {
  if (!rows.length) return "";
  const headers = Object.keys(rows[0]);
  return [
    headers.join(","),
    ...rows.map((row) => headers.map((header) => csvCell(row[header])).join(",")),
  ].join("\n");
}

function csvCell(value) {
  const raw = String(value ?? "");
  const text = /^[=+\-@]/.test(raw) ? `'${raw}` : raw;
  return /[",\n\r]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

async function loadInvoices() {
  const table = document.querySelector("#invoices-table");
  const pagination = document.querySelector("#invoices-pagination");
  table.innerHTML = '<div class="empty-state">Loading invoices...</div>';

  try {
    const result = await listInvoicesPage(invoiceListState);
    const invoices = result.records;
    Object.assign(invoiceListState, { records: invoices, total: result.count });

    if (!invoices.length) {
      table.innerHTML = `<div class="empty-state">${invoiceListState.total ? "No invoices on this page." : "No invoices match these filters."}</div>`;
      renderInvoicePagination();
      return;
    }

    table.innerHTML = `
      <div class="data-row invoice-row data-row-head" style="${tableRowStyle("invoices", invoiceListState.columns)}">
        ${can("manage_finance") ? '<input type="checkbox" data-select-page aria-label="Select this page">' : ""}
        <span>Invoice #</span>
        ${invoiceListState.columns.includes("customer") ? "<span>Customer</span>" : ""}
        ${invoiceListState.columns.includes("load") ? "<span>Load</span>" : ""}
        ${invoiceListState.columns.includes("amount") ? "<span>Amount</span>" : ""}
        ${invoiceListState.columns.includes("due") ? "<span>Due</span>" : ""}
        ${invoiceListState.columns.includes("status") ? "<span>Status</span>" : ""}
        <span>Actions</span>
      </div>
      ${invoices
        .map(
          (invoice) => {
            const { balance, displayStatus } = invoiceFinancialState(invoice, localDateIso());
            return `
            <div class="data-row invoice-row" style="${tableRowStyle("invoices", invoiceListState.columns)}">
              ${can("manage_finance") ? `<input type="checkbox" data-select-row value="${invoice.id}" aria-label="Select invoice ${escapeAttribute(invoice.invoice_no)}" ${invoiceListState.selectedIds.has(invoice.id) ? "checked" : ""}>` : ""}
              <strong>${escapeHtml(invoice.invoice_no)}</strong>
              ${invoiceListState.columns.includes("customer") ? `<span>${escapeHtml(invoice.customers?.name || "-")}</span>` : ""}
              ${invoiceListState.columns.includes("load") ? `<span>${escapeHtml(invoice.loads?.load_no || "-")}</span>` : ""}
              ${invoiceListState.columns.includes("amount") ? `<span>${formatMoney(invoice.amount)}<small class="invoice-balance">${(invoice.invoice_line_items || []).length || (invoice.line_items_snapshot || []).length || 1} line(s) · Balance ${formatMoney(balance)}</small></span>` : ""}
              ${invoiceListState.columns.includes("due") ? `<span>${formatDate(invoice.due_date)}</span>` : ""}
              ${invoiceListState.columns.includes("status") ? `<select class="invoice-status-select" data-invoice-status="${invoice.id}" aria-label="Change status for ${escapeHtml(invoice.invoice_no)}">
                  ${renderInvoiceStatusOptions(displayStatus)}
              </select>` : ""}
              <div class="invoice-row-actions">
                <button type="button" data-preview-invoice="${invoice.id}">Preview</button>
                <button type="button" data-accessorial-invoice="${invoice.id}" ${invoiceHasAccessorials(invoice) ? "" : 'disabled title="No accessorial charges"'}>Accessorials</button>
                <button type="button" data-credit-memo-invoice="${invoice.id}">Credits</button>
                <button type="button" data-email-invoice="${invoice.id}">Email</button>
                <button type="button" data-payment-invoice="${invoice.id}" data-balance="${balance}" ${balance <= 0 || invoice.status === "void" ? "disabled" : ""}>Payment</button>
                <button type="button" data-payment-history="${invoice.id}">History</button>
              </div>
            </div>
          `}
        )
        .join("")}
    `;

    table.querySelectorAll("[data-invoice-status]").forEach((select) => {
      select.addEventListener("change", handleInvoiceStatusChange);
    });
    table.querySelectorAll("[data-preview-invoice]").forEach((button) => {
      button.addEventListener("click", () => renderInvoicePreview(button.dataset.previewInvoice));
    });
    table.querySelectorAll("[data-accessorial-invoice]").forEach((button) => {
      button.addEventListener("click", () => openAccessorialApprovalDialog(invoices.find((invoice) => invoice.id === button.dataset.accessorialInvoice)));
    });
    table.querySelectorAll("[data-credit-memo-invoice]").forEach((button) => {
      button.addEventListener("click", () => openCreditMemoDialog(invoices.find((invoice) => invoice.id === button.dataset.creditMemoInvoice)));
    });
    table.querySelectorAll("[data-email-invoice]").forEach((button) => button.addEventListener("click", async () => {
      const recipient = window.prompt("Recipient email address:");
      if (!recipient) return;
      button.disabled = true;
      button.textContent = "Sending...";
      try {
        await sendTmsInvoice(button.dataset.emailInvoice, recipient.trim());
        await loadInvoices();
      } catch (error) {
        window.alert(error.message);
        button.disabled = false;
        button.textContent = "Email";
      }
    }));
    table.querySelectorAll("[data-payment-invoice]").forEach((button) => button.addEventListener("click", () => openPaymentDialog(button)));
    table.querySelectorAll("[data-payment-history]").forEach((button) => button.addEventListener("click", () => {
      const invoice = invoices.find((item) => item.id === button.dataset.paymentHistory);
      openPaymentHistory(invoice);
    }));
    bindTableRowSelection("invoices");
    renderInvoicePagination();
  } catch (error) {
    table.innerHTML = `
      <div class="empty-state">
        Invoices table is not ready: ${escapeHtml(error.message)}
      </div>
    `;
  }
}

function renderInvoicePagination() {
  const pagination = document.querySelector("#invoices-pagination");
  if (!pagination) return;
  const pageCount = Math.max(1, Math.ceil(invoiceListState.total / invoiceListState.pageSize));
  const start = (invoiceListState.page - 1) * invoiceListState.pageSize;
  pagination.innerHTML = `<span>Showing ${invoiceListState.total ? start + 1 : 0}–${Math.min(start + invoiceListState.records.length, invoiceListState.total)} of ${invoiceListState.total}</span><div><button type="button" data-invoice-page="prev" ${invoiceListState.page === 1 ? "disabled" : ""}>Previous</button><strong>Page ${invoiceListState.page} of ${pageCount}</strong><button type="button" data-invoice-page="next" ${invoiceListState.page >= pageCount ? "disabled" : ""}>Next</button></div>`;
  pagination.querySelectorAll("[data-invoice-page]").forEach((button) => button.addEventListener("click", () => {
    invoiceListState.selectedIds.clear();
    invoiceListState.page += button.dataset.invoicePage === "next" ? 1 : -1;
    loadInvoices();
  }));
}

function handleInvoiceFilters(event) {
  const form = event.currentTarget;
  Object.assign(invoiceListState, {
    search: form.elements.search.value,
    status: form.elements.status.value,
    sortBy: form.elements.sort_by.value,
    sortDirection: form.elements.sort_direction.value,
    pageSize: Number(form.elements.page_size.value),
    page: 1,
  });
  clearTimeout(invoiceFilterTimer);
  invoiceListState.selectedIds.clear();
  clearSelectedTableView("invoices");
  invoiceFilterTimer = setTimeout(loadInvoices, event.target.name === "search" ? 250 : 0);
}

function syncInvoiceFilterForm() {
  const form = document.querySelector("#invoice-filters");
  if (!form) return;
  form.elements.search.value = invoiceListState.search;
  form.elements.status.value = invoiceListState.status;
  form.elements.sort_by.value = invoiceListState.sortBy;
  form.elements.sort_direction.value = invoiceListState.sortDirection;
  form.elements.page_size.value = String(invoiceListState.pageSize);
}

const accessorialItemTypes = new Set(["detention", "layover", "lumper", "tonu", "custom"]);

function isAccessorialItem(item) {
  const type = String(item?.item_type || "").toLowerCase();
  const label = `${item?.description || ""} ${item?.customer_description || ""}`.toLowerCase();
  return accessorialItemTypes.has(type) || /accessorial|detention|layover|lumper|truck ordered|tonu/.test(label);
}

function invoiceHasAccessorials(invoice) {
  return [...(invoice?.invoice_line_items || []), ...(invoice?.line_items_snapshot || [])].some(isAccessorialItem);
}

function openAccessorialApprovalDialog(invoice) {
  const dialog = document.querySelector("#accessorial-approval-dialog");
  const list = document.querySelector("#accessorial-approval-list");
  const items = (invoice?.invoice_line_items || []).filter(isAccessorialItem);
  document.querySelector("#accessorial-approval-summary").textContent = `${invoice?.invoice_no || "Invoice"} · review charges before sending`;
  list.innerHTML = items.length
    ? items.map((item) => renderAccessorialApprovalItem(item, invoice.status)).join("")
    : '<div class="empty-state">This invoice has no accessorial charges.</div>';
  list.querySelectorAll("[data-accessorial-action]").forEach((button) => {
    button.addEventListener("click", () => handleAccessorialApproval(button));
  });
  list.querySelectorAll("[data-accessorial-receipt-form]").forEach((form) => {
    form.addEventListener("submit", (event) => handleAccessorialReceiptUpload(event, invoice));
  });
  dialog.showModal();
}

function renderAccessorialApprovalItem(item, invoiceStatus) {
  const status = item.approval_status || (item.approved ? "approved" : "pending");
  const locked = ["sent", "partial", "paid", "void"].includes(invoiceStatus);
  const decisionMeta = item.approved_at
    ? `${item.approved_by === currentSession?.user?.id ? "Decision by you" : "Decision by accounting user"} · ${formatDateTime(item.approved_at)}`
    : "";
  const receiptState = item.receipt_required
    ? item.document_id ? "Receipt attached" : "Receipt missing"
    : "Receipt not required";
  return `<article class="accessorial-approval-item" data-accessorial-item="${item.id}">
    <div class="accessorial-approval-heading"><div><strong>${escapeHtml(item.customer_description || item.description || formatStatus(item.item_type))}</strong><small>${formatStatus(item.item_type)} · ${formatMoney(item.amount)}</small></div><span class="status-pill" data-state="${status === "approved" ? "ok" : status === "rejected" ? "error" : "neutral"}">${formatStatus(status)}</span></div>
    <p class="accessorial-receipt" data-state="${item.receipt_required && !item.document_id ? "error" : "ok"}">${receiptState}</p>
    ${decisionMeta ? `<small class="accessorial-decision-meta">${escapeHtml(decisionMeta)}</small>` : ""}
    ${!locked && item.receipt_required && !item.document_id ? `<form class="accessorial-receipt-upload" data-accessorial-receipt-form data-line-item-id="${item.id}"><label><span>Receipt document</span><input name="receipt" type="file" accept="application/pdf,image/jpeg,image/png" required></label><button type="submit">Upload Receipt</button><small class="form-message"></small></form>` : ""}
    <label><span>Approval note</span><textarea rows="2" data-accessorial-note ${locked ? "disabled" : ""}>${escapeHtml(item.approval_note || "")}</textarea></label>
    <div class="form-actions"><button type="button" data-accessorial-action="approved" data-line-item-id="${item.id}" ${locked || (item.receipt_required && !item.document_id) ? "disabled" : ""}>Approve</button><button type="button" data-accessorial-action="rejected" data-line-item-id="${item.id}" ${locked ? "disabled" : ""}>Reject</button></div>
    ${locked ? '<small>This invoice is locked; approval is read-only.</small>' : ""}
  </article>`;
}

async function handleAccessorialReceiptUpload(event, invoice) {
  event.preventDefault();
  const form = event.currentTarget;
  const button = form.querySelector('button[type="submit"]');
  const message = form.querySelector(".form-message");
  const file = form.elements.receipt.files[0];
  button.disabled = true;
  message.textContent = "Uploading...";
  let uploadedDocument = null;
  try {
    uploadedDocument = await uploadTmsDocument({
      entityType: "invoice",
      entityId: invoice.id,
      docType: "lumper",
      status: "approved",
      file,
      uploadedBy: currentSession.user.id,
    });
    await attachInvoiceLineItemReceipt(form.dataset.lineItemId, uploadedDocument.id);
  } catch (error) {
    if (uploadedDocument?.id) await deleteTmsDocument(uploadedDocument.id).catch(() => {});
    message.textContent = error.message;
    button.disabled = false;
    return;
  }
  document.querySelector("#accessorial-approval-dialog").close();
  await loadInvoices();
  const refreshed = (await listInvoices()).find((item) => item.id === invoice.id);
  if (refreshed) openAccessorialApprovalDialog(refreshed);
}

async function handleAccessorialApproval(button) {
  const item = button.closest("[data-accessorial-item]");
  const note = item.querySelector("[data-accessorial-note]").value;
  const buttons = Array.from(item.querySelectorAll("button"));
  buttons.forEach((action) => { action.disabled = true; });
  try {
    await updateInvoiceLineItemApproval(button.dataset.lineItemId, button.dataset.accessorialAction, note);
    document.querySelector("#accessorial-approval-dialog").close();
    await loadInvoices();
  } catch (error) {
    window.alert(error.message);
    buttons.forEach((action) => { action.disabled = false; });
  }
}

function creditMemoActions(status) {
  if (status === "draft") return ["approved", "void"];
  if (status === "approved") return ["applied", "void"];
  if (status === "applied") return ["void"];
  return [];
}

function creditMemoActionLabel(status) {
  return status === "approved" ? "Approve" : status === "applied" ? "Apply Credit" : "Void";
}

function openCreditMemoDialog(invoice) {
  const dialog = document.querySelector("#credit-memo-dialog");
  const content = document.querySelector("#credit-memo-content");
  const financial = invoiceFinancialState(invoice, localDateIso());
  const memos = [...(invoice?.invoice_credit_memos || [])].sort((a, b) => String(b.created_at).localeCompare(String(a.created_at)));
  document.querySelector("#credit-memo-summary").textContent = `${invoice.invoice_no} · balance ${formatMoney(financial.balance)}`;
  content.innerHTML = `<div class="credit-memo-list">${memos.length ? memos.map((memo) => `<article class="credit-memo-item"><div><strong>${escapeHtml(memo.credit_memo_no)}</strong><span class="status-pill" data-state="${memo.status === "applied" ? "ok" : memo.status === "void" ? "error" : "neutral"}">${formatStatus(memo.status)}</span></div><p>${escapeHtml(memo.reason)}</p><strong>${formatMoney(memo.amount)}</strong><small>Created ${formatDateTime(memo.created_at)}${memo.applied_at ? ` · Applied ${formatDateTime(memo.applied_at)}` : ""}${memo.voided_at ? ` · Voided ${formatDateTime(memo.voided_at)}` : ""}</small><div class="form-actions">${creditMemoActions(memo.status).map((status) => `<button type="button" data-credit-memo-action="${status}" data-credit-memo-id="${memo.id}">${creditMemoActionLabel(status)}</button>`).join("")}</div></article>`).join("") : '<div class="empty-state">No credit memos for this invoice.</div>'}</div>${financial.balance > 0 && !["draft", "void"].includes(invoice.status) ? `<form id="credit-memo-form" class="record-form"><input name="invoice_id" type="hidden" value="${invoice.id}"><label><span>Credit amount</span><input name="amount" type="number" min="0.01" max="${financial.balance}" step="0.01" required></label><label><span>Reason</span><textarea name="reason" rows="3" maxlength="1000" required placeholder="Rate adjustment, service issue or billing correction"></textarea></label><button type="submit">Create Draft Credit</button><p class="form-message"></p></form>` : '<small class="credit-memo-help">A positive open balance on a sent or partial invoice is required to create a credit memo.</small>'}`;
  content.querySelector("#credit-memo-form")?.addEventListener("submit", handleCreateCreditMemo);
  content.querySelectorAll("[data-credit-memo-action]").forEach((button) => button.addEventListener("click", () => handleCreditMemoStatus(button, invoice.id)));
  dialog.showModal();
}

async function refreshCreditMemoDialog(invoiceId) {
  await loadInvoices();
  const invoice = (await listInvoices()).find((item) => item.id === invoiceId);
  if (invoice) openCreditMemoDialog(invoice);
}

async function handleCreateCreditMemo(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = form.querySelector(".form-message");
  const data = new FormData(form);
  try {
    await createInvoiceCreditMemo(form.elements.invoice_id.value, numberOrZero(data.get("amount")), String(data.get("reason") || "").trim());
    document.querySelector("#credit-memo-dialog").close();
    await refreshCreditMemoDialog(form.elements.invoice_id.value);
  } catch (error) {
    message.textContent = error.message;
  }
}

async function handleCreditMemoStatus(button, invoiceId) {
  button.disabled = true;
  try {
    await updateInvoiceCreditMemoStatus(button.dataset.creditMemoId, button.dataset.creditMemoAction);
    document.querySelector("#credit-memo-dialog").close();
    await refreshCreditMemoDialog(invoiceId);
  } catch (error) {
    window.alert(error.message);
    button.disabled = false;
  }
}

function openPaymentHistory(invoice) {
  let dialog = document.querySelector("#payment-history-dialog");
  if (!dialog) {
    dialog = document.createElement("dialog");
    dialog.id = "payment-history-dialog";
    dialog.className = "crud-dialog payment-history-dialog";
    document.body.append(dialog);
  }
  const payments = [...(invoice.payments || [])].sort((a, b) => String(b.date).localeCompare(String(a.date)));
  const paid = payments.reduce((sum, payment) => sum + numberOrZero(payment.amount), 0);
  dialog.innerHTML = `<article class="panel"><div class="panel-header"><div><h2>Payment History</h2><p>Invoice ${escapeHtml(invoice.invoice_no)} · Paid ${formatMoney(paid)} of ${formatMoney(invoice.amount)}</p></div><button class="dialog-close" type="button" data-close-payment-history>×</button></div><div class="payment-history-list">${payments.length ? payments.map((payment) => `<article><div><strong>${formatMoney(payment.amount)}</strong><span>${formatDate(payment.date)}</span></div><p>${escapeHtml(payment.method || "Method not set")}</p><small>${escapeHtml(payment.reference || "No reference")}</small></article>`).join("") : '<div class="empty-state">No payments recorded.</div>'}</div></article>`;
  dialog.querySelector("[data-close-payment-history]").addEventListener("click", () => dialog.close());
  dialog.showModal();
}

function openPaymentDialog(button) {
  const form = document.querySelector("#payment-form");
  form.reset();
  form.elements.invoice_id.value = button.dataset.paymentInvoice;
  form.elements.amount.value = button.dataset.balance;
  form.elements.amount.max = button.dataset.balance;
  form.elements.date.value = localDateIso();
  document.querySelector("#payment-summary").textContent = `Remaining balance ${formatMoney(button.dataset.balance)}`;
  document.querySelector("#payment-message").textContent = "";
  document.querySelector("#payment-dialog").showModal();
}

async function handlePaymentSubmit(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = document.querySelector("#payment-message");
  const button = form.querySelector("button[type='submit']");
  button.disabled = true;
  message.textContent = "Saving payment...";
  try {
    await createPayment({ invoice_id: form.elements.invoice_id.value, amount: numberOrZero(form.elements.amount.value), date: form.elements.date.value, method: form.elements.method.value, reference: form.elements.reference.value || null });
    document.querySelector("#payment-dialog").close();
    await loadInvoices();
  } catch (error) { message.textContent = error.message; button.disabled = false; }
}

async function renderInvoicePreview(invoiceId) {
  let dialog = document.querySelector("#invoice-preview-dialog");
  if (!dialog) {
    dialog = document.createElement("dialog");
    dialog.id = "invoice-preview-dialog";
    dialog.className = "invoice-preview-dialog";
    document.body.append(dialog);
  }
  dialog.innerHTML = '<div class="empty-state">Loading invoice preview...</div>';
  dialog.showModal();

  try {
    const [invoice, company] = await Promise.all([getInvoiceDetails(invoiceId), getCompanyBillingSettings()]);
    const financials = invoice.load_id ? await getLoadFinancials(invoice.load_id) : null;
    const linehaul = numberOrZero(financials?.linehaul);
    const fsc = numberOrZero(financials?.fsc);
    const accessorials = numberOrZero(financials?.accessorials);
    let rows = Array.isArray(invoice.line_items_snapshot) && invoice.line_items_snapshot.length
      ? invoice.line_items_snapshot
      : [...(invoice.invoice_line_items || [])].sort((a, b) => numberOrZero(a.sort_order) - numberOrZero(b.sort_order));
    rows = rows.filter((row) => row.customer_visible !== false);
    if (!rows.length) {
      rows = [
        { customer_description: "Linehaul", amount: linehaul },
        { customer_description: "Fuel surcharge", amount: fsc },
        { customer_description: "Accessorials", amount: accessorials },
      ].filter((row) => numberOrZero(row.amount) > 0);
    }
    if (!rows.length) rows.push({ customer_description: "Transportation services", amount: numberOrZero(invoice.amount) });

    dialog.innerHTML = `
      <div class="invoice-preview-toolbar">
        <button type="button" data-close-invoice>Close</button>
        <button class="primary-button" type="button" data-print-invoice>Print / Save PDF</button>
      </div>
      <article class="invoice-sheet">
        <header class="invoice-brand-header">
          <div class="invoice-brand"><img src="assets/cyrra-logo.png" alt="Cyrra logo"><div><strong>${escapeHtml(company.legal_name)}</strong><span>${escapeHtml([company.address_line_1, company.address_line_2].filter(Boolean).join(", "))}</span><small>${escapeHtml([company.city, company.state, company.postal_code].filter(Boolean).join(", "))}${company.phone ? ` · ${escapeHtml(company.phone)}` : ""}</small></div></div>
          <div class="invoice-title"><h1>INVOICE</h1><strong>#${escapeHtml(invoice.invoice_no)}</strong><span>${formatStatus(invoice.status)}</span></div>
        </header>
        <section class="invoice-meta-grid">
          <div><span>Bill To</span><strong>${escapeHtml(invoice.customers?.name || "-")}</strong><p>${escapeHtml(invoice.customers?.billing_address || "Billing address not configured")}</p></div>
          <div><span>Invoice Date</span><strong>${formatDate(invoice.created_at)}</strong><span>Due Date</span><strong>${formatDate(invoice.due_date)}</strong></div>
        </section>
        <section class="invoice-load-strip">
          <div><span>Load</span><strong>${escapeHtml(invoice.loads?.load_no || "-")}</strong></div>
          <div><span>Route</span><strong>${escapeHtml(invoice.loads ? routeLabel(invoice.loads) : "-")}</strong></div>
          <div><span>Pickup</span><strong>${formatDate(invoice.loads?.pickup_date)}</strong></div>
          <div><span>Delivery</span><strong>${formatDate(invoice.loads?.delivery_date)}</strong></div>
        </section>
        <table class="invoice-lines"><thead><tr><th>Description</th><th>Amount</th></tr></thead><tbody>${rows.map((row) => `<tr><td><strong>${escapeHtml(row.customer_description || row.description || "Invoice line")}</strong>${row.receipt_required ? "<br><small>Receipt required</small>" : ""}${row.approved ? "<br><small>Approved</small>" : ""}</td><td>${formatMoney(row.amount ?? (numberOrZero(row.quantity || 1) * numberOrZero(row.unit_amount)))}</td></tr>`).join("")}</tbody></table>
        <section class="invoice-total"><span>Total Due</span><strong>${formatMoney(invoice.amount)}</strong></section>
        <footer><p>Thank you for your business.</p><small>${escapeHtml(company.payment_instructions || "Payment instructions to be configured.")}${company.tax_id ? ` · Tax ID: ${escapeHtml(company.tax_id)}` : ""}${company.email ? ` · ${escapeHtml(company.email)}` : ""}</small></footer>
      </article>`;
    dialog.querySelector("[data-close-invoice]").addEventListener("click", () => dialog.close());
    dialog.querySelector("[data-print-invoice]").addEventListener("click", () => {
      document.body.classList.add("invoice-printing");
      window.print();
    });
  } catch (error) {
    dialog.innerHTML = `<div class="empty-state">Invoice preview unavailable: ${escapeHtml(error.message)}</div><button type="button" data-close-invoice>Close</button>`;
    dialog.querySelector("[data-close-invoice]").addEventListener("click", () => dialog.close());
  }
}

async function renderInvoiceBatchPreview(batchId) {
  let dialog = document.querySelector("#invoice-preview-dialog");
  if (!dialog) {
    dialog = document.createElement("dialog");
    dialog.id = "invoice-preview-dialog";
    dialog.className = "invoice-preview-dialog";
    document.body.append(dialog);
  }
  const batch = invoiceBatchState.batches.find((item) => item.id === batchId);
  if (!batch) {
    dialog.innerHTML = '<div class="empty-state">Billing batch preview unavailable.</div><button type="button" data-close-batch-preview>Close</button>';
    dialog.querySelector("[data-close-batch-preview]").addEventListener("click", () => dialog.close());
    dialog.showModal();
    return;
  }

  dialog.innerHTML = '<div class="empty-state">Loading billing statement preview...</div>';
  dialog.showModal();

  try {
    const company = await getCompanyBillingSettings();
    const items = [...(batch.invoice_batch_items || [])].sort((a, b) => String(a.invoices?.invoice_no || "").localeCompare(String(b.invoices?.invoice_no || "")));
    const periodLabel = batch.period_start || batch.period_end
      ? `${formatDate(batch.period_start)} - ${formatDate(batch.period_end)}`
      : "-";
    dialog.innerHTML = `
      <div class="invoice-preview-toolbar">
        <button type="button" data-close-batch-preview>Close</button>
        <button class="primary-button" type="button" data-print-batch-preview>Print / Save PDF</button>
      </div>
      <article class="invoice-sheet statement-sheet">
        <header class="invoice-brand-header">
          <div class="invoice-brand"><img src="assets/cyrra-logo.png" alt="Cyrra logo"><div><strong>${escapeHtml(company.legal_name)}</strong><span>${escapeHtml([company.address_line_1, company.address_line_2].filter(Boolean).join(", "))}</span><small>${escapeHtml([company.city, company.state, company.postal_code].filter(Boolean).join(", "))}${company.phone ? ` · ${escapeHtml(company.phone)}` : ""}</small></div></div>
          <div class="invoice-title"><h1>STATEMENT</h1><strong>#${escapeHtml(batch.batch_no)}</strong><span>${formatStatus(batch.status)}</span></div>
        </header>
        <section class="invoice-meta-grid">
          <div><span>Bill To</span><strong>${escapeHtml(batch.customers?.name || "-")}</strong><p>${escapeHtml(batch.customers?.billing_address || "Billing address not configured")}</p>${batch.customers?.billing_email ? `<small>${escapeHtml(batch.customers.billing_email)}</small>` : ""}</div>
          <div><span>Statement Date</span><strong>${formatDate(batch.created_at)}</strong><span>Due Date</span><strong>${formatDate(batch.due_date)}</strong><span>Period</span><strong>${escapeHtml(periodLabel)}</strong></div>
        </section>
        <section class="invoice-load-strip statement-summary-strip">
          <div><span>Billing Cycle</span><strong>${formatStatus(batch.billing_cycle || "per_load")}</strong></div>
          <div><span>Invoices</span><strong>${batch.invoice_count || items.length}</strong></div>
          <div><span>Status</span><strong>${formatStatus(batch.status)}</strong></div>
          <div><span>Total</span><strong>${formatMoney(batch.total_amount)}</strong></div>
        </section>
        <table class="invoice-lines">
          <thead><tr><th>Invoice</th><th>Status</th><th>Due Date</th><th>Statement Amount</th></tr></thead>
          <tbody>${items.length ? items.map((item) => `<tr><td><strong>${escapeHtml(item.invoices?.invoice_no || "Invoice")}</strong><br><small>Original amount ${formatMoney(item.invoices?.amount)}</small></td><td>${formatStatus(item.invoices?.status || "-")}</td><td>${formatDate(item.invoices?.due_date)}</td><td>${formatMoney(item.amount_snapshot)}</td></tr>`).join("") : '<tr><td colspan="4">No invoices attached to this statement.</td></tr>'}</tbody>
        </table>
        ${batch.notes ? `<section class="statement-note"><span>Notes</span><p>${escapeHtml(batch.notes)}</p></section>` : ""}
        <section class="invoice-total"><span>Statement Total</span><strong>${formatMoney(batch.total_amount)}</strong></section>
        <footer><p>Thank you for your business.</p><small>${escapeHtml(company.payment_instructions || "Payment instructions to be configured.")}${company.tax_id ? ` · Tax ID: ${escapeHtml(company.tax_id)}` : ""}${company.email ? ` · ${escapeHtml(company.email)}` : ""}</small></footer>
      </article>`;
    dialog.querySelector("[data-close-batch-preview]").addEventListener("click", () => dialog.close());
    dialog.querySelector("[data-print-batch-preview]").addEventListener("click", () => {
      document.body.classList.add("invoice-printing");
      window.print();
    });
  } catch (error) {
    dialog.innerHTML = `<div class="empty-state">Billing statement preview unavailable: ${escapeHtml(error.message)}</div><button type="button" data-close-batch-preview>Close</button>`;
    dialog.querySelector("[data-close-batch-preview]").addEventListener("click", () => dialog.close());
  }
}

window.addEventListener("afterprint", () => document.body.classList.remove("invoice-printing"));

async function handleCreateInvoice(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const message = document.querySelector("#invoice-message");
  const formData = new FormData(form);

  message.textContent = "Saving invoice...";

  try {
    const lineItems = collectInvoiceLineItems();
    await createInvoiceWithLineItems({
      invoice_no: formData.get("invoice_no") || null,
      customer_id: formData.get("customer_id"),
      load_id: formData.get("load_id") || null,
      due_date: formData.get("due_date") || null,
      status: formData.get("status"),
    }, lineItems);

    resetInvoiceForm();
    message.textContent = "Invoice saved.";
    await Promise.all([loadInvoiceFormOptions(), loadInvoices()]);
  } catch (error) {
    message.textContent = error.message;
  }
}

async function handleInvoiceStatusChange(event) {
  const select = event.currentTarget;
  const invoiceId = select.dataset.invoiceStatus;
  const status = select.value;

  if (!canManuallySetInvoiceStatus(status)) {
    await loadInvoices();
    window.alert("Partial, Paid and Overdue are calculated from payments, credits and due dates. Use Payment or Credits to change the balance.");
    return;
  }

  select.disabled = true;

  try {
    await updateInvoiceStatus(invoiceId, status);
    await loadInvoices();
  } catch (error) {
    await loadInvoices();
    window.alert(`Invoice status update failed: ${error.message}`);
  }
}

function renderSettlements() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Phase 2", title: "Carrier Pay / Driver Settlements" })}
    <dialog id="settlement-dialog" class="crud-dialog settlement-dialog">
      <article class="panel">
        <div class="panel-header"><div><h2>Create Settlement</h2><p>Select an eligible delivered load and payee cost.</p></div><button id="close-settlement-dialog" class="dialog-close" type="button">×</button></div>
        <form id="settlement-form" class="record-form">
          <label><span>Load / Payee</span><select name="candidate" id="settlement-candidate" required></select></label>
          <label><span>Notes</span><textarea name="notes" rows="3" placeholder="Optional accounting note"></textarea></label>
          <div class="settlement-candidate-summary" id="settlement-candidate-summary"></div>
          <div class="form-actions"><button type="submit">Create Draft Settlement</button></div>
          <p id="settlement-message" class="form-message"></p>
        </form>
      </article>
    </dialog>
    <dialog id="settlement-pay-batch-dialog" class="crud-dialog settlement-pay-batch-dialog">
      <article class="panel">
        <div class="panel-header"><div><h2>Create Pay Batch</h2><p>Group payable settlements for one carrier or driver.</p></div><button id="close-settlement-pay-batch-dialog" class="dialog-close" type="button">×</button></div>
        <form id="settlement-pay-batch-form" class="record-form">
          <label><span>Payee</span><select name="payee_key" id="settlement-pay-batch-payee" required></select></label>
          <label><span>Period start</span><input name="period_start" type="date"></label>
          <label><span>Period end</span><input name="period_end" type="date"></label>
          <label><span>Pay date</span><input name="pay_date" type="date"></label>
          <label class="company-payment-field"><span>Notes</span><textarea name="notes" rows="3" placeholder="Optional pay batch note"></textarea></label>
          <section class="settlement-pay-batch-candidates">
            <div class="inline-section-header"><div><span>Payable settlements</span><small>Select settlements to include.</small></div><strong id="settlement-pay-batch-total">$0.00</strong></div>
            <div id="settlement-pay-batch-candidate-list"></div>
          </section>
          <div class="form-actions"><button type="submit">Create Pay Batch</button></div>
          <p id="settlement-pay-batch-message" class="form-message"></p>
        </form>
      </article>
    </dialog>
    <section id="settlements-root" class="settlements-root"><div class="empty-state">Loading settlements...</div></section>
  `;
  document.querySelector("#close-settlement-dialog").addEventListener("click", () => document.querySelector("#settlement-dialog").close());
  document.querySelector("#settlement-form").addEventListener("submit", handleCreateSettlement);
  document.querySelector("#close-settlement-pay-batch-dialog").addEventListener("click", () => document.querySelector("#settlement-pay-batch-dialog").close());
  document.querySelector("#settlement-pay-batch-form").addEventListener("submit", handleCreateSettlementPayBatch);
  document.querySelector("#settlement-pay-batch-payee").addEventListener("change", renderSettlementPayBatchCandidateList);
  loadSettlementsPage();
}

let settlementCandidateCache = [];
let settlementPayBatchState = { settlements: [], batches: [], candidates: [] };
const settlementListState = { search: "", type: "", status: "", sortBy: "created_at", direction: "desc", page: 1, pageSize: 5 };
const settlementBatchListState = { search: "", status: "", sortBy: "created_at", direction: "desc", page: 1, pageSize: 5 };
let settlementFilterTimer = null;

async function loadSettlementsPage() {
  const root = document.querySelector("#settlements-root");
  try {
    const [loads, settlements, payBatches] = await Promise.all([listSettlementCandidates(), listSettlements(), listSettlementPayBatches()]);
    settlementCandidateCache = buildSettlementCandidates(loads);
    settlementPayBatchState = { settlements, batches: payBatches, candidates: settlementPayBatchCandidates(settlements, payBatches) };
    renderSettlementWorkspace();
  } catch (error) {
    root.innerHTML = `<div class="empty-state">Settlements are not ready: ${escapeHtml(error.message)}</div>`;
  }
}

function settlementPayeeName(record) {
  return record.payee_type === "carrier" ? record.carriers?.name : record.drivers?.name || "-";
}

function filteredSettlementRecords(records, state, isBatch = false) {
  const query = state.search.trim().toLowerCase();
  return records.filter((record) => !query || [isBatch ? record.batch_no : record.settlement_no, settlementPayeeName(record)].some((value) => String(value || "").toLowerCase().includes(query)))
    .filter((record) => !state.type || record.payee_type === state.type)
    .filter((record) => !state.status || record.status === state.status)
    .sort((left, right) => {
      const leftValue = state.sortBy === "amount" ? numberOrZero(left.total_amount) : Date.parse(left[state.sortBy] || "") || 0;
      const rightValue = state.sortBy === "amount" ? numberOrZero(right.total_amount) : Date.parse(right[state.sortBy] || "") || 0;
      return (leftValue - rightValue) * (state.direction === "asc" ? 1 : -1);
    });
}

function pagedSettlementRecords(records, state) {
  const pageCount = Math.max(1, Math.ceil(records.length / state.pageSize));
  state.page = Math.min(state.page, pageCount);
  const start = (state.page - 1) * state.pageSize;
  return { pageRecords: records.slice(start, start + state.pageSize), pageCount, start };
}

function renderSettlementListControls(type) {
  const state = type === "settlements" ? settlementListState : settlementBatchListState;
  return `<div class="settlement-list-controls" data-settlement-list-controls="${type}">
    <label><span>Search</span><input name="search" value="${escapeAttribute(state.search)}" placeholder="${type === "settlements" ? "Settlement, payee..." : "Batch, payee..."}"></label>
    ${type === "settlements" ? `<label><span>Type</span><select name="type"><option value="">All types</option><option value="carrier" ${state.type === "carrier" ? "selected" : ""}>Carrier</option><option value="driver" ${state.type === "driver" ? "selected" : ""}>Driver</option></select></label>` : ""}
    <label><span>Status</span><select name="status"><option value="">All statuses</option>${["draft", "approved", "paid", "void"].map((status) => `<option value="${status}" ${state.status === status ? "selected" : ""}>${formatStatus(status)}</option>`).join("")}</select></label>
    <label><span>Sort by</span><select name="sortBy"><option value="created_at" ${state.sortBy === "created_at" ? "selected" : ""}>Created</option><option value="${type === "settlements" ? "settlement_date" : "pay_date"}" ${state.sortBy === (type === "settlements" ? "settlement_date" : "pay_date") ? "selected" : ""}>${type === "settlements" ? "Settlement date" : "Pay date"}</option><option value="amount" ${state.sortBy === "amount" ? "selected" : ""}>Amount</option></select></label>
    <label><span>Direction</span><select name="direction"><option value="desc" ${state.direction === "desc" ? "selected" : ""}>Descending</option><option value="asc" ${state.direction === "asc" ? "selected" : ""}>Ascending</option></select></label>
    <label><span>Rows</span><select name="pageSize">${[5, 10, 25].map((size) => `<option value="${size}" ${state.pageSize === size ? "selected" : ""}>${size}</option>`).join("")}</select></label>
    <button type="button" data-clear-settlement-filters>Clear</button>
  </div>`;
}

function renderSettlementPagination(type, total, pageCount, start, shown) {
  const state = type === "settlements" ? settlementListState : settlementBatchListState;
  return `<div class="table-pagination" data-settlement-pagination="${type}"><span>Showing ${total ? start + 1 : 0}–${Math.min(start + shown, total)} of ${total}</span><div><button type="button" data-page-move="-1" ${state.page === 1 ? "disabled" : ""}>Previous</button><strong>Page ${state.page} of ${pageCount}</strong><button type="button" data-page-move="1" ${state.page === pageCount ? "disabled" : ""}>Next</button></div></div>`;
}

function renderSettlementWorkspace() {
  const root = document.querySelector("#settlements-root");
  const { settlements, batches: payBatches } = settlementPayBatchState;
  const filteredSettlements = filteredSettlementRecords(settlements, settlementListState);
  const filteredBatches = filteredSettlementRecords(payBatches, settlementBatchListState, true);
  const settlementPage = pagedSettlementRecords(filteredSettlements, settlementListState);
  const batchPage = pagedSettlementRecords(filteredBatches, settlementBatchListState);
  const totals = settlements.reduce((result, settlement) => {
      result[settlement.status] = (result[settlement.status] || 0) + numberOrZero(settlement.total_amount);
      return result;
    }, {});
    root.innerHTML = `
      <section class="report-kpis settlement-kpis">
        ${renderReportKpi("Draft", formatMoney(totals.draft || 0))}
        ${renderReportKpi("Approved / Payable", formatMoney(totals.approved || 0))}
        ${renderReportKpi("Paid", formatMoney(totals.paid || 0))}
        ${renderReportKpi("Eligible costs", settlementCandidateCache.length)}
      </section>
      <section class="panel">
        <div class="panel-header">
          <div><h2>Settlement List</h2><p>Carrier and driver pay statements connected to completed loads.</p></div>
          <div class="table-tools"><button id="reload-settlements" type="button">Refresh</button><button id="open-settlement-dialog" class="primary-button" type="button" ${settlementCandidateCache.length ? "" : `disabled title="No eligible completed load costs"`}>Create Settlement</button></div>
        </div>
        ${settlementCandidateCache.length ? "" : '<p class="panel-action-help">Create Settlement becomes available when a delivered load has an unprocessed carrier or driver cost.</p>'}
        ${renderSettlementListControls("settlements")}
        <div class="data-table">${renderSettlementRows(settlementPage.pageRecords)}</div>
        ${renderSettlementPagination("settlements", filteredSettlements.length, settlementPage.pageCount, settlementPage.start, settlementPage.pageRecords.length)}
      </section>
      <section class="panel settlement-pay-batches-panel">
        <div class="panel-header">
          <div><h2>Pay Batches</h2><p>Group carrier or driver settlements into pay period statements.</p></div>
          <div class="table-tools"><button id="reload-pay-batches" type="button">Refresh</button><button id="open-settlement-pay-batch-dialog" class="primary-button" type="button" ${settlementPayBatchState.candidates.length ? "" : `disabled title="No Draft or Approved settlements are available"`}>Create Pay Batch</button></div>
        </div>
        ${settlementPayBatchState.candidates.length ? "" : '<p class="panel-action-help">Create Pay Batch becomes available when a Draft or Approved settlement is not already in a batch.</p>'}
        ${renderSettlementListControls("batches")}
        <div class="data-table">${renderSettlementPayBatchRows(batchPage.pageRecords)}</div>
        ${renderSettlementPagination("batches", filteredBatches.length, batchPage.pageCount, batchPage.start, batchPage.pageRecords.length)}
      </section>
    `;
    document.querySelector("#reload-settlements").addEventListener("click", loadSettlementsPage);
    document.querySelector("#open-settlement-dialog").addEventListener("click", openSettlementDialog);
    root.querySelectorAll("[data-settlement-status]").forEach((select) => select.addEventListener("change", handleSettlementStatusChange));
    root.querySelectorAll("[data-preview-settlement]").forEach((button) => button.addEventListener("click", () => {
      const settlement = settlements.find((item) => item.id === button.dataset.previewSettlement);
      renderSettlementPreview(settlement);
    }));
    document.querySelector("#reload-pay-batches").addEventListener("click", loadSettlementsPage);
    document.querySelector("#open-settlement-pay-batch-dialog").addEventListener("click", openSettlementPayBatchDialog);
    root.querySelectorAll("[data-pay-batch-status]").forEach((select) => select.addEventListener("change", handleSettlementPayBatchStatusChange));
    root.querySelectorAll("[data-preview-pay-batch]").forEach((button) => button.addEventListener("click", () => {
      const batch = payBatches.find((item) => item.id === button.dataset.previewPayBatch);
      renderSettlementPayBatchPreview(batch);
    }));
    bindSettlementListControls();
}

function bindSettlementListControls() {
  document.querySelectorAll("[data-settlement-list-controls]").forEach((controls) => {
    const type = controls.dataset.settlementListControls;
    const state = type === "settlements" ? settlementListState : settlementBatchListState;
    controls.querySelectorAll("input, select").forEach((field) => field.addEventListener(field.name === "search" ? "input" : "change", () => {
      state[field.name] = field.name === "pageSize" ? Number(field.value) : field.value;
      state.page = 1;
      clearTimeout(settlementFilterTimer);
      settlementFilterTimer = setTimeout(renderSettlementWorkspace, field.name === "search" ? 250 : 0);
    }));
    controls.querySelector("[data-clear-settlement-filters]").addEventListener("click", () => {
      Object.assign(state, { search: "", type: "", status: "", sortBy: "created_at", direction: "desc", page: 1, pageSize: 5 });
      renderSettlementWorkspace();
    });
  });
  document.querySelectorAll("[data-settlement-pagination]").forEach((pagination) => pagination.querySelectorAll("[data-page-move]").forEach((button) => button.addEventListener("click", () => {
    const state = pagination.dataset.settlementPagination === "settlements" ? settlementListState : settlementBatchListState;
    state.page += Number(button.dataset.pageMove);
    renderSettlementWorkspace();
  })));
}

function buildSettlementCandidates(loads) {
  return loads.flatMap((load) => {
    const financials = Array.isArray(load.load_financials) ? load.load_financials[0] : load.load_financials;
    const activeItems = (load.settlement_items || []).filter((item) => item.active);
    const candidates = [];
    if (load.carrier_id && numberOrZero(financials?.carrier_cost) > 0 && !activeItems.some((item) => item.payee_type === "carrier")) {
      candidates.push({ load, payeeType: "carrier", payee: load.carriers?.name || "Carrier", amount: numberOrZero(financials.carrier_cost) });
    }
    if (load.driver_id && numberOrZero(financials?.driver_pay) > 0 && !activeItems.some((item) => item.payee_type === "driver")) {
      candidates.push({ load, payeeType: "driver", payee: load.drivers?.name || "Driver", amount: numberOrZero(financials.driver_pay) });
    }
    return candidates;
  });
}

function openSettlementDialog() {
  const form = document.querySelector("#settlement-form");
  form.reset();
  document.querySelector("#settlement-message").textContent = "";
  const select = document.querySelector("#settlement-candidate");
  select.innerHTML = settlementCandidateCache.map((candidate, index) => `<option value="${index}">${formatStatus(candidate.payeeType)} · Load ${escapeHtml(candidate.load.load_no)} · ${escapeHtml(candidate.payee)} · ${formatMoney(candidate.amount)}</option>`).join("");
  select.onchange = renderSettlementCandidateSummary;
  renderSettlementCandidateSummary();
  document.querySelector("#settlement-dialog").showModal();
}

function renderSettlementCandidateSummary() {
  const index = Number(document.querySelector("#settlement-candidate").value || 0);
  const candidate = settlementCandidateCache[index];
  document.querySelector("#settlement-candidate-summary").innerHTML = candidate ? `
    <span>${escapeHtml(routeLabel(candidate.load))}</span><strong>${formatMoney(candidate.amount)}</strong><small>${formatStatus(candidate.payeeType)} pay to ${escapeHtml(candidate.payee)}</small>
  ` : '<span>No eligible costs.</span>';
}

async function handleCreateSettlement(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const candidate = settlementCandidateCache[Number(form.elements.candidate.value)];
  const message = document.querySelector("#settlement-message");
  const button = form.querySelector("button[type='submit']");
  if (!candidate) return;
  button.disabled = true;
  message.textContent = "Creating settlement...";
  try {
    await createLoadSettlement(candidate.load.id, candidate.payeeType, form.elements.notes.value || null);
    document.querySelector("#settlement-dialog").close();
    await loadSettlementsPage();
  } catch (error) {
    message.textContent = error.message;
  } finally {
    button.disabled = false;
  }
}

function renderSettlementRows(settlements) {
  if (!settlements.length) return '<div class="empty-state">No matching settlements.</div>';
  return `
    <div class="data-row settlement-row data-row-head"><span>Settlement</span><span>Payee</span><span>Type</span><span>Date</span><span>Amount</span><span>Status</span><span>Actions</span></div>
    ${settlements.map((settlement) => `
      <div class="data-row settlement-row">
        <strong>${escapeHtml(settlement.settlement_no)}</strong>
        <span>${escapeHtml(settlement.payee_type === "carrier" ? settlement.carriers?.name : settlement.drivers?.name || "-")}</span>
        <span>${formatStatus(settlement.payee_type)}</span>
        <span>${formatDate(settlement.settlement_date)}</span>
        <strong>${formatMoney(settlement.total_amount)}</strong>
        <select class="workflow-status-select" data-state="${settlement.status}" data-settlement-status="${settlement.id}" aria-label="Change status for ${escapeAttribute(settlement.settlement_no)}" ${["paid", "void"].includes(settlement.status) ? "disabled" : ""}>${settlementStatusOptions(settlement.status).map((status) => `<option value="${status}">${formatStatus(status)}</option>`).join("")}</select>
        <div class="settlement-actions"><button type="button" data-preview-settlement="${settlement.id}">Preview</button></div>
      </div>
    `).join("")}`;
}

function renderSettlementPayBatchRows(batches) {
  if (!batches.length) return '<div class="empty-state">No matching pay batches.</div>';
  return `
    <div class="data-row settlement-pay-batch-row data-row-head"><span>Batch</span><span>Payee</span><span>Period</span><span>Settlements</span><span>Amount</span><span>Status</span><span>Actions</span></div>
    ${batches.map((batch) => `
      <div class="data-row settlement-pay-batch-row">
        <strong>${escapeHtml(batch.batch_no)}</strong>
        <span>${escapeHtml(batch.payee_type === "carrier" ? batch.carriers?.name : batch.drivers?.name || "-")}</span>
        <span>${formatDate(batch.period_start)} - ${formatDate(batch.period_end)}</span>
        <span>${batch.settlement_count} settlement${batch.settlement_count === 1 ? "" : "s"}</span>
        <strong>${formatMoney(batch.total_amount)}</strong>
        <select class="workflow-status-select" data-state="${batch.status}" data-pay-batch-status="${batch.id}" aria-label="Change status for ${escapeAttribute(batch.batch_no)}" ${["paid", "void"].includes(batch.status) ? "disabled" : ""}>${settlementStatusOptions(batch.status).map((status) => `<option value="${status}">${formatStatus(status)}</option>`).join("")}</select>
        <div class="settlement-actions"><button type="button" data-preview-pay-batch="${batch.id}">Preview</button></div>
      </div>
    `).join("")}
  `;
}

function payeeKeyForSettlement(settlement) {
  return `${settlement.payee_type}:${settlement.carrier_id || settlement.driver_id}`;
}

function openSettlementPayBatchDialog() {
  const dialog = document.querySelector("#settlement-pay-batch-dialog");
  const form = document.querySelector("#settlement-pay-batch-form");
  const payeeSelect = document.querySelector("#settlement-pay-batch-payee");
  form.reset();
  document.querySelector("#settlement-pay-batch-message").textContent = "";
  const payeeGroups = new Map();
  settlementPayBatchState.candidates.forEach((settlement) => {
    const key = payeeKeyForSettlement(settlement);
    if (!payeeGroups.has(key)) {
      payeeGroups.set(key, {
        key,
        label: `${formatStatus(settlement.payee_type)} · ${settlement.payee_type === "carrier" ? settlement.carriers?.name : settlement.drivers?.name || "-"}`,
      });
    }
  });
  payeeSelect.innerHTML = [...payeeGroups.values()].map((payee) => `<option value="${payee.key}">${escapeHtml(payee.label)}</option>`).join("");
  renderSettlementPayBatchCandidateList();
  dialog.showModal();
}

function renderSettlementPayBatchCandidateList() {
  const payeeKey = document.querySelector("#settlement-pay-batch-payee")?.value;
  const list = document.querySelector("#settlement-pay-batch-candidate-list");
  if (!list) return;
  const candidates = settlementPayBatchState.candidates.filter((settlement) => payeeKeyForSettlement(settlement) === payeeKey);
  list.innerHTML = candidates.length
    ? candidates.map((settlement) => `<label class="settlement-pay-batch-candidate"><input type="checkbox" name="settlement_ids" value="${settlement.id}" data-amount="${settlement.total_amount}" checked><span><strong>${escapeHtml(settlement.settlement_no)}</strong><small>${formatStatus(settlement.status)} · ${formatDate(settlement.settlement_date)}</small></span><strong>${formatMoney(settlement.total_amount)}</strong></label>`).join("")
    : '<div class="empty-state compact-empty">No payable settlements for this payee.</div>';
  list.querySelectorAll("input").forEach((input) => input.addEventListener("change", updateSettlementPayBatchTotal));
  updateSettlementPayBatchTotal();
}

function updateSettlementPayBatchTotal() {
  const total = Array.from(document.querySelectorAll('#settlement-pay-batch-candidate-list input[name="settlement_ids"]:checked'))
    .reduce((sum, input) => sum + numberOrZero(input.dataset.amount), 0);
  const totalEl = document.querySelector("#settlement-pay-batch-total");
  if (totalEl) totalEl.textContent = formatMoney(total);
}

async function handleCreateSettlementPayBatch(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = document.querySelector("#settlement-pay-batch-message");
  const button = form.querySelector('button[type="submit"]');
  const settlementIds = Array.from(form.querySelectorAll('input[name="settlement_ids"]:checked')).map((input) => input.value);
  if (!settlementIds.length) {
    message.textContent = "Select at least one settlement.";
    return;
  }
  button.disabled = true;
  message.textContent = "Creating pay batch...";
  try {
    const formData = new FormData(form);
    await createSettlementPayBatch({
      period_start: formData.get("period_start") || null,
      period_end: formData.get("period_end") || null,
      pay_date: formData.get("pay_date") || null,
      notes: formData.get("notes") || null,
    }, settlementIds);
    document.querySelector("#settlement-pay-batch-dialog").close();
    await loadSettlementsPage();
  } catch (error) {
    message.textContent = error.message;
  } finally {
    button.disabled = false;
  }
}

async function handleSettlementStatusChange(event) {
  const select = event.currentTarget;
  select.disabled = true;
  try {
    await updateSettlementStatus(select.dataset.settlementStatus, select.value);
    await loadSettlementsPage();
  } catch (error) {
    window.alert(`Settlement update failed: ${error.message}`);
    await loadSettlementsPage();
  }
}

async function handleSettlementPayBatchStatusChange(event) {
  const select = event.currentTarget;
  select.disabled = true;
  try {
    await updateSettlementPayBatchStatus(select.dataset.payBatchStatus, select.value);
    await loadSettlementsPage();
  } catch (error) {
    window.alert(`Pay batch update failed: ${error.message}`);
    await loadSettlementsPage();
  }
}

async function renderSettlementPreview(settlement) {
  let dialog = document.querySelector("#settlement-preview-dialog");
  if (!dialog) {
    dialog = document.createElement("dialog");
    dialog.id = "settlement-preview-dialog";
    dialog.className = "invoice-preview-dialog";
    document.body.append(dialog);
  }
  dialog.innerHTML = '<div class="empty-state">Loading settlement preview...</div>';
  dialog.showModal();
  try {
    const company = await getCompanyBillingSettings();
    const payee = settlement.payee_type === "carrier" ? settlement.carriers?.name : settlement.drivers?.name;
    const items = settlement.settlement_items || [];
    dialog.innerHTML = `
      <div class="invoice-preview-toolbar"><button type="button" data-close-settlement-preview>Close</button><button class="primary-button" type="button" data-print-settlement>Print / Save PDF</button></div>
      <article class="invoice-sheet settlement-sheet">
        <header class="invoice-brand-header">
          <div class="invoice-brand"><img src="assets/cyrra-logo.png" alt="Cyrra logo"><div><strong>${escapeHtml(company.legal_name)}</strong><span>${escapeHtml([company.address_line_1, company.address_line_2].filter(Boolean).join(", "))}</span><small>${escapeHtml([company.city, company.state, company.postal_code].filter(Boolean).join(", "))}</small></div></div>
          <div class="invoice-title"><h1>SETTLEMENT</h1><strong>#${escapeHtml(settlement.settlement_no)}</strong><span>${formatStatus(settlement.status)}</span></div>
        </header>
        <section class="invoice-meta-grid"><div><span>Pay To</span><strong>${escapeHtml(payee || "-")}</strong><p>${formatStatus(settlement.payee_type)}</p></div><div><span>Settlement Date</span><strong>${formatDate(settlement.settlement_date)}</strong><span>Status</span><strong>${formatStatus(settlement.status)}</strong></div></section>
        <table class="invoice-lines"><thead><tr><th>Load / Route</th><th>Amount</th></tr></thead><tbody>${items.map((item) => `<tr><td><strong>Load ${escapeHtml(item.loads?.load_no || "-")}</strong><br><small>${escapeHtml(item.loads ? routeLabel(item.loads) : item.description || "-")}</small></td><td>${formatMoney(item.amount)}</td></tr>`).join("")}</tbody></table>
        <section class="invoice-total"><span>Total Settlement</span><strong>${formatMoney(settlement.total_amount)}</strong></section>
        <footer><p>${escapeHtml(settlement.notes || "Settlement for transportation services.")}</p><small>${escapeHtml(company.email || "")}${company.phone ? ` · ${escapeHtml(company.phone)}` : ""}</small></footer>
      </article>`;
    dialog.querySelector("[data-close-settlement-preview]").addEventListener("click", () => dialog.close());
    dialog.querySelector("[data-print-settlement]").addEventListener("click", () => { document.body.classList.add("invoice-printing"); window.print(); });
  } catch (error) {
    dialog.innerHTML = `<div class="empty-state">Settlement preview unavailable: ${escapeHtml(error.message)}</div><button type="button" data-close-settlement-preview>Close</button>`;
    dialog.querySelector("[data-close-settlement-preview]").addEventListener("click", () => dialog.close());
  }
}

async function renderSettlementPayBatchPreview(batch) {
  let dialog = document.querySelector("#settlement-preview-dialog");
  if (!dialog) {
    dialog = document.createElement("dialog");
    dialog.id = "settlement-preview-dialog";
    dialog.className = "invoice-preview-dialog";
    document.body.append(dialog);
  }
  dialog.innerHTML = '<div class="empty-state">Loading pay batch preview...</div>';
  dialog.showModal();
  try {
    const company = await getCompanyBillingSettings();
    const payee = batch.payee_type === "carrier" ? batch.carriers?.name : batch.drivers?.name;
    const items = batch.settlement_pay_batch_items || [];
    dialog.innerHTML = `
      <div class="invoice-preview-toolbar"><button type="button" data-close-pay-batch-preview>Close</button><button class="primary-button" type="button" data-print-pay-batch>Print / Save PDF</button></div>
      <article class="invoice-sheet settlement-sheet">
        <header class="invoice-brand-header">
          <div class="invoice-brand"><img src="assets/cyrra-logo.png" alt="Cyrra logo"><div><strong>${escapeHtml(company.legal_name)}</strong><span>${escapeHtml([company.address_line_1, company.address_line_2].filter(Boolean).join(", "))}</span><small>${escapeHtml([company.city, company.state, company.postal_code].filter(Boolean).join(", "))}</small></div></div>
          <div class="invoice-title"><h1>PAY BATCH</h1><strong>#${escapeHtml(batch.batch_no)}</strong><span>${formatStatus(batch.status)}</span></div>
        </header>
        <section class="invoice-meta-grid"><div><span>Pay To</span><strong>${escapeHtml(payee || "-")}</strong><p>${formatStatus(batch.payee_type)}</p></div><div><span>Pay Date</span><strong>${formatDate(batch.pay_date)}</strong><span>Period</span><strong>${formatDate(batch.period_start)} - ${formatDate(batch.period_end)}</strong></div></section>
        <table class="invoice-lines"><thead><tr><th>Settlement / Loads</th><th>Amount</th></tr></thead><tbody>${items.map((item) => `<tr><td><strong>${escapeHtml(item.settlements?.settlement_no || "Settlement")}</strong><br><small>${escapeHtml((item.settlements?.settlement_items || []).map((settlementItem) => settlementItem.loads?.load_no ? `Load ${settlementItem.loads.load_no}` : "").filter(Boolean).join(", ") || formatStatus(item.settlements?.status || ""))}</small></td><td>${formatMoney(item.amount_snapshot)}</td></tr>`).join("")}</tbody></table>
        <section class="invoice-total"><span>Total Pay Batch</span><strong>${formatMoney(batch.total_amount)}</strong></section>
        <footer><p>${escapeHtml(batch.notes || "Pay batch for approved transportation settlements.")}</p><small>${escapeHtml(company.email || "")}${company.phone ? ` · ${escapeHtml(company.phone)}` : ""}</small></footer>
      </article>`;
    dialog.querySelector("[data-close-pay-batch-preview]").addEventListener("click", () => dialog.close());
    dialog.querySelector("[data-print-pay-batch]").addEventListener("click", () => { document.body.classList.add("invoice-printing"); window.print(); });
  } catch (error) {
    dialog.innerHTML = `<div class="empty-state">Pay batch preview unavailable: ${escapeHtml(error.message)}</div><button type="button" data-close-pay-batch-preview>Close</button>`;
    dialog.querySelector("[data-close-pay-batch-preview]").addEventListener("click", () => dialog.close());
  }
}

function renderReports() {
  reportsActiveTab = "overview";
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Phase 2", title: "Reports / Analytics" })}
    <section id="reports-root" class="reports-root">
      <div class="empty-state">Loading reports...</div>
    </section>
  `;

  loadReports();
}

async function loadReports() {
  const root = document.querySelector("#reports-root");

  try {
    const [customers, carriers, drivers, trucks, trailers, loads, invoices, profitability, savedViews,scheduledReports,deliveryLogs] = await Promise.all([
      listCustomers(),
      listCarriers(),
      listDrivers(),
      listTrucks(),
      listTrailers(),
      listLoads(),
      listInvoices(),
      listLoadProfitability(),
      listSavedReportViews(),
      listScheduledReports(),
      listReportDeliveryLogs(),
    ]);
    const statusCounts = countBy(loads.map((load) => ({ status: operationalLoadStatus(load) })), "status");
    const billingStatusCounts = countBy(loads.map((load) => ({ status: billingLoadStatus(load) })), "status");
    const today = localDateIso();
    const invoiceReport = invoices.map((invoice) => ({ ...invoice, ...invoiceFinancialState(invoice, today) }));
    const nonVoidInvoices = invoiceReport.filter((invoice) => invoice.status !== "void");
    const receivables = invoiceReport.filter((invoice) => invoice.collectible);
    const followUpInvoices = nonVoidInvoices
      .filter((invoice) => invoice.balance > 0)
      .sort((a, b) => b.daysPastDue - a.daysPastDue || String(a.due_date || "9999").localeCompare(String(b.due_date || "9999")));
    const totalInvoiceAmount = sumBy(nonVoidInvoices, "amount");
    const collectedAmount = sumBy(nonVoidInvoices, "paid");
    const openReceivables = sumBy(receivables, "balance");
    const overdueAmount = sumBy(receivables.filter((invoice) => invoice.daysPastDue > 0), "balance");
    const draftAmount = sumBy(invoiceReport.filter((invoice) => invoice.status === "draft"), "balance");
    const partialCount = receivables.filter((invoice) => invoice.paid > 0).length;
    const agingBuckets = [
      { label: "Current", invoices: receivables.filter((invoice) => invoice.daysPastDue === 0) },
      { label: "1–30 days", invoices: receivables.filter((invoice) => invoice.daysPastDue >= 1 && invoice.daysPastDue <= 30) },
      { label: "31–60 days", invoices: receivables.filter((invoice) => invoice.daysPastDue >= 31 && invoice.daysPastDue <= 60) },
      { label: "61+ days", invoices: receivables.filter((invoice) => invoice.daysPastDue >= 61) },
    ];
    const operationalAnalytics = operationalKpiSummary(loads, profitability);

    root.innerHTML = `
      <nav class="page-tabs report-tabs" aria-label="Report sections">
        ${[["overview","Overview"],["financial","Financial"],["operations","Operations"],["builder","Report Builder"]].map(([value,label])=>`<button type="button" data-report-tab="${value}">${label}</button>`).join("")}
      </nav>
      <section class="report-kpis" data-report-tab-panel="overview">
        ${renderReportKpi("Customers", customers.length)}
        ${renderReportKpi("Carriers", carriers.length)}
        ${renderReportKpi("Drivers", drivers.length)}
        ${renderReportKpi("Trucks", trucks.length)}
        ${renderReportKpi("Trailers", trailers.length)}
        ${renderReportKpi("Loads", loads.length)}
        ${renderReportKpi("Non-void invoice total", formatMoney(totalInvoiceAmount))}
        ${renderReportKpi("Collected", formatMoney(collectedAmount))}
        ${renderReportKpi("Open A/R", formatMoney(openReceivables))}
        ${renderReportKpi("Overdue balance", formatMoney(overdueAmount))}
        ${renderReportKpi("Draft balance", formatMoney(draftAmount))}
        ${renderReportKpi("Partially paid", partialCount)}
        ${renderReportKpi("Ready to bill", billingStatusCounts.ready_to_bill || 0)}
      </section>
      <section class="panel receivables-panel" data-report-tab-panel="financial">
        <div class="panel-header">
          <div>
            <h2>Accounts Receivable Aging</h2>
            <p>Outstanding collectible balances grouped by days past due.</p>
          </div>
          <a class="button-link" href="#billing">Open Billing</a>
        </div>
        <div class="aging-grid">
          ${agingBuckets.map((bucket) => renderAgingBucket(bucket.label, bucket.invoices)).join("")}
        </div>
      </section>
      <section class="panel" data-report-tab-panel="financial">
        <div class="panel-header">
          <div>
            <h2>Invoice Follow-up</h2>
            <p>Drafts and outstanding invoices that still require billing or collection action.</p>
          </div>
        </div>
        <div class="data-table">
          ${renderReceivablesTable(followUpInvoices)}
        </div>
      </section>
      <section class="panel profitability-panel" data-report-tab-panel="financial">
        <div class="panel-header">
          <div>
            <h2>Load Profitability</h2>
            <p>Revenue, carrier cost, driver pay and gross margin by load.</p>
          </div>
        </div>
        <form id="profitability-filters" class="profitability-filters">
          <label><span>Customer</span><select name="customer"><option value="">All customers</option>${customers.map((customer) => `<option value="${customer.id}">${escapeHtml(customer.name)}</option>`).join("")}</select></label>
          <label><span>Driver</span><select name="driver"><option value="">All drivers</option>${drivers.map((driver) => `<option value="${driver.id}">${escapeHtml(driver.name)}</option>`).join("")}</select></label>
          <label><span>Pickup from</span><input name="date_from" type="date"></label>
          <label><span>Pickup to</span><input name="date_to" type="date"></label>
          <label class="checkbox-row profitability-loss-filter"><input name="loss_only" type="checkbox"><span>Negative margin only</span></label>
          <button type="button" id="clear-profitability-filters">Clear</button>
        </form>
        <div id="profitability-results"></div>
      </section>
      <section class="panel profitability-pivot-panel" data-report-tab-panel="financial">
        <div class="panel-header"><div><h2>Profitability Pivot</h2><p>Aggregate load count, revenue, cost and margin by a business dimension.</p></div><div class="panel-header-actions"><button type="button" id="preview-profitability-pivot">Preview / PDF</button><button type="button" id="export-profitability-pivot-xlsx">Export XLSX</button><button type="button" id="export-profitability-pivot">Export CSV</button></div></div>
        <form id="profitability-pivot-controls" class="profitability-filters"><label><span>Group by</span><select name="group_by"><option value="customer">Customer</option><option value="driver">Driver</option><option value="status">Status</option></select></label><label><span>Sort by</span><select name="sort_by"><option value="margin">Margin</option><option value="revenue">Revenue</option><option value="cost">Cost</option><option value="loads">Load count</option><option value="margin_percent">Margin %</option></select></label></form>
        <div id="profitability-pivot-results" class="data-table"></div>
      </section>
      <section class="panel operational-analytics-panel" data-report-tab-panel="operations">
        <div class="panel-header"><div><h2>Operational Performance</h2><p>Service, mileage efficiency and detention signals calculated from recorded load and stop activity.</p></div></div>
        <section class="report-kpis operational-kpi-grid">
          ${renderReportKpi("On-time stops", operationalAnalytics.on_time_percent===null?"No actuals":formatPercent(operationalAnalytics.on_time_percent), `${operationalAnalytics.measured_stops} measured`)}
          ${renderReportKpi("Loaded-mile utilization", operationalAnalytics.utilization_percent===null?"No mileage":formatPercent(operationalAnalytics.utilization_percent))}
          ${renderReportKpi("Revenue / loaded mile", operationalAnalytics.revenue_per_loaded_mile===null?"No mileage":formatMoney(operationalAnalytics.revenue_per_loaded_mile))}
          ${renderReportKpi("Detention exposure", `${operationalAnalytics.detention_hours.toFixed(1)} h`, `${operationalAnalytics.detention_stops} stops over 2h`)}
        </section>
        <div class="driver-performance"><h3>Driver on-time delivery</h3>${renderDriverPerformance(operationalAnalytics.driver_performance)}</div>
      </section>
      <section class="panel report-builder-panel" data-report-tab-panel="builder">
        <div class="panel-header"><div><h2>Report Builder</h2><p>Choose columns and filters, save the view, and export the current result.</p></div></div>
        <form id="report-builder-form" class="record-form"><label><span>Saved view</span><select name="saved_view"><option value="">New report</option>${savedViews.map((view)=>`<option value="${view.id}">${escapeHtml(view.name)}</option>`).join("")}</select></label><label><span>Report name</span><input name="name" required placeholder="Weekly Profitability"></label>
        <div class="report-column-picker">${[["load_no","Load"],["customer","Customer"],["driver","Driver"],["status","Status"],["pickup_date","Pickup"],["revenue","Revenue"],["cost","Cost"],["margin","Margin"],["margin_percent","Margin %"]].map(([value,label])=>`<label class="checkbox-row"><input type="checkbox" name="columns" value="${value}" ${["load_no","customer","revenue","cost","margin"].includes(value)?"checked":""}><span>${label}</span></label>`).join("")}</div>
        <div class="report-builder-filters"><label><span>Customer</span><select name="customer_id"><option value="">All customers</option>${customers.map((item)=>`<option value="${item.id}">${escapeHtml(item.name)}</option>`).join("")}</select></label><label><span>Driver</span><select name="driver_id"><option value="">All drivers</option>${drivers.map((item)=>`<option value="${item.id}">${escapeHtml(item.name)}</option>`).join("")}</select></label><label><span>Status</span><select name="status"><option value="">All statuses</option>${dispatchStatuses.map((item)=>`<option value="${item}">${formatStatus(item)}</option>`).join("")}</select></label><label><span>Group by</span><select name="group_by"><option value="">No grouping</option><option value="customer">Customer</option><option value="driver">Driver</option><option value="status">Status</option></select></label><label><span>Pickup from</span><input name="date_from" type="date"></label><label><span>Pickup to</span><input name="date_to" type="date"></label></div>
        <div class="form-actions"><button type="submit">Save View</button><button id="export-report-builder" type="button">Export CSV</button><button id="delete-report-view" type="button" disabled>Delete View</button></div><p class="form-message"></p></form>
        <div id="report-builder-results" class="data-table"></div>
      </section>
      <section class="panel scheduled-reports-panel" data-report-tab-panel="builder">
        <div class="panel-header"><div><h2>Scheduled Reports</h2><p>Save delivery cadence now; Run Now generates the CSV and records a delivery log.</p></div></div>
        <form id="scheduled-report-form" class="record-form compact-form"><label><span>Schedule name</span><input name="name" required placeholder="Weekly operations"></label><label><span>Saved view</span><select name="report_view_id" required><option value="">Select saved view</option>${savedViews.map((view)=>`<option value="${view.id}">${escapeHtml(view.name)}</option>`).join("")}</select></label><label><span>Frequency</span><select name="frequency"><option value="daily">Daily</option><option value="weekly">Weekly</option><option value="monthly">Monthly</option></select></label><label><span>Next run</span><input name="next_run_at" type="datetime-local" required></label><label><span>Recipients</span><input name="recipients" type="text" placeholder="ops@example.com, finance@example.com" required></label><label class="checkbox-row"><input name="active" type="checkbox" checked><span>Active</span></label><button type="submit" ${savedViews.length?"":"disabled"}>Save Schedule</button><p class="form-message"></p></form>
        <div id="scheduled-report-list" class="scheduled-report-list">${renderScheduledReports(scheduledReports,deliveryLogs)}</div>
      </section>
      <section class="report-grid" data-report-tab-panel="overview">
        <article class="panel">
          <div class="panel-header">
            <div>
              <h2>Loads by Status</h2>
              <p>Operational distribution across the dispatch pipeline.</p>
            </div>
          </div>
          <div class="report-bars">
            ${dispatchStatuses.map((status) => renderStatusBar(status, statusCounts[status] || 0, loads.length)).join("")}
          </div>
        </article>
        <article class="panel">
          <div class="panel-header">
            <div>
              <h2>Invoice Status</h2>
              <p>Billing status totals and collection focus.</p>
            </div>
          </div>
          <div class="report-bars">
            ${invoiceStatuses.map((status) => renderStatusBar(status, countBy(invoices, "status")[status] || 0, invoices.length)).join("")}
          </div>
        </article>
      </section>
      <section class="panel" data-report-tab-panel="overview">
        <div class="panel-header">
          <div>
            <h2>Recent Loads</h2>
            <p>Latest operational records for quick review.</p>
          </div>
        </div>
        <div class="data-table">
          ${renderRecentLoads(loads.slice(0, 8))}
        </div>
      </section>
    `;
    initializeProfitabilityReport(profitability);
    initializeProfitabilityPivot(profitability);
    initializeReportBuilder(profitability,savedViews);
    initializeScheduledReports(scheduledReports);
    initializeReportTabs();
  } catch (error) {
    root.innerHTML = `<div class="empty-state">Reports are not ready: ${escapeHtml(error.message)}</div>`;
  }
}

let profitabilityRecords = [];

function initializeProfitabilityReport(records) {
  profitabilityRecords = records.map((record) => ({ ...record, ...calculateLoadProfitability(record) }));
  const form = document.querySelector("#profitability-filters");
  form.addEventListener("input", renderProfitabilityResults);
  form.addEventListener("change", renderProfitabilityResults);
  document.querySelector("#clear-profitability-filters").addEventListener("click", () => {
    form.reset();
    renderProfitabilityResults();
  });
  renderProfitabilityResults();
}

function renderProfitabilityResults() {
  const form = document.querySelector("#profitability-filters");
  const results = document.querySelector("#profitability-results");
  if (!form || !results) return;
  const data = new FormData(form);
  const customerId = String(data.get("customer") || "");
  const driverId = String(data.get("driver") || "");
  const dateFrom = String(data.get("date_from") || "");
  const dateTo = String(data.get("date_to") || "");
  const lossOnly = data.get("loss_only") === "on";
  const filtered = profitabilityRecords.filter((record) => {
    const load = record.loads;
    if (customerId && load?.customer_id !== customerId) return false;
    if (driverId && load?.driver_id !== driverId) return false;
    if (dateFrom && (!load?.pickup_date || load.pickup_date < dateFrom)) return false;
    if (dateTo && (!load?.pickup_date || load.pickup_date > dateTo)) return false;
    if (lossOnly && record.margin >= 0) return false;
    return true;
  });
  const revenue = sumBy(filtered, "revenue");
  const cost = sumBy(filtered, "cost");
  const margin = sumBy(filtered, "margin");
  const marginPercent = revenue ? (margin / revenue) * 100 : 0;
  const negativeCount = filtered.filter((record) => record.margin < 0).length;
  results.innerHTML = `
    <section class="profitability-kpis">
      ${renderReportKpi("Revenue", formatMoney(revenue))}
      ${renderReportKpi("Total cost", formatMoney(cost))}
      ${renderReportKpi("Gross margin", formatMoney(margin))}
      ${renderReportKpi("Margin %", formatPercent(marginPercent))}
      ${renderReportKpi("Negative margin loads", negativeCount)}
    </section>
    <div class="data-table profitability-table">${renderProfitabilityRows(filtered)}</div>
  `;
}

let profitabilityPivotRecords=[];
function initializeProfitabilityPivot(records){profitabilityPivotRecords=records;const form=document.querySelector("#profitability-pivot-controls");form.addEventListener("change",renderProfitabilityPivot);document.querySelector("#export-profitability-pivot").addEventListener("click",exportProfitabilityPivot);document.querySelector("#export-profitability-pivot-xlsx").addEventListener("click",exportProfitabilityPivotXlsx);document.querySelector("#preview-profitability-pivot").addEventListener("click",previewProfitabilityPivotPdf);renderProfitabilityPivot();}
function currentProfitabilityPivot(){const data=new FormData(document.querySelector("#profitability-pivot-controls"));return profitabilityPivot(profitabilityPivotRecords,String(data.get("group_by")),String(data.get("sort_by")));}
function renderProfitabilityPivot(){const rows=currentProfitabilityPivot();const root=document.querySelector("#profitability-pivot-results");root.innerHTML=`<div class="data-row profitability-pivot-row data-row-head"><span>Group</span><span>Loads</span><span>Revenue</span><span>Cost</span><span>Margin</span><span>Margin %</span></div>${rows.map((row)=>`<div class="data-row profitability-pivot-row"><strong>${escapeHtml(formatStatus(row.label))}</strong><span>${row.loads}</span><span>${formatMoney(row.revenue)}</span><span>${formatMoney(row.cost)}</span><span>${formatMoney(row.margin)}</span><span>${formatPercent(row.margin_percent)}</span></div>`).join("")||'<div class="empty-state">No profitability data available.</div>'}`;}
function exportProfitabilityPivot(){const rows=currentProfitabilityPivot().map((row)=>({group:row.label,loads:row.loads,revenue:row.revenue,cost:row.cost,margin:row.margin,margin_percent:row.margin_percent}));downloadCsv("cyrra-profitability-pivot.csv",rows);}
function exportProfitabilityPivotXlsx(){try{const rows=currentProfitabilityPivot().map((row)=>({group:row.label,loads:row.loads,revenue:row.revenue,cost:row.cost,margin:row.margin,margin_percent:row.margin_percent}));downloadBlob("cyrra-profitability-pivot.xlsx",buildXlsxBlob(rows,"Profitability Pivot"));}catch(error){window.alert(error.message);}}
function previewProfitabilityPivotPdf(){const form=document.querySelector("#profitability-pivot-controls");const groupLabel=form.elements.group_by.options[form.elements.group_by.selectedIndex].textContent;const sortLabel=form.elements.sort_by.options[form.elements.sort_by.selectedIndex].textContent;const rows=currentProfitabilityPivot();let dialog=document.querySelector("#report-print-dialog");if(!dialog){dialog=document.createElement("dialog");dialog.id="report-print-dialog";dialog.className="invoice-preview-dialog";document.body.append(dialog);}dialog.innerHTML=`<div class="invoice-preview-toolbar"><button type="button" data-close-report-print>Close</button><button class="primary-button" type="button" data-print-report>Print / Save PDF</button></div><article class="invoice-sheet report-print-sheet"><header class="invoice-brand-header"><div class="invoice-brand"><img src="assets/cyrra-logo.png" alt="Cyrra logo"><div><strong>Cyrra TMS</strong><span>Reports / Analytics</span><small>Generated ${formatDateTime(new Date().toISOString())}</small></div></div><div class="invoice-title"><h1>REPORT</h1><strong>Profitability Pivot</strong><span>Grouped by ${escapeHtml(groupLabel)}</span></div></header><section class="report-print-meta"><div><span>Group by</span><strong>${escapeHtml(groupLabel)}</strong></div><div><span>Sort by</span><strong>${escapeHtml(sortLabel)}</strong></div><div><span>Rows</span><strong>${rows.length}</strong></div></section><table class="invoice-lines"><thead><tr><th>Group</th><th>Loads</th><th>Revenue</th><th>Cost</th><th>Margin</th><th>Margin %</th></tr></thead><tbody>${rows.map((row)=>`<tr><td><strong>${escapeHtml(formatStatus(row.label))}</strong></td><td>${row.loads}</td><td>${formatMoney(row.revenue)}</td><td>${formatMoney(row.cost)}</td><td>${formatMoney(row.margin)}</td><td>${formatPercent(row.margin_percent)}</td></tr>`).join("")}</tbody></table><footer><p>Cyrra TMS Profitability Report</p><small>Generated from current staging report data and selected pivot configuration.</small></footer></article>`;dialog.querySelector("[data-close-report-print]").addEventListener("click",()=>dialog.close());dialog.querySelector("[data-print-report]").addEventListener("click",()=>{document.body.classList.add("invoice-printing");window.print();});dialog.showModal();}

function renderProfitabilityRows(records) {
  if (!records.length) return '<div class="empty-state">No load financials match the selected filters.</div>';
  return `
    <div class="data-row profitability-row data-row-head"><span>Load</span><span>Customer</span><span>Pickup</span><span>Revenue</span><span>Carrier</span><span>Driver</span><span>Other cost</span><span>Margin</span></div>
    ${records.map((record) => `
      <div class="data-row profitability-row ${record.margin < 0 ? "is-negative" : ""}">
        <span><strong>${escapeHtml(record.loads?.load_no || "-")}</strong><small>${escapeHtml(record.loads ? routeLabel(record.loads) : "-")}</small></span>
        <span>${escapeHtml(record.loads?.customers?.name || "-")}</span>
        <span>${formatDate(record.loads?.pickup_date)}</span>
        <span>${formatMoney(record.revenue)}</span>
        <span>${formatMoney(record.carrier_cost)}</span>
        <span>${formatMoney(record.driver_pay)}</span>
        <span>${formatMoney(record.deductions)}</span>
        <span><strong>${formatMoney(record.margin)}</strong><small>${formatPercent(record.marginPercent)}</small></span>
      </div>
    `).join("")}
  `;
}

let reportBuilderRecords=[]; let reportBuilderViews=[];
function initializeReportTabs(){const buttons=[...document.querySelectorAll("[data-report-tab]")];const panels=[...document.querySelectorAll("[data-report-tab-panel]")];const activate=(tab)=>{reportsActiveTab=tab;buttons.forEach((button)=>{const active=button.dataset.reportTab===tab;button.dataset.active=String(active);button.setAttribute("aria-selected",String(active));});panels.forEach((panel)=>panel.hidden=panel.dataset.reportTabPanel!==tab);};buttons.forEach((button)=>button.addEventListener("click",()=>activate(button.dataset.reportTab)));activate(reportsActiveTab);}
function initializeReportBuilder(records,views){reportBuilderRecords=records.map((r)=>({...r,...calculateLoadProfitability(r)}));reportBuilderViews=views;const form=document.querySelector("#report-builder-form");form.addEventListener("input",renderReportBuilderResults);form.elements.saved_view.addEventListener("change",applySavedReportView);form.addEventListener("submit",handleSaveReportView);document.querySelector("#export-report-builder").addEventListener("click",exportReportBuilderCsv);document.querySelector("#delete-report-view").addEventListener("click",handleDeleteReportView);renderReportBuilderResults();}
function reportViewFromForm(){const form=document.querySelector("#report-builder-form");const data=new FormData(form);return validateSavedReportView({name:data.get("name")||"Unsaved Report",columns:data.getAll("columns"),filters:{customer_id:data.get("customer_id")||"",driver_id:data.get("driver_id")||"",status:data.get("status")||"",date_from:data.get("date_from")||"",date_to:data.get("date_to")||""},group_by:data.get("group_by")||null,sort_by:"pickup_date",sort_direction:"desc"});}
function renderReportBuilderResults(){const root=document.querySelector("#report-builder-results");if(!root)return;try{const view=reportViewFromForm();const rows=applyReportView(reportBuilderRecords,view);root.innerHTML=`<div class="report-builder-summary"><strong>${rows.length} rows</strong><span>${view.group_by?`Grouped by ${formatStatus(view.group_by)}`:"Ungrouped"}</span></div><div class="data-row report-builder-row report-builder-head" style="--report-columns:${view.columns.length}">${view.columns.map((c)=>`<span>${formatStatus(c)}</span>`).join("")}</div>${rows.map((r)=>`<div class="data-row report-builder-row" style="--report-columns:${view.columns.length}">${view.columns.map((c)=>`<span>${["revenue","cost","margin"].includes(c)?formatMoney(reportRecordValue(r,c)):c==="margin_percent"?formatPercent(reportRecordValue(r,c)):escapeHtml(formatStatus(reportRecordValue(r,c)||"-"))}</span>`).join("")}</div>`).join("")||'<div class="empty-state">No rows match this report.</div>'}`;}catch(error){root.innerHTML=`<div class="empty-state">${escapeHtml(error.message)}</div>`;}}
function applySavedReportView(event){const form=event.currentTarget.form;const view=reportBuilderViews.find((v)=>v.id===event.target.value);document.querySelector("#delete-report-view").disabled=!view;if(!view)return;form.elements.name.value=view.name;form.querySelectorAll('[name="columns"]').forEach((input)=>input.checked=view.columns.includes(input.value));for(const [key,value] of Object.entries(view.filters||{})){if(form.elements[key])form.elements[key].value=value||"";}form.elements.group_by.value=view.group_by||"";renderReportBuilderResults();}
async function handleSaveReportView(event){event.preventDefault();const form=event.currentTarget;const message=form.querySelector(".form-message");try{const view=reportViewFromForm();const branchId=currentAccess.find((access)=>access.branch_id)?.branch_id||reportBuilderRecords[0]?.loads?.branch_id;if(!branchId)throw new Error("A branch is required to save this report view.");const payload={...view,branch_id:branchId,dataset:"load_profitability"};if(form.elements.saved_view.value)payload.id=form.elements.saved_view.value;const saved=await saveReportView(payload);await loadReports();const savedViewSelect=document.querySelector('#report-builder-form [name="saved_view"]');savedViewSelect.value=saved.id;savedViewSelect.dispatchEvent(new Event("change",{bubbles:true}));const refreshedMessage=document.querySelector("#report-builder-form .form-message");refreshedMessage.textContent="Report view saved.";}catch(error){message.textContent=error.message;}}
async function handleDeleteReportView(){const form=document.querySelector("#report-builder-form");if(!form.elements.saved_view.value)return;try{await deleteSavedReportView(form.elements.saved_view.value);await loadReports();}catch(error){window.alert(error.message);}}
function exportReportBuilderCsv(){try{const view=reportViewFromForm();const rows=applyReportView(reportBuilderRecords,view).map((r)=>Object.fromEntries(view.columns.map((c)=>[c,reportRecordValue(r,c)??""])));downloadCsv(`cyrra-${view.name.toLowerCase().replace(/[^a-z0-9]+/g,"-")}.csv`,rows);}catch(error){window.alert(error.message);}}

let scheduledReportRecords=[];
function renderScheduledReports(schedules,logs){if(!schedules.length)return '<div class="empty-state">No scheduled reports yet. Save a report view first, then create a schedule.</div>';return schedules.map((schedule)=>{const latest=logs.find((log)=>log.scheduled_report_id===schedule.id);return `<article class="scheduled-report-card"><div><strong>${escapeHtml(schedule.name)}</strong><span>${formatStatus(schedule.frequency)} · CSV · ${schedule.active?"Active":"Paused"}</span><small>${escapeHtml(schedule.saved_report_views?.name||"Saved view")} · next ${formatDateTime(schedule.next_run_at)}</small><small>${schedule.recipients.map(escapeHtml).join(", ")}</small>${latest?`<small>Last log: ${formatStatus(latest.status)} · ${latest.row_count} rows · ${formatDateTime(latest.created_at)}</small>`:""}${schedule.last_error?`<small class="status-text status-text-danger">Last error: ${escapeHtml(schedule.last_error)}</small>`:""}</div><div><button type="button" data-run-scheduled-report="${schedule.id}">Run Now</button><button type="button" data-delete-scheduled-report="${schedule.id}">Delete</button></div></article>`;}).join("");}
function initializeScheduledReports(schedules){scheduledReportRecords=schedules;const form=document.querySelector("#scheduled-report-form");form.addEventListener("submit",handleSaveScheduledReport);document.querySelectorAll("[data-run-scheduled-report]").forEach((button)=>button.addEventListener("click",()=>handleRunScheduledReport(button.dataset.runScheduledReport)));document.querySelectorAll("[data-delete-scheduled-report]").forEach((button)=>button.addEventListener("click",()=>handleDeleteScheduledReport(button.dataset.deleteScheduledReport)));}
async function handleSaveScheduledReport(event){event.preventDefault();const form=event.currentTarget;const message=form.querySelector(".form-message");try{const input=validateScheduledReport(Object.fromEntries(new FormData(form)));const branchId=currentAccess.find((access)=>access.branch_id)?.branch_id||reportBuilderRecords[0]?.loads?.branch_id;if(!branchId)throw new Error("A branch is required to schedule reports.");await saveScheduledReport({...input,branch_id:branchId});reportsActiveTab="builder";await loadReports();}catch(error){message.textContent=error.message;}}
async function handleRunScheduledReport(id){const schedule=scheduledReportRecords.find((item)=>item.id===id);const view=reportBuilderViews.find((item)=>item.id===schedule?.report_view_id);if(!schedule||!view)return;const rows=applyReportView(reportBuilderRecords,view).map((record)=>Object.fromEntries(view.columns.map((column)=>[column,reportRecordValue(record,column)??""])));const fileName=`cyrra-${schedule.name.toLowerCase().replace(/[^a-z0-9]+/g,"-")}.csv`;try{downloadCsv(fileName,rows);const now=new Date().toISOString();await recordReportDelivery({branch_id:schedule.branch_id,scheduled_report_id:schedule.id,status:"generated",export_format:"csv",row_count:rows.length,file_name:fileName});await updateScheduledReportRun(schedule.id,now,nextScheduledReportRun(schedule.next_run_at,schedule.frequency));reportsActiveTab="builder";await loadReports();}catch(error){window.alert(error.message);}}
async function handleDeleteScheduledReport(id){try{await deleteScheduledReport(id);reportsActiveTab="builder";await loadReports();}catch(error){window.alert(error.message);}}

function renderUserAuditEvent(event) {
  const email = event.new_value?.email || event.old_value?.email || "";
  return `
    <article>
      <div><strong>${formatStatus(event.action)}</strong>${email ? `<span>${escapeHtml(email)}</span>` : ""}</div>
      <small>${formatDateTime(event.created_at)} · Actor ${escapeHtml(event.user_id || "system")}</small>
    </article>
  `;
}

function renderReportKpi(label, value, detail="") {
  return `
    <article class="report-kpi">
      <span>${label}</span>
      <strong>${value}</strong>
      ${detail?`<small>${escapeHtml(detail)}</small>`:""}
    </article>
  `;
}

function renderDriverPerformance(drivers){
  if(!drivers.length)return '<div class="empty-state">No assigned driver activity yet.</div>';
  return `<div class="report-bars">${drivers.map((driver)=>{const measured=driver.on_time_percent!==null;const percent=measured?Math.round(driver.on_time_percent):0;return `<div class="report-bar"><div><span>${escapeHtml(driver.name)}</span><strong>${measured?formatPercent(driver.on_time_percent):"No delivery actuals"}</strong></div><i><b style="width:${percent}%"></b></i><small>${driver.loads} loads · ${driver.measured} measured deliveries</small></div>`;}).join("")}</div>`;
}

function renderAgingBucket(label, invoices) {
  return `
    <article>
      <span>${escapeHtml(label)}</span>
      <strong>${formatMoney(sumBy(invoices, "balance"))}</strong>
      <small>${invoices.length} invoice${invoices.length === 1 ? "" : "s"}</small>
    </article>
  `;
}

function renderReceivablesTable(invoices) {
  if (!invoices.length) {
    return '<div class="empty-state">No outstanding invoices. Accounts receivable is clear.</div>';
  }

  return `
    <div class="data-row receivable-row data-row-head">
      <span>Invoice</span><span>Customer</span><span>Status</span><span>Due</span><span>Original</span><span>Paid</span><span>Balance / Aging</span>
    </div>
    ${invoices.map((invoice) => {
      const aging = invoice.status === "draft"
        ? "Draft — not sent"
        : invoice.daysPastDue > 0
          ? `${invoice.daysPastDue} day${invoice.daysPastDue === 1 ? "" : "s"} overdue`
          : "Current";
      return `
        <div class="data-row receivable-row">
          <strong>${escapeHtml(invoice.invoice_no)}</strong>
          <span>${escapeHtml(invoice.customers?.name || "-")}</span>
          <span>${formatStatus(invoice.status)}</span>
          <span>${formatDate(invoice.due_date)}</span>
          <span>${formatMoney(invoice.amount)}</span>
          <span>${formatMoney(invoice.paid)}</span>
          <span><strong>${formatMoney(invoice.balance)}</strong><small class="invoice-aging ${invoice.daysPastDue > 0 ? "is-overdue" : ""}">${escapeHtml(aging)}</small></span>
        </div>
      `;
    }).join("")}
  `;
}

function renderStatusBar(label, count, total) {
  const percent = total ? Math.round((count / total) * 100) : 0;

  return `
    <div class="report-bar">
      <div>
        <span>${formatStatus(label)}</span>
        <strong>${count}</strong>
      </div>
      <i><b style="width: ${percent}%"></b></i>
    </div>
  `;
}

function renderRecentLoads(loads) {
  if (!loads.length) {
    return '<div class="empty-state">No loads yet.</div>';
  }

  return `
    <div class="data-row report-load-row data-row-head">
      <span>Load #</span>
      <span>Status</span>
      <span>Customer</span>
      <span>Route</span>
      <span>Pickup</span>
    </div>
    ${loads
      .map(
        (load) => `
          <div class="data-row report-load-row">
            <strong>${escapeHtml(load.load_no)}</strong>
            <span>${formatStatus(operationalLoadStatus(load))}</span>
            <span>${escapeHtml(load.customers?.name || "-")}</span>
            <span>${escapeHtml(routeLabel(load))}</span>
            <span>${formatDate(load.pickup_date)}</span>
          </div>
        `
      )
      .join("")}
  `;
}

async function renderNotifications() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Back Office", title: "Notification Center", status: "Loading inbox" })}
    <section id="notification-root" class="notification-root">
      <div class="empty-state">Loading notifications...</div>
    </section>
  `;
  const statusEl = document.querySelector("#connection-status");
  const root = document.querySelector("#notification-root");
  try {
    const [notifications, preferences, deliveryLogs, branches] = await Promise.all([
      listNotifications(100),
      listNotificationPreferences(currentSession.user.id),
      listNotificationDeliveryLogs(100),
      listBranches(),
    ]);
    const summary = notificationSummary(notifications);
    const deliverySummary = notificationDeliverySummary(deliveryLogs);
    statusEl.dataset.state = summary.unread ? "warning" : "ok";
    statusEl.textContent = `${summary.unread} unread`;
    root.innerHTML = `
      <section class="notification-summary">
        <article><span>Unread</span><strong>${summary.unread}</strong></article>
        <article><span>Active</span><strong>${summary.total}</strong></article>
        <article><span>Critical</span><strong>${summary.critical}</strong></article>
        <article><span>Delivery</span><strong>${escapeHtml(formatStatus(deliverySummary.state))}</strong></article>
      </section>
      <article class="panel notification-inbox-panel">
        <div class="panel-header">
          <div><h2>Inbox</h2><p>Branch-scoped in-app alerts for work that needs attention.</p></div>
          <div class="table-tools">
            <button id="mark-all-notifications-read" type="button" ${summary.unread ? "" : "disabled"}>Mark all read</button>
            <button id="reload-notifications" type="button">Refresh</button>
            <button id="open-notification-create" class="primary-button" type="button" ${notificationBranchOptions(branches).length ? "" : "disabled"}>Create Notice</button>
          </div>
        </div>
        <p id="notification-message" class="form-message">${escapeHtml(notificationNotice)}</p>
        <div class="notification-list">
          ${notifications.length ? notifications.map(renderNotificationItem).join("") : '<div class="empty-state compact-empty">No notifications yet.</div>'}
        </div>
      </article>
      <section class="notification-settings-grid">
        ${renderNotificationPreferences(preferences)}
        ${renderNotificationDelivery(deliveryLogs, deliverySummary)}
      </section>
      ${renderNotificationComposer(branches)}
    `;
    bindNotificationEvents(notifications, preferences);
    refreshNotificationBadge();
  } catch (error) {
    statusEl.dataset.state = "error";
    statusEl.textContent = "Unavailable";
    root.innerHTML = `<div class="empty-state">Notification Center is not ready: ${escapeHtml(error.message)}</div>`;
  }
}

function notificationBranchOptions(allBranches = []) {
  const assignedBranches = currentAccess
    .filter((access) => access.branch_id)
    .map((access) => ({ id: access.branch_id, name: access.branches?.name || "Assigned branch" }));
  if (assignedBranches.length) {
    return [...new Map(assignedBranches.map((branch) => [branch.id, branch])).values()];
  }
  if (currentRoles.includes("owner")) {
    return allBranches;
  }
  return [];
}

function renderNotificationComposer(allBranches = []) {
  const branches = notificationBranchOptions(allBranches);
  if (!branches.length) {
    return `
      <dialog id="notification-create-dialog" class="crud-dialog notification-create-dialog"><article class="panel notification-compose-panel">
        <div class="panel-header"><div><h2>Create Notice</h2><p>Needs branch-assigned access for manual notices.</p></div><button class="dialog-close" type="button" data-close-notification-create aria-label="Close">×</button></div>
        <div class="empty-state compact-empty">No branch-specific access is available in this session.</div>
      </article></dialog>`;
  }
  return `
    <dialog id="notification-create-dialog" class="crud-dialog notification-create-dialog"><article class="panel notification-compose-panel">
      <div class="panel-header"><div><h2>Create Notice</h2><p>Manual in-app event for operational follow-up.</p></div><button class="dialog-close" type="button" data-close-notification-create aria-label="Close">×</button></div>
      <form id="notification-create-form" class="record-form">
        <label><span>Branch</span><select name="branch_id">${branches.map((branch) => `<option value="${branch.id}">${escapeHtml(branch.name)}</option>`).join("")}</select></label>
        <label><span>Event</span><select name="event_type">${notificationEventTypes.map((type) => `<option value="${type}">${escapeHtml(formatStatus(type))}</option>`).join("")}</select></label>
        <label><span>Severity</span><select name="severity"><option value="info">Info</option><option value="warning">Warning</option><option value="critical">Critical</option></select></label>
        <label><span>Title</span><input name="title" maxlength="160" required></label>
        <label><span>Body</span><textarea name="body" maxlength="1200" rows="3" required></textarea></label>
        <label><span>Action URL</span><input name="action_url" placeholder="#loads"></label>
        <div class="form-actions"><button type="submit">Create Notice</button></div>
        <p id="notification-create-message" class="form-message"></p>
      </form>
    </article></dialog>`;
}

function renderNotificationItem(notification) {
  const notificationLoadId = notification.metadata?.load_id
    || (["load", "loads"].includes(notification.entity_type) ? notification.entity_id : "");
  const actionUrl = notificationLoadId
    ? `#loads/${encodeURIComponent(notificationLoadId)}`
    : notification.action_url;
  return `
    <article class="notification-item" data-notification-id="${notification.id}" data-status="${notification.status}" data-severity="${notification.severity}">
      <div>
        <header>
          <strong>${escapeHtml(notification.title)}</strong>
          <span>${escapeHtml(formatStatus(notification.event_type))}</span>
        </header>
        <p>${escapeHtml(notification.body)}</p>
        <small>${escapeHtml(formatStatus(notification.severity))} · ${formatDateTime(notification.created_at)}${notification.read_at ? ` · Read ${formatDateTime(notification.read_at)}` : ""}</small>
      </div>
      <div class="notification-actions">
        ${actionUrl ? `<a class="button-link" href="${escapeAttribute(actionUrl)}">Open</a>` : ""}
        ${notification.status === "unread" ? '<button type="button" data-mark-notification-read>Mark read</button>' : ""}
      </div>
    </article>`;
}

function renderNotificationPreferences(preferences) {
  const eventDescriptions = {
    assignment: "Load and resource assignment changes.",
    load_message: "Internal notes added to a load conversation.",
    missing_pod: "Delivered loads waiting on POD review.",
    late_load: "Pickup, delivery or ETA risk needs attention.",
    compliance_expiry: "Driver or equipment compliance dates are near or past due.",
    maintenance_expiry: "Scheduled maintenance is due or overdue.",
    unpaid_invoice: "Open invoice or receivable follow-up.",
    system: "Platform and workflow health notices.",
  };
  return `
    <article class="panel notification-preferences-panel">
      <div class="panel-header"><div><h2>Preferences</h2><p>In-app is active now; other channels are logged server-side until adapters are added.</p></div></div>
      <form id="notification-preferences-form" class="notification-preferences-form">
        <div class="notification-preference-list">
          ${notificationEventTypes.map((eventType) => {
            const preference = preferences.find((item) => item.event_type === eventType && !item.branch_id) || {};
            return `
              <label class="notification-preference-row" data-preference-id="${preference.id || ""}" data-event-type="${eventType}">
                <span class="preference-copy">
                  <strong>${escapeHtml(formatStatus(eventType))}</strong>
                  <small>${escapeHtml(eventDescriptions[eventType] || "Notification type")}</small>
                </span>
                <span class="preference-switch">
                  <input name="${eventType}" type="checkbox" ${preference.in_app_enabled ?? true ? "checked" : ""}>
                  <i></i>
                </span>
              </label>`;
          }).join("")}
        </div>
        <div class="notification-preferences-footer">
          <button type="submit">Save Preferences</button>
        </div>
      </form>
    </article>`;
}

function renderNotificationDelivery(deliveryLogs, summary) {
  return `
    <article class="panel notification-delivery-panel">
      <div class="panel-header"><div><h2>Delivery Log</h2><p>Server-side adapter status without exposing credentials.</p></div></div>
      <div class="delivery-summary">
        <span data-state="sent"><i></i><em>Sent</em><strong>${summary.sent}</strong></span>
        <span data-state="retry"><i></i><em>Retry</em><strong>${summary.retry + summary.queued}</strong></span>
        <span data-state="failed"><i></i><em>Failed</em><strong>${summary.failed}</strong></span>
        <span data-state="skipped"><i></i><em>Skipped</em><strong>${summary.skipped}</strong></span>
      </div>
      <div class="notification-delivery-list">
        ${deliveryLogs.length ? deliveryLogs.slice(0, 8).map((log) => `
          <div class="notification-delivery-row">
            <strong>${escapeHtml(formatStatus(log.channel))}</strong>
            <span>${escapeHtml(formatStatus(log.status))} · attempts ${log.attempt_count}</span>
            <small>${escapeHtml(log.last_error || formatDateTime(log.created_at))}</small>
          </div>`).join("") : '<div class="notification-delivery-empty">No delivery events yet.</div>'}
      </div>
    </article>`;
}

function bindNotificationEvents(notifications, preferences) {
  document.querySelector("#reload-notifications")?.addEventListener("click", renderNotifications);
  document.querySelector("#open-notification-create")?.addEventListener("click", () => document.querySelector("#notification-create-dialog")?.showModal());
  document.querySelector("[data-close-notification-create]")?.addEventListener("click", () => document.querySelector("#notification-create-dialog")?.close());
  document.querySelector("#mark-all-notifications-read")?.addEventListener("click", async () => {
    const unreadIds = notifications.filter((notification) => notification.status === "unread").map((notification) => notification.id);
    if (unreadIds.length) await handleMarkNotificationsRead(unreadIds);
  });
  document.querySelectorAll("[data-mark-notification-read]").forEach((button) => {
    button.addEventListener("click", async (event) => {
      const item = event.currentTarget.closest("[data-notification-id]");
      await handleMarkNotificationsRead([item.dataset.notificationId]);
    });
  });
  document.querySelector("#notification-create-form")?.addEventListener("submit", handleCreateNotification);
  document.querySelector("#notification-preferences-form")?.addEventListener("submit", (event) => handleSaveNotificationPreferences(event, preferences));
}

async function handleMarkNotificationsRead(notificationIds) {
  const message = document.querySelector("#notification-message");
  try {
    message.textContent = "Updating notifications...";
    const count = await markNotificationsRead(notificationIds);
    notificationNotice = `${count} notification${count === 1 ? "" : "s"} marked read.`;
    renderNotifications();
  } catch (error) {
    message.dataset.state = "error";
    message.textContent = error.message;
  }
}

async function handleCreateNotification(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = document.querySelector("#notification-create-message");
  const submit = form.querySelector("button[type='submit']");
  const values = Object.fromEntries(new FormData(form).entries());
  try {
    submit.disabled = true;
    message.dataset.state = "";
    message.textContent = "Creating notification...";
    await createNotification({
      branchId: values.branch_id,
      recipientId: currentSession.user.id,
      eventType: values.event_type,
      severity: values.severity,
      title: values.title,
      body: values.body,
      actionUrl: values.action_url,
      metadata: { source: "manual_ui" },
    });
    notificationNotice = "Notification created.";
    document.querySelector("#notification-create-dialog")?.close();
    renderNotifications();
  } catch (error) {
    message.dataset.state = "error";
    message.textContent = error.message;
  } finally {
    submit.disabled = false;
  }
}

async function handleSaveNotificationPreferences(event, existingPreferences) {
  event.preventDefault();
  const message = document.querySelector("#notification-message");
  const rows = Array.from(event.currentTarget.querySelectorAll(".notification-preference-row"));
  try {
    message.dataset.state = "";
    message.textContent = "Saving preferences...";
    await Promise.all(rows.map((row) => {
      const existing = existingPreferences.find((preference) => preference.id === row.dataset.preferenceId);
      return saveNotificationPreference({
        id: existing?.id,
        user_id: currentSession.user.id,
        branch_id: null,
        event_type: row.dataset.eventType,
        in_app_enabled: row.querySelector("input").checked,
        email_enabled: existing?.email_enabled ?? false,
        sms_enabled: existing?.sms_enabled ?? false,
        push_enabled: existing?.push_enabled ?? false,
      });
    }));
    notificationNotice = "Notification preferences saved.";
    renderNotifications();
  } catch (error) {
    message.dataset.state = "error";
    message.textContent = error.message;
  }
}

function renderSettings() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Admin", title: "Settings / Users / Roles" })}
    <section id="settings-root" class="settings-root">
      <div class="empty-state">Loading settings...</div>
    </section>
  `;

  loadSettings();
}

async function loadSettings() {
  const root = document.querySelector("#settings-root");

  try {
    const [directory, company, tenants, integrations] = await Promise.all([
      manageTmsUsers("list"),
      getCompanyBillingSettings(),
      listCompanyTenants(),
      loadIntegrationFoundation(),
    ]);
    if (!(directory.drivers || []).length) {
      try {
        directory.drivers = await listDrivers();
      } catch {
        directory.drivers = [];
      }
    }
    const internalUsers = (directory.users || []).filter(isInternalStaffUser);

    root.innerHTML = `
      <section class="settings-grid">
        <article class="panel settings-user-card">
          <div class="panel-header">
            <div>
              <h2>Current User</h2>
              <p>Supabase Auth session currently controlling this browser.</p>
            </div>
          </div>
          <div class="settings-user">
            <article><span>Email</span><strong>${escapeHtml(currentSession?.user?.email || "-")}</strong></article>
            <article><span>User ID</span><strong>${escapeHtml(currentSession?.user?.id || "-")}</strong></article>
            <article><span>Last sign in</span><strong>${formatDateTime(currentSession?.user?.last_sign_in_at)}</strong></article>
          </div>
        </article>
        <article class="panel">
          <div class="panel-header">
            <div>
              <h2>Add User</h2>
              <p>Create a confirmed account or email an invitation.</p>
            </div>
          </div>
          ${renderUserCreateForm(directory.branches)}
        </article>
      </section>
      <section id="user-directory-panel" class="panel">
        <div class="panel-header">
          <div>
            <h2>Internal Users</h2>
            <p>${internalUsers.length} staff user${internalUsers.length === 1 ? "" : "s"}; driver and portal accounts live in their own tabs.</p>
          </div>
          <button id="reload-settings" type="button">Refresh</button>
        </div>
        <div id="user-management-message" class="form-message">${escapeHtml(settingsNotice)}</div>
        <div class="user-directory">
          ${internalUsers.length ? internalUsers.map((user) => renderManagedUser(user, directory.branches, directory.callerId)).join("") : '<div class="empty-state compact-empty">No internal staff users yet.</div>'}
        </div>
      </section>
      <section id="customer-portal-access-panel" class="panel">
        <div class="panel-header">
          <div>
            <h2>Customer Portal Access</h2>
            <p>Create external customer accounts and scope them to one customer record.</p>
          </div>
          <div class="panel-actions">
            <button type="button" data-open-customer-portal-create>Add Portal User</button>
            <button type="button" data-open-customer-portal-grant>Grant Existing</button>
          </div>
        </div>
        <div class="form-message portal-management-message"></div>
        ${renderSettingsFormDialog("customer-portal-create-dialog", "Add Customer Portal User", "Create login credentials and customer scope.", renderCustomerPortalCreateForm(directory.customers || []))}
        ${renderSettingsFormDialog("customer-portal-grant-dialog", "Grant Customer Portal Access", "Attach an existing Auth user to a customer account.", renderExistingPortalGrantForm(directory.users || [], directory.customers || []))}
        <section class="portal-access-directory settings-access-directory">
          ${(directory.users || []).some((user) => user.portal_access?.length)
            ? directory.users.flatMap((user) => (user.portal_access || []).map((access) => renderPortalManagedAccess(user, access, directory.customers || []))).join("")
            : '<div class="empty-state compact-empty">No customer portal access has been granted yet.</div>'}
        </section>
      </section>
      <section id="carrier-portal-access-panel" class="panel">
        <div class="panel-header"><div><h2>Carrier Portal Access</h2><p>Create external carrier accounts and scope them to one carrier record.</p></div>
          <div class="panel-actions">
            <button type="button" data-open-carrier-portal-create>Add Carrier User</button>
            <button type="button" data-open-carrier-portal-grant>Grant Existing</button>
          </div>
        </div>
        <div class="form-message portal-management-message"></div>
        ${renderSettingsFormDialog("carrier-portal-create-dialog", "Add Carrier Portal User", "Create login credentials and carrier scope.", renderCarrierPortalCreateForm(directory.carriers || []))}
        ${renderSettingsFormDialog("carrier-portal-grant-dialog", "Grant Carrier Portal Access", "Attach an existing Auth user to a carrier account.", renderExistingCarrierPortalGrantForm(directory.users || [], directory.carriers || []))}
        <section class="portal-access-directory settings-access-directory">
          ${(directory.users || []).some((user) => user.carrier_portal_access?.length)
            ? directory.users.flatMap((user) => (user.carrier_portal_access || []).map((access) => renderCarrierPortalManagedAccess(user, access, directory.carriers || []))).join("")
            : '<div class="empty-state compact-empty">No carrier portal access has been granted yet.</div>'}
        </section>
      </section>
      <section id="driver-app-access-panel" class="panel">
        <div class="panel-header"><div><h2>Driver App Access</h2><p>Create or maintain Driver App logins from the driver profile list.</p></div></div>
        <div id="driver-app-management-message" class="form-message"></div>
        <section class="driver-app-directory">
          ${(directory.drivers || []).length
            ? directory.drivers.map((driver) => renderDriverAppAccessRow(driver, directory.users || [])).join("")
            : '<div class="empty-state compact-empty">No driver profiles yet.</div>'}
        </section>
      </section>
      ${renderCompanyTenantPanel(tenants)}
      <section id="company-billing-panel" class="panel">
        <div class="panel-header"><div><h2>Company Billing Details</h2><p>Issuer identity shown on invoice preview and printed PDF.</p></div></div>
        <form id="company-billing-form" class="record-form company-billing-form">
          <label><span>Legal name</span><input name="legal_name" value="${escapeAttribute(company.legal_name)}" required></label>
          <label><span>Address line 1</span><input name="address_line_1" value="${escapeAttribute(company.address_line_1 || "")}"></label>
          <label><span>Address line 2</span><input name="address_line_2" value="${escapeAttribute(company.address_line_2 || "")}"></label>
          <label><span>City</span><input name="city" value="${escapeAttribute(company.city || "")}"></label>
          <label><span>State</span><input name="state" value="${escapeAttribute(company.state || "")}"></label>
          <label><span>Postal code</span><input name="postal_code" value="${escapeAttribute(company.postal_code || "")}"></label>
          <label><span>Country</span><input name="country" value="${escapeAttribute(company.country || "USA")}"></label>
          <label><span>Phone</span><input name="phone" value="${escapeAttribute(company.phone || "")}"></label>
          <label><span>Billing email</span><input name="email" type="email" value="${escapeAttribute(company.email || "")}"></label>
          <label><span>Tax ID</span><input name="tax_id" value="${escapeAttribute(company.tax_id || "")}"></label>
          <label><span>Invoice prefix</span><input name="invoice_number_prefix" value="${escapeAttribute(company.invoice_number_prefix || "INV-")}" required></label>
          <label><span>Next invoice #</span><input name="invoice_number_next" type="number" min="1" step="1" value="${company.invoice_number_next ?? 1001}" required></label>
          <label><span>Number padding</span><input name="invoice_number_padding" type="number" min="1" max="12" step="1" value="${company.invoice_number_padding ?? 4}" required></label>
          <label class="company-payment-field"><span>Payment instructions</span><textarea name="payment_instructions" rows="4">${escapeHtml(company.payment_instructions || "")}</textarea></label>
          <div class="form-actions"><button type="submit">Save Billing Details</button></div>
          <p id="company-billing-message" class="form-message"></p>
        </form>
      </section>
      <section id="user-activity-panel" class="panel">
        <div class="panel-header">
          <div>
            <h2>Recent User Management Activity</h2>
            <p>Latest audited Owner actions from the protected server workflow.</p>
          </div>
        </div>
        <div class="user-audit-list">
          ${directory.activity?.length
            ? directory.activity.map(renderUserAuditEvent).join("")
            : '<div class="empty-state">No user management audit events yet.</div>'}
        </div>
      </section>
      ${renderIntegrationFoundation(integrations, directory.branches)}
    `;

    document.querySelector("#reload-settings").addEventListener("click", loadSettings);
    document.querySelector("#company-billing-form").addEventListener("submit", handleCompanyBillingSave);
    document.querySelectorAll("[data-company-tenant-form]").forEach((form) => form.addEventListener("submit", handleCompanyTenantSave));
    bindIntegrationFoundation();
    bindUserManagement(directory.branches, directory.drivers || []);
    bindDriverAppManagement(directory.drivers || []);
    bindPortalManagement();
    bindCarrierPortalManagement();
    configureCrudFormAccess("#create-user-form", "manage_users", "Add User", resetCreateUserForm, "#user-directory-panel");
    setupPageTabs(
      [
        document.querySelector("#user-directory-panel"),
        document.querySelector("#customer-portal-access-panel"),
        document.querySelector("#carrier-portal-access-panel"),
        document.querySelector("#driver-app-access-panel"),
        document.querySelector("#company-tenant-panel"),
        document.querySelector("#company-billing-panel"),
        document.querySelector("#integration-foundation-panel"),
        document.querySelector("#user-activity-panel"),
      ],
      ["Users", "Customer Portal", "Carrier Portal", "Driver App", "Company", "Billing Details", "Integrations", "Activity"],
      settingsActiveTab,
      (label) => { settingsActiveTab = label; }
    );
  } catch (error) {
    root.innerHTML = `<div class="empty-state">Owner User Management is unavailable: ${escapeHtml(error.message)}</div>`;
  }
}

async function loadIntegrationFoundation() {
  let connections = [];
  let events = [];
  let externalIds = [];
  let rateLimits = [];
  let unavailableMessage = "";

  try {
    [connections, events, externalIds, rateLimits] = await Promise.all([
      listIntegrationConnections(),
      listIntegrationWebhookEvents(),
      listIntegrationExternalIds(),
      listIntegrationRateLimits(),
    ]);
  } catch (error) {
    unavailableMessage = error.message;
  }

  return {
    connections,
    events,
    externalIds,
    rateLimits,
    unavailableMessage,
    summary: integrationHealthSummary(connections, events, rateLimits),
  };
}

function renderCompanyTenantPanel(tenants = []) {
  return `
    <section id="company-tenant-panel" class="panel">
      <div class="panel-header">
        <div>
          <h2>Company / Tenant</h2>
          <p>Tenant identity, locale and branch/terminal structure.</p>
        </div>
      </div>
      <div class="company-tenant-list">
        ${tenants.length ? tenants.map(renderCompanyTenantCard).join("") : '<div class="empty-state">No company tenant records available.</div>'}
      </div>
    </section>
  `;
}

function renderCompanyTenantCard(company) {
  const branches = company.branches || [];
  return `
    <article class="company-tenant-card">
      <form class="record-form company-tenant-form" data-company-tenant-form data-company-id="${company.id}">
        <div class="company-tenant-summary">
          <div>
            <span class="eyebrow-label">Tenant</span>
            <h3>${escapeHtml(company.legal_name)}</h3>
            <small>${escapeHtml(company.dba_name || "No DBA name")} · ${formatStatus(company.status)}</small>
          </div>
          <span class="status-pill" data-state="${company.status === "active" ? "ok" : "warning"}">${formatStatus(company.status)}</span>
        </div>
        <label><span>Legal name</span><input name="legal_name" value="${escapeAttribute(company.legal_name || "")}" required></label>
        <label><span>DBA name</span><input name="dba_name" value="${escapeAttribute(company.dba_name || "")}"></label>
        <label><span>Timezone</span><input name="timezone" value="${escapeAttribute(company.timezone || "America/Chicago")}" required></label>
        <label><span>Locale</span><input name="locale" value="${escapeAttribute(company.locale || "en-US")}" required></label>
        <label><span>Country</span><input name="country" value="${escapeAttribute(company.country || "USA")}" required></label>
        <label><span>Status</span><select name="status"><option value="active" ${company.status === "active" ? "selected" : ""}>Active</option><option value="inactive" ${company.status === "inactive" ? "selected" : ""}>Inactive</option></select></label>
        <div class="form-actions"><button type="submit">Save Company</button></div>
        <p class="form-message" data-company-tenant-message></p>
      </form>
      <section class="tenant-branches">
        <div class="inline-section-header"><div><span>Branches / Terminals</span><small>${branches.length} linked branch${branches.length === 1 ? "" : "es"}</small></div></div>
        <div class="tenant-branch-list">
          ${branches.length ? branches.map(renderTenantBranch).join("") : '<div class="empty-state compact-empty">No branches linked to this company yet.</div>'}
        </div>
      </section>
    </article>
  `;
}

function renderTenantBranch(branch) {
  return `
    <article class="tenant-branch-card">
      <div><strong>${escapeHtml(branch.name)}</strong><small>${escapeHtml(branch.branch_code || formatStatus(branch.terminal_type || "terminal"))}</small></div>
      <span>${escapeHtml(branch.address || "-")}</span>
      <small>${escapeHtml(branch.timezone || "America/Chicago")} · ${formatStatus(branch.status)}</small>
    </article>
  `;
}

function renderIntegrationFoundation(integrations, branches = []) {
  const { connections, events, externalIds, rateLimits, summary, unavailableMessage } = integrations;
  const mockControls = appConfig.name === "staging"
    ? `<div class="integration-mock-controls">
        <label><span>Mock branch</span><select id="mock-telematics-branch">${branchOptions(branches)}</select></label>
        <button id="generate-mock-telematics" type="button" ${branches.length ? "" : "disabled"}>Generate Mock Snapshot</button>
      </div>`
    : "";
  return `
    <section id="integration-foundation-panel" class="panel">
      <div class="panel-header">
        <div>
          <h2>Integration Foundation</h2>
          <p>Provider-neutral connections, webhook inbox, external IDs, retry state and rate limits.</p>
        </div>
        <div class="integration-header-actions">
          ${mockControls}
          <span class="status-pill" data-state="${summary.state}">${formatStatus(summary.state)}</span>
        </div>
      </div>
      ${mockControls ? `<p id="mock-telematics-message" class="form-message integration-mock-message">${escapeHtml(integrationNotice)}</p>` : ""}
      <div class="integration-summary">
        <article><span>Connections</span><strong>${summary.activeConnections}/${summary.totalConnections}</strong><small>Active / total</small></article>
        <article><span>Inbox</span><strong>${summary.retry}</strong><small>Retry queued</small></article>
        <article><span>Dead letter</span><strong>${summary.deadLetter}</strong><small>Needs review</small></article>
        <article><span>Rate buckets</span><strong>${summary.exhaustedBuckets}</strong><small>Exhausted windows</small></article>
      </div>
      ${unavailableMessage ? `<div class="integration-unavailable">Integration foundation tables are not available on this environment yet: ${escapeHtml(unavailableMessage)}</div>` : ""}
      <div class="integration-grid">
        <article class="integration-section">
          <header><h3>Connections</h3><small>Credential values stay server-side.</small></header>
          <div class="integration-list">
            ${connections.length ? connections.map(renderIntegrationConnection).join("") : '<div class="empty-state compact-empty">No integration connections configured.</div>'}
          </div>
        </article>
        <article class="integration-section">
          <header><h3>Webhook Inbox</h3><small>Idempotent event intake and retry queue.</small></header>
          <div class="integration-list">
            ${events.length ? events.map(renderIntegrationEvent).join("") : '<div class="empty-state compact-empty">No webhook events received.</div>'}
          </div>
        </article>
        <article class="integration-section">
          <header><h3>External ID Map</h3><small>Provider IDs linked to local TMS records.</small></header>
          <div class="integration-list">
            ${externalIds.length ? externalIds.slice(0, 8).map(renderExternalIdMap).join("") : '<div class="empty-state compact-empty">No external IDs mapped.</div>'}
          </div>
        </article>
        <article class="integration-section">
          <header><h3>Rate Limits</h3><small>Server-side adapter buckets.</small></header>
          <div class="integration-list">
            ${rateLimits.length ? rateLimits.slice(0, 8).map(renderRateLimitBucket).join("") : '<div class="empty-state compact-empty">No rate limit buckets used.</div>'}
          </div>
        </article>
      </div>
    </section>
  `;
}

function renderIntegrationConnection(connection) {
  const state = ["error", "expired", "revoked"].includes(connection.status) || ["error", "expired", "revoked"].includes(connection.credential_status)
    ? "error"
    : connection.status === "active" && connection.credential_status === "configured"
      ? "ok"
      : "warning";
  return `
    <div class="integration-row">
      <span><strong>${escapeHtml(connection.display_name)}</strong><small>${escapeHtml(connection.provider)} · ${formatStatus(connection.integration_type)} · ${escapeHtml(connection.branches?.name || "Branch")}</small></span>
      <span class="status-pill" data-state="${state}">${formatStatus(connection.status)}</span>
      <small>${formatStatus(connection.credential_status)}${connection.credential_hint ? ` · ${escapeHtml(connection.credential_hint)}` : ""}</small>
      <small>Last success ${formatDateTime(connection.last_success_at)}</small>
    </div>
  `;
}

function renderIntegrationEvent(event) {
  const state = event.status === "dead_letter" ? "error" : event.status === "retry" ? "warning" : event.status === "processed" ? "ok" : "";
  const canRetry = ["dead_letter", "retry"].includes(event.status);
  return `
    <div class="integration-row" data-event-id="${event.id}">
      <span><strong>${escapeHtml(event.event_type)}</strong><small>${escapeHtml(event.provider)} · ${escapeHtml(event.external_event_id || event.idempotency_key)}</small></span>
      <span class="status-pill" data-state="${state}">${formatStatus(event.status)}</span>
      <small>${event.retry_count}/${event.max_retries} retries · Next ${formatDateTime(event.next_retry_at)}</small>
      <div class="integration-actions">
        ${canRetry ? `<button type="button" data-integration-event-status="retry">Retry</button><button type="button" data-integration-event-status="ignored">Ignore</button>` : ""}
      </div>
    </div>
  `;
}

function renderExternalIdMap(item) {
  return `
    <div class="integration-row">
      <span><strong>${formatStatus(item.entity_type)}</strong><small>${escapeHtml(item.provider)} · ${escapeHtml(item.external_id)}</small></span>
      <small>${escapeHtml(item.entity_id)}</small>
      <small>Seen ${formatDateTime(item.last_seen_at || item.updated_at)}</small>
    </div>
  `;
}

function renderRateLimitBucket(bucket) {
  const exhausted = Number(bucket.used_count) >= Number(bucket.limit_count);
  return `
    <div class="integration-row">
      <span><strong>${escapeHtml(bucket.bucket_key)}</strong><small>${escapeHtml(bucket.provider)} · ${bucket.window_seconds}s window</small></span>
      <span class="status-pill" data-state="${exhausted ? "warning" : "ok"}">${bucket.used_count}/${bucket.limit_count}</span>
      <small>Reset ${formatDateTime(bucket.reset_at)}</small>
    </div>
  `;
}

function bindIntegrationFoundation() {
  const mockButton = document.querySelector("#generate-mock-telematics");
  mockButton?.addEventListener("click", async () => {
    const message = document.querySelector("#mock-telematics-message");
    const branchId = document.querySelector("#mock-telematics-branch")?.value || "";
    const originalLabel = mockButton.textContent;
    mockButton.disabled = true;
    mockButton.textContent = "Generating...";
    message.dataset.state = "";
    message.textContent = "Generating staging GPS and HOS snapshots...";
    try {
      const result = await generateMockTelematics(branchId);
      integrationNotice = `Mock snapshot generated for ${result.branch.name}: ${result.positions} GPS and ${result.hos} HOS records.`;
      settingsActiveTab = "Integrations";
      await loadSettings();
    } catch (error) {
      message.dataset.state = "error";
      message.textContent = error.message;
      mockButton.disabled = false;
      mockButton.textContent = originalLabel;
    }
  });

  document.querySelectorAll("[data-integration-event-status]").forEach((button) => {
    button.addEventListener("click", async () => {
      const row = button.closest("[data-event-id]");
      const originalLabel = button.textContent;
      button.disabled = true;
      button.textContent = "Saving...";
      try {
        await updateIntegrationWebhookStatus(row.dataset.eventId, button.dataset.integrationEventStatus);
        settingsNotice = "Integration event updated.";
        settingsActiveTab = "Integrations";
        await loadSettings();
      } catch (error) {
        button.textContent = originalLabel;
        button.disabled = false;
        const message = document.querySelector("#user-management-message");
        if (message) {
          message.dataset.state = "error";
          message.textContent = error.message;
        }
      }
    });
  });
}

async function handleCompanyBillingSave(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = document.querySelector("#company-billing-message");
  const button = form.querySelector("button[type='submit']");
  button.disabled = true;
  message.textContent = "Saving billing details...";
  try {
    const formData = new FormData(form);
    const numbering = validateInvoiceNumberingSettings({
      prefix: formData.get("invoice_number_prefix"),
      nextNumber: formData.get("invoice_number_next"),
      padding: formData.get("invoice_number_padding"),
    });
    await updateCompanyBillingSettings({
      ...Object.fromEntries(formData),
      ...numbering,
    });
    message.textContent = "Billing details saved.";
  } catch (error) {
    message.textContent = error.message;
  } finally {
    button.disabled = false;
  }
}

async function handleCompanyTenantSave(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = form.querySelector("[data-company-tenant-message]");
  const button = form.querySelector("button[type='submit']");
  const formData = new FormData(form);
  button.disabled = true;
  message.textContent = "Saving company profile...";
  try {
    const profile = validateCompanyTenantProfile({
      legalName: formData.get("legal_name"),
      timezone: formData.get("timezone"),
      locale: formData.get("locale"),
      country: formData.get("country"),
      status: formData.get("status"),
    });
    await updateCompanyTenantProfile(form.dataset.companyId, {
      ...profile,
      dba_name: String(formData.get("dba_name") || "").trim(),
    });
    settingsNotice = "Company profile saved.";
    settingsActiveTab = "Company";
    await loadSettings();
  } catch (error) {
    message.textContent = error.message;
  } finally {
    button.disabled = false;
  }
}

function resetCreateUserForm() {
  const form = document.querySelector("#create-user-form");
  form?.reset();
  if (form) closeFormDialog(form);
}

function branchOptions(branches, selectedId = "") {
  return branches
    .map((branch) => `<option value="${branch.id}" ${branch.id === selectedId ? "selected" : ""}>${escapeHtml(branch.name)}</option>`)
    .join("");
}

function customerOptions(customers, selectedId = "") {
  return customers
    .map((customer) => `<option value="${customer.id}" ${customer.id === selectedId ? "selected" : ""}>${escapeHtml(customer.name)}${customer.status === "inactive" ? " (inactive)" : ""}</option>`)
    .join("");
}

const internalUserRoles = new Set(["owner", "admin", "dispatcher", "accounting", "safety"]);

function isInternalStaffUser(user) {
  return (user.access || []).some((access) => internalUserRoles.has(access.role));
}

function driverOptions(drivers = []) {
  return drivers
    .map((driver) => {
      const linked = Boolean(driver.auth_user_id);
      const label = [
        driver.name || "Unnamed driver",
        driver.email || "No email",
        driver.branches?.name || "Branch",
        linked ? "linked" : "",
      ].filter(Boolean).join(" · ");
      return `<option value="${driver.id}" data-branch-id="${escapeAttribute(driver.branch_id || "")}" ${linked ? "disabled" : ""}>${escapeHtml(label)}</option>`;
    })
    .join("");
}

function renderUserCreateForm(branches) {
  return `
    <form id="create-user-form" class="record-form compact-form">
      <label>First name<input name="firstName" type="text" required autocomplete="off"></label>
      <label>Last name<input name="lastName" type="text" required autocomplete="off"></label>
      <label>Email<input name="email" type="email" required autocomplete="off"></label>
      <label>Delivery
        <select name="mode"><option value="invite">Email invitation</option><option value="create">Create with password</option></select>
      </label>
      <label class="password-field" hidden>Temporary password<input name="password" type="password" minlength="8" autocomplete="new-password"></label>
      <label>Role
        <select name="role">${["admin", "dispatcher", "accounting", "safety", "owner"].map((role) => `<option value="${role}">${formatStatus(role)}</option>`).join("")}</select>
      </label>
      <label class="branch-field">Branch<select name="branchId" required>${branchOptions(branches)}</select></label>
      <label>Status<select name="status"><option value="active">Active</option><option value="inactive">Inactive</option></select></label>
      <div class="form-actions"><button type="submit">Add user</button><span class="form-message"></span></div>
    </form>
  `;
}

function renderSettingsFormDialog(id, title, description, formMarkup) {
  return `<dialog id="${id}" class="crud-dialog settings-form-dialog">
    <article class="panel">
      <div class="panel-header"><div><h2>${escapeHtml(title)}</h2><p>${escapeHtml(description)}</p></div><button class="dialog-close" type="button" data-close-settings-dialog aria-label="Close">×</button></div>
      ${formMarkup}
    </article>
  </dialog>`;
}

function renderCustomerPortalCreateForm(customers = []) {
  return `
    <form id="create-portal-user-form" class="record-form compact-form">
      <label>First name<input name="firstName" type="text" required autocomplete="off"></label>
      <label>Last name<input name="lastName" type="text" required autocomplete="off"></label>
      <label>Email<input name="email" type="email" required autocomplete="off"></label>
      <label>Delivery
        <select name="mode"><option value="invite">Email invitation</option><option value="create">Create with password</option></select>
      </label>
      <label class="password-field" hidden>Temporary password<input name="password" type="password" minlength="8" autocomplete="new-password"></label>
      <label>Customer<select name="customerId" required>${customers.length ? customerOptions(customers) : '<option value="">No customers available</option>'}</select></label>
      <label>Status<select name="status"><option value="active">Active</option><option value="inactive">Inactive</option></select></label>
      <div class="form-actions"><button type="submit" ${customers.length ? "" : "disabled"}>Add portal user</button><span class="form-message"></span></div>
    </form>
  `;
}

function renderExistingPortalGrantForm(users = [], customers = []) {
  const userOptions = users
    .map((user) => `<option value="${user.id}">${escapeHtml(user.email || user.id)}</option>`)
    .join("");
  return `
    <form id="grant-portal-access-form" class="record-form compact-form">
      <label>User<select name="userId" required>${userOptions || '<option value="">No users available</option>'}</select></label>
      <label>Customer<select name="customerId" required>${customers.length ? customerOptions(customers) : '<option value="">No customers available</option>'}</select></label>
      <label>Status<select name="status"><option value="active">Active</option><option value="inactive">Inactive</option></select></label>
      <div class="form-actions"><button type="submit" ${users.length && customers.length ? "" : "disabled"}>Grant portal access</button><span class="form-message"></span></div>
    </form>
  `;
}

function renderPortalManagedAccess(user, access, customers = []) {
  return `
    <article class="portal-access-row" data-user-id="${user.id}" data-access-id="${access.id}" data-email="${escapeAttribute(user.email || "")}">
      <div>
        <strong>${escapeHtml(user.email || "No email")}</strong>
        <span>${escapeHtml([user.first_name, user.last_name].filter(Boolean).join(" ") || "Portal user")}</span>
        <small>${escapeHtml(access.customers?.name || "Customer")} · created ${formatDateTime(access.created_at)}</small>
      </div>
      <form class="portal-access-form">
        <select name="customerId" required>${customerOptions(customers, access.customer_id)}</select>
        <select name="status"><option value="active" ${access.status === "active" ? "selected" : ""}>Active</option><option value="inactive" ${access.status === "inactive" ? "selected" : ""}>Inactive</option></select>
        <button type="submit">Save portal access</button>
      </form>
      <div class="portal-login-actions">
        <button type="button" data-customer-portal-edit-user>Edit login</button>
        <button type="button" data-customer-portal-reset-password>Send reset</button>
      </div>
    </article>
    ${renderPortalLoginDialog(user, "Customer Portal Login", "Identity, email and temporary password for this customer portal user.", "customer-portal-identity-form")}`;
}

function carrierOptions(carriers = [], selectedId = "") {
  return carriers.map((carrier) => `<option value="${carrier.id}" ${carrier.id === selectedId ? "selected" : ""}>${escapeHtml(carrier.name)}${carrier.status === "inactive" ? " (inactive)" : ""}</option>`).join("");
}

function renderCarrierPortalCreateForm(carriers = []) {
  return `<form id="create-carrier-portal-user-form" class="record-form compact-form">
    <label>First name<input name="firstName" required autocomplete="off"></label><label>Last name<input name="lastName" required autocomplete="off"></label>
    <label>Email<input name="email" type="email" required autocomplete="off"></label>
    <label>Delivery<select name="mode"><option value="invite">Email invitation</option><option value="create">Create with password</option></select></label>
    <label class="password-field" hidden>Temporary password<input name="password" type="password" minlength="8" autocomplete="new-password"></label>
    <label>Carrier<select name="carrierId" required>${carriers.length ? carrierOptions(carriers) : '<option value="">No carriers available</option>'}</select></label>
    <label>Status<select name="status"><option value="active">Active</option><option value="inactive">Inactive</option></select></label>
    <div class="form-actions"><button type="submit" ${carriers.length ? "" : "disabled"}>Add carrier portal user</button><span class="form-message"></span></div></form>`;
}

function renderExistingCarrierPortalGrantForm(users = [], carriers = []) {
  const options = users.map((user) => `<option value="${user.id}">${escapeHtml(user.email || user.id)}</option>`).join("");
  return `<form id="grant-carrier-portal-access-form" class="record-form compact-form">
    <label>User<select name="userId" required>${options || '<option value="">No users available</option>'}</select></label>
    <label>Carrier<select name="carrierId" required>${carriers.length ? carrierOptions(carriers) : '<option value="">No carriers available</option>'}</select></label>
    <label>Status<select name="status"><option value="active">Active</option><option value="inactive">Inactive</option></select></label>
    <div class="form-actions"><button type="submit" ${users.length && carriers.length ? "" : "disabled"}>Grant carrier access</button><span class="form-message"></span></div></form>`;
}

function renderCarrierPortalManagedAccess(user, access, carriers = []) {
  return `<article class="portal-access-row carrier-access-row" data-user-id="${user.id}" data-access-id="${access.id}" data-email="${escapeAttribute(user.email || "")}"><div class="portal-access-identity"><span class="portal-user-avatar">${escapeHtml((user.first_name?.[0] || user.email?.[0] || "C").toUpperCase())}</span><div><strong>${escapeHtml(user.email || "No email")}</strong><span>${escapeHtml([user.first_name,user.last_name].filter(Boolean).join(" ") || "Portal user")}</span><small>Created ${formatDateTime(access.created_at)}</small></div><span class="status-pill" data-state="${access.status === "active" ? "ok" : ""}">${formatStatus(access.status)}</span></div>
    <form class="portal-access-form carrier-portal-access-form"><label><span>Carrier account</span><select name="carrierId" required>${carrierOptions(carriers, access.carrier_id)}</select></label><label><span>Access status</span><select name="status"><option value="active" ${access.status === "active" ? "selected" : ""}>Active</option><option value="inactive" ${access.status === "inactive" ? "selected" : ""}>Inactive</option></select></label><button type="submit">Save Access</button></form>
    <div class="portal-login-actions"><button type="button" data-carrier-portal-edit-user>Edit login</button><button type="button" data-carrier-portal-reset-password>Send reset</button></div></article>
    ${renderPortalLoginDialog(user, "Carrier Portal Login", "Identity, email and temporary password for this carrier portal user.", "carrier-portal-identity-form")}`;
}

function renderPortalLoginDialog(user, title, description, formClass) {
  return `<dialog class="crud-dialog identity-dialog">
    <article class="panel">
      <div class="panel-header"><div><h2>${escapeHtml(title)}</h2><p>${escapeHtml(description)}</p></div><button class="dialog-close cancel-user-edit" type="button" aria-label="Close">×</button></div>
      <form class="record-form identity-form ${formClass}" data-user-id="${user.id}">
        <label><span>First name</span><input name="firstName" value="${escapeAttribute(user.first_name || "")}" required></label>
        <label><span>Last name</span><input name="lastName" value="${escapeAttribute(user.last_name || "")}" required></label>
        <label><span>Email</span><input name="email" type="email" value="${escapeAttribute(user.email || "")}" required></label>
        <label><span>New temporary password</span><input name="password" type="password" minlength="8" placeholder="Leave blank to keep current"></label>
        <p class="form-help">Use this when you want to set a new temporary password directly. Passwords are never stored in portal access records.</p>
        <div class="form-actions"><button type="submit">Save login</button><button class="cancel-user-edit" type="button">Cancel</button></div>
        <p class="form-message"></p>
      </form>
    </article>
  </dialog>`;
}

function driverLinkedUser(driver, users = []) {
  if (driver.auth_user_id) return users.find((user) => user.id === driver.auth_user_id) || null;
  const driverEmail = String(driver.email || "").toLowerCase();
  if (!driverEmail) return null;
  return users.find((user) => String(user.email || "").toLowerCase() === driverEmail) || null;
}

function renderDriverAppAccessRow(driver, users = []) {
  const user = driverLinkedUser(driver, users);
  const hardLinked = Boolean(driver.auth_user_id);
  const emailMatched = !hardLinked && Boolean(user);
  const hasLogin = hardLinked || emailMatched;
  const access = user?.access?.find((item) => item.role === "driver") || user?.access?.[0];
  const { firstName, lastName } = driverNameParts(driver.name);
  const stateLabel = hardLinked ? "Linked" : emailMatched ? "Email match" : "No login";
  const stateHelp = hardLinked
    ? escapeHtml(user?.email || "Auth user linked")
    : emailMatched
      ? "Login works by matching this driver's email. Use Edit login to maintain it."
      : "Create credentials for Driver App access.";
  return `
    <article class="driver-app-access-row" data-driver-id="${driver.id}" data-user-id="${escapeAttribute(user?.id || "")}" data-email="${escapeAttribute(user?.email || driver.email || "")}">
      <div class="driver-app-profile">
        <span class="portal-user-avatar">${escapeHtml(userInitials(driver.name || driver.email || "D"))}</span>
        <div>
          <strong>${escapeHtml(driver.name || "Unnamed driver")}</strong>
          <span>${escapeHtml(driver.branches?.name || "Branch not set")}</span>
          <small>${escapeHtml(driver.email || "No driver email on profile")}</small>
        </div>
      </div>
      <div class="driver-app-login-state">
        <span class="status-pill" data-state="${hardLinked ? "ok" : emailMatched ? "warning" : ""}">${stateLabel}</span>
        <small>${stateHelp}</small>
        ${access ? `<small>${escapeHtml(formatStatus(access.role))} · ${escapeHtml(formatStatus(access.status))}</small>` : ""}
      </div>
      ${hasLogin ? `
        <div class="driver-app-actions">
          <button type="button" data-driver-edit-user ${user ? "" : "disabled"}>Edit login</button>
          <button type="button" data-driver-reset-password ${user ? "" : "disabled"}>Send reset</button>
        </div>
        ${user ? `<dialog class="crud-dialog identity-dialog">
          <article class="panel">
            <div class="panel-header"><div><h2>Driver App Login</h2><p>Identity, email and temporary password for ${escapeHtml(driver.name || "driver")}.</p></div><button class="dialog-close cancel-user-edit" type="button" aria-label="Close">×</button></div>
            <form class="record-form identity-form driver-identity-form" data-user-id="${user.id}">
              <label><span>First name</span><input name="firstName" value="${escapeAttribute(user.first_name || firstName)}" required></label>
              <label><span>Last name</span><input name="lastName" value="${escapeAttribute(user.last_name || lastName)}" required></label>
              <label><span>Email</span><input name="email" type="email" value="${escapeAttribute(user.email || driver.email || "")}" required></label>
              <label><span>New temporary password</span><input name="password" type="password" minlength="8" placeholder="Leave blank to keep current"></label>
              <p class="form-help">Use this when you want to set a new temporary password directly. The password is not stored in the driver profile.</p>
              <div class="form-actions"><button type="submit">Save login</button><button class="cancel-user-edit" type="button">Cancel</button></div>
              <p class="form-message"></p>
            </form>
          </article>
        </dialog>` : ""}
      ` : `
        <form class="driver-app-create-form">
          <input name="driverId" type="hidden" value="${escapeAttribute(driver.id)}">
          <input name="branchId" type="hidden" value="${escapeAttribute(driver.branch_id || "")}">
          <input name="firstName" type="hidden" value="${escapeAttribute(firstName)}">
          <input name="lastName" type="hidden" value="${escapeAttribute(lastName)}">
          <label><span>Login email</span><input name="email" type="email" value="${escapeAttribute(driver.email || "")}" required autocomplete="off"></label>
          <label><span>Temporary password</span><input name="password" type="password" minlength="8" required autocomplete="new-password"></label>
          <button type="submit" ${driver.branch_id ? "" : "disabled"}>Create login</button>
        </form>
      `}
    </article>`;
}

function renderManagedUser(user, branches, callerId) {
  const access = user.access[0];
  const currentRole = access?.role || "admin";
  const currentStatus = access?.status || "active";
  return `
    <article class="managed-user" data-user-id="${user.id}" data-email="${escapeHtml(user.email || "")}">
      <div class="managed-user-identity">
        <strong>${escapeHtml([user.first_name, user.last_name].filter(Boolean).join(" ") || "Name not set")}${user.id === callerId ? " (you)" : ""}</strong>
        <span>${escapeHtml(user.email || "No email")}</span>
        <small>Created ${formatDateTime(user.created_at)} · Last sign-in ${formatDateTime(user.last_sign_in_at)}</small>
      </div>
      <form class="access-form" ${access ? `data-access-id="${access.id}"` : `data-new-access="true"`}>
          <label><span>Role</span><select name="role">${["owner", "admin", "dispatcher", "accounting", "safety", "driver"].map((role) => `<option value="${role}" ${role === currentRole ? "selected" : ""}>${formatStatus(role)}</option>`).join("")}</select></label>
          <label><span>Branch</span><select name="branchId" ${currentRole === "owner" ? "disabled" : ""}>${branchOptions(branches, access?.branch_id)}</select></label>
          <label><span>Status</span><select name="status"><option value="active" ${currentStatus === "active" ? "selected" : ""}>Active</option><option value="inactive" ${currentStatus === "inactive" ? "selected" : ""}>Inactive</option></select></label>
          <button type="submit">${access ? "Save access" : "Grant access"}</button>
        </form>
      <div class="managed-user-actions">
        <button class="edit-user" type="button">Edit user</button>
        <button class="reset-user-password" type="button">Send reset</button>
        <button class="delete-user danger-button" type="button" ${user.id === callerId ? "disabled" : ""}>Delete</button>
      </div>
    </article>
    <dialog class="crud-dialog identity-dialog">
      <article class="panel">
        <div class="panel-header"><div><h2>Edit User</h2><p>Identity, email and temporary password.</p></div><button class="dialog-close cancel-user-edit" type="button" aria-label="Close">×</button></div>
        <form class="record-form identity-form" data-user-id="${user.id}">
          <label><span>First name</span><input name="firstName" value="${escapeAttribute(user.first_name)}" required></label>
          <label><span>Last name</span><input name="lastName" value="${escapeAttribute(user.last_name)}" required></label>
          <label><span>Email</span><input name="email" type="email" value="${escapeAttribute(user.email || "")}" required></label>
          <label><span>New temporary password</span><input name="password" type="password" minlength="8" placeholder="Leave blank to keep current"></label>
          <p class="form-help">Leave the password field blank unless you want to replace the user's current password.</p>
          <div class="form-actions">
            <button type="submit">Save user</button>
            <button class="cancel-user-edit" type="button">Cancel</button>
          </div>
          <p class="form-message"></p>
        </form>
      </article>
    </dialog>
  `;
}

function bindUserManagement() {
  const createForm = document.querySelector("#create-user-form");
  const userPanel = document.querySelector("#user-directory-panel");
  const mode = createForm.querySelector('[name="mode"]');
  const role = createForm.querySelector('[name="role"]');
  const passwordField = createForm.querySelector(".password-field");
  const branchField = createForm.querySelector(".branch-field");
  const branchSelect = branchField.querySelector("select");
  const syncCreateFields = () => {
    passwordField.hidden = mode.value !== "create";
    passwordField.querySelector("input").required = mode.value === "create";
    branchField.hidden = role.value === "owner";
    branchSelect.required = role.value !== "owner";
  };
  mode.addEventListener("change", syncCreateFields);
  role.addEventListener("change", syncCreateFields);
  syncCreateFields();

  createForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const message = createForm.querySelector(".form-message");
    const values = Object.fromEntries(new FormData(createForm));
    message.textContent = "Adding user...";
    try {
      await manageTmsUsers(values.mode, values);
      settingsNotice = `User ${values.email} added successfully.`;
      await loadSettings();
    } catch (error) {
      message.textContent = error.message;
    }
  });

  userPanel?.querySelectorAll(".access-form").forEach((form) => {
    const roleSelect = form.querySelector('[name="role"]');
    const branchSelect = form.querySelector('[name="branchId"]');
    const syncBranch = () => { branchSelect.disabled = roleSelect.value === "owner"; };
    roleSelect.addEventListener("change", syncBranch);
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      const values = Object.fromEntries(new FormData(form));
      const button = form.querySelector('button[type="submit"]');
      const originalLabel = button.textContent;
      button.disabled = true;
      button.textContent = "Saving...";
      try {
        await manageTmsUsers(form.dataset.newAccess ? "addAccess" : "updateAccess", {
          accessId: form.dataset.accessId,
          userId: form.closest(".managed-user").dataset.userId,
          ...values,
        });
        settingsNotice = "Access updated successfully.";
        await loadSettings();
      } catch (error) {
        const message = document.querySelector("#user-management-message");
        message.dataset.state = "error";
        message.textContent = error.message;
        button.disabled = false;
        button.textContent = originalLabel;
      }
    });
  });

  userPanel?.querySelectorAll(".reset-user-password").forEach((button) => button.addEventListener("click", async () => {
    const card = button.closest(".managed-user");
    try {
      await manageTmsUsers("resetPassword", { userId: card.dataset.userId, email: card.dataset.email });
      settingsNotice = `Password reset email sent to ${card.dataset.email}.`;
      document.querySelector("#user-management-message").textContent = settingsNotice;
    } catch (error) {
      const message = document.querySelector("#user-management-message");
      message.dataset.state = "error";
      message.textContent = error.message;
    }
  }));

  userPanel?.querySelectorAll(".edit-user").forEach((button) => button.addEventListener("click", () => {
    button.closest(".managed-user").nextElementSibling.showModal();
  }));

  userPanel?.querySelectorAll(".cancel-user-edit").forEach((button) => button.addEventListener("click", () => {
    button.closest("dialog")?.close();
  }));

  userPanel?.querySelectorAll(".identity-form").forEach((form) => form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const button = form.querySelector('button[type="submit"]');
    const message = form.querySelector(".form-message");
    message.dataset.state = "";
    message.textContent = "";
    button.disabled = true;
    button.textContent = "Saving...";
    try {
      await manageTmsUsers("updateUser", { userId: form.dataset.userId, ...Object.fromEntries(new FormData(form)) });
      settingsNotice = "User details updated successfully.";
      await loadSettings();
    } catch (error) {
      message.dataset.state = "error";
      message.textContent = error.message;
      button.disabled = false;
      button.textContent = "Save user";
    }
  }));

  userPanel?.querySelectorAll(".delete-user").forEach((button) => button.addEventListener("click", async () => {
    const card = button.closest(".managed-user");
    if (!window.confirm(`Permanently delete ${card.dataset.email}?`)) return;
    try {
      await manageTmsUsers("delete", { userId: card.dataset.userId });
      settingsNotice = `User ${card.dataset.email} deleted successfully.`;
      await loadSettings();
    } catch (error) {
      const message = document.querySelector("#user-management-message");
      message.dataset.state = "error";
      message.textContent = error.message;
    }
  }));
}

function bindDriverAppManagement() {
  const panel = document.querySelector("#driver-app-access-panel");
  if (!panel) return;
  const message = panel.querySelector("#driver-app-management-message");

  panel.querySelectorAll(".driver-app-create-form").forEach((form) => {
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      const button = form.querySelector('button[type="submit"]');
      const values = Object.fromEntries(new FormData(form));
      button.disabled = true;
      message.dataset.state = "";
      message.textContent = "Creating Driver App login...";
      try {
        await manageTmsUsers("create", { ...values, role: "driver", status: "active" });
        settingsNotice = `Driver App login created for ${values.email}.`;
        settingsActiveTab = "Driver App";
        await loadSettings();
      } catch (error) {
        message.dataset.state = "error";
        message.textContent = error.message;
        button.disabled = false;
      }
    });
  });

  panel.querySelectorAll("[data-driver-edit-user]").forEach((button) => {
    button.addEventListener("click", () => button.closest(".driver-app-access-row").querySelector("dialog")?.showModal());
  });

  panel.querySelectorAll(".driver-identity-form").forEach((form) => {
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      const button = form.querySelector('button[type="submit"]');
      const formMessage = form.querySelector(".form-message");
      button.disabled = true;
      button.textContent = "Saving...";
      formMessage.dataset.state = "";
      formMessage.textContent = "";
      try {
        await manageTmsUsers("updateUser", { userId: form.dataset.userId, ...Object.fromEntries(new FormData(form)) });
        settingsNotice = "Driver App login updated successfully.";
        settingsActiveTab = "Driver App";
        await loadSettings();
      } catch (error) {
        formMessage.dataset.state = "error";
        formMessage.textContent = error.message;
        button.disabled = false;
        button.textContent = "Save login";
      }
    });
  });

  panel.querySelectorAll("[data-driver-reset-password]").forEach((button) => {
    button.addEventListener("click", async () => {
      const row = button.closest(".driver-app-access-row");
      const email = row.dataset.email;
      button.disabled = true;
      message.dataset.state = "";
      message.textContent = "Sending password reset...";
      try {
        await manageTmsUsers("resetPassword", { userId: row.dataset.userId, email });
        message.textContent = `Password reset email sent to ${email}.`;
      } catch (error) {
        message.dataset.state = "error";
        message.textContent = error.message;
        button.disabled = false;
      }
    });
  });

  panel.querySelectorAll(".cancel-user-edit").forEach((button) => button.addEventListener("click", () => {
    button.closest("dialog")?.close();
  }));
}

function bindPortalManagement() {
  const panel = document.querySelector("#customer-portal-access-panel");
  panel?.querySelector("[data-open-customer-portal-create]")?.addEventListener("click", () => document.querySelector("#customer-portal-create-dialog")?.showModal());
  panel?.querySelector("[data-open-customer-portal-grant]")?.addEventListener("click", () => document.querySelector("#customer-portal-grant-dialog")?.showModal());
  panel?.querySelectorAll("[data-close-settings-dialog]").forEach((button) => button.addEventListener("click", () => button.closest("dialog")?.close()));
  const createForm = document.querySelector("#create-portal-user-form");
  if (createForm) {
    const mode = createForm.querySelector('[name="mode"]');
    const passwordField = createForm.querySelector(".password-field");
    const syncPassword = () => {
      passwordField.hidden = mode.value !== "create";
      passwordField.querySelector("input").required = mode.value === "create";
    };
    mode.addEventListener("change", syncPassword);
    syncPassword();
    createForm.addEventListener("submit", async (event) => {
      event.preventDefault();
      const message = createForm.querySelector(".form-message");
      const values = Object.fromEntries(new FormData(createForm));
      message.dataset.state = "";
      message.textContent = "Adding portal user...";
      try {
        await manageTmsUsers("createCustomerPortal", values);
        settingsNotice = `Customer portal user ${values.email} added successfully.`;
        settingsActiveTab = "Customer Portal";
        await loadSettings();
      } catch (error) {
        message.dataset.state = "error";
        message.textContent = error.message;
      }
    });
  }

  const grantForm = document.querySelector("#grant-portal-access-form");
  grantForm?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const message = grantForm.querySelector(".form-message");
    const values = Object.fromEntries(new FormData(grantForm));
    message.dataset.state = "";
    message.textContent = "Granting portal access...";
    try {
      await manageTmsUsers("addCustomerPortalAccess", values);
      settingsNotice = "Customer portal access granted successfully.";
      settingsActiveTab = "Customer Portal";
      await loadSettings();
    } catch (error) {
      message.dataset.state = "error";
      message.textContent = error.message;
    }
  });

  panel?.querySelectorAll(".portal-access-form").forEach((form) => {
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      const row = form.closest(".portal-access-row");
      const values = Object.fromEntries(new FormData(form));
      const button = form.querySelector('button[type="submit"]');
      const originalLabel = button.textContent;
      button.disabled = true;
      button.textContent = "Saving...";
      try {
        await manageTmsUsers("updateCustomerPortalAccess", {
          accessId: row.dataset.accessId,
          userId: row.dataset.userId,
          ...values,
        });
        settingsNotice = "Customer portal access updated successfully.";
        settingsActiveTab = "Customer Portal";
        await loadSettings();
      } catch (error) {
        const message = document.querySelector("#user-management-message");
        message.dataset.state = "error";
        message.textContent = error.message;
        button.disabled = false;
        button.textContent = originalLabel;
      }
    });
  });

  bindPortalLoginControls({
    panelSelector: "#customer-portal-access-panel",
    editSelector: "[data-customer-portal-edit-user]",
    resetSelector: "[data-customer-portal-reset-password]",
    formSelector: ".customer-portal-identity-form",
    activeTab: "Customer Portal",
    label: "Customer portal",
  });
}

function bindCarrierPortalManagement() {
  const panel = document.querySelector("#carrier-portal-access-panel");
  panel?.querySelector("[data-open-carrier-portal-create]")?.addEventListener("click", () => document.querySelector("#carrier-portal-create-dialog")?.showModal());
  panel?.querySelector("[data-open-carrier-portal-grant]")?.addEventListener("click", () => document.querySelector("#carrier-portal-grant-dialog")?.showModal());
  panel?.querySelectorAll("[data-close-settings-dialog]").forEach((button) => button.addEventListener("click", () => button.closest("dialog")?.close()));
  const createForm = document.querySelector("#create-carrier-portal-user-form");
  if (createForm) {
    const mode = createForm.querySelector('[name="mode"]');
    const passwordField = createForm.querySelector(".password-field");
    const syncPassword = () => { passwordField.hidden = mode.value !== "create"; passwordField.querySelector("input").required = mode.value === "create"; };
    mode.addEventListener("change", syncPassword); syncPassword();
    createForm.addEventListener("submit", async (event) => {
      event.preventDefault(); const values = Object.fromEntries(new FormData(createForm)); const message = createForm.querySelector(".form-message");
      message.textContent = "Adding carrier portal user...";
      try { await manageTmsUsers("createCarrierPortal", values); settingsNotice = `Carrier portal user ${values.email} added successfully.`; settingsActiveTab = "Carrier Portal"; await loadSettings(); }
      catch (error) { message.dataset.state = "error"; message.textContent = error.message; }
    });
  }
  document.querySelector("#grant-carrier-portal-access-form")?.addEventListener("submit", async (event) => {
    event.preventDefault(); const form = event.currentTarget; const values = Object.fromEntries(new FormData(form)); const message = form.querySelector(".form-message");
    message.textContent = "Granting carrier portal access...";
    try { await manageTmsUsers("addCarrierPortalAccess", values); settingsNotice = "Carrier portal access granted successfully."; settingsActiveTab = "Carrier Portal"; await loadSettings(); }
    catch (error) { message.dataset.state = "error"; message.textContent = error.message; }
  });
  document.querySelectorAll("#carrier-portal-access-panel .carrier-portal-access-form").forEach((form) => form.addEventListener("submit", async (event) => {
    event.preventDefault(); const row = form.closest(".portal-access-row"); const button = form.querySelector('button[type="submit"]'); button.disabled = true;
    try { await manageTmsUsers("updateCarrierPortalAccess", { accessId: row.dataset.accessId, userId: row.dataset.userId, ...Object.fromEntries(new FormData(form)) }); settingsNotice = "Carrier portal access updated successfully."; settingsActiveTab = "Carrier Portal"; await loadSettings(); }
    catch (error) { document.querySelector("#user-management-message").textContent = error.message; button.disabled = false; }
  }));

  bindPortalLoginControls({
    panelSelector: "#carrier-portal-access-panel",
    editSelector: "[data-carrier-portal-edit-user]",
    resetSelector: "[data-carrier-portal-reset-password]",
    formSelector: ".carrier-portal-identity-form",
    activeTab: "Carrier Portal",
    label: "Carrier portal",
  });
}

function bindPortalLoginControls({ panelSelector, editSelector, resetSelector, formSelector, activeTab, label }) {
  const panel = document.querySelector(panelSelector);
  if (!panel) return;
  const message = panel.querySelector(".portal-management-message");
  panel.querySelectorAll(editSelector).forEach((button) => {
    button.addEventListener("click", () => button.closest(".portal-access-row")?.nextElementSibling?.showModal());
  });
  panel.querySelectorAll(resetSelector).forEach((button) => {
    button.addEventListener("click", async () => {
      const row = button.closest(".portal-access-row");
      button.disabled = true;
      message.dataset.state = "";
      message.textContent = "Sending password reset...";
      try {
        await manageTmsUsers("resetPassword", { userId: row.dataset.userId, email: row.dataset.email });
        message.textContent = `Password reset email sent to ${row.dataset.email}.`;
      } catch (error) {
        message.dataset.state = "error";
        message.textContent = error.message;
        button.disabled = false;
      }
    });
  });
  panel.querySelectorAll(formSelector).forEach((form) => {
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      const button = form.querySelector('button[type="submit"]');
      const formMessage = form.querySelector(".form-message");
      button.disabled = true;
      button.textContent = "Saving...";
      formMessage.dataset.state = "";
      formMessage.textContent = "";
      try {
        await manageTmsUsers("updateUser", { userId: form.dataset.userId, ...Object.fromEntries(new FormData(form)) });
        settingsNotice = `${label} login updated successfully.`;
        settingsActiveTab = activeTab;
        await loadSettings();
      } catch (error) {
        formMessage.dataset.state = "error";
        formMessage.textContent = error.message;
        button.disabled = false;
        button.textContent = "Save login";
      }
    });
  });
  panel.querySelectorAll(".cancel-user-edit").forEach((button) => button.addEventListener("click", () => {
    button.closest("dialog")?.close();
  }));
}

async function renderQuotes() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Sales", title: "Rate Engine / Quotes" })}
    <dialog id="quote-dialog" class="crud-dialog quote-dialog">
      <article class="panel">
        <div class="panel-header"><div><h2>Create Quote</h2><p>Totals and estimated margin are recalculated by the database.</p></div><button id="close-quote-dialog" class="dialog-close" type="button" aria-label="Close">×</button></div>
        <form id="quote-form" class="record-form">
          <label><span>Quote number</span><input name="quote_no" placeholder="Q-2026-1001" required></label>
          <label><span>Customer</span><select name="customer_id" required></select></label>
          <label><span>Rate agreement</span><select name="rate_agreement_id"><option value="">Manual pricing</option></select></label>
          <label><span>Origin</span><input name="origin" placeholder="Chicago, IL" required></label>
          <label><span>Destination</span><input name="destination" placeholder="Dallas, TX" required></label>
          <label><span>Pickup</span><input name="pickup_date" type="date"></label>
          <label><span>Delivery</span><input name="delivery_date" type="date"></label>
          <label><span>Equipment</span><input name="equipment_type" placeholder="Dry Van"></label>
          <label><span>Loaded miles</span><input name="loaded_miles" type="number" min="0" step="0.1" value="0"></label>
          <label><span>Linehaul</span><input name="linehaul" type="number" min="0" step="0.01" value="0" required></label>
          <label><span>Fuel surcharge</span><input name="fuel_surcharge" type="number" min="0" step="0.01" value="0"></label>
          <label><span>Accessorials</span><input name="accessorials" type="number" min="0" step="0.01" value="0"></label>
          <label><span>Discount</span><input name="discount" type="number" min="0" step="0.01" value="0"></label>
          <label><span>Estimated carrier cost</span><input name="estimated_carrier_cost" type="number" min="0" step="0.01" value="0"></label>
          <div id="quote-preview" class="form-help"></div>
          <label><span>Valid until</span><input name="valid_until" type="date"></label>
          <label><span>Customer notes</span><textarea name="customer_notes"></textarea></label>
          <button class="primary-button" type="submit">Save Draft Quote</button>
          <p id="quote-message" class="form-message"></p>
        </form>
      </article>
    </dialog>
    <dialog id="rate-agreement-dialog" class="crud-dialog rate-agreement-dialog">
      <article class="panel">
        <div class="panel-header"><div><h2>Add Rate Agreement</h2><p>Save reusable lane, mileage, fuel and minimum-charge rules.</p></div><button id="close-rate-agreement-dialog" class="dialog-close" type="button" aria-label="Close">×</button></div>
        <form id="rate-agreement-form" class="record-form">
          <label><span>Customer</span><select name="customer_id" required></select></label>
          <label><span>Agreement name</span><input name="name" required placeholder="Chicago–Dallas 2026"></label>
          <label><span>Origin pattern</span><input name="origin_pattern" placeholder="Chicago, IL"></label>
          <label><span>Destination pattern</span><input name="destination_pattern" placeholder="Dallas, TX"></label>
          <label><span>Equipment</span><input name="equipment_type" placeholder="Dry Van"></label>
          <label><span>Fixed linehaul</span><input name="linehaul_rate" type="number" min="0" step="0.01" value="0"></label>
          <label><span>Rate per mile</span><input name="rate_per_mile" type="number" min="0" step="0.0001" value="0"></label>
          <label><span>Fuel surcharge %</span><input name="fuel_surcharge_percent" type="number" min="0" step="0.01" value="0"></label>
          <label><span>Minimum charge</span><input name="minimum_charge" type="number" min="0" step="0.01" value="0"></label>
          <div class="form-actions"><button type="submit">Save Agreement</button></div><p id="rate-agreement-message" class="form-message"></p>
        </form>
      </article>
    </dialog>
    <dialog id="recurring-template-dialog" class="crud-dialog quote-dialog"><article class="panel">
      <div class="panel-header"><div><h2>Add Recurring Load Template</h2><p>Create a reusable schedule; loads are generated manually in this foundation.</p></div><button id="close-recurring-template-dialog" class="dialog-close" type="button">×</button></div>
      <form id="recurring-template-form" class="record-form">
        <label><span>Template name</span><input name="name" required placeholder="Weekly Chicago–Dallas"></label><label><span>Customer</span><select name="customer_id" required></select></label>
        <label><span>Frequency</span><select name="frequency"><option value="weekly">Weekly</option><option value="daily">Daily</option><option value="biweekly">Biweekly</option><option value="monthly">Monthly</option></select></label><label><span>Interval</span><input name="interval_count" type="number" min="1" max="52" value="1"></label>
        <label><span>Next run</span><input name="next_run_date" type="date" required></label><label><span>Delivery offset days</span><input name="delivery_offset_days" type="number" min="0" max="90" value="1"></label>
        <label><span>Origin</span><input name="origin" required></label><label><span>Destination</span><input name="destination" required></label><label><span>Equipment</span><input name="equipment_type"></label><label><span>Commodity</span><input name="commodity"></label>
        <label><span>Loaded miles</span><input name="loaded_miles" type="number" min="0" step="0.1" value="0"></label><label><span>Linehaul</span><input name="linehaul" type="number" min="0" step="0.01" value="0"></label><label><span>Fuel surcharge</span><input name="fuel_surcharge" type="number" min="0" step="0.01" value="0"></label><label><span>Accessorials</span><input name="accessorials" type="number" min="0" step="0.01" value="0"></label>
        <div class="form-actions"><button type="submit">Save Template</button></div><p class="form-message"></p>
      </form></article></dialog>
    <section class="panel">
      <div class="panel-header"><div><h2>Quote Pipeline</h2><p>Customer pricing, estimated margin and controlled quote-to-load conversion.</p></div><div class="table-tools"><button id="reload-quotes" type="button">Refresh</button><button id="open-quote-dialog" class="primary-button" type="button">Create Quote</button></div></div>
      <div id="quotes-list" class="data-table"><div class="empty-state">Loading quotes...</div></div>
    </section>
    <section class="panel review-panel">
      <div class="panel-header"><div><h2>Customer Request Inbox</h2><p>Review portal quote/shipment requests, update status and reply in the customer thread.</p></div></div>
      <div id="customer-request-inbox" class="portal-load-list"><div class="empty-state">Loading customer requests...</div></div>
    </section>
    <section class="panel review-panel">
      <div class="panel-header"><div><h2>Customer Rate Agreements</h2><p>Lane, mileage, fuel and minimum-charge rules available to the quote calculator.</p></div><button id="open-rate-agreement-dialog" type="button">Add Rate Agreement</button></div>
      <div id="rate-agreements-list" class="data-table"><div class="empty-state">Loading agreements...</div></div>
    </section>
    <section class="panel review-panel">
      <div class="panel-header"><div><h2>Recurring Load Templates</h2><p>Reusable schedules with controlled manual load generation.</p></div><button id="open-recurring-template-dialog" type="button">Add Template</button></div>
      <div id="recurring-templates-list" class="data-table"><div class="empty-state">Loading templates...</div></div>
    </section>`;
  document.querySelector("#reload-quotes").addEventListener("click", loadQuotesWorkspace);
  document.querySelector("#open-quote-dialog").addEventListener("click", () => document.querySelector("#quote-dialog").showModal());
  document.querySelector("#close-quote-dialog").addEventListener("click", () => document.querySelector("#quote-dialog").close());
  document.querySelector("#open-rate-agreement-dialog").addEventListener("click", () => document.querySelector("#rate-agreement-dialog").showModal());
  document.querySelector("#close-rate-agreement-dialog").addEventListener("click", () => document.querySelector("#rate-agreement-dialog").close());
  document.querySelector("#open-recurring-template-dialog").addEventListener("click", () => document.querySelector("#recurring-template-dialog").showModal());
  document.querySelector("#close-recurring-template-dialog").addEventListener("click", () => document.querySelector("#recurring-template-dialog").close());
  document.querySelector("#quote-form").addEventListener("submit", handleCreateQuote);
  document.querySelector("#quote-form").addEventListener("input", updateQuotePreview);
  document.querySelector("#quote-form [name=rate_agreement_id]").addEventListener("change", applySelectedRateAgreement);
  document.querySelector("#rate-agreement-form").addEventListener("submit", handleCreateRateAgreement);
  document.querySelector("#recurring-template-form").addEventListener("submit", handleCreateRecurringTemplate);
  await loadQuotesWorkspace();
}

let quoteWorkspace = { customers: [], agreements: [], quotes: [], templates: [], requests: [], requestMessages: [] };

async function loadQuotesWorkspace() {
  const list = document.querySelector("#quotes-list");
  try {
    const [customers, agreements, quotes, templates, requests, requestMessages] = await Promise.all([listCustomers(), listCustomerRateAgreements(), listFreightQuotes(), listRecurringLoadTemplates(), listCustomerShipmentRequests(), listCustomerPortalMessages()]);
    quoteWorkspace = { customers, agreements, quotes, templates, requests, requestMessages };
    const options = '<option value="">Select customer</option>' + customers.map((item) => `<option value="${item.id}">${escapeHtml(item.name)}</option>`).join("");
    document.querySelector("#quote-form [name=customer_id]").innerHTML = options;
    document.querySelector("#rate-agreement-form [name=customer_id]").innerHTML = options;
    document.querySelector("#recurring-template-form [name=customer_id]").innerHTML = options;
    document.querySelector("#quote-form [name=rate_agreement_id]").innerHTML = '<option value="">Manual pricing</option>' + agreements.filter((item) => item.active).map((item) => `<option value="${item.id}">${escapeHtml(item.name)} · ${escapeHtml(item.customers?.name || "")}</option>`).join("");
    list.innerHTML = quotes.length ? quotes.map(renderQuoteRow).join("") : '<div class="empty-state">No quotes yet.</div>';
    document.querySelector("#rate-agreements-list").innerHTML = agreements.length ? agreements.map(renderRateAgreementRow).join("") : '<div class="empty-state">No rate agreements yet.</div>';
    document.querySelector("#recurring-templates-list").innerHTML = templates.length ? templates.map(renderRecurringTemplateRow).join("") : '<div class="empty-state">No recurring templates yet.</div>';
    document.querySelector("#customer-request-inbox").innerHTML=requests.length?requests.map((request)=>renderOperationsCustomerRequest(request,requestMessages.filter((message)=>message.request_id===request.id))).join(""):'<div class="empty-state">No customer requests yet.</div>';
    document.querySelectorAll("[data-quote-status]").forEach((button) => button.addEventListener("click", handleQuoteStatus));
    document.querySelectorAll("[data-convert-quote]").forEach((button) => button.addEventListener("click", handleConvertQuote));
    document.querySelectorAll("[data-generate-template]").forEach((button) => button.addEventListener("click", handleGenerateRecurringLoad));
    document.querySelectorAll("[data-request-status]").forEach((select)=>select.addEventListener("change",handleCustomerRequestStatus));
    document.querySelectorAll("[data-operations-request-message]").forEach((form)=>form.addEventListener("submit",handleOperationsRequestMessage));
    updateQuotePreview();
  } catch (error) { list.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`; }
}

function renderOperationsCustomerRequest(request,messages=[]){return `<article class="portal-load-card customer-request-card"><header><div><strong>${escapeHtml(request.request_no)}</strong><span>${escapeHtml(request.customers?.name||"Customer")} · ${escapeHtml(formatStatus(request.request_type))}</span><small>${escapeHtml(request.origin)} → ${escapeHtml(request.destination)} · Pickup ${formatDate(request.pickup_date)}</small></div><label class="request-status-control"><span>Status</span><select data-request-status="${request.id}">${["submitted","reviewing","approved","rejected","converted"].map((status)=>`<option value="${status}" ${status===request.status?"selected":""}>${formatStatus(status)}</option>`).join("")}</select></label></header>${request.notes?`<p class="customer-request-note">${escapeHtml(request.notes)}</p>`:""}<section class="portal-mini-section customer-request-thread"><h3>Conversation</h3><div class="customer-request-messages">${messages.map((message)=>`<article class="customer-request-message ${message.sender_id===currentSession?.user?.id?"is-own":""}"><div><strong>${message.sender_id===currentSession?.user?.id?"You":"Customer"}</strong><small>${formatDateTime(message.created_at)}</small></div><p>${escapeHtml(message.body)}</p></article>`).join("")||'<small>No messages yet.</small>'}</div><form data-operations-request-message="${request.id}" class="customer-request-message-form"><input name="body" required maxlength="2000" placeholder="Reply to customer"><button type="submit">Send Reply</button></form></section></article>`;}
async function handleCustomerRequestStatus(event){try{await updateCustomerShipmentRequest(event.currentTarget.dataset.requestStatus,{status:event.currentTarget.value});await loadQuotesWorkspace();}catch(error){window.alert(error.message);}}
async function handleOperationsRequestMessage(event){event.preventDefault();const form=event.currentTarget;try{const body=String(new FormData(form).get("body")||"").trim();if(!body)throw new Error("Reply is required.");await createCustomerPortalMessage({request_id:form.dataset.operationsRequestMessage,body});await loadQuotesWorkspace();}catch(error){window.alert(error.message);}}

function renderQuoteRow(quote) {
  const actions = quote.status === "draft" ? `<button data-quote-status="${quote.id}:sent">Mark Sent</button><button data-quote-status="${quote.id}:accepted">Accept</button>`
    : quote.status === "sent" ? `<button data-quote-status="${quote.id}:accepted">Accept</button><button data-quote-status="${quote.id}:rejected">Reject</button>`
    : quote.status === "accepted" ? `<button data-convert-quote="${quote.id}">Convert to Load</button>` : "";
  const marginTone = Number(quote.estimated_margin || 0) < 0 ? "negative" : "positive";
  return `<article class="quote-pipeline-row">
    <div class="quote-identity"><strong>${escapeHtml(quote.quote_no)}</strong><span>${escapeHtml(quote.customers?.name || "Customer")}</span><small>${escapeHtml(quote.origin)} <b>→</b> ${escapeHtml(quote.destination)}</small></div>
    <div class="quote-money"><strong>${formatMoney(quote.customer_total)}</strong><span>Customer total</span></div>
    <div class="quote-margin" data-tone="${marginTone}"><strong>${formatMoney(quote.estimated_margin)}</strong><span>${Number(quote.estimated_margin_percent || 0).toFixed(1)}% margin</span></div>
    <div class="quote-status"><em data-status="${escapeHtml(quote.status)}">${escapeHtml(formatStatus(quote.status))}</em></div>
    <div class="quote-actions">${quote.loads?.load_no ? `<a href="#loads/${quote.converted_load_id}">Open Load ${escapeHtml(quote.loads.load_no)}</a>` : actions}</div>
  </article>`;
}

function renderRateAgreementRow(item) {
  const pricing = Number(item.linehaul_rate || 0) > 0 ? `${formatMoney(item.linehaul_rate)} fixed` : `${formatMoney(item.rate_per_mile)}/mi`;
  return `<article class="rate-agreement-row">
    <div class="quote-identity"><strong>${escapeHtml(item.name)}</strong><span>${escapeHtml(item.customers?.name || "Customer")}</span><small>${escapeHtml(item.origin_pattern || "Any origin")} <b>→</b> ${escapeHtml(item.destination_pattern || "Any destination")}</small></div>
    <div class="agreement-rule"><strong>${pricing}</strong><span>${escapeHtml(item.equipment_type || "Any equipment")}</span></div>
    <div class="agreement-rule"><strong>${Number(item.fuel_surcharge_percent || 0).toFixed(2)}%</strong><span>Fuel surcharge</span></div>
    <div class="agreement-rule"><strong>${formatMoney(item.minimum_charge)}</strong><span>Minimum charge</span></div>
    <div class="quote-status"><em data-status="${item.active ? "active" : "inactive"}">${item.active ? "Active" : "Inactive"}</em></div>
  </article>`;
}

function renderRecurringTemplateRow(item) {
  return `<article class="rate-agreement-row"><div class="quote-identity"><strong>${escapeHtml(item.name)}</strong><span>${escapeHtml(item.customers?.name || "Customer")}</span><small>${escapeHtml(item.origin)} <b>→</b> ${escapeHtml(item.destination)}</small></div><div class="agreement-rule"><strong>${formatStatus(item.frequency)}</strong><span>Every ${item.interval_count} cycle${item.interval_count === 1 ? "" : "s"}</span></div><div class="agreement-rule"><strong>${formatDate(item.next_run_date)}</strong><span>Next run</span></div><div class="agreement-rule"><strong>${formatMoney(numberOrZero(item.linehaul)+numberOrZero(item.fuel_surcharge)+numberOrZero(item.accessorials))}</strong><span>Planned revenue</span></div><div class="quote-actions">${item.loads?.load_no ? `<a href="#loads/${item.last_generated_load_id}">Last ${escapeHtml(item.loads.load_no)}</a>` : ""}<button type="button" data-generate-template="${item.id}" ${item.active ? "" : "disabled"}>Generate Load</button></div></article>`;
}

function updateQuotePreview() {
  const form = document.querySelector("#quote-form"); if (!form) return;
  const pricing = calculateQuotePricing(Object.fromEntries(new FormData(form)));
  document.querySelector("#quote-preview").textContent = `Customer total ${formatMoney(pricing.customer_total)} · Estimated margin ${formatMoney(pricing.estimated_margin)} (${pricing.estimated_margin_percent.toFixed(1)}%)`;
}

function applySelectedRateAgreement(event) {
  const agreement = quoteWorkspace.agreements.find((item) => item.id === event.target.value); if (!agreement) return;
  const form = event.target.form; const miles = numberOrZero(form.elements.loaded_miles.value);
  const pricing = calculateQuotePricing({ loaded_miles: miles, linehaul: agreement.linehaul_rate, rate_per_mile: agreement.rate_per_mile, minimum_charge: agreement.minimum_charge, fuel_surcharge_percent: agreement.fuel_surcharge_percent });
  form.elements.customer_id.value = agreement.customer_id; form.elements.origin.value ||= agreement.origin_pattern || ""; form.elements.destination.value ||= agreement.destination_pattern || ""; form.elements.equipment_type.value ||= agreement.equipment_type || ""; form.elements.linehaul.value = pricing.linehaul; form.elements.fuel_surcharge.value = pricing.fuel_surcharge; updateQuotePreview();
}

async function handleCreateQuote(event) {
  event.preventDefault(); const form = event.currentTarget; const message = document.querySelector("#quote-message");
  try { const raw = Object.fromEntries(new FormData(form)); const payload = validateFreightQuoteInput(raw); const customer = quoteWorkspace.customers.find((item) => item.id === payload.customer_id); await createFreightQuote({ ...raw, ...payload, branch_id: customer?.branch_id, rate_agreement_id: raw.rate_agreement_id || null, equipment_type: raw.equipment_type || null, valid_until: raw.valid_until || null, customer_notes: raw.customer_notes || null }); form.reset(); message.textContent = ""; await loadQuotesWorkspace(); document.querySelector("#quote-dialog").close(); }
  catch (error) { message.textContent = error.message; }
}

async function handleCreateRateAgreement(event) {
  event.preventDefault(); const form = event.currentTarget; const raw = Object.fromEntries(new FormData(form)); const customer = quoteWorkspace.customers.find((item) => item.id === raw.customer_id); const message = document.querySelector("#rate-agreement-message");
  try { await createCustomerRateAgreement({ branch_id: customer?.branch_id, customer_id: raw.customer_id, name: raw.name.trim(), origin_pattern: raw.origin_pattern || null, destination_pattern: raw.destination_pattern || null, equipment_type: raw.equipment_type || null, linehaul_rate: numberOrZero(raw.linehaul_rate), rate_per_mile: numberOrZero(raw.rate_per_mile), fuel_surcharge_percent: numberOrZero(raw.fuel_surcharge_percent), minimum_charge: numberOrZero(raw.minimum_charge) }); form.reset(); message.textContent = ""; await loadQuotesWorkspace(); document.querySelector("#rate-agreement-dialog").close(); }
  catch (error) { message.textContent = error.message; }
}

async function handleCreateRecurringTemplate(event) {
  event.preventDefault(); const form=event.currentTarget; const raw=Object.fromEntries(new FormData(form)); const message=form.querySelector(".form-message");
  try { const input=validateRecurringLoadTemplate(raw); const customer=quoteWorkspace.customers.find((item)=>item.id===input.customer_id); await createRecurringLoadTemplate({...raw,...input,branch_id:customer?.branch_id,equipment_type:raw.equipment_type||null,commodity:raw.commodity||null,loaded_miles:numberOrZero(raw.loaded_miles),linehaul:numberOrZero(raw.linehaul),fuel_surcharge:numberOrZero(raw.fuel_surcharge),accessorials:numberOrZero(raw.accessorials)}); form.reset(); message.textContent=""; await loadQuotesWorkspace(); document.querySelector("#recurring-template-dialog").close(); }
  catch(error) { message.textContent=error.message; }
}

async function handleGenerateRecurringLoad(event) {
  const template=quoteWorkspace.templates.find((item)=>item.id===event.currentTarget.dataset.generateTemplate); if(!template)return;
  const suggested=`${template.name.replace(/[^a-z0-9]+/gi,"-").replace(/^-|-$/g,"").toUpperCase()}-${template.next_run_date.replaceAll("-","")}`;
  const loadNo=window.prompt("New load number",suggested); if(!loadNo)return;
  try { await generateLoadFromRecurringTemplate(template.id,loadNo,template.next_run_date); await loadQuotesWorkspace(); }
  catch(error) { window.alert(error.message); }
}

async function handleQuoteStatus(event) { const [id, status] = event.currentTarget.dataset.quoteStatus.split(":"); try { await updateFreightQuoteStatus(id, status); await loadQuotesWorkspace(); } catch (error) { window.alert(error.message); } }
async function handleConvertQuote(event) { const quote = quoteWorkspace.quotes.find((item) => item.id === event.currentTarget.dataset.convertQuote); const loadNo = window.prompt("New load number", quote ? `${quote.quote_no}-LOAD` : ""); if (!loadNo) return; try { await convertFreightQuoteToLoad(event.currentTarget.dataset.convertQuote, loadNo); await loadQuotesWorkspace(); } catch (error) { window.alert(error.message); } }

async function renderLoads() {
  await ensureTableSavedViews("loads");
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Phase 1", title: "Orders / Loads" })}
    <dialog id="duplicate-load-dialog" class="crud-dialog duplicate-load-dialog">
      <article class="panel">
        <div class="panel-header"><div><h2>Duplicate Load</h2><p id="duplicate-load-summary"></p></div><button id="close-duplicate-load" class="dialog-close" type="button" aria-label="Close">×</button></div>
        <form id="duplicate-load-form" class="record-form">
          <input name="source_load_id" type="hidden">
          <label><span>New load number</span><input name="load_no" required></label>
          <label><span>Pickup date</span><input name="pickup_date" type="date"></label>
          <label><span>Delivery date</span><input name="delivery_date" type="date"></label>
          <label class="checkbox-row duplicate-stops"><input name="copy_stops" type="checkbox" checked><span>Copy stops and shift appointments to the new pickup date</span></label>
          <p class="form-help">The new load starts as New and unassigned. Financials, documents, invoices, settlements and messages are not copied.</p>
          <div class="form-actions"><button class="primary-button" type="submit">Create Duplicate</button></div>
          <p id="duplicate-load-message" class="form-message"></p>
        </form>
      </article>
    </dialog>
    <section class="content-grid load-grid">
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2>All Loads</h2>
            <p>Core order board before the full dispatch kanban is added.</p>
          </div>
          <button id="reload-loads" type="button">Refresh</button>
        </div>
        <div class="load-saved-views" aria-label="Saved load views">
          ${[
            ["all", "All"],
            ["unassigned", "Unassigned"],
            ["today", "Today"],
            ["in_transit", "In Transit"],
            ["ready_to_bill", "Ready to Bill"],
          ].map(([value, label]) => `<button type="button" data-load-view="${value}">${label}</button>`).join("")}
        </div>
        <form id="load-filters" class="load-filters">
          <label><span>Search</span><input name="search" type="search" placeholder="Load, route, customer, driver..."></label>
          <label><span>Operational status</span><select name="status"><option value="">All statuses</option>${dispatchStatuses.map((status) => `<option value="${status}">${formatStatus(status)}</option>`).join("")}</select></label>
          <label><span>Customer</span><select name="customer_id"><option value="">All customers</option></select></label>
          <label><span>Driver</span><select name="driver_id"><option value="">All drivers</option></select></label>
          <label><span>Pickup from</span><input name="date_from" type="date"></label>
          <label><span>Pickup to</span><input name="date_to" type="date"></label>
          <label><span>Sort by</span><select name="sort_by"><option value="created_at">Created</option><option value="load_no">Load #</option><option value="pickup_date">Pickup</option><option value="delivery_date">Delivery</option><option value="operational_status">Status</option></select></label>
          <label><span>Direction</span><select name="sort_direction"><option value="desc">Descending</option><option value="asc">Ascending</option></select></label>
          <label><span>Rows</span><select name="page_size"><option value="5">5</option><option value="10" selected>10</option><option value="25">25</option></select></label>
          <button type="reset">Clear</button>
        </form>
        ${renderTableViewControls("loads")}
        <div id="loads-table" class="data-table"></div>
        <div id="loads-pagination" class="table-pagination"></div>
      </article>
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2 id="load-form-title">Add Order</h2>
            <p>Creates a basic load with customer, route and assignment.</p>
          </div>
        </div>
        <form id="load-form" class="record-form">
          <input name="id" type="hidden" />
          <label>
            <span>Load number</span>
            <input name="load_no" type="text" placeholder="MM-1001" required />
          </label>
          <label><span>Reference number</span><input name="reference_no" type="text" /></label>
          <label><span>Broker reference</span><input name="broker_reference" type="text" /></label>
          <label>
            <span>Customer</span>
            <select name="customer_id" id="load-customer" required></select>
          </label>
          <label>
            <span>Carrier</span>
            <select name="carrier_id" id="load-carrier"></select>
          </label>
          <label>
            <span>Origin</span>
            <input name="origin" type="text" placeholder="Chicago, IL" />
          </label>
          <label>
            <span>Destination</span>
            <input name="destination" type="text" placeholder="Dallas, TX" />
          </label>
          <label><span>Commodity</span><input name="commodity" type="text" /></label>
          <label><span>Weight (lb)</span><input name="weight_lbs" type="number" min="0" step="0.01" /></label>
          <label><span>Pieces</span><input name="pieces" type="number" min="0" step="1" /></label>
          <label><span>Pallets</span><input name="pallets" type="number" min="0" step="1" /></label>
          <label><span>Required equipment</span><input name="equipment_type" type="text" placeholder="Dry Van, Reefer, Flatbed..." /></label>
          <label><span>Minimum temperature (°F)</span><input name="temperature_min_f" type="number" step="0.1" /></label>
          <label><span>Maximum temperature (°F)</span><input name="temperature_max_f" type="number" step="0.1" /></label>
          <label class="checkbox-row"><input name="hazmat" type="checkbox" /><span>Hazmat load</span></label>
          <label><span>Loaded miles</span><input name="loaded_miles" type="number" min="0" step="0.1" /></label>
          <label><span>Empty miles</span><input name="empty_miles" type="number" min="0" step="0.1" /></label>
          <label>
            <span>Pickup date</span>
            <input name="pickup_date" type="date" />
          </label>
          <label>
            <span>Delivery date</span>
            <input name="delivery_date" type="date" />
          </label>
          <label>
            <span>Driver</span>
            <select name="driver_id" id="load-driver"></select>
          </label>
          <label>
            <span>Truck</span>
            <select name="truck_id" id="load-truck"></select>
          </label>
          <label>
            <span>Trailer</span>
            <select name="trailer_id" id="load-trailer"></select>
          </label>
          <label>
            <span>Operational status</span>
            <select name="status">
              ${dispatchStatuses.map((status) => `<option value="${status}">${formatStatus(status)}</option>`).join("")}
            </select>
          </label>
          <label><span>Internal notes</span><textarea name="internal_notes" rows="3"></textarea></label>
          <label><span>Customer-visible notes</span><textarea name="customer_notes" rows="3"></textarea></label>
          <div class="form-actions">
            <button id="load-submit" type="submit">Save Load</button>
            <button id="cancel-load-edit" type="button" hidden>Cancel edit</button>
          </div>
          <p id="load-message" class="form-message"></p>
        </form>
      </article>
    </section>
    <section id="load-details-panel" class="panel load-details-panel" hidden></section>
  `;

  document.querySelector("#reload-loads").addEventListener("click", loadLoads);
  document.querySelector("#load-filters").addEventListener("input", handleLoadFilters);
  document.querySelector("#load-filters").addEventListener("reset", () => {
    setTimeout(() => {
      Object.assign(loadListState, { search: "", status: "", customerId: "", driverId: "", dateFrom: "", dateTo: "", view: "all", page: 1, pageSize: 10, sortBy: "created_at", sortDirection: "desc" });
      globalSearch.value = "";
      clearSelectedTableView("loads");
      loadLoads();
    });
  });
  document.querySelectorAll("[data-load-view]").forEach((button) => button.addEventListener("click", handleLoadSavedView));
  bindTableViewControls("loads");
  document.querySelector("#load-form").addEventListener("submit", handleCreateLoad);
  document.querySelector("#cancel-load-edit").addEventListener("click", resetLoadForm);
  document.querySelector("#close-duplicate-load").addEventListener("click", closeDuplicateLoad);
  document.querySelector("#duplicate-load-form").addEventListener("submit", handleDuplicateLoad);
  configureCrudFormAccess("#load-form", "manage_operations", "Add Order", resetLoadForm);
  loadLoadFormOptions();
  const linkedLoadId = loadIdFromHash();
  loadLoads().then(() => {
    if (linkedLoadId) renderLoadDetails(linkedLoadId);
  });
}

async function loadLoadFormOptions() {
  const customerSelect = document.querySelector("#load-customer");
  const carrierSelect = document.querySelector("#load-carrier");
  const driverSelect = document.querySelector("#load-driver");
  const truckSelect = document.querySelector("#load-truck");
  const trailerSelect = document.querySelector("#load-trailer");

  customerSelect.innerHTML = '<option value="">Loading customers...</option>';
  carrierSelect.innerHTML = '<option value="">Loading carriers...</option>';
  driverSelect.innerHTML = '<option value="">Loading drivers...</option>';
  truckSelect.innerHTML = '<option value="">Loading trucks...</option>';
  trailerSelect.innerHTML = '<option value="">Loading trailers...</option>';

  try {
    const [customers, carriers, drivers, trucks, trailers] = await Promise.all([
      listCustomers(),
      listCarriers(),
      listDrivers(),
      listTrucks(),
      listTrailers(),
    ]);

    customerSelect.innerHTML = customers.length
      ? customers.map((customer) => `<option value="${customer.id}">${escapeHtml(customer.name)}</option>`).join("")
      : '<option value="">Add a customer first</option>';
    carrierSelect.innerHTML = [
      '<option value="">No outside carrier</option>',
      ...carriers.map((carrier) => `<option value="${carrier.id}">${escapeHtml(carrier.name)}</option>`),
    ].join("");
    driverSelect.innerHTML = [
      '<option value="">Unassigned</option>',
      ...drivers.map((driver) => `<option value="${driver.id}">${escapeHtml(driver.name)}</option>`),
    ].join("");
    truckSelect.innerHTML = [
      '<option value="">Unassigned</option>',
      ...trucks.map((truck) => `<option value="${truck.id}">${escapeHtml(truck.unit_no)}</option>`),
    ].join("");
    trailerSelect.innerHTML = [
      '<option value="">Unassigned</option>',
      ...trailers.map((trailer) => `<option value="${trailer.id}">${escapeHtml(trailer.unit_no)}</option>`),
    ].join("");
  } catch (error) {
    customerSelect.innerHTML = `<option value="">${escapeHtml(error.message)}</option>`;
    carrierSelect.innerHTML = '<option value="">Unavailable</option>';
    driverSelect.innerHTML = '<option value="">Unavailable</option>';
    truckSelect.innerHTML = '<option value="">Unavailable</option>';
    trailerSelect.innerHTML = '<option value="">Unavailable</option>';
  }
}

async function loadLoads() {
  const table = document.querySelector("#loads-table");
  table.innerHTML = '<div class="empty-state">Loading loads...</div>';

  try {
    const [result, customers, drivers] = await Promise.all([
      listLoadsPage({ ...loadListState, today: localDateIso() }),
      listCustomers(),
      listDrivers(),
    ]);
    const loads = result.records;
    Object.assign(loadListState, { loads, total: result.count });
    populateLoadFilterOptions(customers, drivers);
    syncLoadFilterForm();

    if (!loads.length) {
      table.innerHTML = '<div class="empty-state">No loads match these filters.</div>';
    }

    renderFilteredLoads();
  } catch (error) {
    table.innerHTML = `
      <div class="empty-state">
        Loads table is not ready: ${escapeHtml(error.message)}
      </div>
    `;
  }
}

function renderFilteredLoads() {
  const table = document.querySelector("#loads-table");
  const pagination = document.querySelector("#loads-pagination");
  if (!table || !pagination) return;
  const loads = loadListState.loads;
  const pageCount = Math.max(1, Math.ceil(loadListState.total / loadListState.pageSize));
  loadListState.page = Math.min(loadListState.page, pageCount);
  const start = (loadListState.page - 1) * loadListState.pageSize;
  const pageLoads = loads;

  document.querySelectorAll("[data-load-view]").forEach((button) => {
    button.dataset.active = String(button.dataset.loadView === loadListState.view);
  });

  if (!pageLoads.length) {
    table.innerHTML = '<div class="empty-state">No loads match these filters.</div>';
  } else {
    table.innerHTML = `
      <div class="data-row load-row data-row-head" style="${tableRowStyle("loads", loadListState.columns)}">
        ${can("manage_operations") ? '<input type="checkbox" data-select-page aria-label="Select this page">' : ""}
        <span>Load #</span>
        ${loadListState.columns.includes("status") ? "<span>Status</span>" : ""}
        ${loadListState.columns.includes("customer") ? "<span>Customer</span>" : ""}
        ${loadListState.columns.includes("route") ? "<span>Carrier / Route</span>" : ""}
        ${loadListState.columns.includes("pickup") ? "<span>Pickup</span>" : ""}
        <span></span>
      </div>
      ${pageLoads
        .map(
          (load) => `
            <div class="data-row load-row" data-operational-status="${escapeAttribute(operationalLoadStatus(load))}" style="${tableRowStyle("loads", loadListState.columns)}">
              ${can("manage_operations") ? `<input type="checkbox" data-select-row value="${load.id}" aria-label="Select load ${escapeAttribute(load.load_no)}" ${loadListState.selectedIds.has(load.id) ? "checked" : ""}>` : ""}
              <span class="load-number-cell"><strong>${escapeHtml(load.load_no)}</strong>${load.reference_no ? `<small>${escapeHtml(load.reference_no)}</small>` : ""}</span>
              ${loadListState.columns.includes("status") ? `<span class="load-status-stack"><strong class="load-operation-badge" data-state="${loadStatusTone(operationalLoadStatus(load))}">${formatStatus(operationalLoadStatus(load))}</strong><small>${formatStatus(billingLoadStatus(load))}</small></span>` : ""}
              ${loadListState.columns.includes("customer") ? `<span>${escapeHtml(load.customers?.name || "-")}</span>` : ""}
              ${loadListState.columns.includes("route") ? `<span class="load-route-cell">${load.carriers?.name ? `<strong>${escapeHtml(load.carriers.name)}</strong>` : ""}<small>${escapeHtml(routeLabel(load))}</small></span>` : ""}
              ${loadListState.columns.includes("pickup") ? `<span>${loadPickupWindow(load)}</span>` : ""}
              <div class="row-actions">
                <button class="row-action load-open-action" type="button" data-load-id="${load.id}">Open</button>
                ${can("manage_operations") ? `<button class="row-action" type="button" data-edit-load="${load.id}">Edit</button><button class="row-action" type="button" data-duplicate-load="${load.id}">Duplicate</button>` : ""}
              </div>
            </div>
          `
        )
        .join("")}
    `;

    table.querySelectorAll("[data-load-id]").forEach((button) => {
      button.addEventListener("click", () => renderLoadDetails(button.dataset.loadId));
    });
    table.querySelectorAll("[data-edit-load]").forEach((button) => {
      const load = loadListState.loads.find((item) => item.id === button.dataset.editLoad);
      button.addEventListener("click", () => fillLoadForm(load));
    });
    table.querySelectorAll("[data-duplicate-load]").forEach((button) => {
      const load = loadListState.loads.find((item) => item.id === button.dataset.duplicateLoad);
      button.addEventListener("click", () => openDuplicateLoad(load));
    });
    bindTableRowSelection("loads");
  }

  pagination.innerHTML = `<span>Showing ${loadListState.total ? start + 1 : 0}–${Math.min(start + pageLoads.length, loadListState.total)} of ${loadListState.total}</span><div><button type="button" data-load-page="prev" ${loadListState.page === 1 ? "disabled" : ""}>Previous</button><strong>Page ${loadListState.page} of ${pageCount}</strong><button type="button" data-load-page="next" ${loadListState.page === pageCount ? "disabled" : ""}>Next</button></div>`;
  pagination.querySelectorAll("[data-load-page]").forEach((button) => button.addEventListener("click", () => {
    loadListState.selectedIds.clear();
    loadListState.page += button.dataset.loadPage === "next" ? 1 : -1;
    loadLoads();
  }));
}

function loadStatusTone(status) {
  if (["delivered"].includes(status)) return "ok";
  if (["cancelled", "exception"].includes(status)) return "error";
  if (["dispatched", "en_route", "arrived_pickup", "loading", "picked_up", "in_transit", "arrived_delivery", "unloading"].includes(status)) return "live";
  return "";
}

const driverQueueKey = "cyrra-driver-offline-queue";

function driverQueuedItems() {
  try {
    return JSON.parse(localStorage.getItem(driverQueueKey) || "[]");
  } catch {
    return [];
  }
}

function setDriverQueuedItems(items) {
  localStorage.setItem(driverQueueKey, JSON.stringify(items));
}

function queueDriverAction(item) {
  const items = driverQueuedItems();
  items.push({ id: crypto.randomUUID(), created_at: new Date().toISOString(), ...item });
  setDriverQueuedItems(items);
  return items.length;
}

async function flushDriverQueue() {
  const items = driverQueuedItems();
  if (!items.length) return { sent: 0, remaining: 0 };
  const remaining = [];
  let sent = 0;
  for (const item of items) {
    try {
      if (item.type === "message") {
        await createLoadMessage(item.loadId, currentSession.user.id, item.body);
      } else if (item.type === "problem") {
        await createDriverProblemReport(item.report);
      } else {
        remaining.push(item);
        continue;
      }
      sent += 1;
    } catch {
      remaining.push(item);
    }
  }
  setDriverQueuedItems(remaining);
  return { sent, remaining: remaining.length };
}

function driverLoadSortValue(load) {
  return Date.parse(load.pickup_date || load.delivery_date || load.created_at || "") || 0;
}

function renderDriverStopList(stops = []) {
  if (!stops.length) return '<div class="empty-state compact-empty">No stops assigned yet.</div>';
  return [...stops]
    .sort((a, b) => a.stop_order - b.stop_order)
    .map((stop) => `
      <article class="driver-stop">
        <span>${escapeHtml(formatStatus(stop.type))}</span>
        <div>
          <strong>${escapeHtml(stop.facility || "Stop")}</strong>
          <small>${escapeHtml(stop.address || "Address pending")}</small>
          <small>${formatDateTime(stop.appointment_from)}${stop.appointment_to ? ` - ${formatDateTime(stop.appointment_to)}` : ""}</small>
        </div>
      </article>
    `).join("");
}

function renderDriverMessages(messages = []) {
  if (!messages.length) return '<div class="empty-state compact-empty">No dispatch messages yet.</div>';
  return messages.slice(-3).map((message) => `<article class="driver-message ${message.sender_id === currentSession?.user?.id ? "is-own" : ""}">
    <strong>${message.sender_id === currentSession?.user?.id ? "You" : "Dispatch"}</strong>
    <p>${escapeHtml(message.body)}</p>
    <small>${formatDateTime(message.created_at)}</small>
  </article>`).join("");
}

function renderDriverProblemReports(reports = []) {
  if (!reports.length) return '<div class="empty-state compact-empty">No problem reports for this load.</div>';
  return reports.slice(0, 3).map((report) => `<article class="driver-report-row" data-severity="${escapeAttribute(report.severity)}">
    <div><strong>${escapeHtml(formatStatus(report.category))}</strong><span>${escapeHtml(formatStatus(report.status))}</span></div>
    <p>${escapeHtml(report.description)}</p>
    <small>${formatDateTime(report.created_at)}</small>
  </article>`).join("");
}

function renderDriverLoadCard(load, messages, reports) {
  const statusOptions = driverStatusOptions(load);
  const opStatus = operationalLoadStatus(load);
  return `
    <article class="driver-load-card" data-load-id="${load.id}">
      <header>
        <div>
          <span class="eyebrow">${escapeHtml(load.load_no)}</span>
          <h2>${escapeHtml(routeLabel(load))}</h2>
        </div>
        <span class="status-pill">${escapeHtml(formatStatus(opStatus))}</span>
      </header>
      <div class="driver-load-meta">
        <span><strong>Pickup</strong>${loadPickupWindow(load)}</span>
        <span><strong>Truck</strong>${escapeHtml(load.trucks?.unit_no || "Not assigned")}</span>
        <span><strong>Trailer</strong>${escapeHtml(load.trailers?.unit_no || "Not assigned")}</span>
      </div>
      <section>
        <h3>Stops</h3>
        <div class="driver-stop-list">${renderDriverStopList(load.load_stops)}</div>
      </section>
      <section class="driver-card-grid">
        <form class="driver-status-form compact-form" data-driver-status-form>
          <input type="hidden" name="load_id" value="${load.id}">
          <label>Status
            <select name="status" ${statusOptions.length ? "" : "disabled"}>
              ${statusOptions.length ? statusOptions.map((status) => `<option value="${status}">${formatStatus(status)}</option>`).join("") : `<option value="">No driver action available</option>`}
            </select>
          </label>
          <button type="submit" ${statusOptions.length ? "" : "disabled"}>Update Status</button>
          <span class="form-message"></span>
        </form>
        <form class="driver-pod-form compact-form" data-driver-pod-form>
          <input type="hidden" name="load_id" value="${load.id}">
          <label>POD / Camera Upload<input name="file" type="file" accept="application/pdf,image/jpeg,image/png" required></label>
          <button type="submit">Upload POD</button>
          <span class="form-message"></span>
        </form>
      </section>
      <section class="driver-card-grid">
        <div class="driver-thread">
          <h3>Dispatch Messages</h3>
          <div class="driver-message-list">${renderDriverMessages(messages)}</div>
          <form class="driver-message-form" data-driver-message-form>
            <input type="hidden" name="load_id" value="${load.id}">
            <textarea name="body" rows="2" placeholder="Message dispatch..." required></textarea>
            <button type="submit">Send</button>
            <span class="form-message"></span>
          </form>
        </div>
        <div class="driver-thread">
          <h3>Problem Report</h3>
          <div class="driver-report-list">${renderDriverProblemReports(reports)}</div>
          <form class="driver-problem-form" data-driver-problem-form>
            <input type="hidden" name="load_id" value="${load.id}">
            <label>Category<select name="category">
              ${["delay", "damage", "detention", "paperwork", "safety", "other"].map((category) => `<option value="${category}">${formatStatus(category)}</option>`).join("")}
            </select></label>
            <label>Severity<select name="severity"><option value="info">Info</option><option value="warning">Warning</option><option value="critical">Critical</option></select></label>
            <textarea name="description" rows="3" placeholder="Describe what happened..." required></textarea>
            <button type="submit">Report Problem</button>
            <span class="form-message"></span>
          </form>
        </div>
      </section>
    </article>
  `;
}

async function renderDriverPortal() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Driver PWA", title: "My Driver Work", status: "Loading" })}
    <section class="driver-portal">
      <article class="panel">
        <div class="panel-header">
          <div><h2>Assigned Loads</h2><p>Mobile workflow for stops, status, POD, problems and dispatch messages.</p></div>
          <div class="table-tools"><button id="driver-flush-queue" type="button">Sync Queue</button><button id="driver-refresh" type="button">Refresh</button></div>
        </div>
        <div id="driver-portal-status" class="driver-portal-status">Loading driver profile...</div>
        <div id="driver-loads" class="driver-load-list"></div>
      </article>
    </section>
  `;

  document.querySelector("#driver-refresh").addEventListener("click", renderDriverPortal);
  document.querySelector("#driver-flush-queue").addEventListener("click", async () => {
    const status = document.querySelector("#driver-portal-status");
    status.textContent = "Syncing queued driver actions...";
    const result = await flushDriverQueue();
    status.textContent = `Queue sync: ${result.sent} sent, ${result.remaining} remaining.`;
  });

  const status = document.querySelector("#driver-portal-status");
  const list = document.querySelector("#driver-loads");

  try {
    const profile = await getCurrentDriverProfile(currentSession?.user?.id, currentSession?.user?.email);
    if (!profile) {
      status.textContent = "No driver profile is linked to this login. Match the Auth user email to a driver profile or set driver.auth_user_id.";
      list.innerHTML = '<div class="empty-state">Driver profile is not linked yet.</div>';
      return;
    }

    const queueResult = await flushDriverQueue();
    const loads = (await listLoads())
      .filter((load) => load.driver_id === profile.id && !["cancelled"].includes(operationalLoadStatus(load)))
      .sort((a, b) => driverLoadSortValue(a) - driverLoadSortValue(b));

    const connectionStatus = document.querySelector("#connection-status");
    if (connectionStatus) connectionStatus.textContent = "Ready";
    status.textContent = `${profile.name} · ${profile.branches?.name || "Branch"} · ${loads.length} assigned load${loads.length === 1 ? "" : "s"} · queue ${queueResult.remaining}`;
    if (!loads.length) {
      list.innerHTML = '<div class="empty-state">No assigned driver loads right now.</div>';
      return;
    }

    const [messageResults, reportResults] = await Promise.all([
      Promise.all(loads.map((load) => listLoadMessages(load.id).catch(() => []))),
      Promise.all(loads.map((load) => listDriverProblemReports(load.id).catch(() => []))),
    ]);
    list.innerHTML = loads.map((load, index) => renderDriverLoadCard(load, messageResults[index], reportResults[index])).join("");
    bindDriverPortalActions();
  } catch (error) {
    status.textContent = "Driver App is unavailable.";
    list.innerHTML = `<div class="empty-state">${escapeHtml(error.message)}</div>`;
  }
}

function bindDriverPortalActions() {
  document.querySelectorAll("[data-driver-status-form]").forEach((form) => form.addEventListener("submit", handleDriverStatusUpdate));
  document.querySelectorAll("[data-driver-pod-form]").forEach((form) => form.addEventListener("submit", handleDriverPodUpload));
  document.querySelectorAll("[data-driver-message-form]").forEach((form) => form.addEventListener("submit", handleDriverMessageSend));
  document.querySelectorAll("[data-driver-problem-form]").forEach((form) => form.addEventListener("submit", handleDriverProblemReport));
}

async function handleDriverStatusUpdate(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = form.querySelector(".form-message");
  const button = form.querySelector("button[type='submit']");
  button.disabled = true;
  message.textContent = "Updating...";
  try {
    await driverUpdateLoadStatus(form.elements.load_id.value, form.elements.status.value);
    message.textContent = "Status updated.";
    await renderDriverPortal();
  } catch (error) {
    message.textContent = error.message;
    button.disabled = false;
  }
}

async function handleDriverPodUpload(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = form.querySelector(".form-message");
  const button = form.querySelector("button[type='submit']");
  button.disabled = true;
  message.textContent = "Uploading...";
  try {
    await uploadTmsDocument({
      entityType: "load",
      entityId: form.elements.load_id.value,
      docType: "pod",
      status: "pending",
      file: form.elements.file.files[0],
      uploadedBy: currentSession.user.id,
    });
    message.textContent = "POD uploaded for review.";
    form.reset();
  } catch (error) {
    message.textContent = error.message;
  } finally {
    button.disabled = false;
  }
}

async function handleDriverMessageSend(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = form.querySelector(".form-message");
  const button = form.querySelector("button[type='submit']");
  button.disabled = true;
  try {
    const body = validateMessageBody(form.elements.body.value);
    await createLoadMessage(form.elements.load_id.value, currentSession.user.id, body);
    message.textContent = "Message sent.";
    form.reset();
    await renderDriverPortal();
  } catch (error) {
    const body = form.elements.body.value.trim();
    if (body) {
      const count = queueDriverAction({ type: "message", loadId: form.elements.load_id.value, body });
      message.textContent = `Message queued locally (${count}).`;
      form.reset();
    } else {
      message.textContent = error.message;
    }
  } finally {
    button.disabled = false;
  }
}

async function handleDriverProblemReport(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = form.querySelector(".form-message");
  const button = form.querySelector("button[type='submit']");
  button.disabled = true;
  try {
    const report = validateDriverProblemReport({
      category: form.elements.category.value,
      severity: form.elements.severity.value,
      description: form.elements.description.value,
    });
    const payload = { load_id: form.elements.load_id.value, reported_by: currentSession.user.id, ...report };
    await createDriverProblemReport(payload);
    message.textContent = "Problem reported.";
    form.reset();
    await renderDriverPortal();
  } catch (error) {
    const description = form.elements.description.value.trim();
    if (description) {
      const report = {
        load_id: form.elements.load_id.value,
        reported_by: currentSession.user.id,
        category: form.elements.category.value,
        severity: form.elements.severity.value,
        description,
      };
      const count = queueDriverAction({ type: "problem", report });
      message.textContent = `Problem report queued locally (${count}).`;
      form.reset();
    } else {
      message.textContent = error.message;
    }
  } finally {
    button.disabled = false;
  }
}

function openDuplicateLoad(load) {
  if (!load) return;
  const form = document.querySelector("#duplicate-load-form");
  form.reset();
  form.elements.source_load_id.value = load.id;
  form.elements.load_no.value = `${load.load_no}-COPY`;
  form.elements.pickup_date.value = load.pickup_date || "";
  form.elements.delivery_date.value = load.delivery_date || "";
  form.elements.copy_stops.checked = true;
  document.querySelector("#duplicate-load-summary").textContent = `Source ${load.load_no} · ${routeLabel(load)}`;
  document.querySelector("#duplicate-load-message").textContent = "";
  document.querySelector("#duplicate-load-dialog").showModal();
}

function closeDuplicateLoad() {
  const dialog = document.querySelector("#duplicate-load-dialog");
  if (dialog?.open) dialog.close();
}

async function handleDuplicateLoad(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = document.querySelector("#duplicate-load-message");
  const button = form.querySelector("button[type='submit']");
  button.disabled = true;
  message.textContent = "Creating duplicate load...";
  try {
    const input = validateDuplicateLoadInput({
      loadNo: form.elements.load_no.value,
      pickupDate: form.elements.pickup_date.value,
      deliveryDate: form.elements.delivery_date.value,
    });
    const duplicatedLoadId = await duplicateLoad({
      sourceLoadId: form.elements.source_load_id.value,
      ...input,
      copyStops: form.elements.copy_stops.checked,
    });
    closeDuplicateLoad();
    await loadLoads();
    await renderLoadDetails(duplicatedLoadId);
  } catch (error) {
    message.textContent = error.message;
  } finally {
    button.disabled = false;
  }
}

function populateLoadFilterOptions(customers, drivers) {
  const form = document.querySelector("#load-filters");
  form.elements.customer_id.innerHTML = '<option value="">All customers</option>' + customers.map(({ id, name }) => `<option value="${id}">${escapeHtml(name)}</option>`).join("");
  form.elements.driver_id.innerHTML = '<option value="">All drivers</option>' + drivers.map(({ id, name }) => `<option value="${id}">${escapeHtml(name)}</option>`).join("");
}

function syncLoadFilterForm() {
  const form = document.querySelector("#load-filters");
  form.elements.search.value = loadListState.search;
  form.elements.status.value = loadListState.status;
  form.elements.customer_id.value = loadListState.customerId;
  form.elements.driver_id.value = loadListState.driverId;
  form.elements.date_from.value = loadListState.dateFrom;
  form.elements.date_to.value = loadListState.dateTo;
  form.elements.sort_by.value = loadListState.sortBy;
  form.elements.sort_direction.value = loadListState.sortDirection;
  form.elements.page_size.value = String(loadListState.pageSize);
  globalSearch.value = loadListState.search;
}

function handleLoadFilters(event) {
  const form = event.currentTarget;
  Object.assign(loadListState, {
    search: form.elements.search.value,
    status: form.elements.status.value,
    customerId: form.elements.customer_id.value,
    driverId: form.elements.driver_id.value,
    dateFrom: form.elements.date_from.value,
    dateTo: form.elements.date_to.value,
    sortBy: form.elements.sort_by.value,
    sortDirection: form.elements.sort_direction.value,
    pageSize: Number(form.elements.page_size.value),
    view: "all",
    page: 1,
  });
  globalSearch.value = loadListState.search;
  loadListState.selectedIds.clear();
  clearSelectedTableView("loads");
  clearTimeout(loadFilterTimer);
  loadFilterTimer = setTimeout(loadLoads, event.target.name === "search" ? 250 : 0);
}

function handleLoadSavedView(event) {
  Object.assign(loadListState, { view: event.currentTarget.dataset.loadView, status: "", page: 1 });
  syncLoadFilterForm();
  loadLoads();
}

async function fillLoadForm(load) {
  const form = document.querySelector("#load-form");
  await loadLoadFormOptions();
  form.elements.id.value = load.id;
  form.elements.load_no.value = load.load_no || "";
  form.elements.reference_no.value = load.reference_no || "";
  form.elements.broker_reference.value = load.broker_reference || "";
  form.elements.customer_id.value = load.customer_id || "";
  form.elements.carrier_id.value = load.carrier_id || "";
  form.elements.origin.value = load.origin || "";
  form.elements.destination.value = load.destination || "";
  form.elements.commodity.value = load.commodity || "";
  form.elements.weight_lbs.value = load.weight_lbs ?? "";
  form.elements.pieces.value = load.pieces ?? "";
  form.elements.pallets.value = load.pallets ?? "";
  form.elements.equipment_type.value = load.equipment_type || "";
  form.elements.temperature_min_f.value = load.temperature_min_f ?? "";
  form.elements.temperature_max_f.value = load.temperature_max_f ?? "";
  form.elements.hazmat.checked = Boolean(load.hazmat);
  form.elements.loaded_miles.value = load.loaded_miles ?? "";
  form.elements.empty_miles.value = load.empty_miles ?? "";
  form.elements.pickup_date.value = load.pickup_date || "";
  form.elements.delivery_date.value = load.delivery_date || "";
  form.elements.driver_id.value = load.driver_id || "";
  form.elements.truck_id.value = load.truck_id || "";
  form.elements.trailer_id.value = load.trailer_id || "";
  form.elements.status.value = operationalLoadStatus(load);
  form.elements.internal_notes.value = load.internal_notes || load.notes || "";
  form.elements.customer_notes.value = load.customer_notes || "";
  document.querySelector("#load-form-title").textContent = "Edit Order";
  document.querySelector("#load-submit").textContent = "Update Load";
  document.querySelector("#cancel-load-edit").hidden = false;
  document.querySelector("#load-message").textContent = "";
  openFormDialog(form);
  form.scrollIntoView({ behavior: "smooth", block: "start" });
}

function resetLoadForm() {
  const form = document.querySelector("#load-form");
  form.reset();
  form.elements.id.value = "";
  document.querySelector("#load-form-title").textContent = "Add Order";
  document.querySelector("#load-submit").textContent = "Save Load";
  document.querySelector("#cancel-load-edit").hidden = true;
  document.querySelector("#load-message").textContent = "";
  closeFormDialog(form);
}

async function renderLoadDetails(loadId, activeTab = "Overview") {
  const panel = document.querySelector("#load-details-panel");
  panel.hidden = false;
  panel.innerHTML = '<div class="empty-state">Loading load details...</div>';

  try {
    const [load, stops, financials, documents, invoices, activity, messages, facilities, trackingLinks, trackingLogs, carrierTenders, carrierBids, carriers] = await Promise.all([
      getLoad(loadId),
      listLoadStops(loadId),
      getLoadFinancials(loadId),
      listLoadDocuments(loadId),
      listInvoices(),
      can("view_audit") ? listLoadActivity(loadId) : Promise.resolve([]),
      listLoadMessages(loadId),
      listFacilities(),
      can("view_tracking") ? listPublicTrackingLinks(loadId) : Promise.resolve([]),
      can("view_tracking") ? listPublicTrackingAccessLogs(loadId) : Promise.resolve([]),
      can("manage_operations") ? listCarrierTenders(loadId) : Promise.resolve([]),
      can("manage_operations") ? listCarrierLoadBids(loadId) : Promise.resolve([]),
      can("manage_operations") ? listCarriers() : Promise.resolve([]),
    ]);
    const approvedPod = documents.some((document) => document.doc_type === "pod" && document.status === "approved");
    const activeInvoice = invoices.find((invoice) => invoice.load_id === load.id && invoice.status !== "void");
    const opStatus = operationalLoadStatus(load);
    const docStatus = loadDocumentStatus(load, documents);
    const billStatus = billingLoadStatus(load, activeInvoice);
    const invoiceAmount = financials
      ? numberOrZero(financials.linehaul) + numberOrZero(financials.fsc) + numberOrZero(financials.accessorials)
      : 0;
    const pickupStop = stops.find((stop) => stop.type === "pickup") || stops[0];
    const deliveryStop = [...stops].reverse().find((stop) => stop.type === "delivery") || stops.at(-1);

    panel.innerHTML = `
      <div class="panel-header">
        <div>
          <h2>Load ${escapeHtml(load.load_no)}</h2>
          <p>${escapeHtml(routeLabel(load))}</p>
        </div>
        <span class="status-pill" data-state="ok">${formatStatus(opStatus)}</span>
      </div>
      <div class="detail-layout">
        <section class="billing-readiness">
          <div>
            <span class="status-pill" data-state="${approvedPod ? "ok" : "error"}">${formatStatus(docStatus)}</span>
            <strong>${activeInvoice ? `Invoice ${escapeHtml(activeInvoice.invoice_no)} · ${formatStatus(activeInvoice.status)}` : formatStatus(billStatus)}</strong>
            <small>${billingReadinessMessage(load, approvedPod, activeInvoice)}</small>
          </div>
          <div class="billing-readiness-actions">
            ${opStatus === "delivered" && billStatus === "not_ready" && can("manage_operations") ? `<button type="button" data-ready-to-bill="${load.id}" ${approvedPod ? "" : "disabled title=\"Upload and approve a POD first\""}>${approvedPod ? "Mark Ready to Bill" : "Approve POD First"}</button>` : ""}
            ${billStatus === "ready_to_bill" && !activeInvoice && can("manage_finance") ? `<button class="primary-button" type="button" data-create-load-invoice="${load.id}" data-customer-id="${load.customer_id}" data-load-no="${escapeAttribute(load.load_no)}" data-amount="${invoiceAmount}" ${invoiceAmount > 0 ? "" : "disabled title=\"Save positive load revenue first\""}>${invoiceAmount > 0 ? "Create Draft Invoice" : "Add Financials First"}</button>` : ""}
            ${activeInvoice && can("view_billing") ? '<a class="button-link" href="#billing">Open Billing</a>' : ""}
          </div>
          <p id="billing-readiness-message" class="form-message"></p>
        </section>
        <section class="detail-summary">
          <article><span>Customer</span><strong>${escapeHtml(load.customers?.name || "-")}</strong></article>
          <article><span>Carrier</span><strong>${escapeHtml(load.carriers?.name || "No outside carrier")}</strong></article>
          <article><span>Driver</span><strong>${escapeHtml(load.drivers?.name || "Unassigned")}</strong></article>
          <article><span>Truck</span><strong>${escapeHtml(load.trucks?.unit_no || "Unassigned")}</strong></article>
          <article><span>Trailer</span><strong>${escapeHtml(load.trailers?.unit_no || "Unassigned")}</strong></article>
          <article><span>Operational status</span><strong>${formatStatus(opStatus)}</strong></article>
          <article><span>Document status</span><strong>${formatStatus(docStatus)}</strong></article>
          <article><span>Billing status</span><strong>${formatStatus(billStatus)}</strong></article>
          <article><span>Pickup appointment</span><strong>${formatStopWindow(pickupStop, load.pickup_date)}</strong></article>
          <article><span>Delivery appointment</span><strong>${formatStopWindow(deliveryStop, load.delivery_date)}</strong></article>
          <article><span>Reference</span><strong>${escapeHtml(load.reference_no || "-")}</strong></article>
          <article><span>Broker reference</span><strong>${escapeHtml(load.broker_reference || "-")}</strong></article>
          <article><span>Commodity</span><strong>${escapeHtml(load.commodity || "-")}</strong></article>
          <article><span>Weight</span><strong>${load.weight_lbs == null ? "-" : `${formatNumber(load.weight_lbs)} lb`}</strong></article>
          <article><span>Pieces / Pallets</span><strong>${formatNumber(load.pieces)} / ${formatNumber(load.pallets)}</strong></article>
          <article><span>Equipment</span><strong>${escapeHtml(load.equipment_type || "-")}</strong></article>
          <article><span>Temperature</span><strong>${formatTemperature(load)}</strong></article>
          <article><span>Hazmat</span><strong>${load.hazmat ? "Yes" : "No"}</strong></article>
          <article><span>Miles (loaded / empty / total)</span><strong>${formatNumber(load.loaded_miles)} / ${formatNumber(load.empty_miles)} / ${formatNumber(loadTotalMiles(load))}</strong></article>
          <article><span>Internal notes</span><strong>${escapeHtml(load.internal_notes || load.notes || "-")}</strong></article>
          <article><span>Customer-visible notes</span><strong>${escapeHtml(load.customer_notes || "-")}</strong></article>
        </section>
        <section class="stops-section">
          <div class="subsection-header">
            <div>
              <h3>Stops</h3>
              <p>Pickup, delivery and intermediate appointments.</p>
            </div>
          </div>
          <div id="stops-list" class="stops-list">${renderStops(stops)}</div>
        </section>
        <section class="stop-form-wrap" ${can("manage_operations") ? "" : "hidden"}>
          <div class="subsection-header">
            <div>
              <h3 id="stop-form-title">Add Stop</h3>
              <p>Create or update pickup, delivery and intermediate appointments.</p>
            </div>
          </div>
          <form id="stop-form" class="record-form compact-form">
            <input name="load_id" type="hidden" value="${load.id}" />
            <input name="id" type="hidden" />
            <input name="stop_order" type="hidden" value="${stops.length + 1}" />
            <label>
              <span>Type</span>
              <select name="type">
                <option value="pickup">Pickup</option>
                <option value="delivery">Delivery</option>
                <option value="stop">Stop</option>
              </select>
            </label>
            <label>
              <span>Address book facility</span>
              <select name="facility_id"><option value="">Custom location</option>${facilities.filter((facility) => facility.status === "active").map((facility) => `<option value="${facility.id}">${escapeHtml(facility.name)} · ${escapeHtml(`${facility.city}, ${facility.state}`)}</option>`).join("")}</select>
            </label>
            <label>
              <span>Facility name</span>
              <input name="facility" type="text" />
            </label>
            <label>
              <span>Address</span>
              <input name="address" type="text" />
            </label>
            <label><span>Contact name</span><input name="contact_name" type="text" /></label>
            <label><span>Contact phone</span><input name="contact_phone" type="tel" /></label>
            <label>
              <span>Appointment from</span>
              <input name="appointment_from" type="datetime-local" />
            </label>
            <label>
              <span>Appointment to</span>
              <input name="appointment_to" type="datetime-local" />
            </label>
            <label><span>Actual arrival</span><input name="actual_arrival" type="datetime-local" /></label>
            <label><span>Actual departure</span><input name="actual_departure" type="datetime-local" /></label>
            <label>
              <span>Reference</span>
              <input name="reference" type="text" />
            </label>
            <label>
              <span>Instructions</span>
              <input name="instructions" type="text" />
            </label>
            <div class="form-actions"><button id="stop-submit" type="submit">Save Stop</button><button id="cancel-stop-edit" type="button" hidden>Cancel edit</button></div>
            <p id="stop-message" class="form-message"></p>
          </form>
        </section>
        <section class="financials-section" ${can("manage_finance") ? "" : "hidden"}>
          <div class="subsection-header">
            <div>
              <h3>Financials</h3>
              <p>Revenue, cost and margin for this load.</p>
            </div>
          </div>
          ${renderFinancialSummary(financials)}
          <form id="financials-form" class="record-form compact-form" ${can("manage_finance") ? "" : "hidden"}>
            <input name="load_id" type="hidden" value="${load.id}" />
            <label>
              <span>Linehaul</span>
              <input name="linehaul" type="number" min="0" step="0.01" value="${financials?.linehaul ?? 0}" />
            </label>
            <label>
              <span>FSC</span>
              <input name="fsc" type="number" min="0" step="0.01" value="${financials?.fsc ?? 0}" />
            </label>
            <label>
              <span>Accessorials</span>
              <input name="accessorials" type="number" min="0" step="0.01" value="${financials?.accessorials ?? 0}" />
            </label>
            <label>
              <span>Deductions</span>
              <input name="deductions" type="number" min="0" step="0.01" value="${financials?.deductions ?? 0}" />
            </label>
            <label>
              <span>Carrier cost</span>
              <input name="carrier_cost" type="number" min="0" step="0.01" value="${financials?.carrier_cost ?? 0}" />
            </label>
            <label>
              <span>Driver pay</span>
              <input name="driver_pay" type="number" min="0" step="0.01" value="${financials?.driver_pay ?? 0}" />
            </label>
            <button type="submit">Save Financials</button>
            <p id="financials-message" class="form-message"></p>
          </form>
        </section>
        <section class="documents-section">
          <div class="subsection-header">
            <div>
              <h3>Documents</h3>
              <p>Private PDF and image uploads, maximum 15 MB.</p>
            </div>
          </div>
          <div id="documents-list" class="documents-list">${renderDocuments(documents)}</div>
          <form id="document-form" class="record-form compact-form">
            <input name="entity_id" type="hidden" value="${load.id}" />
            <label>
              <span>Document type</span>
              <select name="doc_type">
                <option value="rate_con">Rate confirmation</option>
                <option value="bol">BOL</option>
                <option value="pod">POD</option>
                <option value="fuel">Fuel</option>
                <option value="lumper">Lumper</option>
                <option value="invoice">Invoice</option>
                <option value="photo">Photo</option>
              </select>
            </label>
            <label>
              <span>File</span>
              <input name="file" type="file" accept="application/pdf,image/jpeg,image/png" required />
            </label>
            <label>
              <span>Status</span>
              <select name="status">
                <option value="pending">Pending</option>
                <option value="approved">Approved</option>
                <option value="needs_review">Needs review</option>
                <option value="rejected">Rejected</option>
              </select>
            </label>
            <button type="submit">Save Document</button>
            <p id="document-message" class="form-message"></p>
          </form>
        </section>
        <section class="messages-section">
          <div class="subsection-header"><div><h3>Messages</h3><p>Internal communication connected to this load.</p></div></div>
          <div id="load-message-thread" class="load-message-thread">${renderLoadMessages(messages)}</div>
          <form id="load-message-form" class="load-message-form">
            <input name="load_id" type="hidden" value="${load.id}">
            <label><span>New message</span><textarea name="body" rows="3" maxlength="4000" placeholder="Write an internal update..." required></textarea></label>
            <div><button class="primary-button" type="submit">Send Message</button><small><span id="load-message-count">0</span>/4000</small></div>
            <p id="load-message-feedback" class="form-message"></p>
          </form>
        </section>
        ${can("manage_operations") ? `<section class="carrier-tender-section">
          <div class="subsection-header"><div><h3>Carrier Bids & Tender</h3><p>Compare carrier bids, select a winner, then send the standard tender.</p></div></div>
          ${renderCarrierBidManager(load, financials, carrierBids, carriers)}
          ${renderCarrierTenderManager(load, carrierTenders)}
        </section>` : ""}
        ${can("view_tracking") ? `<section class="public-tracking-section">
          <div class="subsection-header"><div><h3>Public Tracking</h3><p>Time-limited customer-safe shipment link with privacy controls.</p></div></div>
          ${renderPublicTrackingManager(load, trackingLinks, trackingLogs)}
        </section>` : ""}
        ${can("view_audit") ? `<section class="activity-section">
          <div class="subsection-header"><div><h3>Activity</h3><p>Protected audit history for this load and its related records.</p></div></div>
          <div class="load-activity-list">${renderLoadActivity(activity)}</div>
        </section>` : ""}
      </div>
    `;

    document.querySelector("#stop-form")?.addEventListener("submit", handleCreateStop);
    document.querySelector('#stop-form [name="facility_id"]')?.addEventListener("change", (event) => populateStopFromFacility(event.currentTarget, facilities));
    document.querySelector("#cancel-stop-edit")?.addEventListener("click", () => resetStopForm(stops.length + 1));
    bindStopActions(stops, loadId);
    document.querySelector("#financials-form")?.addEventListener("submit", handleSaveFinancials);
    document.querySelector("#document-form").addEventListener("submit", handleCreateDocument);
    document.querySelector("#load-message-form").addEventListener("submit", handleCreateLoadMessage);
    document.querySelector("#load-message-form textarea").addEventListener("input", updateLoadMessageCount);
    document.querySelector("#carrier-tender-form")?.addEventListener("submit", handleCreateCarrierTender);
    document.querySelector("#carrier-bid-invite-form")?.addEventListener("submit", handleCarrierBidInvite);
    document.querySelectorAll("[data-select-carrier-bid]").forEach((button) => button.addEventListener("click", handleCarrierBidSelect));
    document.querySelector("#public-tracking-form")?.addEventListener("submit", handleCreatePublicTrackingLink);
    document.querySelectorAll("[data-revoke-tracking-link]").forEach((button) => button.addEventListener("click", handleRevokePublicTrackingLink));
    document.querySelector("[data-ready-to-bill]")?.addEventListener("click", handleMarkReadyToBill);
    document.querySelector("[data-create-load-invoice]")?.addEventListener("click", handleCreateLoadInvoice);
    bindDocumentDeleteButtons(() => renderLoadDetails(loadId, "Documents"));
    setupLoadDetailTabs(panel, activeTab);
    panel.scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (error) {
    panel.innerHTML = `<div class="empty-state">Load details are not ready: ${escapeHtml(error.message)}</div>`;
  }
}

function defaultTenderExpiryValue() {
  const date = new Date(Date.now() + 24 * 60 * 60 * 1000);
  return new Date(date.getTime() - date.getTimezoneOffset() * 60000).toISOString().slice(0, 16);
}

function renderCarrierBidManager(load, financials, bids = [], carriers = []) {
  const revenue = numberOrZero(financials?.linehaul)+numberOrZero(financials?.fsc)+numberOrZero(financials?.accessorials)-numberOrZero(financials?.deductions);
  const compared = carrierBidComparison(bids, revenue);
  const available = carriers.filter((carrier) => carrier.status === "active" && !bids.some((bid) => bid.carrier_id === carrier.id));
  return `<section class="carrier-bid-manager"><form id="carrier-bid-invite-form" class="record-form compact-form"><input name="load_id" type="hidden" value="${load.id}"><input name="branch_id" type="hidden" value="${load.branch_id}"><label><span>Invite carrier</span><select name="carrier_id" required><option value="">Select carrier</option>${available.map((carrier) => `<option value="${carrier.id}">${escapeHtml(carrier.name)}</option>`).join("")}</select></label><label><span>Target amount</span><input name="target_amount" type="number" min="0" step="0.01"></label><label><span>Expires</span><input name="expires_at" type="datetime-local" value="${defaultTenderExpiryValue()}" required></label><button type="submit" ${available.length ? "" : "disabled"}>Invite to Bid</button><p class="form-message"></p></form>
    <div class="carrier-bid-grid">${compared.length ? compared.map((bid) => `<article class="carrier-bid-card"><header><strong>${escapeHtml(bid.carriers?.name || "Carrier")}</strong><em data-status="${bid.status}">${formatStatus(bid.status)}</em></header><div><span>Bid<strong>${bid.bid_amount ? formatMoney(bid.bid_amount) : "Waiting"}</strong></span><span>Transit<strong>${bid.transit_days ? `${bid.transit_days} days` : "-"}</strong></span><span>Margin<strong>${bid.bid_amount ? `${formatMoney(bid.margin)} (${bid.marginPercent.toFixed(1)}%)` : "-"}</strong></span></div>${bid.carrier_note ? `<p>${escapeHtml(bid.carrier_note)}</p>` : ""}${bid.status === "submitted" ? `<button type="button" data-select-carrier-bid="${bid.id}">Select Carrier</button>` : ""}</article>`).join("") : '<div class="empty-state compact-empty">No carriers invited to bid.</div>'}</div></section>`;
}

async function handleCarrierBidInvite(event) {
  event.preventDefault(); const form=event.currentTarget; const message=form.querySelector(".form-message");
  try { const values=Object.fromEntries(new FormData(form)); const input=validateCarrierBidInput(values); await createCarrierLoadBid({load_id:values.load_id,branch_id:values.branch_id,carrier_id:input.carrier_id,target_amount:input.target_amount || null,expires_at:new Date(input.expires_at).toISOString(),status:"invited"}); await renderLoadDetails(values.load_id,"Tender"); }
  catch(error) { message.textContent=error.message; }
}

async function handleCarrierBidSelect(event) {
  const bidId=event.currentTarget.dataset.selectCarrierBid;
  try { const selected=await selectCarrierLoadBid(bidId); await renderLoadDetails(selected.load_id,"Tender"); }
  catch(error) { window.alert(error.message); }
}

function renderCarrierTenderManager(load, tenders = []) {
  const current = tenders[0];
  if (!load.carrier_id) return '<div class="empty-state compact-empty">Assign an outside carrier before creating a tender.</div>';
  return `<form id="carrier-tender-form" class="record-form compact-form">
      <input name="load_id" type="hidden" value="${load.id}"><input name="branch_id" type="hidden" value="${load.branch_id}"><input name="carrier_id" type="hidden" value="${load.carrier_id}">
      <label><span>Carrier</span><input value="${escapeAttribute(load.carriers?.name || "Assigned carrier")}" disabled></label>
      <label><span>Expires</span><input name="expires_at" type="datetime-local" value="${defaultTenderExpiryValue()}" required></label>
      <button type="submit">${current ? "Resend Tender" : "Create Tender"}</button><p id="carrier-tender-message" class="form-message"></p>
    </form>
    <div class="portal-access-directory">${current ? `<article class="portal-access-row"><div><strong>${escapeHtml(current.carriers?.name || load.carriers?.name || "Carrier")}</strong><span>${escapeHtml(formatStatus(current.status))}</span><small>Offered ${formatDateTime(current.offered_at)} · expires ${formatDateTime(current.expires_at)}${current.responded_at ? ` · responded ${formatDateTime(current.responded_at)}` : ""}</small></div></article>` : '<div class="empty-state compact-empty">No tender has been sent.</div>'}</div>`;
}

async function handleCreateCarrierTender(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = form.querySelector("#carrier-tender-message");
  const values = Object.fromEntries(new FormData(form));
  message.textContent = "Sending tender...";
  try {
    await createCarrierTender({
      load_id: values.load_id, branch_id: values.branch_id, carrier_id: values.carrier_id,
      status: "pending", offered_at: new Date().toISOString(), expires_at: new Date(values.expires_at).toISOString(), responded_at: null, response_note: null,
      created_by: currentSession.user.id,
    });
    await renderLoadDetails(values.load_id, "Tender");
  } catch (error) {
    message.dataset.state = "error";
    message.textContent = error.message;
  }
}

function billingReadinessMessage(load, approvedPod, activeInvoice) {
  const opStatus = operationalLoadStatus(load);
  const billStatus = billingLoadStatus(load, activeInvoice);
  if (activeInvoice) return "This load already has an active invoice; duplicate creation is blocked.";
  if (!approvedPod) return "Upload a POD and set its document status to Approved before billing.";
  if (opStatus !== "delivered") return "Complete delivery before starting the billing workflow.";
  if (billStatus === "not_ready") return "Billing prerequisites are complete. The load can move to Ready to Bill.";
  if (billStatus === "ready_to_bill") return "Ready for a draft invoice after positive load revenue is saved.";
  if (["drafted", "invoiced", "partial", "paid"].includes(billStatus)) return `Billing lifecycle is ${formatStatus(billStatus)}.`;
  return "Complete delivery before starting the billing workflow.";
}

async function handleMarkReadyToBill(event) {
  const button = event.currentTarget;
  const message = document.querySelector("#billing-readiness-message");
  button.disabled = true;
  message.textContent = "Checking POD and updating load...";
  try {
    await updateLoadBillingStatus(button.dataset.readyToBill, "ready_to_bill");
    await renderLoadDetails(button.dataset.readyToBill);
  } catch (error) {
    message.textContent = error.message;
    button.disabled = false;
  }
}

async function handleCreateLoadInvoice(event) {
  const button = event.currentTarget;
  const message = document.querySelector("#billing-readiness-message");
  button.disabled = true;
  message.textContent = "Creating draft invoice...";
  try {
    const financials = await getLoadFinancials(button.dataset.createLoadInvoice);
    const lineItems = validateInvoiceLineItems([
      numberOrZero(financials?.linehaul) > 0 ? { item_type: "linehaul", description: `Linehaul - Load ${button.dataset.loadNo}`, customer_description: `Linehaul - Load ${button.dataset.loadNo}`, quantity: 1, unit_amount: numberOrZero(financials.linehaul), approved: true } : null,
      numberOrZero(financials?.fsc) > 0 ? { item_type: "fsc", description: "Fuel surcharge", customer_description: "Fuel surcharge", quantity: 1, unit_amount: numberOrZero(financials.fsc), approved: true } : null,
      numberOrZero(financials?.accessorials) > 0 ? { item_type: "custom", description: "Accessorials", customer_description: "Accessorials", quantity: 1, unit_amount: numberOrZero(financials.accessorials), approved: false, receipt_required: true } : null,
    ].filter(Boolean));
    await createInvoiceWithLineItems({
      invoice_no: null,
      customer_id: button.dataset.customerId,
      load_id: button.dataset.createLoadInvoice,
      due_date: localDateIso(30),
      status: "draft",
    }, lineItems);
    await renderLoadDetails(button.dataset.createLoadInvoice);
  } catch (error) {
    message.textContent = error.message;
    button.disabled = false;
  }
}

function defaultTrackingExpiryValue() {
  return new Date(Date.now() + 7 * 86400000).toISOString().slice(0, 16);
}

function publicTrackingUrl(actionUrl) {
  return `${window.location.origin}${window.location.pathname}${actionUrl}`;
}

function trackingLinkState(link) {
  if (link.revoked_at) return "revoked";
  if (Date.parse(link.expires_at) <= Date.now()) return "expired";
  return "active";
}

function renderPublicTrackingManager(load, links = [], logs = []) {
  const activeLinks = links.filter((link) => trackingLinkState(link) === "active").length;
  return `
    <div class="public-tracking-manager">
      <form id="public-tracking-form" class="record-form compact-form">
        <input name="load_id" type="hidden" value="${load.id}">
        <label><span>Label</span><input name="label" type="text" maxlength="120" placeholder="Customer ETA link"></label>
        <label><span>Expires</span><input name="expires_at" type="datetime-local" value="${defaultTrackingExpiryValue()}" required></label>
        <fieldset class="tracking-privacy-options">
          <legend>Public fields</legend>
          <label><input name="show_status" type="checkbox" checked>Status</label>
          <label><input name="show_eta" type="checkbox" checked>Next stop / ETA</label>
          <label><input name="show_location" type="checkbox">Latest GPS location</label>
          <label><input name="show_driver" type="checkbox">Driver contact</label>
          <label><input name="show_equipment" type="checkbox">Equipment</label>
        </fieldset>
        <button type="submit">Create Tracking Link</button>
        <p id="public-tracking-message" class="form-message">${publicTrackingNotice}</p>
      </form>
      <section class="tracking-link-list">
        <header><strong>${activeLinks} active tracking link${activeLinks === 1 ? "" : "s"}</strong><span>${links.length} total</span></header>
        ${links.length ? links.map(renderTrackingLinkRow).join("") : '<div class="empty-state compact-empty">No public tracking links yet.</div>'}
      </section>
      <section class="tracking-access-list">
        <header><strong>Access Log</strong><span>Latest 20 public opens</span></header>
        ${logs.length ? logs.map((log) => `<article class="tracking-access-row" data-state="${escapeAttribute(log.access_status)}">
          <strong>${escapeHtml(formatStatus(log.access_status))}</strong>
          <span>${formatDateTime(log.accessed_at)}</span>
          <small>${escapeHtml(log.user_agent || "No user agent")}</small>
        </article>`).join("") : '<div class="empty-state compact-empty">No public access yet.</div>'}
      </section>
    </div>
  `;
}

function renderTrackingLinkRow(link) {
  const state = trackingLinkState(link);
  const visibleFields = [
    link.show_status ? "Status" : null,
    link.show_eta ? "ETA" : null,
    link.show_location ? "Location" : null,
    link.show_driver ? "Driver" : null,
    link.show_equipment ? "Equipment" : null,
  ].filter(Boolean).join(", ");
  return `<article class="tracking-link-row" data-state="${state}">
    <div>
      <strong>${escapeHtml(link.label || "Public tracking link")}</strong>
      <small>${escapeHtml(visibleFields || "No fields")} · expires ${formatDateTime(link.expires_at)}</small>
    </div>
    <span class="status-pill">${formatStatus(state)}</span>
    <button type="button" data-revoke-tracking-link="${link.id}" ${state === "active" ? "" : "disabled"}>Revoke</button>
  </article>`;
}

async function handleCreatePublicTrackingLink(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const message = document.querySelector("#public-tracking-message");
  const button = form.querySelector("button[type='submit']");
  button.disabled = true;
  message.textContent = "Creating public tracking link...";
  try {
    const input = validatePublicTrackingLinkInput({
      label: form.elements.label.value,
      expiresAt: form.elements.expires_at.value,
      privacy: {
        show_status: form.elements.show_status.checked,
        show_eta: form.elements.show_eta.checked,
        show_location: form.elements.show_location.checked,
        show_driver: form.elements.show_driver.checked,
        show_equipment: form.elements.show_equipment.checked,
      },
    });
    const link = await createPublicTrackingLink({
      loadId: form.elements.load_id.value,
      label: input.label,
      expiresAt: new Date(input.expiresAt).toISOString(),
      privacy: input.privacy,
    });
    publicTrackingNotice = `Tracking link created: <a href="${escapeAttribute(publicTrackingUrl(link.action_url))}" target="_blank" rel="noopener">${escapeHtml(publicTrackingUrl(link.action_url))}</a>`;
    await renderLoadDetails(form.elements.load_id.value, "Tracking");
  } catch (error) {
    message.textContent = error.message;
    button.disabled = false;
  }
}

async function handleRevokePublicTrackingLink(event) {
  const button = event.currentTarget;
  if (!window.confirm("Revoke this public tracking link?")) return;
  button.disabled = true;
  try {
    await revokePublicTrackingLink(button.dataset.revokeTrackingLink);
    publicTrackingNotice = "Tracking link revoked.";
    const loadId = document.querySelector("#public-tracking-form")?.elements.load_id.value;
    if (loadId) await renderLoadDetails(loadId, "Tracking");
  } catch (error) {
    window.alert(error.message);
    button.disabled = false;
  }
}

function setupLoadDetailTabs(panel, activeTab = "Overview") {
  const layout = panel.querySelector(".detail-layout");
  const summary = layout?.querySelector(".detail-summary");
  const stops = layout?.querySelector(".stops-section");
  const stopForm = layout?.querySelector(".stop-form-wrap");
  const financials = layout?.querySelector(".financials-section");
  const documents = layout?.querySelector(".documents-section");
  const messages = layout?.querySelector(".messages-section");
  const carrierTender = layout?.querySelector(".carrier-tender-section");
  const publicTracking = layout?.querySelector(".public-tracking-section");
  const activity = layout?.querySelector(".activity-section");
  if (!layout || !summary || !stops || !stopForm || !financials || !documents || !messages) return;

  const overviewTab = document.createElement("section");
  overviewTab.className = "load-detail-tab";
  overviewTab.append(summary);
  const stopsTab = document.createElement("section");
  stopsTab.className = "load-detail-tab tab-stack";
  stopsTab.append(stops, stopForm);
  layout.prepend(overviewTab);
  financials.before(stopsTab);
  financials.classList.add("load-detail-tab");
  documents.classList.add("load-detail-tab");
  messages.classList.add("load-detail-tab");
  activity?.classList.add("load-detail-tab");
  const tabs = [overviewTab, stopsTab];
  const labels = ["Overview", "Stops"];
  if (can("manage_finance")) {
    tabs.push(financials);
    labels.push("Financials");
  }
  tabs.push(documents);
  labels.push("Documents");
  tabs.push(messages);
  labels.push("Messages");
  if (carrierTender) {
    carrierTender.classList.add("load-detail-tab");
    tabs.push(carrierTender);
    labels.push("Tender");
  }
  if (publicTracking) {
    publicTracking.classList.add("load-detail-tab");
    tabs.push(publicTracking);
    labels.push("Tracking");
  }
  if (activity) {
    tabs.push(activity);
    labels.push("Activity");
  }
  setupPageTabs(tabs, labels, activeTab);
}

function messageActor(message) {
  if (!message.sender_id) return "System";
  return message.sender_id === currentSession?.user?.id ? "You" : `User ${message.sender_id.slice(0, 8)}`;
}

function renderLoadMessages(messages) {
  if (!messages.length) return '<div class="empty-state compact-empty">No messages yet. Start the load conversation below.</div>';
  return messages.map((message) => `<article class="load-message ${message.sender_id === currentSession?.user?.id ? "is-own" : ""}">
    <header><strong>${escapeHtml(messageActor(message))}</strong><time>${formatDateTime(message.created_at)}</time></header>
    <p>${escapeHtml(message.body)}</p>
  </article>`).join("");
}

function updateLoadMessageCount(event) {
  document.querySelector("#load-message-count").textContent = String(event.currentTarget.value.length);
}

async function handleCreateLoadMessage(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const feedback = document.querySelector("#load-message-feedback");
  const button = form.querySelector("button[type='submit']");
  const loadId = form.elements.load_id.value;
  button.disabled = true;
  feedback.textContent = "Sending message...";
  try {
    const body = validateMessageBody(form.elements.body.value);
    await createLoadMessage(loadId, currentSession.user.id, body);
    const messages = await listLoadMessages(loadId);
    document.querySelector("#load-message-thread").innerHTML = renderLoadMessages(messages);
    form.reset();
    document.querySelector("#load-message-count").textContent = "0";
    feedback.textContent = "Message sent.";
    document.querySelector("#load-message-thread").lastElementChild?.scrollIntoView({ behavior: "smooth", block: "nearest" });
  } catch (error) {
    feedback.textContent = error.message;
  } finally {
    button.disabled = false;
  }
}

const loadActivityLabels = {
  loads: "Load",
  load_stops: "Stop",
  load_financials: "Financials",
  documents: "Document",
  invoices: "Invoice",
};

function auditValue(field, value) {
  if (value === null || value === undefined || value === "") return "Not set";
  if (field.endsWith("_id")) return value ? "Assigned" : "Unassigned";
  if (["linehaul", "fsc", "accessorials", "deductions", "carrier_cost", "driver_pay", "amount"].includes(field)) return formatMoney(value);
  if (field.includes("date") || field.startsWith("appointment_")) return formatDateTime(value);
  return formatStatus(String(value));
}

function renderLoadActivity(activity) {
  const visibleActivity = activity.filter((log) => log.action !== "update" || loadAuditChanges(log).length);
  if (!visibleActivity.length) return '<div class="empty-state compact-empty">No audited activity for this load yet.</div>';
  return visibleActivity.map((log) => {
    const changes = loadAuditChanges(log);
    const actor = log.user_id === currentSession?.user?.id ? "You" : log.user_id ? `User ${log.user_id.slice(0, 8)}` : "System";
    const label = loadActivityLabels[log.entity_type] || formatStatus(log.entity_type);
    return `<article class="load-activity-item">
      <div class="load-activity-marker" data-action="${escapeAttribute(log.action)}"></div>
      <div>
        <header><strong>${escapeHtml(label)} ${formatStatus(log.action)}</strong><time>${formatDateTime(log.created_at)}</time></header>
        <small>${escapeHtml(actor)}</small>
        ${changes.length ? `<ul>${changes.map((change) => `<li><span>${formatStatus(change.field)}</span><del>${escapeHtml(auditValue(change.field, change.oldValue))}</del><strong>→ ${escapeHtml(auditValue(change.field, change.newValue))}</strong></li>`).join("")}</ul>` : ""}
      </div>
    </article>`;
  }).join("");
}

function renderStops(stops) {
  if (!stops.length) {
    return '<div class="empty-state">No stops yet.</div>';
  }

  return stops
    .map(
      (stop, index) => `
        <article class="stop-card">
          <div>
            <strong>${stop.stop_order}. ${formatStatus(stop.type)}</strong>
            <span>${escapeHtml(stop.facility || "Facility not set")}</span>
          </div>
          <p>${escapeHtml(stop.address || "-")}</p>
          <small>Appointment: ${formatDateTime(stop.appointment_from)} → ${formatDateTime(stop.appointment_to)}</small>
          ${(stop.contact_name || stop.contact_phone) ? `<small>Contact: ${escapeHtml([stop.contact_name, stop.contact_phone].filter(Boolean).join(" · "))}</small>` : ""}
          ${(stop.actual_arrival || stop.actual_departure) ? `<small>Actual: ${formatDateTime(stop.actual_arrival)} → ${formatDateTime(stop.actual_departure)}</small>` : ""}
          ${stop.reference ? `<small>Reference: ${escapeHtml(stop.reference)}</small>` : ""}
          ${stop.instructions ? `<p>${escapeHtml(stop.instructions)}</p>` : ""}
          ${can("manage_operations") ? `<div class="stop-actions">
            <button type="button" data-move-stop="up" data-stop-id="${stop.id}" ${index === 0 ? "disabled" : ""}>↑</button>
            <button type="button" data-move-stop="down" data-stop-id="${stop.id}" ${index === stops.length - 1 ? "disabled" : ""}>↓</button>
            <button type="button" data-edit-stop="${stop.id}">Edit</button>
            <button class="danger-button" type="button" data-delete-stop="${stop.id}">Delete</button>
          </div>` : ""}
        </article>
      `
    )
    .join("");
}

function renderFinancialSummary(financials) {
  if (!financials) {
    return '<div class="empty-state compact-empty">No financials yet.</div>';
  }

  return `
    <div class="financial-summary">
      <article><span>Revenue</span><strong>${formatMoney(financials.linehaul + financials.fsc + financials.accessorials)}</strong></article>
      <article><span>Cost</span><strong>${formatMoney(financials.carrier_cost + financials.driver_pay + financials.deductions)}</strong></article>
      <article><span>Margin</span><strong>${formatMoney(financials.margin)}</strong></article>
    </div>
  `;
}

function renderDocuments(documents) {
  if (!documents.length) {
    return '<div class="empty-state compact-empty">No documents yet.</div>';
  }

  return documents
    .map(
      (document) => `
        <article class="document-card">
          <div>
            <strong>${formatStatus(document.doc_type)}</strong>
            <select class="document-status-select" data-document-status="${document.id}" aria-label="Document status">
              ${["pending", "approved", "needs_review", "rejected"].map((status) => `<option value="${status}" ${status === document.status ? "selected" : ""}>${formatStatus(status)}</option>`).join("")}
            </select>
          </div>
          ${
            document.signed_url
              ? `<a href="${escapeAttribute(document.signed_url)}" target="_blank" rel="noreferrer">Open document</a>`
              : "<small>File unavailable</small>"
          }
          <small>${escapeHtml(document.file_name || "Legacy link")}${document.file_size ? ` · ${formatFileSize(document.file_size)}` : ""}</small>
          <button class="delete-document danger-button" type="button" data-document-id="${document.id}">Delete</button>
        </article>
      `
    )
    .join("");
}

function bindDocumentDeleteButtons(refresh) {
  document.querySelectorAll("[data-document-status]").forEach((select) => select.addEventListener("change", async () => {
    select.disabled = true;
    try {
      await updateTmsDocumentStatus(select.dataset.documentStatus, select.value);
      await refresh();
    } catch (error) {
      select.disabled = false;
      window.alert(error.message);
    }
  }));
  document.querySelectorAll(".delete-document").forEach((button) => button.addEventListener("click", async () => {
    if (!window.confirm("Delete this document and its stored file?")) return;
    button.disabled = true;
    button.textContent = "Deleting...";
    try {
      await deleteTmsDocument(button.dataset.documentId);
      await refresh();
    } catch (error) {
      button.disabled = false;
      button.textContent = "Delete";
      window.alert(error.message);
    }
  }));
}

function populateStopFromFacility(select, facilities) {
  const facility = facilities.find((item) => item.id === select.value);
  if (!facility) return;
  const form = select.form;
  form.elements.facility.value = facility.name;
  form.elements.address.value = [facility.address_line1, facility.address_line2, `${facility.city}, ${facility.state} ${facility.postal_code || ""}`.trim(), facility.country !== "US" ? facility.country : null].filter(Boolean).join(", ");
  form.elements.contact_name.value = facility.contact_name || "";
  form.elements.contact_phone.value = facility.contact_phone || "";
  if (!form.elements.instructions.value) form.elements.instructions.value = facility.instructions || "";
}

async function handleCreateStop(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const message = document.querySelector("#stop-message");
  const formData = new FormData(form);
  const loadId = formData.get("load_id");
  const stopId = formData.get("id");

  message.textContent = stopId ? "Updating stop..." : "Saving stop...";

  try {
    const appointmentFrom = formData.get("appointment_from");
    const appointmentTo = formData.get("appointment_to");
    const actualArrival = formData.get("actual_arrival");
    const actualDeparture = formData.get("actual_departure");
    validateStopWindows({ appointmentFrom, appointmentTo, actualArrival, actualDeparture });
    const stop = {
      type: formData.get("type"),
      facility_id: formData.get("facility_id") || null,
      facility: formData.get("facility") || null,
      address: formData.get("address") || null,
      contact_name: formData.get("contact_name") || null,
      contact_phone: formData.get("contact_phone") || null,
      appointment_from: dateTimeOrNull(appointmentFrom),
      appointment_to: dateTimeOrNull(appointmentTo),
      actual_arrival: dateTimeOrNull(actualArrival),
      actual_departure: dateTimeOrNull(actualDeparture),
      reference: formData.get("reference") || null,
      instructions: formData.get("instructions") || null,
    };

    if (stopId) {
      await updateLoadStop(stopId, stop);
    } else {
      await createLoadStop({ ...stop, load_id: loadId, stop_order: numberOrNull(formData.get("stop_order")) || 1 });
    }

    message.textContent = stopId ? "Stop updated." : "Stop saved.";
    await renderLoadDetails(loadId, "Stops");
  } catch (error) {
    message.textContent = error.message;
  }
}

function bindStopActions(stops, loadId) {
  document.querySelectorAll("[data-edit-stop]").forEach((button) => button.addEventListener("click", () => {
    const stop = stops.find((item) => item.id === button.dataset.editStop);
    if (!stop) return;
    const form = document.querySelector("#stop-form");
    form.elements.id.value = stop.id;
    form.elements.stop_order.value = stop.stop_order;
    form.elements.type.value = stop.type;
    form.elements.facility_id.value = stop.facility_id || "";
    form.elements.facility.value = stop.facility || "";
    form.elements.address.value = stop.address || "";
    form.elements.contact_name.value = stop.contact_name || "";
    form.elements.contact_phone.value = stop.contact_phone || "";
    form.elements.appointment_from.value = dateTimeInputValue(stop.appointment_from);
    form.elements.appointment_to.value = dateTimeInputValue(stop.appointment_to);
    form.elements.actual_arrival.value = dateTimeInputValue(stop.actual_arrival);
    form.elements.actual_departure.value = dateTimeInputValue(stop.actual_departure);
    form.elements.reference.value = stop.reference || "";
    form.elements.instructions.value = stop.instructions || "";
    document.querySelector("#stop-form-title").textContent = `Edit Stop ${stop.stop_order}`;
    document.querySelector("#stop-submit").textContent = "Update Stop";
    document.querySelector("#cancel-stop-edit").hidden = false;
    form.scrollIntoView({ behavior: "smooth", block: "start" });
  }));

  document.querySelectorAll("[data-move-stop]").forEach((button) => button.addEventListener("click", async () => {
    button.disabled = true;
    try {
      await moveLoadStop(button.dataset.stopId, button.dataset.moveStop);
      await renderLoadDetails(loadId, "Stops");
    } catch (error) {
      window.alert(error.message);
      button.disabled = false;
    }
  }));

  document.querySelectorAll("[data-delete-stop]").forEach((button) => button.addEventListener("click", async () => {
    if (!window.confirm("Delete this stop? Remaining stops will be renumbered.")) return;
    button.disabled = true;
    try {
      await deleteLoadStop(button.dataset.deleteStop);
      await renderLoadDetails(loadId, "Stops");
    } catch (error) {
      window.alert(error.message);
      button.disabled = false;
    }
  }));
}

function resetStopForm(nextOrder) {
  const form = document.querySelector("#stop-form");
  form.reset();
  form.elements.id.value = "";
  form.elements.stop_order.value = nextOrder;
  document.querySelector("#stop-form-title").textContent = "Add Stop";
  document.querySelector("#stop-submit").textContent = "Save Stop";
  document.querySelector("#cancel-stop-edit").hidden = true;
  document.querySelector("#stop-message").textContent = "";
}

async function handleSaveFinancials(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const message = document.querySelector("#financials-message");
  const formData = new FormData(form);
  const loadId = formData.get("load_id");

  message.textContent = "Saving financials...";

  try {
    await saveLoadFinancials({
      load_id: loadId,
      linehaul: numberOrZero(formData.get("linehaul")),
      fsc: numberOrZero(formData.get("fsc")),
      accessorials: numberOrZero(formData.get("accessorials")),
      deductions: numberOrZero(formData.get("deductions")),
      carrier_cost: numberOrZero(formData.get("carrier_cost")),
      driver_pay: numberOrZero(formData.get("driver_pay")),
    });

    message.textContent = "Financials saved.";
    await renderLoadDetails(loadId, "Financials");
  } catch (error) {
    message.textContent = error.message;
  }
}

async function handleCreateDocument(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const message = document.querySelector("#document-message");
  const formData = new FormData(form);
  const loadId = formData.get("entity_id");

  message.textContent = "Uploading document...";

  try {
    await uploadTmsDocument({
      entityType: "load",
      entityId: loadId,
      docType: formData.get("doc_type"),
      status: formData.get("status"),
      file: formData.get("file"),
      uploadedBy: currentSession?.user?.id || null,
    });

    message.textContent = "Document uploaded securely.";
    await renderLoadDetails(loadId, "Documents");
  } catch (error) {
    message.textContent = error.message;
  }
}

async function handleCreateLoad(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const message = document.querySelector("#load-message");
  const formData = new FormData(form);
  const loadId = formData.get("id");

  message.textContent = loadId ? "Updating load..." : "Saving load...";

  try {
    const payload = validateLoadOperationalDetails({
      load_no: formData.get("load_no"),
      reference_no: formData.get("reference_no") || null,
      broker_reference: formData.get("broker_reference") || null,
      customer_id: formData.get("customer_id"),
      carrier_id: formData.get("carrier_id") || null,
      driver_id: formData.get("driver_id") || null,
      truck_id: formData.get("truck_id") || null,
      trailer_id: formData.get("trailer_id") || null,
      status: formData.get("status"),
      origin: formData.get("origin") || null,
      destination: formData.get("destination") || null,
      commodity: formData.get("commodity") || null,
      weight_lbs: numberOrNull(formData.get("weight_lbs")),
      pieces: numberOrNull(formData.get("pieces")),
      pallets: numberOrNull(formData.get("pallets")),
      equipment_type: formData.get("equipment_type") || null,
      temperature_min_f: numberOrNull(formData.get("temperature_min_f")),
      temperature_max_f: numberOrNull(formData.get("temperature_max_f")),
      hazmat: formData.get("hazmat") === "on",
      loaded_miles: numberOrNull(formData.get("loaded_miles")),
      empty_miles: numberOrNull(formData.get("empty_miles")),
      pickup_date: formData.get("pickup_date") || null,
      delivery_date: formData.get("delivery_date") || null,
      internal_notes: formData.get("internal_notes") || null,
      customer_notes: formData.get("customer_notes") || null,
      notes: formData.get("internal_notes") || null,
    });

    if (loadId) {
      await updateLoad(loadId, payload);
    } else {
      await createLoad({
        ...payload,
        dispatcher_id: currentSession?.user?.id || null,
      });
    }

    resetLoadForm();
    message.textContent = loadId ? "Load updated." : "Load saved.";
    await Promise.all([loadLoadFormOptions(), loadLoads()]);
  } catch (error) {
    message.textContent = error.message;
  }
}

function formatStatus(value) {
  return String(value || "-")
    .split("_")
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ");
}

function formatDate(value) {
  if (!value) {
    return "-";
  }

  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(value))
    ? new Date(`${value}T00:00:00`)
    : new Date(value);

  if (Number.isNaN(date.getTime())) {
    return "-";
  }

  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "2-digit",
    year: "numeric",
  }).format(date);
}

function formatDateTime(value) {
  if (!value) {
    return "-";
  }

  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "2-digit",
    hour: "numeric",
    minute: "2-digit",
  }).format(new Date(value));
}

function dateTimeInputValue(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  const local = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 16);
}

function formatNumber(value) {
  if (value === null || value === undefined || value === "") {
    return "-";
  }

  return Number(value).toLocaleString("en-US");
}

function formatFileSize(bytes) {
  const value = Number(bytes) || 0;
  if (value < 1024) return `${value} B`;
  if (value < 1024 * 1024) return `${(value / 1024).toFixed(1)} KB`;
  return `${(value / (1024 * 1024)).toFixed(1)} MB`;
}

function formatMoney(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
  }).format(Number(value || 0));
}

function formatPercent(value) {
  return `${Number(value || 0).toLocaleString("en-US", { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`;
}

function countBy(items, key) {
  return items.reduce((counts, item) => {
    const value = item[key] || "unknown";
    counts[value] = (counts[value] || 0) + 1;
    return counts;
  }, {});
}

function sumBy(items, key) {
  return items.reduce((total, item) => total + Number(item[key] || 0), 0);
}

function numberOrNull(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }

  return Number(value);
}

function numberOrZero(value) {
  if (value === null || value === undefined || value === "") {
    return 0;
  }

  return Number(value);
}

function dateTimeOrNull(value) {
  if (!value) {
    return null;
  }

  return new Date(value).toISOString();
}

function routeLabel(load) {
  const origin = load.origin || "-";
  const destination = load.destination || "-";

  return `${origin} → ${destination}`;
}

function orderedLoadStops(load) {
  return [...(load.load_stops || [])].sort((a, b) => a.stop_order - b.stop_order);
}

function loadPickupWindow(load) {
  const pickup = orderedLoadStops(load).find((stop) => stop.type === "pickup") || orderedLoadStops(load)[0];
  if (!pickup?.appointment_from) return formatDate(load.pickup_date);
  return `<span>${formatDateTime(pickup.appointment_from)}</span>${pickup.appointment_to ? `<small> to ${formatDateTime(pickup.appointment_to)}</small>` : ""}`;
}

function formatStopWindow(stop, fallbackDate) {
  if (!stop?.appointment_from) return formatDate(fallbackDate);
  return `${formatDateTime(stop.appointment_from)}${stop.appointment_to ? ` → ${formatDateTime(stop.appointment_to)}` : ""}`;
}

function formatTemperature(load) {
  if (load.temperature_min_f == null && load.temperature_max_f == null) return "-";
  if (load.temperature_min_f == null) return `Up to ${formatNumber(load.temperature_max_f)} °F`;
  if (load.temperature_max_f == null) return `${formatNumber(load.temperature_min_f)} °F minimum`;
  return `${formatNumber(load.temperature_min_f)}–${formatNumber(load.temperature_max_f)} °F`;
}

function renderRows(results) {
  return results
    .map((result) => {
      const state = result.ok ? "ok" : "error";
      const count = result.ok ? result.count : "-";

      return `
        <div class="table-row" data-state="${state}">
          <span>${result.table}</span>
          <strong>${count}</strong>
          <small>${escapeHtml(result.message)}</small>
        </div>
      `;
    })
    .join("");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttribute(value) {
  return escapeHtml(value).replaceAll("`", "&#096;");
}

function closeGlobalSearchResults() {
  globalSearchResults.hidden = true;
  globalSearchResults.innerHTML = "";
  globalSearch.setAttribute("aria-expanded", "false");
  globalSearchActiveIndex = -1;
}

function scheduleGlobalSearch() {
  clearTimeout(globalSearchTimer);
  const query = globalSearch.value.trim();
  if (query.length < 2) {
    closeGlobalSearchResults();
    return;
  }
  globalSearchTimer = setTimeout(() => runGlobalSearch(query), 220);
}

async function safeGlobalSearchLoad(enabled, loader) {
  if (!enabled) return [];
  try { return await loader(); } catch { return []; }
}

async function loadGlobalSearchRecords() {
  if (globalSearchCache.records.length && Date.now() - globalSearchCache.loadedAt < 60000) return globalSearchCache.records;
  const [loads, customers, drivers, carriers, trucks, trailers, invoices] = await Promise.all([
    safeGlobalSearchLoad(can("view_loads"), listLoads),
    safeGlobalSearchLoad(can("view_customers"), listCustomers),
    safeGlobalSearchLoad(can("view_drivers"), listDrivers),
    safeGlobalSearchLoad(can("view_carriers"), listCarriers),
    safeGlobalSearchLoad(can("view_equipment"), listTrucks),
    safeGlobalSearchLoad(can("view_equipment"), listTrailers),
    safeGlobalSearchLoad(can("view_billing"), listInvoices),
  ]);
  const records = [
    ...loads.map((item) => ({ type: "Load", title: item.load_no, detail: `${item.customers?.name || "No customer"} · ${routeLabel(item)}`, keywords: [item.reference_no,item.broker_reference,item.commodity,item.drivers?.name,item.trucks?.unit_no,item.status,item.operational_status,item.billing_status].filter(Boolean).join(" "), href: `#loads/${item.id}` })),
    ...customers.map((item) => ({ type: "Customer", title: item.name, detail: item.billing_email || item.billing_address || "Customer profile", keywords: [item.email,item.phone,item.mc_number].filter(Boolean).join(" "), href: "#customers" })),
    ...drivers.map((item) => ({ type: "Driver", title: item.name, detail: item.status ? `Driver · ${formatStatus(item.status)}` : "Driver profile", keywords: [item.email,item.phone,item.license_number].filter(Boolean).join(" "), href: "#drivers" })),
    ...carriers.map((item) => ({ type: "Carrier", title: item.name, detail: item.mc_number ? `MC ${item.mc_number}` : "Carrier profile", keywords: [item.dot_number,item.email,item.phone].filter(Boolean).join(" "), href: "#carriers" })),
    ...trucks.map((item) => ({ type: "Truck", title: item.unit_no, detail: item.status ? `Truck · ${formatStatus(item.status)}` : "Truck profile", keywords: [item.vin,item.plate].filter(Boolean).join(" "), href: "#equipment" })),
    ...trailers.map((item) => ({ type: "Trailer", title: item.unit_no, detail: item.status ? `Trailer · ${formatStatus(item.status)}` : "Trailer profile", keywords: [item.vin,item.plate].filter(Boolean).join(" "), href: "#equipment" })),
    ...invoices.map((item) => ({ type: "Invoice", title: item.invoice_no, detail: `${item.customers?.name || "Customer"} · ${formatMoney(item.amount)} · ${formatStatus(item.status)}`, keywords: [item.loads?.load_no,item.status].filter(Boolean).join(" "), href: "#billing" })),
  ];
  globalSearchCache = { records, loadedAt: Date.now() };
  return records;
}

async function runGlobalSearch(query) {
  globalSearchResults.hidden = false;
  globalSearchResults.innerHTML = '<div class="global-search-state">Searching...</div>';
  globalSearch.setAttribute("aria-expanded", "true");
  const records = await loadGlobalSearchRecords();
  if (globalSearch.value.trim() !== query) return;
  const results = rankGlobalSearchRecords(records, query, 10);
  globalSearchActiveIndex = results.length ? 0 : -1;
  globalSearchResults.innerHTML = results.length
    ? results.map((item, index) => `<a href="${escapeAttribute(item.href)}" role="option" data-global-search-result data-index="${index}" aria-selected="${index === 0}"><span>${escapeHtml(item.type)}</span><strong>${escapeHtml(item.title)}</strong><small>${escapeHtml(item.detail)}</small></a>`).join("")
    : '<div class="global-search-state">No matching records.</div>';
  globalSearchResults.querySelectorAll("[data-global-search-result]").forEach((result) => result.addEventListener("click", closeGlobalSearchResults));
}

function updateGlobalSearchActiveResult() {
  const results = Array.from(globalSearchResults.querySelectorAll("[data-global-search-result]"));
  results.forEach((result, index) => result.setAttribute("aria-selected", String(index === globalSearchActiveIndex)));
  results[globalSearchActiveIndex]?.scrollIntoView({ block: "nearest" });
}

function handleGlobalSearchKeydown(event) {
  const results = Array.from(globalSearchResults.querySelectorAll("[data-global-search-result]"));
  if (event.key === "Escape") {
    closeGlobalSearchResults();
    return;
  }
  if (["ArrowDown", "ArrowUp"].includes(event.key) && results.length) {
    event.preventDefault();
    const direction = event.key === "ArrowDown" ? 1 : -1;
    globalSearchActiveIndex = (globalSearchActiveIndex + direction + results.length) % results.length;
    updateGlobalSearchActiveResult();
    return;
  }
  if (event.key !== "Enter") return;
  event.preventDefault();
  if (results[globalSearchActiveIndex]) {
    results[globalSearchActiveIndex].click();
    return;
  }
  loadListState.search = globalSearch.value;
  loadListState.page = 1;
  window.location.hash = "loads";
}

loginForm.addEventListener("submit", handleLogin);
signupButton.addEventListener("click", handleSignup);
logoutButton.addEventListener("click", async () => {
  await signOut();
});
refreshButton.addEventListener("click", () => {
  globalSearchCache = { records: [], loadedAt: 0 };
  renderRoute();
});
addOrderButton.addEventListener("click", () => {
  window.location.hash = "loads";
});
globalSearch.addEventListener("input", scheduleGlobalSearch);
globalSearch.addEventListener("focus", () => {
  if (globalSearch.value.trim().length >= 2) scheduleGlobalSearch();
});
globalSearch.addEventListener("keydown", handleGlobalSearchKeydown);
document.addEventListener("click", (event) => {
  if (!event.target.closest(".search-box")) closeGlobalSearchResults();
});
window.addEventListener("hashchange", renderRoute);

onAuthStateChange(setAuthView);

getSession()
  .then(setAuthView)
  .catch(() => setAuthView(null));
