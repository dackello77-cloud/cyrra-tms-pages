import {
  checkExistingTables,
  checkTable,
  checkTmsTables,
  createCarrier,
  createCustomer,
  createDocument,
  createDriver,
  createLoad,
  createLoadDocument,
  createLoadStop,
  createInvoice,
  createTrailer,
  createTruck,
  getLoadFinancials,
  getLoad,
  getSession,
  listCarriers,
  listCustomers,
  listDocuments,
  listDrivers,
  listLoadDocuments,
  listLoads,
  listLoadStops,
  listInvoices,
  listTrailers,
  listTrucks,
  listTableRows,
  onAuthStateChange,
  saveLoadFinancials,
  signIn,
  signOut,
  signUp,
  updateInvoiceStatus,
  updateCarrier,
  updateCustomer,
  updateDriver,
  updateLoad,
  updateLoadStatus,
  updateTrailer,
  updateTruck,
} from "./api.js";

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
const mapboxToken = "";

let currentSession = null;
let fleetMapInstance = null;

function route() {
  return window.location.hash.replace("#", "") || "dashboard";
}

function setAuthView(session) {
  currentSession = session;
  authScreen.hidden = Boolean(session);
  appShell.hidden = !session;

  if (session) {
    renderRoute();
  }
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
    link.classList.toggle("active", link.hash === `#${activeRoute}`);
  });
}

function renderRoute() {
  if (!currentSession) {
    return;
  }

  const activeRoute = route();
  updateActiveNav(activeRoute);

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

  if (activeRoute === "tracking") {
    renderFleetMap();
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

  if (activeRoute === "dispatch") {
    renderDispatchBoard();
    return;
  }

  if (activeRoute === "billing") {
    renderBilling();
    return;
  }

  if (activeRoute === "reports") {
    renderReports();
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
            <a class="panel-link" href="#dispatch">Open Dispatch</a>
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
        <a class="panel metric-card metric-link" href="#tracking">
          <div><h3>Fleet availability</h3><small>Drivers / trucks / trailers</small></div>
          <strong id="dash-fleet-available">0</strong>
          <div id="dash-fleet-breakdown" class="mini-list"></div>
        </a>
        <a class="panel metric-card metric-link" href="#billing">
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

      <article class="panel">
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
          <a class="panel-link" href="#billing">Billing</a>
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
    const [results, tmsResults, drivers, trucks, trailers, loads, documents, invoices] = await Promise.all([
      checkExistingTables(),
      checkTmsTables(),
      safeDashboardList(listDrivers),
      safeDashboardList(listTrucks),
      safeDashboardList(listTrailers),
      safeDashboardList(listLoads),
      safeDashboardList(listDocuments),
      safeDashboardList(listInvoices),
    ]);
    const hasConnection = results.some((result) => result.ok);
    const tmsReady = tmsResults.every((result) => result.ok);
    const setupPercent = Math.round((results.filter((result) => result.ok).length / results.length) * 100);
    const tmsPercent = Math.round((tmsResults.filter((result) => result.ok).length / tmsResults.length) * 100);
    const activeLoads = loads.filter((load) => !["delivered", "billed", "cancelled"].includes(load.status));
    const unassignedLoads = loads.filter((load) => !load.driver_id || !load.truck_id || !load.trailer_id);
    const today = localDateIso();
    const todayLoads = loads.filter((load) => load.pickup_date === today || load.delivery_date === today);
    const docIssues = documents.filter((document) => ["pending", "needs_review", "rejected"].includes(document.status));
    const statusCounts = countBy(loads, "status");
    const availableDrivers = drivers.filter((driver) => driver.status === "available").length;
    const availableTrucks = trucks.filter((truck) => truck.status === "available").length;
    const availableTrailers = trailers.filter((trailer) => trailer.status === "available").length;
    const unpaidInvoices = invoices.filter((invoice) => !["paid", "void"].includes(invoice.status));
    const overdueInvoices = invoices.filter((invoice) => invoice.status === "overdue");
    const readyToBill = loads.filter((load) => load.status === "ready_to_bill").length;

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
    alertsEl.innerHTML = renderDashboardAlerts({ unassignedLoads, todayLoads, readyToBill, docIssues, overdueInvoices });
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

function renderDashboardAlerts({ unassignedLoads, todayLoads, readyToBill, docIssues, overdueInvoices }) {
  const alerts = [
    { label: "Loads missing assignment", count: unassignedLoads.length, href: "#loads", tone: unassignedLoads.length ? "warn" : "ok" },
    { label: "Pickup or delivery today", count: todayLoads.length, href: "#loads", tone: todayLoads.length ? "info" : "ok" },
    { label: "Ready to bill", count: readyToBill, href: "#billing", tone: readyToBill ? "info" : "ok" },
    { label: "Documents needing attention", count: docIssues.length, href: "#documents", tone: docIssues.length ? "warn" : "ok" },
    { label: "Overdue invoices", count: overdueInvoices.length, href: "#billing", tone: overdueInvoices.length ? "danger" : "ok" },
  ];

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
      status: formatStatus(load.status),
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
    ...invoices.slice(0, 4).map((invoice) => ({
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

function renderCustomers() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Phase 1", title: "Customers" })}
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
  `;

  document.querySelector("#reload-customers").addEventListener("click", loadCustomers);
  document.querySelector("#customer-form").addEventListener("submit", handleCreateCustomer);
  document.querySelector("#cancel-customer-edit").addEventListener("click", resetCustomerForm);
  loadCustomers();
}

async function loadCustomers() {
  const table = document.querySelector("#customers-table");
  table.innerHTML = '<div class="empty-state">Loading customers...</div>';

  try {
    const customers = await listCustomers();

    if (!customers.length) {
      table.innerHTML = '<div class="empty-state">No customers yet.</div>';
      return;
    }

    table.innerHTML = `
      <div class="data-row customer-row data-row-head">
        <span>Name</span>
        <span>Credit</span>
        <span>Portal</span>
        <span>Status</span>
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
              <button class="row-action" type="button" data-edit-customer="${customer.id}">Edit</button>
            </div>
          `
        )
        .join("")}
    `;

    table.querySelectorAll("[data-edit-customer]").forEach((button) => {
      const customer = customers.find((item) => item.id === button.dataset.editCustomer);
      button.addEventListener("click", () => fillCustomerForm(customer));
    });
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
  form.elements.credit_status.value = customer.credit_status || "ok";
  form.elements.portal_enabled.checked = Boolean(customer.portal_enabled);
  document.querySelector("#customer-form-title").textContent = "Edit Customer";
  document.querySelector("#customer-submit").textContent = "Update Customer";
  document.querySelector("#cancel-customer-edit").hidden = false;
  document.querySelector("#customer-message").textContent = "";
}

function resetCustomerForm() {
  const form = document.querySelector("#customer-form");
  form.reset();
  form.elements.id.value = "";
  document.querySelector("#customer-form-title").textContent = "Add Customer";
  document.querySelector("#customer-submit").textContent = "Save Customer";
  document.querySelector("#cancel-customer-edit").hidden = true;
  document.querySelector("#customer-message").textContent = "";
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
      credit_status: formData.get("credit_status"),
      portal_enabled: formData.get("portal_enabled") === "on",
    };

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

function renderCarriers() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Phase 1", title: "Carriers" })}
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
              <button class="row-action" type="button" data-edit-carrier="${carrier.id}">Edit</button>
            </div>
          `
        )
        .join("")}
    `;

    table.querySelectorAll("[data-edit-carrier]").forEach((button) => {
      const carrier = carriers.find((item) => item.id === button.dataset.editCarrier);
      button.addEventListener("click", () => fillCarrierForm(carrier));
    });
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
  form.elements.status.value = carrier.status || "active";
  document.querySelector("#carrier-form-title").textContent = "Edit Carrier";
  document.querySelector("#carrier-submit").textContent = "Update Carrier";
  document.querySelector("#cancel-carrier-edit").hidden = false;
  document.querySelector("#carrier-message").textContent = "";
}

function resetCarrierForm() {
  const form = document.querySelector("#carrier-form");
  form.reset();
  form.elements.id.value = "";
  document.querySelector("#carrier-form-title").textContent = "Add Carrier";
  document.querySelector("#carrier-submit").textContent = "Save Carrier";
  document.querySelector("#cancel-carrier-edit").hidden = true;
  document.querySelector("#carrier-message").textContent = "";
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
      status: formData.get("status"),
    };

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
          <label>
            <span>CDL number</span>
            <input name="cdl_no" type="text" />
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
        <span>Medical</span>
        <span></span>
      </div>
      ${drivers
        .map(
          (driver) => `
            <div class="data-row driver-row">
              <strong>${escapeHtml(driver.name)}</strong>
              <span>${formatStatus(driver.status)}</span>
              <span>${escapeHtml(driver.phone || "-")}</span>
              <span>${formatDate(driver.medical_expiry)}</span>
              <button class="row-action" type="button" data-edit-driver="${driver.id}">Edit</button>
            </div>
          `
        )
        .join("")}
    `;

    table.querySelectorAll("[data-edit-driver]").forEach((button) => {
      const driver = drivers.find((item) => item.id === button.dataset.editDriver);
      button.addEventListener("click", () => fillDriverForm(driver));
    });
  } catch (error) {
    table.innerHTML = `
      <div class="empty-state">
        Drivers table is not ready: ${escapeHtml(error.message)}
      </div>
    `;
  }
}

function fillDriverForm(driver) {
  const form = document.querySelector("#driver-form");
  form.elements.id.value = driver.id;
  form.elements.name.value = driver.name || "";
  form.elements.phone.value = driver.phone || "";
  form.elements.email.value = driver.email || "";
  form.elements.cdl_no.value = driver.cdl_no || "";
  form.elements.medical_expiry.value = driver.medical_expiry || "";
  form.elements.status.value = driver.status || "available";
  document.querySelector("#driver-form-title").textContent = "Edit Driver";
  document.querySelector("#driver-submit").textContent = "Update Driver";
  document.querySelector("#cancel-driver-edit").hidden = false;
  document.querySelector("#driver-message").textContent = "";
}

function resetDriverForm() {
  const form = document.querySelector("#driver-form");
  form.reset();
  form.elements.id.value = "";
  document.querySelector("#driver-form-title").textContent = "Add Driver";
  document.querySelector("#driver-submit").textContent = "Save Driver";
  document.querySelector("#cancel-driver-edit").hidden = true;
  document.querySelector("#driver-message").textContent = "";
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
      cdl_no: formData.get("cdl_no") || null,
      medical_expiry: formData.get("medical_expiry") || null,
      status: formData.get("status"),
    };

    if (driverId) {
      await updateDriver(driverId, payload);
    } else {
      await createDriver(payload);
    }

    resetDriverForm();
    message.textContent = driverId ? "Driver updated." : "Driver saved.";
    await loadDrivers();
  } catch (error) {
    message.textContent = error.message;
  }
}

function renderEquipment() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Phase 1", title: "Equipment" })}
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
              <button class="row-action" type="button" data-edit-truck="${truck.id}">Edit</button>
            </div>
          `
        )
        .join("")}
    `;

    table.querySelectorAll("[data-edit-truck]").forEach((button) => {
      const truck = trucks.find((item) => item.id === button.dataset.editTruck);
      button.addEventListener("click", () => fillTruckForm(truck));
    });
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
  form.elements.status.value = truck.status || "available";
  document.querySelector("#truck-form-title").textContent = "Edit Truck";
  document.querySelector("#truck-submit").textContent = "Update Truck";
  document.querySelector("#cancel-truck-edit").hidden = false;
  document.querySelector("#truck-message").textContent = "";
}

function resetTruckForm() {
  const form = document.querySelector("#truck-form");
  form.reset();
  form.elements.id.value = "";
  document.querySelector("#truck-form-title").textContent = "Add Truck";
  document.querySelector("#truck-submit").textContent = "Save Truck";
  document.querySelector("#cancel-truck-edit").hidden = true;
  document.querySelector("#truck-message").textContent = "";
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
              <button class="row-action" type="button" data-edit-trailer="${trailer.id}">Edit</button>
            </div>
          `
        )
        .join("")}
    `;

    table.querySelectorAll("[data-edit-trailer]").forEach((button) => {
      const trailer = trailers.find((item) => item.id === button.dataset.editTrailer);
      button.addEventListener("click", () => fillTrailerForm(trailer));
    });
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
  form.elements.status.value = trailer.status || "available";
  document.querySelector("#trailer-form-title").textContent = "Edit Trailer";
  document.querySelector("#trailer-submit").textContent = "Update Trailer";
  document.querySelector("#cancel-trailer-edit").hidden = false;
  document.querySelector("#trailer-message").textContent = "";
}

function resetTrailerForm() {
  const form = document.querySelector("#trailer-form");
  form.reset();
  form.elements.id.value = "";
  document.querySelector("#trailer-form-title").textContent = "Add Trailer";
  document.querySelector("#trailer-submit").textContent = "Save Trailer";
  document.querySelector("#cancel-trailer-edit").hidden = true;
  document.querySelector("#trailer-message").textContent = "";
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

function renderFleetMap() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Operations", title: "Fleet Map" })}
    <section id="fleet-map-root" class="fleet-map-root">
      <div class="empty-state">Loading fleet map...</div>
    </section>
  `;

  loadFleetMap();
}

async function loadFleetMap() {
  const root = document.querySelector("#fleet-map-root");

  try {
    const [loads, drivers, trucks, trailers] = await Promise.all([
      listLoads(),
      listDrivers(),
      listTrucks(),
      listTrailers(),
    ]);
    const activeLoads = loads.filter((load) => !["delivered", "billed", "cancelled"].includes(load.status));
    const availableTrucks = trucks.filter((truck) => truck.status === "available").length;
    const availableTrailers = trailers.filter((trailer) => trailer.status === "available").length;
    const availableDrivers = drivers.filter((driver) => driver.status === "available").length;

    root.innerHTML = `
      <section class="fleet-kpis">
        ${renderReportKpi("Active loads", activeLoads.length)}
        ${renderReportKpi("Available drivers", availableDrivers)}
        ${renderReportKpi("Available trucks", availableTrucks)}
        ${renderReportKpi("Available trailers", availableTrailers)}
      </section>
      <section class="fleet-map-panel panel">
        <div class="panel-header">
          <div>
            <h2>Mapbox Fleet Map</h2>
            <p>Origin, destination and route view from active load locations.</p>
          </div>
          <button id="reload-fleet-map" type="button">Refresh</button>
        </div>
        <div class="mapbox-shell">
          <div id="fleet-map" class="mapbox-map" aria-label="Fleet map"></div>
          <div id="fleet-map-message" class="mapbox-message">Preparing Mapbox...</div>
        </div>
        <div class="fleet-map-canvas route-list">
          ${activeLoads.length ? activeLoads.slice(0, 10).map(renderFleetRoute).join("") : '<div class="empty-state compact-empty">No active loads to map.</div>'}
        </div>
      </section>
      <section class="fleet-columns">
        ${renderFleetAssetColumn("Drivers", drivers, "name")}
        ${renderFleetAssetColumn("Trucks", trucks, "unit_no")}
        ${renderFleetAssetColumn("Trailers", trailers, "unit_no")}
      </section>
    `;

    document.querySelector("#reload-fleet-map").addEventListener("click", loadFleetMap);
    renderMapboxFleetMap(activeLoads);
  } catch (error) {
    root.innerHTML = `<div class="empty-state">Fleet map is not ready: ${escapeHtml(error.message)}</div>`;
  }
}

function renderFleetRoute(load) {
  return `
    <article class="fleet-route-card">
      <div>
        <strong>${escapeHtml(load.load_no)}</strong>
        <span>${formatStatus(load.status)}</span>
      </div>
      <p>${escapeHtml(routeLabel(load))}</p>
      <small>
        ${escapeHtml(load.drivers?.name || "No driver")} ·
        ${escapeHtml(load.trucks?.unit_no || "No truck")} ·
        ${escapeHtml(load.trailers?.unit_no || "No trailer")}
      </small>
    </article>
  `;
}

async function renderMapboxFleetMap(activeLoads) {
  const mapEl = document.querySelector("#fleet-map");
  const messageEl = document.querySelector("#fleet-map-message");

  if (!mapEl || !messageEl) {
    return;
  }

  if (!mapboxToken) {
    messageEl.textContent = "Map is not configured for this public deployment.";
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
  messageEl.textContent = activeLoads.length ? "Geocoding active load locations..." : "No active loads to show on map.";

  fleetMapInstance.on("load", async () => {
    const mappedRoutes = await buildMappedRoutes(activeLoads.slice(0, 12));

    if (!mappedRoutes.length) {
      messageEl.textContent = "Add origin and destination on loads to draw routes.";
      return;
    }

    const bounds = new window.mapboxgl.LngLatBounds();
    const routeFeatures = [];

    mappedRoutes.forEach((route, index) => {
      const color = mapRouteColor(index);
      bounds.extend(route.origin.coordinates);
      bounds.extend(route.destination.coordinates);
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
    fleetMapInstance.fitBounds(bounds, { padding: 72, maxZoom: 8 });
    messageEl.textContent = `${mappedRoutes.length} active route${mappedRoutes.length === 1 ? "" : "s"} mapped`;
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

function renderFleetAssetColumn(title, rows, labelKey) {
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
                .map(
                  (row) => `
                    <div>
                      <strong>${escapeHtml(row[labelKey] || "-")}</strong>
                      <span>${formatStatus(row.status)}</span>
                    </div>
                  `
                )
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
            <p>Creates document metadata using an existing file URL.</p>
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
            <span>Entity ID</span>
            <input name="entity_id" type="text" placeholder="UUID from related record" required />
          </label>
          <label>
            <span>Document type</span>
            <input name="doc_type" type="text" placeholder="rate_confirmation, bol, insurance..." required />
          </label>
          <label>
            <span>File URL</span>
            <input name="file_url" type="url" />
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
  loadDocumentCenter();
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
              ${
                document.file_url
                  ? `<a class="row-action" href="${escapeAttribute(document.file_url)}" target="_blank" rel="noreferrer">Open</a>`
                  : "<span>-</span>"
              }
            </div>
          `
        )
        .join("")}
    `;
  } catch (error) {
    table.innerHTML = `<div class="empty-state">Documents table is not ready: ${escapeHtml(error.message)}</div>`;
  }
}

async function handleCreateCenterDocument(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const message = document.querySelector("#document-center-message");
  const formData = new FormData(form);

  message.textContent = "Saving document...";

  try {
    await createDocument({
      entity_type: formData.get("entity_type"),
      entity_id: formData.get("entity_id"),
      doc_type: formData.get("doc_type"),
      file_url: formData.get("file_url") || null,
      status: formData.get("status"),
      uploaded_by: currentSession?.user?.id || null,
    });

    form.reset();
    message.textContent = "Document saved.";
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
  "picked_up",
  "in_transit",
  "delivered",
  "ready_to_bill",
];

function renderDispatchBoard() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Phase 1", title: "Dispatch Board" })}
    <section class="panel">
      <div class="panel-header">
        <div>
          <h2>Status Pipeline</h2>
          <p>Move loads through the operating workflow with quick status updates.</p>
        </div>
        <button id="reload-dispatch" type="button">Refresh</button>
      </div>
      <div id="dispatch-board" class="dispatch-board"></div>
    </section>
  `;

  document.querySelector("#reload-dispatch").addEventListener("click", loadDispatchBoard);
  loadDispatchBoard();
}

async function loadDispatchBoard() {
  const board = document.querySelector("#dispatch-board");
  board.innerHTML = '<div class="empty-state">Loading dispatch board...</div>';

  try {
    const loads = await listLoads();
    const grouped = groupLoadsByStatus(loads);

    board.innerHTML = dispatchStatuses
      .map(
        (status) => `
          <section class="dispatch-column">
            <header>
              <h3>${formatStatus(status)}</h3>
              <span>${grouped[status].length}</span>
            </header>
            <div class="dispatch-cards">
              ${
                grouped[status].length
                  ? grouped[status].map(renderDispatchCard).join("")
                  : '<div class="empty-state compact-empty">No loads</div>'
              }
            </div>
          </section>
        `
      )
      .join("");

    board.querySelectorAll("[data-dispatch-status]").forEach((select) => {
      select.addEventListener("change", handleDispatchStatusChange);
    });
  } catch (error) {
    board.innerHTML = `<div class="empty-state">Dispatch board is not ready: ${escapeHtml(error.message)}</div>`;
  }
}

function groupLoadsByStatus(loads) {
  const grouped = Object.fromEntries(dispatchStatuses.map((status) => [status, []]));

  loads.forEach((load) => {
    const status = dispatchStatuses.includes(load.status) ? load.status : "new";
    grouped[status].push(load);
  });

  return grouped;
}

function renderDispatchCard(load) {
  return `
    <article class="dispatch-card">
      <div class="dispatch-card-head">
        <strong>${escapeHtml(load.load_no)}</strong>
        <span>${formatDate(load.pickup_date)}</span>
      </div>
      <p>${escapeHtml(load.customers?.name || "No customer")}</p>
      <small>${escapeHtml(routeLabel(load))}</small>
      <div class="dispatch-meta">
        <span>${escapeHtml(load.carriers?.name || "Own fleet")}</span>
        <span>${escapeHtml(load.drivers?.name || "No driver")}</span>
        <span>${escapeHtml(load.trucks?.unit_no || "No truck")}</span>
      </div>
      <select data-dispatch-status="${load.id}" aria-label="Change status for ${escapeHtml(load.load_no)}">
        ${dispatchStatuses
          .map((status) => `<option value="${status}" ${status === load.status ? "selected" : ""}>${formatStatus(status)}</option>`)
          .join("")}
      </select>
    </article>
  `;
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

function renderBilling() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Phase 2", title: "Billing / Invoices" })}
    <section class="content-grid load-grid">
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2>Invoice List</h2>
            <p>Customer invoices connected to delivered or ready-to-bill loads.</p>
          </div>
          <button id="reload-invoices" type="button">Refresh</button>
        </div>
        <div id="invoices-table" class="data-table"></div>
      </article>
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2>Create Invoice</h2>
            <p>Creates a basic invoice record for accounting follow-up.</p>
          </div>
        </div>
        <form id="invoice-form" class="record-form">
          <label>
            <span>Invoice number</span>
            <input name="invoice_no" type="text" placeholder="INV-1001" required />
          </label>
          <label>
            <span>Customer</span>
            <select name="customer_id" id="invoice-customer" required></select>
          </label>
          <label>
            <span>Load</span>
            <select name="load_id" id="invoice-load"></select>
          </label>
          <label>
            <span>Amount</span>
            <input name="amount" type="number" min="0" step="0.01" required />
          </label>
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
  `;

  document.querySelector("#reload-invoices").addEventListener("click", loadInvoices);
  document.querySelector("#invoice-form").addEventListener("submit", handleCreateInvoice);
  loadInvoiceFormOptions();
  loadInvoices();
}

async function loadInvoiceFormOptions() {
  const customerSelect = document.querySelector("#invoice-customer");
  const loadSelect = document.querySelector("#invoice-load");

  customerSelect.innerHTML = '<option value="">Loading customers...</option>';
  loadSelect.innerHTML = '<option value="">Loading loads...</option>';

  try {
    const [customers, loads] = await Promise.all([listCustomers(), listLoads()]);

    customerSelect.innerHTML = customers.length
      ? customers.map((customer) => `<option value="${customer.id}">${escapeHtml(customer.name)}</option>`).join("")
      : '<option value="">Add a customer first</option>';
    loadSelect.innerHTML = [
      '<option value="">No load selected</option>',
      ...loads.map((load) => `<option value="${load.id}">${escapeHtml(load.load_no)} - ${escapeHtml(routeLabel(load))}</option>`),
    ].join("");
  } catch (error) {
    customerSelect.innerHTML = `<option value="">${escapeHtml(error.message)}</option>`;
    loadSelect.innerHTML = '<option value="">Unavailable</option>';
  }
}

async function loadInvoices() {
  const table = document.querySelector("#invoices-table");
  table.innerHTML = '<div class="empty-state">Loading invoices...</div>';

  try {
    const invoices = await listInvoices();

    if (!invoices.length) {
      table.innerHTML = '<div class="empty-state">No invoices yet.</div>';
      return;
    }

    table.innerHTML = `
      <div class="data-row invoice-row data-row-head">
        <span>Invoice #</span>
        <span>Customer</span>
        <span>Load</span>
        <span>Amount</span>
        <span>Due</span>
        <span>Status</span>
      </div>
      ${invoices
        .map(
          (invoice) => `
            <div class="data-row invoice-row">
              <strong>${escapeHtml(invoice.invoice_no)}</strong>
              <span>${escapeHtml(invoice.customers?.name || "-")}</span>
              <span>${escapeHtml(invoice.loads?.load_no || "-")}</span>
              <span>${formatMoney(invoice.amount)}</span>
              <span>${formatDate(invoice.due_date)}</span>
              <select data-invoice-status="${invoice.id}" aria-label="Change status for ${escapeHtml(invoice.invoice_no)}">
                ${invoiceStatuses
                  .map(
                    (status) =>
                      `<option value="${status}" ${status === invoice.status ? "selected" : ""}>${formatStatus(status)}</option>`
                  )
                  .join("")}
              </select>
            </div>
          `
        )
        .join("")}
    `;

    table.querySelectorAll("[data-invoice-status]").forEach((select) => {
      select.addEventListener("change", handleInvoiceStatusChange);
    });
  } catch (error) {
    table.innerHTML = `
      <div class="empty-state">
        Invoices table is not ready: ${escapeHtml(error.message)}
      </div>
    `;
  }
}

async function handleCreateInvoice(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const message = document.querySelector("#invoice-message");
  const formData = new FormData(form);

  message.textContent = "Saving invoice...";

  try {
    await createInvoice({
      invoice_no: formData.get("invoice_no"),
      customer_id: formData.get("customer_id"),
      load_id: formData.get("load_id") || null,
      amount: numberOrZero(formData.get("amount")),
      due_date: formData.get("due_date") || null,
      status: formData.get("status"),
      sent_at: formData.get("status") === "sent" ? new Date().toISOString() : null,
    });

    form.reset();
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

  select.disabled = true;

  try {
    await updateInvoiceStatus(invoiceId, status);
    await loadInvoices();
  } catch (error) {
    select.disabled = false;
    window.alert(`Invoice status update failed: ${error.message}`);
  }
}

function renderReports() {
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
    const [customers, carriers, drivers, trucks, trailers, loads, invoices] = await Promise.all([
      listCustomers(),
      listCarriers(),
      listDrivers(),
      listTrucks(),
      listTrailers(),
      listLoads(),
      listInvoices(),
    ]);
    const statusCounts = countBy(loads, "status");
    const totalInvoiceAmount = sumBy(invoices, "amount");
    const unpaidAmount = sumBy(
      invoices.filter((invoice) => !["paid", "void"].includes(invoice.status)),
      "amount"
    );
    const overdueAmount = sumBy(
      invoices.filter((invoice) => invoice.status === "overdue"),
      "amount"
    );

    root.innerHTML = `
      <section class="report-kpis">
        ${renderReportKpi("Customers", customers.length)}
        ${renderReportKpi("Carriers", carriers.length)}
        ${renderReportKpi("Drivers", drivers.length)}
        ${renderReportKpi("Trucks", trucks.length)}
        ${renderReportKpi("Trailers", trailers.length)}
        ${renderReportKpi("Loads", loads.length)}
        ${renderReportKpi("Invoice total", formatMoney(totalInvoiceAmount))}
        ${renderReportKpi("Unpaid", formatMoney(unpaidAmount))}
        ${renderReportKpi("Overdue", formatMoney(overdueAmount))}
        ${renderReportKpi("Ready to bill", statusCounts.ready_to_bill || 0)}
      </section>
      <section class="report-grid">
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
      <section class="panel">
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
  } catch (error) {
    root.innerHTML = `<div class="empty-state">Reports are not ready: ${escapeHtml(error.message)}</div>`;
  }
}

function renderReportKpi(label, value) {
  return `
    <article class="report-kpi">
      <span>${label}</span>
      <strong>${value}</strong>
    </article>
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
            <span>${formatStatus(load.status)}</span>
            <span>${escapeHtml(load.customers?.name || "-")}</span>
            <span>${escapeHtml(routeLabel(load))}</span>
            <span>${formatDate(load.pickup_date)}</span>
          </div>
        `
      )
      .join("")}
  `;
}

const accessTables = ["profiles", "roles", "permissions", "user_roles", "settings"];

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
    const [tableChecks, tableRows] = await Promise.all([
      Promise.all(accessTables.map(checkTable)),
      Promise.all(accessTables.map((table) => listTableRows(table, 12).catch((error) => ({ error })))),
    ]);

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
              <h2>Access Tables</h2>
              <p>Readiness check for profiles, roles and permissions.</p>
            </div>
            <button id="reload-settings" type="button">Refresh</button>
          </div>
          <div class="settings-status">
            ${tableChecks.map(renderSettingsStatusRow).join("")}
          </div>
        </article>
      </section>
      <section class="settings-table-grid">
        ${accessTables
          .map((table, index) => renderSettingsTable(table, tableRows[index]))
          .join("")}
      </section>
    `;

    document.querySelector("#reload-settings").addEventListener("click", loadSettings);
  } catch (error) {
    root.innerHTML = `<div class="empty-state">Settings are not ready: ${escapeHtml(error.message)}</div>`;
  }
}

function renderSettingsStatusRow(result) {
  return `
    <div class="settings-status-row" data-state="${result.ok ? "ok" : "error"}">
      <span>${result.table}</span>
      <strong>${result.ok ? result.count : "-"}</strong>
      <small>${escapeHtml(result.message)}</small>
    </div>
  `;
}

function renderSettingsTable(table, rowsOrError) {
  if (rowsOrError?.error) {
    return `
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2>${table}</h2>
            <p>Unable to read table.</p>
          </div>
        </div>
        <div class="empty-state">${escapeHtml(rowsOrError.error.message)}</div>
      </article>
    `;
  }

  const rows = rowsOrError || [];

  return `
    <article class="panel">
      <div class="panel-header">
        <div>
          <h2>${table}</h2>
          <p>${rows.length} visible record${rows.length === 1 ? "" : "s"}</p>
        </div>
      </div>
      <div class="settings-json-list">
        ${
          rows.length
            ? rows.map((row) => `<pre>${escapeHtml(JSON.stringify(row, null, 2))}</pre>`).join("")
            : '<div class="empty-state">No visible records.</div>'
        }
      </div>
    </article>
  `;
}

function renderLoads() {
  pageRoot.innerHTML = `
    ${renderPageHeader({ eyebrow: "Phase 1", title: "Orders / Loads" })}
    <section class="content-grid load-grid">
      <article class="panel">
        <div class="panel-header">
          <div>
            <h2>All Loads</h2>
            <p>Core order board before the full dispatch kanban is added.</p>
          </div>
          <button id="reload-loads" type="button">Refresh</button>
        </div>
        <div id="loads-table" class="data-table"></div>
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
            <span>Status</span>
            <select name="status">
              <option value="new">New</option>
              <option value="booked">Booked</option>
              <option value="assigned">Assigned</option>
              <option value="dispatched">Dispatched</option>
              <option value="picked_up">Picked up</option>
              <option value="in_transit">In transit</option>
              <option value="delivered">Delivered</option>
              <option value="ready_to_bill">Ready to bill</option>
              <option value="billed">Billed</option>
              <option value="cancelled">Cancelled</option>
            </select>
          </label>
          <label>
            <span>Notes</span>
            <input name="notes" type="text" />
          </label>
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
  document.querySelector("#load-form").addEventListener("submit", handleCreateLoad);
  document.querySelector("#cancel-load-edit").addEventListener("click", resetLoadForm);
  loadLoadFormOptions();
  loadLoads();
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
    const loads = await listLoads();

    if (!loads.length) {
      table.innerHTML = '<div class="empty-state">No loads yet.</div>';
      return;
    }

    table.innerHTML = `
      <div class="data-row load-row data-row-head">
        <span>Load #</span>
        <span>Status</span>
        <span>Customer</span>
        <span>Carrier / Route</span>
        <span>Pickup</span>
        <span></span>
      </div>
      ${loads
        .map(
          (load) => `
            <div class="data-row load-row">
              <strong>${escapeHtml(load.load_no)}</strong>
              <span>${formatStatus(load.status)}</span>
              <span>${escapeHtml(load.customers?.name || "-")}</span>
              <span>${escapeHtml(load.carriers?.name ? `${load.carriers.name} · ${routeLabel(load)}` : routeLabel(load))}</span>
              <span>${formatDate(load.pickup_date)}</span>
              <div class="row-actions">
                <button class="row-action" type="button" data-load-id="${load.id}">Open</button>
                <button class="row-action" type="button" data-edit-load="${load.id}">Edit</button>
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
      const load = loads.find((item) => item.id === button.dataset.editLoad);
      button.addEventListener("click", () => fillLoadForm(load));
    });
  } catch (error) {
    table.innerHTML = `
      <div class="empty-state">
        Loads table is not ready: ${escapeHtml(error.message)}
      </div>
    `;
  }
}

async function fillLoadForm(load) {
  const form = document.querySelector("#load-form");
  await loadLoadFormOptions();
  form.elements.id.value = load.id;
  form.elements.load_no.value = load.load_no || "";
  form.elements.customer_id.value = load.customer_id || "";
  form.elements.carrier_id.value = load.carrier_id || "";
  form.elements.origin.value = load.origin || "";
  form.elements.destination.value = load.destination || "";
  form.elements.pickup_date.value = load.pickup_date || "";
  form.elements.delivery_date.value = load.delivery_date || "";
  form.elements.driver_id.value = load.driver_id || "";
  form.elements.truck_id.value = load.truck_id || "";
  form.elements.trailer_id.value = load.trailer_id || "";
  form.elements.status.value = load.status || "new";
  form.elements.notes.value = load.notes || "";
  document.querySelector("#load-form-title").textContent = "Edit Order";
  document.querySelector("#load-submit").textContent = "Update Load";
  document.querySelector("#cancel-load-edit").hidden = false;
  document.querySelector("#load-message").textContent = "";
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
}

async function renderLoadDetails(loadId) {
  const panel = document.querySelector("#load-details-panel");
  panel.hidden = false;
  panel.innerHTML = '<div class="empty-state">Loading load details...</div>';

  try {
    const [load, stops, financials, documents] = await Promise.all([
      getLoad(loadId),
      listLoadStops(loadId),
      getLoadFinancials(loadId),
      listLoadDocuments(loadId),
    ]);

    panel.innerHTML = `
      <div class="panel-header">
        <div>
          <h2>Load ${escapeHtml(load.load_no)}</h2>
          <p>${escapeHtml(routeLabel(load))}</p>
        </div>
        <span class="status-pill" data-state="ok">${formatStatus(load.status)}</span>
      </div>
      <div class="detail-layout">
        <section class="detail-summary">
          <article><span>Customer</span><strong>${escapeHtml(load.customers?.name || "-")}</strong></article>
          <article><span>Carrier</span><strong>${escapeHtml(load.carriers?.name || "No outside carrier")}</strong></article>
          <article><span>Driver</span><strong>${escapeHtml(load.drivers?.name || "Unassigned")}</strong></article>
          <article><span>Truck</span><strong>${escapeHtml(load.trucks?.unit_no || "Unassigned")}</strong></article>
          <article><span>Trailer</span><strong>${escapeHtml(load.trailers?.unit_no || "Unassigned")}</strong></article>
          <article><span>Pickup</span><strong>${formatDate(load.pickup_date)}</strong></article>
          <article><span>Delivery</span><strong>${formatDate(load.delivery_date)}</strong></article>
          <article><span>Notes</span><strong>${escapeHtml(load.notes || "-")}</strong></article>
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
        <section class="stop-form-wrap">
          <div class="subsection-header">
            <div>
              <h3>Add Stop</h3>
              <p>Creates a stop under this load.</p>
            </div>
          </div>
          <form id="stop-form" class="record-form compact-form">
            <input name="load_id" type="hidden" value="${load.id}" />
            <label>
              <span>Type</span>
              <select name="type">
                <option value="pickup">Pickup</option>
                <option value="delivery">Delivery</option>
                <option value="stop">Stop</option>
              </select>
            </label>
            <label>
              <span>Order</span>
              <input name="stop_order" type="number" min="1" value="${stops.length + 1}" />
            </label>
            <label>
              <span>Facility</span>
              <input name="facility" type="text" />
            </label>
            <label>
              <span>Address</span>
              <input name="address" type="text" />
            </label>
            <label>
              <span>Appointment from</span>
              <input name="appointment_from" type="datetime-local" />
            </label>
            <label>
              <span>Appointment to</span>
              <input name="appointment_to" type="datetime-local" />
            </label>
            <label>
              <span>Reference</span>
              <input name="reference" type="text" />
            </label>
            <label>
              <span>Instructions</span>
              <input name="instructions" type="text" />
            </label>
            <button type="submit">Save Stop</button>
            <p id="stop-message" class="form-message"></p>
          </form>
        </section>
        <section class="financials-section">
          <div class="subsection-header">
            <div>
              <h3>Financials</h3>
              <p>Revenue, cost and margin for this load.</p>
            </div>
          </div>
          ${renderFinancialSummary(financials)}
          <form id="financials-form" class="record-form compact-form">
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
              <p>Rate con, BOL, POD, fuel, lumper and invoice links.</p>
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
              <span>File URL</span>
              <input name="file_url" type="url" placeholder="https://..." />
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
      </div>
    `;

    document.querySelector("#stop-form").addEventListener("submit", handleCreateStop);
    document.querySelector("#financials-form").addEventListener("submit", handleSaveFinancials);
    document.querySelector("#document-form").addEventListener("submit", handleCreateDocument);
    panel.scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (error) {
    panel.innerHTML = `<div class="empty-state">Load details are not ready: ${escapeHtml(error.message)}</div>`;
  }
}

function renderStops(stops) {
  if (!stops.length) {
    return '<div class="empty-state">No stops yet.</div>';
  }

  return stops
    .map(
      (stop) => `
        <article class="stop-card">
          <div>
            <strong>${stop.stop_order}. ${formatStatus(stop.type)}</strong>
            <span>${escapeHtml(stop.facility || "Facility not set")}</span>
          </div>
          <p>${escapeHtml(stop.address || "-")}</p>
          <small>${formatDateTime(stop.appointment_from)} → ${formatDateTime(stop.appointment_to)}</small>
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
            <span>${formatStatus(document.status)}</span>
          </div>
          ${
            document.file_url
              ? `<a href="${escapeAttribute(document.file_url)}" target="_blank" rel="noreferrer">Open document</a>`
              : "<small>No URL</small>"
          }
        </article>
      `
    )
    .join("");
}

async function handleCreateStop(event) {
  event.preventDefault();

  const form = event.currentTarget;
  const message = document.querySelector("#stop-message");
  const formData = new FormData(form);
  const loadId = formData.get("load_id");

  message.textContent = "Saving stop...";

  try {
    await createLoadStop({
      load_id: loadId,
      type: formData.get("type"),
      stop_order: numberOrNull(formData.get("stop_order")) || 1,
      facility: formData.get("facility") || null,
      address: formData.get("address") || null,
      appointment_from: dateTimeOrNull(formData.get("appointment_from")),
      appointment_to: dateTimeOrNull(formData.get("appointment_to")),
      reference: formData.get("reference") || null,
      instructions: formData.get("instructions") || null,
    });

    message.textContent = "Stop saved.";
    await renderLoadDetails(loadId);
  } catch (error) {
    message.textContent = error.message;
  }
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
    await renderLoadDetails(loadId);
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

  message.textContent = "Saving document...";

  try {
    await createLoadDocument({
      entity_id: loadId,
      doc_type: formData.get("doc_type"),
      file_url: formData.get("file_url") || null,
      status: formData.get("status"),
      uploaded_by: currentSession?.user?.id || null,
    });

    message.textContent = "Document saved.";
    await renderLoadDetails(loadId);
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
    const payload = {
      load_no: formData.get("load_no"),
      customer_id: formData.get("customer_id"),
      carrier_id: formData.get("carrier_id") || null,
      driver_id: formData.get("driver_id") || null,
      truck_id: formData.get("truck_id") || null,
      trailer_id: formData.get("trailer_id") || null,
      status: formData.get("status"),
      origin: formData.get("origin") || null,
      destination: formData.get("destination") || null,
      pickup_date: formData.get("pickup_date") || null,
      delivery_date: formData.get("delivery_date") || null,
      notes: formData.get("notes") || null,
    };

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

  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "2-digit",
    year: "numeric",
  }).format(new Date(`${value}T00:00:00`));
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

function formatNumber(value) {
  if (value === null || value === undefined || value === "") {
    return "-";
  }

  return Number(value).toLocaleString("en-US");
}

function formatMoney(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
  }).format(Number(value || 0));
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

loginForm.addEventListener("submit", handleLogin);
signupButton.addEventListener("click", handleSignup);
logoutButton.addEventListener("click", async () => {
  await signOut();
});
refreshButton.addEventListener("click", renderRoute);
addOrderButton.addEventListener("click", () => {
  window.location.hash = "loads";
});
window.addEventListener("hashchange", renderRoute);

onAuthStateChange(setAuthView);

getSession()
  .then(setAuthView)
  .catch(() => setAuthView(null));
