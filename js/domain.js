export function numeric(value) {
  if (value === null || value === undefined || value === "") return 0;
  const result = Number(value);
  return Number.isFinite(result) ? result : 0;
}

export function rankGlobalSearchRecords(records = [], query = "", limit = 8) {
  const needle = String(query).trim().toLowerCase();
  if (needle.length < 2) return [];
  return records
    .map((record) => {
      const title = String(record.title || "").toLowerCase();
      const detail = String(record.detail || "").toLowerCase();
      const keywords = String(record.keywords || "").toLowerCase();
      let score = 0;
      if (title === needle) score = 100;
      else if (title.startsWith(needle)) score = 80;
      else if (title.includes(needle)) score = 60;
      else if (detail.includes(needle)) score = 40;
      else if (keywords.includes(needle)) score = 20;
      return { ...record, score };
    })
    .filter((record) => record.score > 0)
    .sort((a, b) => b.score - a.score || String(a.title).localeCompare(String(b.title)))
    .slice(0, limit);
}

const tableSavedViewColumns = {
  loads: new Set(["status", "customer", "route", "pickup"]),
  invoices: new Set(["customer", "load", "amount", "due", "status"]),
};

const tableSavedViewSorts = {
  loads: new Set(["created_at", "load_no", "pickup_date", "delivery_date", "operational_status"]),
  invoices: new Set(["created_at", "invoice_no", "amount", "due_date", "status"]),
};

const tableSavedViewStatuses = {
  loads: new Set(["", "new", "booked", "assigned", "dispatched", "in_transit", "delivered", "ready_to_bill", "billed"]),
  invoices: new Set(["", "draft", "sent", "partial", "paid", "overdue", "void"]),
};

export function validateTableSavedView(type, input = {}) {
  const tableType = String(type || input.table_type || "").trim();
  if (!tableSavedViewColumns[tableType]) throw new Error("Unsupported saved table view.");
  const name = String(input.name || "").trim();
  if (name.length < 2) throw new Error("Saved view name must be at least 2 characters.");
  if (name.length > 80) throw new Error("Saved view name must be 80 characters or fewer.");
  const settings = input.settings || input;
  const allowedColumns = tableSavedViewColumns[tableType];
  const columns = [...new Set((settings.columns || []).filter((column) => allowedColumns.has(column)))];
  if (!columns.length) throw new Error("Choose at least one visible column.");
  const sortBy = tableSavedViewSorts[tableType].has(settings.sortBy) ? settings.sortBy : "created_at";
  const sortDirection = settings.sortDirection === "asc" ? "asc" : "desc";
  const pageSize = Math.min(100, Math.max(1, Number(settings.pageSize) || (tableType === "loads" ? 10 : 5)));
  const normalizedSettings = {
    search: String(settings.search || "").trim().slice(0, 120),
    status: tableSavedViewStatuses[tableType].has(settings.status) ? settings.status : "",
    pageSize,
    sortBy,
    sortDirection,
    columns,
  };
  if (tableType === "loads") {
    normalizedSettings.customerId = String(settings.customerId || "").trim();
    normalizedSettings.driverId = String(settings.driverId || "").trim();
    normalizedSettings.dateFrom = settings.dateFrom || "";
    normalizedSettings.dateTo = settings.dateTo || "";
    normalizedSettings.view = ["all", "unassigned", "today", "in_transit", "ready_to_bill"].includes(settings.view) ? settings.view : "all";
  }
  return {
    table_type: tableType,
    name,
    visibility: input.visibility === "shared" ? "shared" : "private",
    settings: normalizedSettings,
  };
}

export function differenceInCalendarDays(laterDate, earlierDate) {
  const later = Date.parse(`${laterDate}T00:00:00Z`);
  const earlier = Date.parse(`${earlierDate}T00:00:00Z`);
  if (Number.isNaN(later) || Number.isNaN(earlier)) return 0;
  return Math.max(0, Math.floor((later - earlier) / 86400000));
}

export function invoiceFinancialState(invoice, today) {
  const paid = (invoice.payments || []).reduce((sum, payment) => sum + numeric(payment.amount), 0);
  const credited = (invoice.invoice_credit_memos || [])
    .filter((memo) => memo.status === "applied")
    .reduce((sum, memo) => sum + numeric(memo.amount), 0);
  const balance = Math.max(0, numeric(invoice.amount) - paid - credited);
  const collectible = !["draft", "void"].includes(invoice.status) && balance > 0;
  const daysPastDue = collectible && invoice.due_date && invoice.due_date < today
    ? differenceInCalendarDays(today, invoice.due_date)
    : 0;
  const displayStatus = daysPastDue > 0 ? "overdue" : invoice.status;
  return { paid, credited, balance, collectible, daysPastDue, displayStatus };
}

export const manualInvoiceStatusOptions = ["draft", "sent", "void"];

export function canManuallySetInvoiceStatus(status) {
  return manualInvoiceStatusOptions.includes(status);
}

export function invoiceBatchCandidates(invoices = [], today = "") {
  return invoices
    .map((invoice) => ({ ...invoice, financialState: invoiceFinancialState(invoice, today) }))
    .filter((invoice) => !["paid", "void"].includes(invoice.status) && invoice.financialState.balance > 0 && !invoice.batch_id);
}

export function collectionFollowUpState(record = {}, today = "") {
  const daysPastDue = numeric(record.daysPastDue);
  const followUpDate = record.follow_up_date || record.followUpDate || null;
  const followUpDue = Boolean(followUpDate && today && followUpDate <= today);
  let severity = "current";
  if (daysPastDue >= 30 || followUpDue) severity = "urgent";
  else if (daysPastDue > 0) severity = "overdue";
  else if (followUpDate) severity = "scheduled";
  return { severity, followUpDue };
}

const invoiceLineItemTypes = new Set(["linehaul", "fsc", "detention", "layover", "lumper", "tonu", "discount", "tax", "credit_memo", "custom"]);

export function invoiceLineItemsTotal(items = []) {
  return items.reduce((sum, item) => sum + (numeric(item.quantity || 1) * numeric(item.unit_amount)), 0);
}

export function validateInvoiceLineItems(items = []) {
  if (!items.length) throw new Error("At least one invoice line item is required.");
  return items.map((item, index) => {
    const itemType = String(item.item_type || "custom").trim();
    const description = String(item.description || "").trim();
    const customerDescription = String(item.customer_description || description).trim();
    const quantity = numeric(item.quantity || 1);
    const unitAmount = numeric(item.unit_amount);
    if (!invoiceLineItemTypes.has(itemType)) throw new Error("Unsupported invoice line item type.");
    if (!description) throw new Error("Invoice line item description is required.");
    if (quantity <= 0) throw new Error("Invoice line item quantity must be greater than zero.");
    if (unitAmount === 0) throw new Error("Invoice line item amount cannot be zero.");
    return {
      item_type: itemType,
      description,
      customer_description: customerDescription || description,
      quantity,
      unit_amount: unitAmount,
      receipt_required: Boolean(item.receipt_required),
      customer_visible: item.customer_visible !== false,
      approved: Boolean(item.approved),
      sort_order: index + 1,
    };
  });
}

const customerBillingCycles = new Set(["per_load", "weekly", "biweekly", "monthly"]);

export function validateCustomerBillingCycle({ billingCycle = "per_load", billingCycleDay = null } = {}) {
  const cycle = String(billingCycle || "per_load").trim();
  if (!customerBillingCycles.has(cycle)) throw new Error("Unsupported customer billing cycle.");
  const day = billingCycleDay === "" || billingCycleDay === null || billingCycleDay === undefined
    ? null
    : Number(billingCycleDay);
  if (day !== null && (!Number.isInteger(day) || day < 1 || day > 31)) {
    throw new Error("Billing cycle day must be between 1 and 31.");
  }
  return { billing_cycle: cycle, billing_cycle_day: day };
}

export function validateInvoiceNumberingSettings({ prefix, nextNumber, padding }) {
  const invoicePrefix = String(prefix || "").trim();
  const invoiceNextNumber = Number(nextNumber);
  const invoicePadding = Number(padding);
  if (!invoicePrefix) throw new Error("Invoice prefix is required.");
  if (!Number.isInteger(invoiceNextNumber) || invoiceNextNumber < 1) throw new Error("Next invoice number must be a positive integer.");
  if (!Number.isInteger(invoicePadding) || invoicePadding < 1 || invoicePadding > 12) throw new Error("Invoice number padding must be between 1 and 12.");
  return {
    invoice_number_prefix: invoicePrefix,
    invoice_number_next: invoiceNextNumber,
    invoice_number_padding: invoicePadding,
  };
}

export function validateCompanyTenantProfile({ legalName, timezone, locale, country, status = "active" } = {}) {
  const legal_name = String(legalName || "").trim();
  const tenantTimezone = String(timezone || "America/Chicago").trim();
  const tenantLocale = String(locale || "en-US").trim();
  const tenantCountry = String(country || "USA").trim();
  const tenantStatus = String(status || "active").trim();
  if (!legal_name) throw new Error("Company legal name is required.");
  if (!tenantTimezone) throw new Error("Company timezone is required.");
  if (!tenantLocale) throw new Error("Company locale is required.");
  if (!tenantCountry) throw new Error("Company country is required.");
  if (!["active", "inactive"].includes(tenantStatus)) throw new Error("Unsupported company status.");
  return {
    legal_name,
    timezone: tenantTimezone,
    locale: tenantLocale,
    country: tenantCountry,
    status: tenantStatus,
  };
}

export function operationalLoadStatus(load) {
  if (!load) return "new";
  if (load.operational_status) return load.operational_status;
  if (["ready_to_bill", "billed"].includes(load.status)) return "delivered";
  return load.status || "new";
}

export function billingLoadStatus(load, activeInvoice = null) {
  if (load?.billing_status) return load.billing_status;
  if (activeInvoice?.status === "void") return "void";
  if (activeInvoice?.status === "paid") return "paid";
  if (activeInvoice?.status === "partial") return "partial";
  if (["sent", "overdue"].includes(activeInvoice?.status)) return "invoiced";
  if (activeInvoice?.status === "draft") return "drafted";
  if (load?.status === "billed") return "invoiced";
  if (load?.status === "ready_to_bill") return "ready_to_bill";
  return "not_ready";
}

export function loadDocumentStatus(load, documents = []) {
  if (load?.document_status) return load.document_status;
  const hasPod = documents.some((document) => document.doc_type === "pod");
  const approvedPod = documents.some((document) => document.doc_type === "pod" && document.status === "approved");
  if (approvedPod) return "pod_approved";
  return hasPod ? "pod_pending" : "pod_missing";
}

export function isLoadOperationallyClosed(load) {
  return ["delivered", "cancelled"].includes(operationalLoadStatus(load));
}

export function calculateLoadProfitability(record) {
  const revenue = numeric(record.linehaul) + numeric(record.fsc) + numeric(record.accessorials);
  const cost = numeric(record.deductions) + numeric(record.carrier_cost) + numeric(record.driver_pay);
  const margin = revenue - cost;
  return { revenue, cost, margin, marginPercent: revenue ? (margin / revenue) * 100 : 0 };
}

const settlementTransitions = {
  draft: ["approved", "void"],
  approved: ["paid", "void"],
  paid: [],
  void: [],
};

export function settlementStatusOptions(status) {
  return [status, ...(settlementTransitions[status] || [])];
}

export function settlementPayBatchCandidates(settlements = [], existingBatches = []) {
  const batchedSettlementIds = new Set(existingBatches.flatMap((batch) => (batch.settlement_pay_batch_items || []).map((item) => item.settlement_id)));
  return settlements.filter((settlement) => ["draft", "approved"].includes(settlement.status)
    && numeric(settlement.total_amount) > 0
    && !batchedSettlementIds.has(settlement.id));
}

export function getMaintenanceDueState(record, equipment, today, dueSoonDate) {
  if (!["open", "scheduled"].includes(record.status)) return "closed";
  const dueMiles = record.due_miles === null || record.due_miles === undefined ? null : numeric(record.due_miles);
  const odometer = numeric(equipment?.odometer);
  if ((record.due_date && record.due_date < today) || (dueMiles !== null && odometer >= dueMiles)) return "overdue";
  if ((record.due_date && record.due_date <= dueSoonDate) || (dueMiles !== null && odometer + 1000 >= dueMiles)) return "due_soon";
  return "current";
}

export function getComplianceDueState(expiryDate, today, dueSoonDate) {
  if (!expiryDate) return "missing";
  if (expiryDate < today) return "expired";
  if (expiryDate <= dueSoonDate) return "due_soon";
  return "current";
}

export function getDispatchComplianceState(resource, resourceType, today, dueSoonDate) {
  const requirements = resourceType === "driver"
    ? [["CDL", resource.cdl_expiry], ["Medical", resource.medical_expiry]]
    : [["Registration", resource.registration_expiry], ["Insurance", resource.insurance_expiry]];
  const details = requirements.map(([label, expiryDate]) => ({
    label,
    expiryDate,
    state: getComplianceDueState(expiryDate, today, dueSoonDate),
  }));
  const priority = ["expired", "due_soon", "missing", "current"];
  return {
    state: priority.find((state) => details.some((detail) => detail.state === state)) || "current",
    details,
  };
}

export const tmsPermissionRoles = {
  view_dashboard: ["owner", "admin", "dispatcher", "accounting", "safety"],
  view_loads: ["owner", "admin", "dispatcher", "accounting", "safety"],
  view_quotes: ["owner", "admin", "dispatcher", "accounting"],
  view_dispatch: ["owner", "admin", "dispatcher"],
  view_tracking: ["owner", "admin", "dispatcher", "safety"],
  view_drivers: ["owner", "admin", "dispatcher", "safety"],
  view_equipment: ["owner", "admin", "dispatcher", "safety"],
  view_maintenance: ["owner", "admin", "dispatcher", "safety"],
  view_customers: ["owner", "admin", "dispatcher", "accounting"],
  view_carriers: ["owner", "admin", "dispatcher", "accounting", "safety"],
  view_documents: ["owner", "admin", "dispatcher", "accounting", "safety"],
  view_billing: ["owner", "admin", "accounting"],
  view_settlements: ["owner", "admin", "accounting"],
  view_reports: ["owner", "admin", "dispatcher", "accounting", "safety"],
  view_notifications: ["owner", "admin", "dispatcher", "accounting", "safety"],
  view_driver: ["owner", "admin", "dispatcher", "safety", "driver"],
  view_customer_portal: ["customer_portal"],
  view_carrier_portal: ["carrier_portal"],
  view_settings: ["owner"],
  view_audit: ["owner", "admin"],
  manage_operations: ["owner", "admin", "dispatcher"],
  manage_quotes: ["owner", "admin", "dispatcher", "accounting"],
  manage_carriers: ["owner", "admin", "dispatcher", "safety"],
  manage_assets: ["owner", "admin", "dispatcher", "safety"],
  manage_safety: ["owner", "admin", "safety"],
  manage_finance: ["owner", "admin", "accounting"],
  manage_documents: ["owner", "admin", "dispatcher", "accounting", "safety"],
  manage_driver_work: ["driver"],
  manage_users: ["owner"],
};

export function calculateQuotePricing(input = {}) {
  const miles = numeric(input.loaded_miles);
  const ratePerMile = numeric(input.rate_per_mile);
  const fixedLinehaul = numeric(input.linehaul);
  const minimumCharge = numeric(input.minimum_charge);
  const linehaul = Math.max(minimumCharge, fixedLinehaul || miles * ratePerMile);
  const fuelSurcharge = input.fuel_surcharge === null || input.fuel_surcharge === undefined || input.fuel_surcharge === ""
    ? linehaul * numeric(input.fuel_surcharge_percent) / 100
    : numeric(input.fuel_surcharge);
  const accessorials = numeric(input.accessorials);
  const discount = numeric(input.discount);
  const estimatedCarrierCost = numeric(input.estimated_carrier_cost);
  const customerTotal = Math.max(0, linehaul + fuelSurcharge + accessorials - discount);
  const estimatedMargin = customerTotal - estimatedCarrierCost;
  return {
    linehaul: Number(linehaul.toFixed(2)),
    fuel_surcharge: Number(fuelSurcharge.toFixed(2)),
    customer_total: Number(customerTotal.toFixed(2)),
    estimated_margin: Number(estimatedMargin.toFixed(2)),
    estimated_margin_percent: customerTotal ? Number((estimatedMargin / customerTotal * 100).toFixed(2)) : 0,
  };
}

export function validateFreightQuoteInput(input = {}) {
  const normalized = {
    quote_no: String(input.quote_no || "").trim(),
    customer_id: String(input.customer_id || "").trim(),
    origin: String(input.origin || "").trim(),
    destination: String(input.destination || "").trim(),
    pickup_date: input.pickup_date || null,
    delivery_date: input.delivery_date || null,
    loaded_miles: numeric(input.loaded_miles),
    linehaul: numeric(input.linehaul),
    fuel_surcharge: numeric(input.fuel_surcharge),
    accessorials: numeric(input.accessorials),
    discount: numeric(input.discount),
    estimated_carrier_cost: numeric(input.estimated_carrier_cost),
  };
  if (!normalized.quote_no) throw new Error("Quote number is required");
  if (!normalized.customer_id) throw new Error("Customer is required");
  if (!normalized.origin || !normalized.destination) throw new Error("Origin and destination are required");
  if (normalized.pickup_date && normalized.delivery_date && normalized.delivery_date < normalized.pickup_date) throw new Error("Delivery date cannot be before pickup date");
  for (const field of ["loaded_miles", "linehaul", "fuel_surcharge", "accessorials", "discount", "estimated_carrier_cost"]) {
    if (normalized[field] < 0) throw new Error("Quote amounts and miles cannot be negative");
  }
  if (calculateQuotePricing(normalized).customer_total <= 0) throw new Error("Quote total must be greater than zero");
  return normalized;
}

export function customerPortalSummary(loads = [], invoices = []) {
  const activeStatuses = new Set(["assigned", "dispatched", "en_route", "arrived_pickup", "loading", "picked_up", "in_transit", "arrived_delivery", "unloading", "exception"]);
  const delivered = loads.filter((load) => operationalLoadStatus(load) === "delivered").length;
  const active = loads.filter((load) => activeStatuses.has(operationalLoadStatus(load))).length;
  const openInvoices = invoices.filter((invoice) => !["paid", "void", "draft"].includes(invoice.status)).length;
  return { shipments: loads.length, active, delivered, openInvoices };
}

export function carrierPortalSummary(loads = [], tenders = [], settlements = []) {
  return {
    loads: loads.length,
    pendingTenders: tenders.filter((item) => item.status === "pending").length,
    active: loads.filter((load) => !["delivered", "cancelled"].includes(operationalLoadStatus(load))).length,
    payable: settlements.filter((item) => ["approved", "paid"].includes(item.status)).reduce((sum, item) => sum + numeric(item.total_amount), 0),
  };
}

export function carrierBidComparison(bids = [], revenue = 0) {
  const customerRevenue = numeric(revenue);
  return bids.map((bid) => {
    const amount = numeric(bid.bid_amount);
    const margin = customerRevenue - amount;
    return { ...bid, margin, marginPercent: customerRevenue ? Number((margin / customerRevenue * 100).toFixed(2)) : 0 };
  }).sort((a, b) => {
    if (a.status === "submitted" && b.status !== "submitted") return -1;
    if (b.status === "submitted" && a.status !== "submitted") return 1;
    return numeric(a.bid_amount) - numeric(b.bid_amount);
  });
}

export function validateCarrierBidInput(input = {}) {
  const normalized = { carrier_id: String(input.carrier_id || "").trim(), target_amount: numeric(input.target_amount), expires_at: input.expires_at || null };
  if (!normalized.carrier_id) throw new Error("Carrier is required");
  if (normalized.target_amount < 0) throw new Error("Target amount cannot be negative");
  if (normalized.expires_at && new Date(normalized.expires_at).getTime() <= Date.now()) throw new Error("Bid expiry must be in the future");
  return normalized;
}

export function nextRecurringLoadDate(dateValue, frequency, intervalCount = 1) {
  const date = new Date(`${dateValue}T00:00:00Z`); const interval = Math.max(1, Number(intervalCount) || 1);
  if (frequency === "daily") date.setUTCDate(date.getUTCDate() + interval);
  else if (frequency === "weekly") date.setUTCDate(date.getUTCDate() + 7 * interval);
  else if (frequency === "biweekly") date.setUTCDate(date.getUTCDate() + 14 * interval);
  else if (frequency === "monthly") date.setUTCMonth(date.getUTCMonth() + interval);
  else throw new Error("Choose a valid recurring frequency");
  return date.toISOString().slice(0,10);
}

export function validateRecurringLoadTemplate(input = {}) {
  const normalized = { name:String(input.name||"").trim(),customer_id:String(input.customer_id||"").trim(),frequency:String(input.frequency||"").trim(),interval_count:Number(input.interval_count)||1,next_run_date:input.next_run_date||null,delivery_offset_days:Number(input.delivery_offset_days)||0,origin:String(input.origin||"").trim(),destination:String(input.destination||"").trim() };
  if (!normalized.name) throw new Error("Template name is required");
  if (!normalized.customer_id) throw new Error("Customer is required");
  if (!["daily","weekly","biweekly","monthly"].includes(normalized.frequency)) throw new Error("Choose a valid recurring frequency");
  if (!normalized.next_run_date) throw new Error("Next run date is required");
  if (!normalized.origin || !normalized.destination) throw new Error("Origin and destination are required");
  if (normalized.interval_count<1 || normalized.interval_count>52) throw new Error("Interval must be between 1 and 52");
  if (normalized.delivery_offset_days<0 || normalized.delivery_offset_days>90) throw new Error("Delivery offset must be between 0 and 90 days");
  return normalized;
}

export const reportBuilderColumns = ["load_no","customer","driver","status","pickup_date","revenue","cost","margin","margin_percent"];

export function validateSavedReportView(input={}) {
  const name=String(input.name||"").trim(); const columns=[...new Set(input.columns||[])].filter((item)=>reportBuilderColumns.includes(item));
  const groupBy=input.group_by||null;
  if(!name) throw new Error("Report name is required");
  if(!columns.length) throw new Error("Choose at least one report column");
  if(groupBy && !["customer","driver","status"].includes(groupBy)) throw new Error("Choose a valid report grouping");
  return {name,columns,filters:input.filters||{},group_by:groupBy,sort_by:reportBuilderColumns.includes(input.sort_by)?input.sort_by:"pickup_date",sort_direction:input.sort_direction==="asc"?"asc":"desc"};
}

export function applyReportView(records=[],view={}) {
  const f=view.filters||{};
  const rows=records.filter((record)=>{
    const load=record.loads||{};
    if(f.customer_id && load.customer_id!==f.customer_id)return false;
    if(f.driver_id && load.driver_id!==f.driver_id)return false;
    if(f.status && String(load.operational_status||load.status)!==f.status)return false;
    if(f.date_from && (!load.pickup_date||load.pickup_date<f.date_from))return false;
    if(f.date_to && (!load.pickup_date||load.pickup_date>f.date_to))return false;
    return true;
  });
  const key=view.sort_by||"pickup_date"; const direction=view.sort_direction==="asc"?1:-1;
  return rows.sort((a,b)=>String(reportRecordValue(a,key)??"").localeCompare(String(reportRecordValue(b,key)??""),undefined,{numeric:true})*direction);
}

export function reportRecordValue(record,key) {
  const load=record.loads||{};
  return ({load_no:load.load_no,customer:load.customers?.name,driver:load.drivers?.name,status:load.operational_status||load.status,pickup_date:load.pickup_date,revenue:record.revenue,cost:record.cost,margin:record.margin,margin_percent:record.marginPercent})[key];
}

export function operationalKpiSummary(loads=[],financials=[]) {
  const financialByLoad=new Map(financials.map((record)=>[record.load_id,calculateLoadProfitability(record)]));
  let loadedMiles=0; let totalMiles=0; let revenueMiles=0; let revenue=0; let measuredStops=0; let onTimeStops=0; let detentionStops=0; let detentionHours=0;
  const drivers=new Map();
  for(const load of loads){
    const loaded=numeric(load.loaded_miles); const empty=numeric(load.empty_miles); const financial=financialByLoad.get(load.id); loadedMiles+=loaded; totalMiles+=loaded+empty;if(financial){revenue+=financial.revenue;revenueMiles+=loaded;}
    const driverName=load.drivers?.name; const driver=driverName?(drivers.get(load.driver_id)||{id:load.driver_id,name:driverName,loads:0,measured:0,on_time:0}):null;
    if(driver){driver.loads+=1;drivers.set(load.driver_id,driver);}
    for(const stop of load.load_stops||[]){
      const arrival=Date.parse(stop.actual_arrival); const deadline=Date.parse(stop.appointment_to||stop.appointment_from);
      if(Number.isFinite(arrival)&&Number.isFinite(deadline)){measuredStops+=1;if(arrival<=deadline)onTimeStops+=1;if(driver&&stop.type==="delivery"){driver.measured+=1;if(arrival<=deadline)driver.on_time+=1;}}
      const departure=Date.parse(stop.actual_departure);
      if(Number.isFinite(arrival)&&Number.isFinite(departure)){const durationHours=Math.max(0,(departure-arrival)/3600000);if(durationHours>2){detentionStops+=1;detentionHours+=durationHours-2;}}
    }
  }
  return {on_time_percent:measuredStops?(onTimeStops/measuredStops)*100:null,measured_stops:measuredStops,detention_stops:detentionStops,detention_hours:detentionHours,utilization_percent:totalMiles?(loadedMiles/totalMiles)*100:null,revenue_per_loaded_mile:revenueMiles?revenue/revenueMiles:null,driver_performance:[...drivers.values()].map((driver)=>({...driver,on_time_percent:driver.measured?(driver.on_time/driver.measured)*100:null})).sort((a,b)=>(b.on_time_percent??-1)-(a.on_time_percent??-1)||b.loads-a.loads).slice(0,5)};
}

export function nextScheduledReportRun(value,frequency){const date=new Date(value);if(Number.isNaN(date.getTime()))throw new Error("Next run date is required");if(frequency==="daily")date.setUTCDate(date.getUTCDate()+1);else if(frequency==="weekly")date.setUTCDate(date.getUTCDate()+7);else if(frequency==="monthly")date.setUTCMonth(date.getUTCMonth()+1);else throw new Error("Choose a valid report frequency");return date.toISOString();}
export function validateScheduledReport(input={}){const name=String(input.name||"").trim();const reportViewId=String(input.report_view_id||"").trim();const frequency=String(input.frequency||"");const recipients=String(input.recipients||"").split(",").map((item)=>item.trim().toLowerCase()).filter(Boolean);if(!name)throw new Error("Schedule name is required");if(!reportViewId)throw new Error("Saved view is required");if(!["daily","weekly","monthly"].includes(frequency))throw new Error("Choose a valid report frequency");if(!input.next_run_at||Number.isNaN(Date.parse(input.next_run_at)))throw new Error("Next run date is required");if(!recipients.length||recipients.some((email)=>!/^\S+@\S+\.\S+$/.test(email)))throw new Error("Enter valid recipient emails");return{name,report_view_id:reportViewId,frequency,recipients,export_format:"csv",next_run_at:new Date(input.next_run_at).toISOString(),active:Boolean(input.active)};}

export function profitabilityPivot(records=[],groupBy="customer",sortBy="margin"){if(!["customer","driver","status"].includes(groupBy))throw new Error("Choose a valid pivot grouping");const groups=new Map();for(const record of records){const financial={...record,...calculateLoadProfitability(record)};const load=record.loads||{};const label=groupBy==="customer"?(load.customers?.name||"Unassigned customer"):groupBy==="driver"?(load.drivers?.name||"Unassigned driver"):(load.operational_status||load.status||"Unknown");const row=groups.get(label)||{label,loads:0,revenue:0,cost:0,margin:0,margin_percent:0};row.loads+=1;row.revenue+=financial.revenue;row.cost+=financial.cost;row.margin+=financial.margin;row.margin_percent=row.revenue?(row.margin/row.revenue)*100:0;groups.set(label,row);}const key=["loads","revenue","cost","margin","margin_percent"].includes(sortBy)?sortBy:"margin";return[...groups.values()].sort((a,b)=>b[key]-a[key]||a.label.localeCompare(b.label));}

const loadAuditFields = {
  loads: ["status", "operational_status", "document_status", "billing_status", "customer_id", "carrier_id", "driver_id", "truck_id", "trailer_id", "origin", "destination", "pickup_date", "delivery_date", "reference_no", "broker_reference", "commodity", "weight_lbs", "pieces", "pallets", "equipment_type", "temperature_min_f", "temperature_max_f", "hazmat", "loaded_miles", "empty_miles", "internal_notes", "customer_notes"],
  load_stops: ["type", "stop_order", "facility_id", "facility", "address", "contact_name", "contact_phone", "appointment_from", "appointment_to", "actual_arrival", "actual_departure", "reference", "instructions"],
  load_financials: ["linehaul", "fsc", "accessorials", "deductions", "carrier_cost", "driver_pay"],
  documents: ["doc_type", "status", "file_name"],
  invoices: ["invoice_no", "amount", "status", "due_date"],
};

export function loadAuditChanges(log) {
  if (log.action !== "update") return [];
  const oldValue = log.old_value || {};
  const newValue = log.new_value || {};
  return (loadAuditFields[log.entity_type] || [])
    .filter((field) => oldValue[field] !== newValue[field])
    .map((field) => ({ field, oldValue: oldValue[field] ?? null, newValue: newValue[field] ?? null }));
}

export function validateMessageBody(body) {
  const normalized = String(body ?? "").trim();
  if (!normalized) throw new Error("Enter a message before sending");
  if (normalized.length > 4000) throw new Error("Message must be 4,000 characters or fewer");
  return normalized;
}

const driverStatusTransitions = {
  assigned: ["dispatched", "exception"],
  dispatched: ["en_route", "arrived_pickup", "exception"],
  en_route: ["arrived_pickup", "exception"],
  arrived_pickup: ["loading", "picked_up", "exception"],
  loading: ["picked_up", "exception"],
  picked_up: ["in_transit", "arrived_delivery", "exception"],
  in_transit: ["arrived_delivery", "exception"],
  arrived_delivery: ["unloading", "delivered", "exception"],
  unloading: ["delivered", "exception"],
  exception: ["en_route", "in_transit", "arrived_delivery", "delivered"],
};

export function driverStatusOptions(load) {
  const current = operationalLoadStatus(load);
  return driverStatusTransitions[current] || [];
}

export function carrierStatusHelp(load) {
  const current = operationalLoadStatus(load);
  if (current === "delivered") return "Delivery completed. No further carrier status updates are required.";
  if (current === "cancelled") return "This load was cancelled. No carrier status updates are available.";
  return "Dispatch must move this load to Assigned before carrier status updates begin.";
}

export function validateDriverProblemReport({ category, severity, description }) {
  const normalized = {
    category: String(category || "other").trim(),
    severity: String(severity || "info").trim(),
    description: String(description || "").trim(),
  };
  if (!["delay", "damage", "detention", "paperwork", "safety", "other"].includes(normalized.category)) {
    throw new Error("Choose a valid problem category");
  }
  if (!["info", "warning", "critical"].includes(normalized.severity)) {
    throw new Error("Choose a valid problem severity");
  }
  if (normalized.description.length < 3) throw new Error("Describe the problem before sending");
  if (normalized.description.length > 1200) throw new Error("Problem report must be 1,200 characters or fewer");
  return normalized;
}

export function validatePublicTrackingLinkInput({ label, expiresAt, privacy }, now = new Date()) {
  const normalized = {
    label: String(label || "").trim(),
    expiresAt: String(expiresAt || "").trim(),
    privacy: {
      show_status: Boolean(privacy?.show_status),
      show_eta: Boolean(privacy?.show_eta),
      show_location: Boolean(privacy?.show_location),
      show_driver: Boolean(privacy?.show_driver),
      show_equipment: Boolean(privacy?.show_equipment),
    },
  };
  if (normalized.label.length > 120) throw new Error("Tracking link label must be 120 characters or fewer");
  const expires = Date.parse(normalized.expiresAt);
  if (Number.isNaN(expires)) throw new Error("Choose a valid tracking link expiration");
  if (expires <= now.getTime()) throw new Error("Tracking link expiration must be in the future");
  if (expires > now.getTime() + 30 * 86400000) throw new Error("Tracking link cannot expire more than 30 days from now");
  if (!Object.values(normalized.privacy).some(Boolean)) throw new Error("Enable at least one public tracking field");
  return normalized;
}

export function validateDuplicateLoadInput({ loadNo, pickupDate, deliveryDate }) {
  const normalizedLoadNo = String(loadNo ?? "").trim();
  if (!normalizedLoadNo) throw new Error("New load number is required");
  if (pickupDate && deliveryDate && deliveryDate < pickupDate) throw new Error("Delivery date cannot be before pickup date");
  return { loadNo: normalizedLoadNo, pickupDate: pickupDate || null, deliveryDate: deliveryDate || null };
}

export function validateStopWindows({ appointmentFrom, appointmentTo, actualArrival, actualDeparture }) {
  if (appointmentFrom && appointmentTo && Date.parse(appointmentTo) < Date.parse(appointmentFrom)) {
    throw new Error("Appointment end cannot be before appointment start");
  }
  if (actualArrival && actualDeparture && Date.parse(actualDeparture) < Date.parse(actualArrival)) {
    throw new Error("Actual departure cannot be before actual arrival");
  }
}

export function validateLoadOperationalDetails(details) {
  const nonnegative = ["weight_lbs", "pieces", "pallets", "loaded_miles", "empty_miles"];
  for (const field of nonnegative) {
    if (details[field] !== null && details[field] !== undefined && details[field] !== "" && Number(details[field]) < 0) {
      throw new Error(`${field.replaceAll("_", " ")} cannot be negative`);
    }
  }
  if (details.temperature_min_f !== null && details.temperature_min_f !== undefined && details.temperature_min_f !== ""
      && details.temperature_max_f !== null && details.temperature_max_f !== undefined && details.temperature_max_f !== ""
      && Number(details.temperature_max_f) < Number(details.temperature_min_f)) {
    throw new Error("Maximum temperature cannot be below minimum temperature");
  }
  return details;
}

export function loadTotalMiles(load) {
  return numeric(load.loaded_miles) + numeric(load.empty_miles);
}

export function integrationHealthSummary(connections = [], events = [], rateLimits = []) {
  const deadLetter = events.filter((event) => event.status === "dead_letter").length;
  const retry = events.filter((event) => event.status === "retry").length;
  const erroredConnections = connections.filter((connection) => connection.status === "error" || connection.credential_status === "error").length;
  const expiredCredentials = connections.filter((connection) => ["expired", "revoked"].includes(connection.credential_status)).length;
  const exhaustedBuckets = rateLimits.filter((bucket) => numeric(bucket.used_count) >= numeric(bucket.limit_count)).length;
  const state = deadLetter || erroredConnections || expiredCredentials
    ? "error"
    : retry || exhaustedBuckets
      ? "warning"
      : "ok";
  return {
    state,
    activeConnections: connections.filter((connection) => connection.status === "active").length,
    totalConnections: connections.length,
    retry,
    deadLetter,
    erroredConnections,
    expiredCredentials,
    exhaustedBuckets,
  };
}

export const notificationEventTypes = [
  "assignment",
  "load_message",
  "missing_pod",
  "late_load",
  "compliance_expiry",
  "maintenance_expiry",
  "unpaid_invoice",
  "system",
];

export function notificationSummary(notifications = []) {
  const active = notifications.filter((notification) => notification.status !== "archived");
  return {
    total: active.length,
    unread: active.filter((notification) => notification.status === "unread").length,
    critical: active.filter((notification) => notification.severity === "critical").length,
    warning: active.filter((notification) => notification.severity === "warning").length,
  };
}

export function shouldDeliverNotification(preferences = [], eventType, channel = "in_app", branchId = null) {
  const field = `${channel}_enabled`;
  const matches = preferences.filter((preference) => preference.event_type === eventType);
  const branchPreference = matches.find((preference) => preference.branch_id === branchId);
  const globalPreference = matches.find((preference) => !preference.branch_id);
  const preference = branchPreference || globalPreference;
  return preference?.[field] ?? channel === "in_app";
}

export function notificationDeliverySummary(deliveryLogs = []) {
  const counts = deliveryLogs.reduce((summary, log) => {
    summary[log.status] = (summary[log.status] || 0) + 1;
    return summary;
  }, {});
  const retry = (counts.retry || 0) + (counts.queued || 0);
  const failed = counts.failed || 0;
  return {
    state: failed ? "error" : retry ? "warning" : "ok",
    queued: counts.queued || 0,
    sent: counts.sent || 0,
    retry: counts.retry || 0,
    failed,
    skipped: counts.skipped || 0,
  };
}

export function telematicsFreshness(record, nowIso, staleMinutes = 15, offlineMinutes = 60) {
  if (!record?.event_at) return { state: "missing", ageMinutes: null, label: "No live signal" };
  const now = Date.parse(nowIso || new Date().toISOString());
  const eventAt = Date.parse(record.event_at);
  if (Number.isNaN(now) || Number.isNaN(eventAt)) return { state: "unknown", ageMinutes: null, label: "Signal time unknown" };
  const ageMinutes = Math.max(0, Math.floor((now - eventAt) / 60000));
  const state = ageMinutes <= staleMinutes ? "current" : ageMinutes <= offlineMinutes ? "stale" : "offline";
  const label = state === "current" ? "Live" : state === "stale" ? "Stale signal" : "Offline";
  return { state, ageMinutes, label };
}

export function latestTelematicsRecords(records = [], identityFields = []) {
  const sorted = [...records].sort((left, right) => {
    const leftTime = Date.parse(left?.event_at || "");
    const rightTime = Date.parse(right?.event_at || "");
    return (Number.isNaN(rightTime) ? 0 : rightTime) - (Number.isNaN(leftTime) ? 0 : leftTime);
  });
  const seen = new Set();

  return sorted.filter((record, index) => {
    const identityField = identityFields.find((field) => record?.[field] !== null && record?.[field] !== undefined && record?.[field] !== "");
    const identity = identityField ? `${identityField}:${record[identityField]}` : `record:${record?.id || index}`;
    if (seen.has(identity)) return false;
    seen.add(identity);
    return true;
  });
}

export function hosRiskState(snapshot, warningSeconds = 3600) {
  if (!snapshot) return { state: "unknown", label: "HOS unavailable", minimumSeconds: null };
  const values = [snapshot.drive_seconds_remaining, snapshot.shift_seconds_remaining, snapshot.cycle_seconds_remaining]
    .filter((value) => value !== null && value !== undefined && value !== "")
    .map(Number)
    .filter(Number.isFinite);
  const minimumSeconds = values.length ? Math.min(...values) : null;
  if (snapshot.violation || minimumSeconds === 0) return { state: "violation", label: "HOS violation", minimumSeconds };
  if (minimumSeconds !== null && minimumSeconds <= warningSeconds) return { state: "warning", label: "HOS low", minimumSeconds };
  return { state: "ok", label: "HOS ok", minimumSeconds };
}

export function nextOperationalStop(load) {
  return [...(load?.load_stops || [])]
    .filter((stop) => !stop.actual_departure)
    .sort((left, right) => numeric(left.stop_order) - numeric(right.stop_order))[0] || null;
}

export function haversineMiles(from, to) {
  const fromLat = Number(from?.latitude);
  const fromLng = Number(from?.longitude);
  const toLat = Number(to?.latitude);
  const toLng = Number(to?.longitude);
  if (![fromLat, fromLng, toLat, toLng].every(Number.isFinite)) return null;
  const radians = (degrees) => degrees * Math.PI / 180;
  const latitudeDelta = radians(toLat - fromLat);
  const longitudeDelta = radians(toLng - fromLng);
  const value = Math.sin(latitudeDelta / 2) ** 2
    + Math.cos(radians(fromLat)) * Math.cos(radians(toLat)) * Math.sin(longitudeDelta / 2) ** 2;
  return 3958.8 * 2 * Math.atan2(Math.sqrt(value), Math.sqrt(1 - value));
}

export function estimateStopArrival({ position, destination, stop, hosState = "ok", now = new Date().toISOString(), geofenceMiles = 0.25, fallbackSpeedMph = 50 }) {
  const distanceMiles = haversineMiles(position, destination);
  if (distanceMiles === null) return { state: "unknown", distanceMiles: null, etaAt: null, deltaMinutes: null, insideGeofence: false };
  const insideGeofence = distanceMiles <= geofenceMiles;
  const reportedSpeed = Number(position?.speed_mph);
  const speedMph = Number.isFinite(reportedSpeed) && reportedSpeed >= 10 ? reportedSpeed : fallbackSpeedMph;
  const travelMinutes = insideGeofence ? 0 : Math.ceil((distanceMiles / speedMph) * 60);
  const nowTimestamp = validTimestamp(now);
  const etaTimestamp = nowTimestamp === null ? null : nowTimestamp + travelMinutes * 60_000;
  const appointmentTimestamp = validTimestamp(stop?.appointment_to) ?? validTimestamp(stop?.appointment_from);
  const deltaMinutes = etaTimestamp === null || appointmentTimestamp === null
    ? null
    : Math.round((etaTimestamp - appointmentTimestamp) / 60_000);
  let state = insideGeofence ? "arrived" : "tracking";
  if (!insideGeofence && deltaMinutes !== null) {
    if (deltaMinutes > 30) state = "late";
    else if (deltaMinutes > 0) state = "at_risk";
    else state = "on_time";
  }
  if (!insideGeofence && hosState === "violation") state = "late";
  else if (!insideGeofence && hosState === "warning" && ["tracking", "on_time"].includes(state)) state = "at_risk";
  return {
    state,
    distanceMiles,
    etaAt: etaTimestamp === null ? null : new Date(etaTimestamp).toISOString(),
    deltaMinutes,
    insideGeofence,
  };
}

export function applyEtaSignalFreshness(prediction, freshness) {
  if (freshness?.state === "offline") {
    return { ...prediction, state: "offline", etaAt: null, deltaMinutes: null };
  }
  if (freshness?.state === "missing") {
    return { ...prediction, state: "unknown", etaAt: null, deltaMinutes: null };
  }
  if (freshness?.state === "stale") return { ...prediction, state: "stale" };
  return prediction;
}

export function exceptionTypeForPrediction(prediction) {
  if (["violation", "warning"].includes(prediction?.hos?.state)) return "hos_risk";
  if (["offline", "stale"].includes(prediction?.state)) return "offline_signal";
  if (["late", "at_risk"].includes(prediction?.state)) return prediction?.stop?.type === "pickup" ? "late_pickup" : "late_delivery";
  if (prediction?.insideGeofence) return "geofence_review";
  return null;
}

export function validateOperationalExceptionInput({ reason, resolution = null, closing = false }) {
  const normalizedReason = String(reason || "").trim();
  const normalizedResolution = String(resolution || "").trim();
  if (normalizedReason.length < 3) throw new Error("Exception reason must be at least 3 characters");
  if (normalizedReason.length > 1000) throw new Error("Exception reason must be 1,000 characters or fewer");
  if (closing && normalizedResolution.length < 3) throw new Error("Resolution must be at least 3 characters");
  if (normalizedResolution.length > 2000) throw new Error("Resolution must be 2,000 characters or fewer");
  return { reason: normalizedReason, resolution: normalizedResolution || null };
}

function validTimestamp(value) {
  const timestamp = value ? Date.parse(value) : Number.NaN;
  return Number.isNaN(timestamp) ? null : timestamp;
}

export function dispatchLoadWindow(load) {
  const stops = Array.isArray(load.load_stops) ? load.load_stops : [];
  const starts = stops.map((stop) => validTimestamp(stop.appointment_from)).filter((value) => value !== null);
  const ends = stops
    .map((stop) => validTimestamp(stop.appointment_to) ?? validTimestamp(stop.appointment_from))
    .filter((value) => value !== null);
  const pickupFallback = validTimestamp(load.pickup_date ? `${load.pickup_date}T00:00:00` : null);
  const deliveryFallback = validTimestamp(load.delivery_date ? `${load.delivery_date}T23:59:59` : null);
  const start = starts.length ? Math.min(...starts) : pickupFallback;
  const end = ends.length ? Math.max(...ends) : (deliveryFallback ?? start);
  return { start, end: end ?? start };
}

export function dispatchAssignmentConflicts(loads) {
  const conflicts = Object.fromEntries(loads.map((load) => [load.id, []]));
  const activeLoads = loads.filter((load) => operationalLoadStatus(load) !== "cancelled");

  for (let index = 0; index < activeLoads.length; index += 1) {
    const current = activeLoads[index];
    const currentWindow = dispatchLoadWindow(current);
    if (currentWindow.start === null || currentWindow.end === null) continue;

    for (let otherIndex = index + 1; otherIndex < activeLoads.length; otherIndex += 1) {
      const other = activeLoads[otherIndex];
      const otherWindow = dispatchLoadWindow(other);
      if (otherWindow.start === null || otherWindow.end === null) continue;
      if (currentWindow.start > otherWindow.end || otherWindow.start > currentWindow.end) continue;

      const shared = [];
      if (current.driver_id && current.driver_id === other.driver_id) shared.push("driver");
      if (current.truck_id && current.truck_id === other.truck_id) shared.push("truck");
      if (!shared.length) continue;

      conflicts[current.id].push({ loadId: other.id, loadNo: other.load_no, resources: shared });
      conflicts[other.id].push({ loadId: current.id, loadNo: current.load_no, resources: shared });
    }
  }

  return conflicts;
}

export function matchesDispatchSavedView(load, view, context = {}) {
  if (!view || view === "all") return true;
  if (view === "my_loads") return Boolean(context.userId) && load.dispatcher_id === context.userId;
  if (view === "unassigned") return !load.driver_id || !load.truck_id;
  if (view === "today_pickup") {
    const pickupStop = (load.load_stops || [])
      .filter((stop) => stop.type === "pickup" && stop.appointment_from)
      .sort((left, right) => numeric(left.stop_order) - numeric(right.stop_order))[0];
    const date = pickupStop?.appointment_from?.slice(0, 10) || load.pickup_date;
    return Boolean(context.today) && date === context.today;
  }
  if (view === "missing_pod") {
    if (operationalLoadStatus(load) !== "delivered") return false;
    return !(context.documents || []).some((document) => document.entity_id === load.id
      && document.doc_type === "pod" && document.status === "approved");
  }
  return true;
}

export function validateProfileFinancialTerms({ paymentTermsDays, payRate }) {
  if (paymentTermsDays !== null && paymentTermsDays !== undefined && paymentTermsDays !== "") {
    const terms = Number(paymentTermsDays);
    if (!Number.isInteger(terms) || terms < 0 || terms > 365) throw new Error("Payment terms must be between 0 and 365 days");
  }
  if (payRate !== null && payRate !== undefined && payRate !== "" && Number(payRate) < 0) throw new Error("Pay rate cannot be negative");
}

export function validateFacilityInput(facility) {
  for (const field of ["name", "address_line1", "city", "state"]) {
    if (!String(facility[field] || "").trim()) throw new Error(`${field.replaceAll("_", " ")} is required`);
  }
  return { ...facility, name: facility.name.trim(), address_line1: facility.address_line1.trim(), city: facility.city.trim(), state: facility.state.trim() };
}

export function hasTmsPermission(roles, permission) {
  const allowed = tmsPermissionRoles[permission] || [];
  return roles.some((role) => allowed.includes(role));
}
