import { supabase } from "./supabase.js";

export const existingTables = [
  "profiles",
  "roles",
  "permissions",
  "user_roles",
  "sites",
  "pages",
  "posts",
  "media",
  "settings",
];

export const tmsTables = [
  "branches",
  "customers",
  "customer_contacts",
  "carriers",
  "drivers",
  "trucks",
  "trailers",
  "loads",
  "load_stops",
  "load_financials",
  "documents",
  "invoices",
  "payments",
  "maintenance",
  "messages",
  "audit_logs",
];

export async function checkTable(table) {
  const { count, error } = await supabase.from(table).select("*", {
    count: "exact",
    head: true,
  });

  return {
    table,
    count: count ?? 0,
    ok: !error,
    message: error?.message ?? "OK",
  };
}

export async function checkExistingTables() {
  return Promise.all(existingTables.map(checkTable));
}

export async function checkTmsTables() {
  return Promise.all(tmsTables.map(checkTable));
}

export async function listTableRows(table, limit = 20) {
  const { data, error } = await supabase.from(table).select("*").limit(limit);

  if (error) {
    throw error;
  }

  return data ?? [];
}

export async function getSession() {
  const { data, error } = await supabase.auth.getSession();

  if (error) {
    throw error;
  }

  return data.session;
}

export async function signIn(email, password) {
  const { data, error } = await supabase.auth.signInWithPassword({
    email,
    password,
  });

  if (error) {
    throw error;
  }

  return data.session;
}

export async function signUp(email, password) {
  const { data, error } = await supabase.auth.signUp({
    email,
    password,
  });

  if (error) {
    throw error;
  }

  return data.session;
}

export async function signOut() {
  const { error } = await supabase.auth.signOut();

  if (error) {
    throw error;
  }
}

export function onAuthStateChange(callback) {
  return supabase.auth.onAuthStateChange((_event, session) => callback(session));
}

export async function listCustomers() {
  const { data, error } = await supabase
    .from("customers")
    .select("id, name, billing_address, credit_status, portal_enabled, status, created_at")
    .order("created_at", { ascending: false });

  if (error) {
    throw error;
  }

  return data ?? [];
}

export async function createCustomer(customer) {
  const { data, error } = await supabase
    .from("customers")
    .insert(customer)
    .select("id, name, billing_address, credit_status, portal_enabled, status, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function updateCustomer(customerId, customer) {
  const { data, error } = await supabase
    .from("customers")
    .update(customer)
    .eq("id", customerId)
    .select("id, name, billing_address, credit_status, portal_enabled, status, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function listCarriers() {
  const { data, error } = await supabase
    .from("carriers")
    .select("id, mc, dot, name, status, safety_rating, insurance_expiry, created_at")
    .order("created_at", { ascending: false });

  if (error) {
    throw error;
  }

  return data ?? [];
}

export async function createCarrier(carrier) {
  const { data, error } = await supabase
    .from("carriers")
    .insert(carrier)
    .select("id, mc, dot, name, status, safety_rating, insurance_expiry, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function updateCarrier(carrierId, carrier) {
  const { data, error } = await supabase
    .from("carriers")
    .update(carrier)
    .eq("id", carrierId)
    .select("id, mc, dot, name, status, safety_rating, insurance_expiry, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function listDrivers() {
  const { data, error } = await supabase
    .from("drivers")
    .select("id, name, phone, email, cdl_no, medical_expiry, status, created_at")
    .order("created_at", { ascending: false });

  if (error) {
    throw error;
  }

  return data ?? [];
}

export async function createDriver(driver) {
  const { data, error } = await supabase
    .from("drivers")
    .insert(driver)
    .select("id, name, phone, email, cdl_no, medical_expiry, status, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function updateDriver(driverId, driver) {
  const { data, error } = await supabase
    .from("drivers")
    .update(driver)
    .eq("id", driverId)
    .select("id, name, phone, email, cdl_no, medical_expiry, status, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function listTrucks() {
  const { data, error } = await supabase
    .from("trucks")
    .select("id, unit_no, vin, plate, make, model, year, status, odometer, created_at")
    .order("created_at", { ascending: false });

  if (error) {
    throw error;
  }

  return data ?? [];
}

export async function createTruck(truck) {
  const { data, error } = await supabase
    .from("trucks")
    .insert(truck)
    .select("id, unit_no, vin, plate, make, model, year, status, odometer, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function updateTruck(truckId, truck) {
  const { data, error } = await supabase
    .from("trucks")
    .update(truck)
    .eq("id", truckId)
    .select("id, unit_no, vin, plate, make, model, year, status, odometer, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function listTrailers() {
  const { data, error } = await supabase
    .from("trailers")
    .select("id, unit_no, vin, plate, type, status, created_at")
    .order("created_at", { ascending: false });

  if (error) {
    throw error;
  }

  return data ?? [];
}

export async function createTrailer(trailer) {
  const { data, error } = await supabase
    .from("trailers")
    .insert(trailer)
    .select("id, unit_no, vin, plate, type, status, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function updateTrailer(trailerId, trailer) {
  const { data, error } = await supabase
    .from("trailers")
    .update(trailer)
    .eq("id", trailerId)
    .select("id, unit_no, vin, plate, type, status, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function listLoads() {
  const { data, error } = await supabase
    .from("loads")
    .select(
      "id, load_no, status, origin, destination, pickup_date, delivery_date, notes, customer_id, carrier_id, driver_id, truck_id, trailer_id, customers(name), carriers(name), drivers(name), trucks(unit_no), trailers(unit_no), created_at"
    )
    .order("created_at", { ascending: false });

  if (error) {
    throw error;
  }

  return data ?? [];
}

export async function createLoad(load) {
  const { data, error } = await supabase
    .from("loads")
    .insert(load)
    .select("id, load_no, status, origin, destination, pickup_date, delivery_date, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function updateLoad(loadId, load) {
  const { data, error } = await supabase
    .from("loads")
    .update(load)
    .eq("id", loadId)
    .select("id, load_no, status, origin, destination, pickup_date, delivery_date, customer_id, carrier_id, driver_id, truck_id, trailer_id, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function updateLoadStatus(loadId, status) {
  const { data, error } = await supabase
    .from("loads")
    .update({ status })
    .eq("id", loadId)
    .select("id, load_no, status")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function getLoad(loadId) {
  const { data, error } = await supabase
    .from("loads")
    .select(
      "id, load_no, status, origin, destination, pickup_date, delivery_date, notes, customer_id, carrier_id, driver_id, truck_id, trailer_id, customers(name), carriers(name), drivers(name), trucks(unit_no), trailers(unit_no), created_at"
    )
    .eq("id", loadId)
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function listLoadStops(loadId) {
  const { data, error } = await supabase
    .from("load_stops")
    .select(
      "id, load_id, stop_order, type, facility, address, contact_name, contact_phone, appointment_from, appointment_to, actual_arrival, actual_departure, instructions, reference, created_at"
    )
    .eq("load_id", loadId)
    .order("stop_order", { ascending: true })
    .order("created_at", { ascending: true });

  if (error) {
    throw error;
  }

  return data ?? [];
}

export async function createLoadStop(stop) {
  const { data, error } = await supabase
    .from("load_stops")
    .insert(stop)
    .select(
      "id, load_id, stop_order, type, facility, address, contact_name, contact_phone, appointment_from, appointment_to, actual_arrival, actual_departure, instructions, reference, created_at"
    )
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function getLoadFinancials(loadId) {
  const { data, error } = await supabase
    .from("load_financials")
    .select("load_id, linehaul, fsc, accessorials, deductions, carrier_cost, driver_pay, margin, locked, created_at, updated_at")
    .eq("load_id", loadId)
    .maybeSingle();

  if (error) {
    throw error;
  }

  return data;
}

export async function saveLoadFinancials(financials) {
  const { data, error } = await supabase
    .from("load_financials")
    .upsert(financials, { onConflict: "load_id" })
    .select("load_id, linehaul, fsc, accessorials, deductions, carrier_cost, driver_pay, margin, locked, created_at, updated_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function listLoadDocuments(loadId) {
  const { data, error } = await supabase
    .from("documents")
    .select("id, entity_type, entity_id, doc_type, file_url, status, uploaded_by, ocr_text, created_at")
    .eq("entity_type", "load")
    .eq("entity_id", loadId)
    .order("created_at", { ascending: false });

  if (error) {
    throw error;
  }

  return data ?? [];
}

export async function createLoadDocument(document) {
  const { data, error } = await supabase
    .from("documents")
    .insert({
      ...document,
      entity_type: "load",
    })
    .select("id, entity_type, entity_id, doc_type, file_url, status, uploaded_by, ocr_text, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function listDocuments() {
  const { data, error } = await supabase
    .from("documents")
    .select("id, entity_type, entity_id, doc_type, file_url, status, uploaded_by, ocr_text, created_at")
    .order("created_at", { ascending: false });

  if (error) {
    throw error;
  }

  return data ?? [];
}

export async function createDocument(document) {
  const { data, error } = await supabase
    .from("documents")
    .insert(document)
    .select("id, entity_type, entity_id, doc_type, file_url, status, uploaded_by, ocr_text, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function listInvoices() {
  const { data, error } = await supabase
    .from("invoices")
    .select("id, invoice_no, customer_id, load_id, amount, status, due_date, sent_at, created_at, customers(name), loads(load_no)")
    .order("created_at", { ascending: false });

  if (error) {
    throw error;
  }

  return data ?? [];
}

export async function createInvoice(invoice) {
  const { data, error } = await supabase
    .from("invoices")
    .insert(invoice)
    .select("id, invoice_no, customer_id, load_id, amount, status, due_date, sent_at, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function updateInvoiceStatus(invoiceId, status) {
  const updates = {
    status,
    sent_at: status === "sent" ? new Date().toISOString() : undefined,
  };

  const { data, error } = await supabase
    .from("invoices")
    .update(updates)
    .eq("id", invoiceId)
    .select("id, invoice_no, status")
    .single();

  if (error) {
    throw error;
  }

  return data;
}
