import { supabase } from "./supabase.js";

const documentBucket = "tms-documents";
const maxDocumentSize = 15 * 1024 * 1024;
const allowedDocumentTypes = new Set(["application/pdf", "image/jpeg", "image/png"]);

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
  "companies",
  "branches",
  "customers",
  "customer_contacts",
  "carriers",
  "drivers",
  "trucks",
  "trailers",
  "loads",
  "customer_rate_agreements",
  "freight_quotes",
  "recurring_load_templates",
  "saved_report_views",
  "table_saved_views",
  "scheduled_reports",
  "report_delivery_logs",
  "load_stops",
  "load_financials",
  "documents",
  "invoices",
  "invoice_line_items",
  "invoice_credit_memos",
  "invoice_batches",
  "invoice_batch_items",
  "payments",
  "settlements",
  "settlement_items",
  "settlement_pay_batches",
  "settlement_pay_batch_items",
  "maintenance",
  "messages",
  "telematics_position_events",
  "telematics_hos_snapshots",
  "notifications",
  "notification_preferences",
  "notification_delivery_logs",
  "driver_problem_reports",
  "public_tracking_links",
  "public_tracking_access_logs",
  "customer_portal_access",
  "customer_shipment_requests",
  "customer_portal_messages",
  "carrier_portal_access",
  "carrier_load_tenders",
  "carrier_load_bids",
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

export async function listBranches() {
  const { data, error } = await supabase
    .from("branches")
    .select("id,company_id,name,company_name,address,timezone,status,branch_code,terminal_type")
    .order("name");
  if (error) throw error;
  return data ?? [];
}

export async function listCompanyTenants() {
  const { data, error } = await supabase
    .from("companies")
    .select("id,legal_name,dba_name,status,timezone,locale,country,created_at,updated_at,branches(id,name,company_name,address,timezone,status,branch_code,terminal_type)")
    .order("legal_name");
  if (error) throw error;
  return data ?? [];
}

export async function listTableSavedViews(tableType) {
  const { data, error } = await supabase
    .from("table_saved_views")
    .select("id,branch_id,table_type,visibility,name,settings,created_by,created_at,updated_at,branches(name)")
    .eq("table_type", tableType)
    .order("visibility", { ascending: true })
    .order("name", { ascending: true });
  if (error) throw error;
  return data ?? [];
}

export async function saveTableSavedView(view) {
  const { data, error } = await supabase
    .from("table_saved_views")
    .upsert(view, { onConflict: "branch_id,table_type,created_by,name" })
    .select("id,branch_id,table_type,visibility,name,settings,created_by,created_at,updated_at,branches(name)")
    .single();
  if (error) throw error;
  return data;
}

export async function deleteTableSavedView(viewId) {
  const { error } = await supabase.from("table_saved_views").delete().eq("id", viewId);
  if (error) throw error;
}

export async function updateCompanyTenantProfile(companyId, profile) {
  const { data, error } = await supabase.rpc("update_company_profile", {
    target_company: companyId,
    profile_payload: profile,
  });
  if (error) throw error;
  return data;
}

export async function listTableRows(table, limit = 20) {
  const { data, error } = await supabase.from(table).select("*").limit(limit);

  if (error) {
    throw error;
  }

  return data ?? [];
}

export async function manageTmsUsers(action, payload = {}) {
  const { data, error } = await supabase.functions.invoke("manage-tms-users", {
    body: { action, ...payload },
  });

  if (error) {
    let message = error.message;

    try {
      const details = await error.context?.json();
      message = details?.error || message;
    } catch {
      // Keep the Supabase client message when the response is not JSON.
    }

    throw new Error(message);
  }

  if (data?.error) {
    throw new Error(data.error);
  }

  return data;
}

export async function generateMockTelematics(branchId, target = {}) {
  const { data, error } = await supabase.functions.invoke("mock-telematics", {
    body: { branchId, ...target },
  });
  if (error) {
    let message = error.message;
    try {
      const details = await error.context?.json();
      message = details?.error || message;
    } catch {
      // Keep the Supabase client message when the response is not JSON.
    }
    throw new Error(message);
  }
  if (data?.error) throw new Error(data.error);
  return data;
}

export async function getCompanyBillingSettings() {
  const { data, error } = await supabase
    .from("company_billing_settings")
    .select("legal_name,address_line_1,address_line_2,city,state,postal_code,country,phone,email,tax_id,payment_instructions,invoice_number_prefix,invoice_number_next,invoice_number_padding,updated_at")
    .eq("id", true)
    .single();
  if (error) throw error;
  return data;
}

export async function updateCompanyBillingSettings(settings) {
  const { data, error } = await supabase
    .from("company_billing_settings")
    .upsert({ id: true, ...settings })
    .select("legal_name,address_line_1,address_line_2,city,state,postal_code,country,phone,email,tax_id,payment_instructions,invoice_number_prefix,invoice_number_next,invoice_number_padding,updated_at")
    .single();
  if (error) throw error;
  return data;
}

export async function listIntegrationConnections() {
  const { data, error } = await supabase
    .from("integration_connections")
    .select("id,branch_id,provider,integration_type,display_name,status,credential_status,credential_hint,last_sync_at,last_success_at,last_error_at,last_error,updated_at,branches(name)")
    .order("updated_at", { ascending: false });
  if (error) throw error;
  return data ?? [];
}

export async function listIntegrationWebhookEvents(limit = 20) {
  const { data, error } = await supabase
    .from("integration_webhook_events")
    .select("id,branch_id,connection_id,provider,event_type,idempotency_key,external_event_id,status,received_at,processed_at,next_retry_at,retry_count,max_retries,last_error")
    .order("received_at", { ascending: false })
    .limit(limit);
  if (error) throw error;
  return data ?? [];
}

export async function listIntegrationExternalIds(limit = 50) {
  const { data, error } = await supabase
    .from("integration_external_id_map")
    .select("id,branch_id,connection_id,provider,entity_type,entity_id,external_id,last_seen_at,updated_at")
    .order("updated_at", { ascending: false })
    .limit(limit);
  if (error) throw error;
  return data ?? [];
}

export async function listIntegrationRateLimits(limit = 20) {
  const { data, error } = await supabase
    .from("integration_rate_limits")
    .select("id,branch_id,connection_id,provider,bucket_key,window_start,window_seconds,limit_count,used_count,reset_at,updated_at")
    .order("reset_at", { ascending: false })
    .limit(limit);
  if (error) throw error;
  return data ?? [];
}

export async function listTelematicsPositions(limit = 100) {
  const { data, error } = await supabase
    .from("telematics_position_events")
    .select("id,branch_id,connection_id,provider,truck_id,driver_id,load_id,external_vehicle_id,latitude,longitude,speed_mph,heading_degrees,odometer_miles,accuracy_meters,ignition_status,event_at,received_at,trucks(unit_no),drivers(name),loads(load_no,origin,destination,operational_status,billing_status,status)")
    .order("event_at", { ascending: false })
    .limit(limit);
  if (error) throw error;
  return data ?? [];
}

export async function listTelematicsHosSnapshots(limit = 100) {
  const { data, error } = await supabase
    .from("telematics_hos_snapshots")
    .select("id,branch_id,connection_id,provider,driver_id,external_driver_id,duty_status,drive_seconds_remaining,shift_seconds_remaining,cycle_seconds_remaining,break_seconds_remaining,violation,event_at,received_at,drivers(name)")
    .order("event_at", { ascending: false })
    .limit(limit);
  if (error) throw error;
  return data ?? [];
}

export async function listOperationalExceptions(limit = 100) {
  const { data, error } = await supabase
    .from("operational_exceptions")
    .select("id,branch_id,load_id,stop_id,exception_type,severity,status,source,reason,owner_id,created_by,acknowledged_at,acknowledged_by,resolved_at,resolved_by,resolution,detected_at,created_at,updated_at")
    .order("detected_at", { ascending: false })
    .limit(limit);
  if (error) throw error;
  return data ?? [];
}

export async function createOperationalException(exception) {
  const { data, error } = await supabase
    .from("operational_exceptions")
    .insert(exception)
    .select("id,branch_id,load_id,stop_id,exception_type,severity,status,source,reason,owner_id,detected_at,created_at,updated_at")
    .single();
  if (error) throw error;
  return data;
}

export async function updateOperationalException(exceptionId, updates) {
  const { data, error } = await supabase
    .from("operational_exceptions")
    .update(updates)
    .eq("id", exceptionId)
    .select("id,branch_id,load_id,stop_id,exception_type,severity,status,source,reason,owner_id,acknowledged_at,resolved_at,resolution,updated_at")
    .single();
  if (error) throw error;
  return data;
}

export async function listLoadGeofenceEvents(limit = 100) {
  const { data, error } = await supabase
    .from("load_geofence_events")
    .select("id,branch_id,load_id,stop_id,event_type,status,source_position_id,distance_miles,occurred_at,created_by,confirmed_at,confirmed_by,confirmation_note,created_at,updated_at")
    .order("occurred_at", { ascending: false })
    .limit(limit);
  if (error) throw error;
  return data ?? [];
}

export async function createLoadGeofenceEvent(event) {
  const { data, error } = await supabase
    .from("load_geofence_events")
    .insert(event)
    .select("id,branch_id,load_id,stop_id,event_type,status,source_position_id,distance_miles,occurred_at,created_at")
    .single();
  if (error) throw error;
  return data;
}

export async function confirmLoadGeofenceEvent(eventId, decision, note) {
  const { error } = await supabase.rpc("confirm_tms_geofence_event", {
    target_event_id: eventId,
    target_decision: decision,
    target_note: note || null,
  });
  if (error) throw error;
}

export async function updateIntegrationWebhookStatus(eventId, status) {
  const { data, error } = await supabase
    .from("integration_webhook_events")
    .update({
      status,
      next_retry_at: status === "retry" ? new Date().toISOString() : null,
      last_error: status === "retry" ? "Queued manually from integration health" : null,
    })
    .eq("id", eventId)
    .select("id,status,next_retry_at,last_error")
    .single();
  if (error) throw error;
  return data;
}

export async function listNotifications(limit = 100) {
  const { data, error } = await supabase
    .from("notifications")
    .select("id,branch_id,recipient_id,actor_id,event_type,severity,title,body,entity_type,entity_id,action_url,status,read_at,archived_at,expires_at,metadata,created_at,updated_at")
    .order("created_at", { ascending: false })
    .limit(limit);
  if (error) throw error;
  return data ?? [];
}

export async function listNotificationPreferences(userId) {
  const { data, error } = await supabase
    .from("notification_preferences")
    .select("id,user_id,branch_id,event_type,in_app_enabled,email_enabled,sms_enabled,push_enabled,quiet_hours_start,quiet_hours_end,created_at,updated_at")
    .eq("user_id", userId)
    .order("event_type");
  if (error) throw error;
  return data ?? [];
}

export async function saveNotificationPreference(preference) {
  const fields = "id,user_id,branch_id,event_type,in_app_enabled,email_enabled,sms_enabled,push_enabled,quiet_hours_start,quiet_hours_end,updated_at";
  const request = preference.id
    ? supabase.from("notification_preferences").update(preference).eq("id", preference.id)
    : supabase.from("notification_preferences").insert(preference);
  const { data, error } = await request.select(fields).single();
  if (error) throw error;
  return data;
}

export async function listNotificationDeliveryLogs(limit = 100) {
  const { data, error } = await supabase
    .from("notification_delivery_logs")
    .select("id,notification_id,branch_id,channel,provider,status,attempt_count,next_retry_at,sent_at,last_error,created_at,updated_at")
    .order("created_at", { ascending: false })
    .limit(limit);
  if (error) throw error;
  return data ?? [];
}

export async function markNotificationsRead(notificationIds) {
  const { data, error } = await supabase.rpc("mark_tms_notifications_read", {
    target_notification_ids: notificationIds,
  });
  if (error) throw error;
  return data ?? 0;
}

export async function createNotification(notification) {
  const { data, error } = await supabase.rpc("create_tms_notification", {
    target_branch_id: notification.branchId,
    target_recipient_id: notification.recipientId || null,
    target_event_type: notification.eventType,
    target_severity: notification.severity || "info",
    target_title: notification.title,
    target_body: notification.body,
    target_entity_type: notification.entityType || null,
    target_entity_id: notification.entityId || null,
    target_action_url: notification.actionUrl || null,
    target_metadata: notification.metadata || {},
  });
  if (error) throw error;
  return data;
}

export async function sendTmsInvoice(invoiceId, recipient) {
  const { data, error } = await supabase.functions.invoke("send-tms-invoice", { body: { invoiceId, recipient } });
  if (error) {
    let message = error.message;
    try {
      const details = await error.context?.json();
      message = details?.error || message;
    } catch {
      // Keep the Supabase client message when the response is not JSON.
    }
    throw new Error(message);
  }
  if (data?.error) throw new Error(data.error);
  return data;
}

export async function sendTmsInvoiceBatch(batchId, recipient) {
  const { data, error } = await supabase.functions.invoke("send-tms-invoice", { body: { batchId, recipient } });
  if (error) {
    let message = error.message;
    try {
      const details = await error.context?.json();
      message = details?.error || message;
    } catch {
      // Keep the Supabase client message when the response is not JSON.
    }
    throw new Error(message);
  }
  if (data?.error) throw new Error(data.error);
  return data;
}

export async function getSession() {
  const { data, error } = await supabase.auth.getSession();

  if (error) {
    throw error;
  }

  return data.session;
}

export async function getCurrentTmsAccess(userId) {
  const { data, error } = await supabase
    .from("tms_user_access")
    .select("id,user_id,branch_id,role,status,branches(name)")
    .eq("user_id", userId)
    .eq("status", "active");
  if (error) throw error;
  return data ?? [];
}

export async function getCurrentCustomerPortalAccess(userId) {
  const { data, error } = await supabase
    .from("customer_portal_access")
    .select("id,user_id,branch_id,customer_id,status,customers(name)")
    .eq("user_id", userId)
    .eq("status", "active");
  if (error) throw error;
  return data ?? [];
}

export async function listCustomerShipmentRequests(){const {data,error}=await supabase.from("customer_shipment_requests").select("id,branch_id,customer_id,request_no,request_type,status,origin,destination,pickup_date,delivery_date,equipment_type,commodity,loaded_miles,notes,linked_quote_id,linked_load_id,created_at,updated_at,customers(name)").order("created_at",{ascending:false});if(error)throw error;return data??[];}
export async function createCustomerShipmentRequest(request){const {data,error}=await supabase.from("customer_shipment_requests").insert(request).select().single();if(error)throw error;return data;}
export async function updateCustomerShipmentRequest(id,changes){const {data,error}=await supabase.from("customer_shipment_requests").update(changes).eq("id",id).select().single();if(error)throw error;return data;}
export async function listCustomerPortalMessages(){const {data,error}=await supabase.from("customer_portal_messages").select("id,branch_id,customer_id,request_id,load_id,sender_id,body,created_at").order("created_at");if(error)throw error;return data??[];}
export async function createCustomerPortalMessage(message){const {data,error}=await supabase.from("customer_portal_messages").insert(message).select().single();if(error)throw error;return data;}

export async function getCurrentCarrierPortalAccess(userId) {
  const { data, error } = await supabase.from("carrier_portal_access")
    .select("id,user_id,branch_id,carrier_id,status,carriers(name)")
    .eq("user_id", userId).eq("status", "active");
  if (error) throw error;
  return data ?? [];
}

export async function listCarrierPortalTenders() {
  const { data, error } = await supabase.from("carrier_load_tenders")
    .select("id,load_id,carrier_id,status,offered_at,expires_at,responded_at,response_note,loads(load_no,origin,destination,pickup_date,delivery_date)")
    .order("offered_at", { ascending: false });
  if (error) throw error;
  return data ?? [];
}

export async function listCarrierTenders(loadId) {
  const { data, error } = await supabase.from("carrier_load_tenders")
    .select("id,load_id,carrier_id,status,offered_at,expires_at,responded_at,response_note,carriers(name)")
    .eq("load_id", loadId).order("offered_at", { ascending: false });
  if (error) throw error;
  return data ?? [];
}

export async function createCarrierTender(tender) {
  const { data, error } = await supabase.from("carrier_load_tenders")
    .upsert(tender, { onConflict: "load_id,carrier_id" })
    .select("id,load_id,carrier_id,status,offered_at,expires_at,responded_at,response_note").single();
  if (error) throw error;
  return data;
}

export async function respondCarrierTender(tenderId, status, note = null) {
  const { data, error } = await supabase.rpc("carrier_portal_respond_tender", { target_tender_id: tenderId, target_status: status, target_note: note });
  if (error) throw error;
  return data;
}

export async function listCarrierLoadBids(loadId = null) {
  let query = supabase.from("carrier_load_bids")
    .select("id,branch_id,load_id,carrier_id,status,target_amount,bid_amount,transit_days,expires_at,carrier_note,internal_note,submitted_at,selected_at,created_at,carriers(name),loads(load_no,origin,destination,load_financials(linehaul,fsc,accessorials,deductions))")
    .order("created_at", { ascending: false });
  if (loadId) query = query.eq("load_id", loadId);
  const { data, error } = await query;
  if (error) throw error;
  return data ?? [];
}

export async function createCarrierLoadBid(bid) {
  const { data, error } = await supabase.from("carrier_load_bids").upsert(bid, { onConflict: "load_id,carrier_id" }).select().single();
  if (error) throw error;
  return data;
}

export async function submitCarrierLoadBid(bidId, amount, transitDays, note = null) {
  const { data, error } = await supabase.rpc("carrier_portal_submit_bid", { target_bid_id: bidId, target_amount: amount, target_transit_days: transitDays, target_note: note });
  if (error) throw error;
  return data;
}

export async function selectCarrierLoadBid(bidId) {
  const { data, error } = await supabase.rpc("select_carrier_load_bid", { target_bid_id: bidId });
  if (error) throw error;
  return data;
}

export async function carrierPortalUpdateLoadStatus(loadId, status) {
  const { data, error } = await supabase.rpc("carrier_portal_update_load_status", { target_load_id: loadId, target_status: status });
  if (error) throw error;
  return data;
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
    .select("id, branch_id, name, billing_address, billing_email, billing_phone, payment_terms_days, billing_cycle, billing_cycle_day, notes, credit_status, portal_enabled, status, created_at")
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
    .select("id, name, billing_address, billing_email, billing_phone, payment_terms_days, billing_cycle, billing_cycle_day, notes, credit_status, portal_enabled, status, created_at")
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
    .select("id, name, billing_address, billing_email, billing_phone, payment_terms_days, billing_cycle, billing_cycle_day, notes, credit_status, portal_enabled, status, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function listCustomerContacts(customerId) {
  const { data, error } = await supabase
    .from("customer_contacts")
    .select("id,customer_id,name,email,phone,role,created_at")
    .eq("customer_id", customerId)
    .order("name");
  if (error) throw error;
  return data ?? [];
}

export async function createCustomerContact(contact) {
  const { data, error } = await supabase
    .from("customer_contacts")
    .insert(contact)
    .select("id,customer_id,name,email,phone,role,created_at")
    .single();
  if (error) throw error;
  return data;
}

export async function listFacilities() {
  const { data, error } = await supabase
    .from("facilities")
    .select("id,customer_id,name,facility_type,address_line1,address_line2,city,state,postal_code,country,contact_name,contact_phone,contact_email,hours,instructions,status,customers(name),created_at")
    .order("name");
  if (error) throw error;
  return data ?? [];
}

export async function createFacility(facility) {
  const { data, error } = await supabase
    .from("facilities")
    .insert(facility)
    .select("id,customer_id,name,facility_type,address_line1,address_line2,city,state,postal_code,country,contact_name,contact_phone,contact_email,hours,instructions,status,created_at")
    .single();
  if (error) throw error;
  return data;
}

export async function updateFacility(facilityId, facility) {
  const { data, error } = await supabase
    .from("facilities")
    .update(facility)
    .eq("id", facilityId)
    .select("id,customer_id,name,facility_type,address_line1,address_line2,city,state,postal_code,country,contact_name,contact_phone,contact_email,hours,instructions,status,created_at")
    .single();
  if (error) throw error;
  return data;
}

export async function listCarriers() {
  const { data, error } = await supabase
    .from("carriers")
    .select("id, mc, dot, name, status, safety_rating, insurance_expiry, contact_name, contact_email, contact_phone, payment_terms_days, w9_status, pay_type, pay_rate, notes, created_at")
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
    .select("id, mc, dot, name, status, safety_rating, insurance_expiry, contact_name, contact_email, contact_phone, payment_terms_days, w9_status, pay_type, pay_rate, notes, created_at")
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
    .select("id, mc, dot, name, status, safety_rating, insurance_expiry, contact_name, contact_email, contact_phone, payment_terms_days, w9_status, pay_type, pay_rate, notes, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function listDrivers() {
  const { data, error } = await supabase
    .from("drivers")
    .select("id, branch_id, auth_user_id, name, phone, email, address, hire_date, emergency_contact_name, emergency_contact_phone, emergency_contact_relation, cdl_no, cdl_expiry, medical_expiry, pay_type, pay_rate, status, created_at, branches(name)")
    .order("created_at", { ascending: false });

  if (error) {
    throw error;
  }

  return data ?? [];
}

export async function getCurrentDriverProfile(userId, email) {
  if (!userId && !email) return null;

  const fields = "id,branch_id,auth_user_id,name,phone,email,status,cdl_expiry,medical_expiry,branches(name)";
  if (userId) {
    const { data, error } = await supabase
      .from("drivers")
      .select(fields)
      .eq("auth_user_id", userId)
      .maybeSingle();
    if (error) throw error;
    if (data) return data;
  }

  if (email) {
    const { data, error } = await supabase
      .from("drivers")
      .select(fields)
      .ilike("email", email)
      .limit(1);
    if (error) throw error;
    return data?.[0] || null;
  }

  return null;
}

export async function createDriver(driver) {
  const { data, error } = await supabase
    .from("drivers")
    .insert(driver)
    .select("id, branch_id, auth_user_id, name, phone, email, address, hire_date, emergency_contact_name, emergency_contact_phone, emergency_contact_relation, cdl_no, cdl_expiry, medical_expiry, pay_type, pay_rate, status, created_at, branches(name)")
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
    .select("id, branch_id, auth_user_id, name, phone, email, address, hire_date, emergency_contact_name, emergency_contact_phone, emergency_contact_relation, cdl_no, cdl_expiry, medical_expiry, pay_type, pay_rate, status, created_at, branches(name)")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function listTrucks() {
  const { data, error } = await supabase
    .from("trucks")
    .select("id, unit_no, vin, plate, make, model, year, status, odometer, registration_expiry, insurance_expiry, created_at")
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
    .select("id, unit_no, vin, plate, make, model, year, status, odometer, registration_expiry, insurance_expiry, created_at")
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
    .select("id, unit_no, vin, plate, make, model, year, status, odometer, registration_expiry, insurance_expiry, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function listTrailers() {
  const { data, error } = await supabase
    .from("trailers")
    .select("id, unit_no, vin, plate, type, status, registration_expiry, insurance_expiry, created_at")
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
    .select("id, unit_no, vin, plate, type, status, registration_expiry, insurance_expiry, created_at")
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
    .select("id, unit_no, vin, plate, type, status, registration_expiry, insurance_expiry, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function listMaintenance() {
  const { data, error } = await supabase
    .from("maintenance")
    .select("id,equipment_type,equipment_id,service_type,due_date,due_miles,scheduled_date,completed_date,vendor,cost,service_reference,out_of_service,status,notes,created_at,updated_at")
    .order("created_at", { ascending: false });
  if (error) throw error;
  return data ?? [];
}

export async function createMaintenance(maintenance) {
  const { data, error } = await supabase
    .from("maintenance")
    .insert(maintenance)
    .select("id,equipment_type,equipment_id,service_type,due_date,due_miles,scheduled_date,completed_date,vendor,cost,service_reference,out_of_service,status,notes,created_at,updated_at")
    .single();
  if (error) throw error;
  return data;
}

export async function updateMaintenance(maintenanceId, maintenance) {
  const { data, error } = await supabase
    .from("maintenance")
    .update(maintenance)
    .eq("id", maintenanceId)
    .select("id,equipment_type,equipment_id,service_type,due_date,due_miles,scheduled_date,completed_date,vendor,cost,service_reference,out_of_service,status,notes,created_at,updated_at")
    .single();
  if (error) throw error;
  return data;
}

export async function updateMaintenanceStatus(maintenanceId, status) {
  const { data, error } = await supabase
    .from("maintenance")
    .update({ status })
    .eq("id", maintenanceId)
    .select("id,status,scheduled_date,completed_date,out_of_service")
    .single();
  if (error) throw error;
  return data;
}

export async function listLoads() {
  const { data, error } = await supabase
    .from("loads")
    .select(
      "id, branch_id, load_no, status, operational_status, document_status, billing_status, origin, destination, pickup_date, delivery_date, notes, reference_no, broker_reference, commodity, weight_lbs, pieces, pallets, equipment_type, temperature_min_f, temperature_max_f, hazmat, loaded_miles, empty_miles, internal_notes, customer_notes, customer_id, carrier_id, driver_id, truck_id, trailer_id, dispatcher_id, customers(name), carriers(name), drivers(name), trucks(unit_no), trailers(unit_no), load_stops(id,facility_id,stop_order,type,facility,address,appointment_from,appointment_to,actual_arrival,actual_departure), created_at"
    )
    .order("created_at", { ascending: false });

  if (error) {
    throw error;
  }

  return data ?? [];
}

const loadListSelect = "id, branch_id, load_no, status, operational_status, document_status, billing_status, origin, destination, pickup_date, delivery_date, notes, reference_no, broker_reference, commodity, weight_lbs, pieces, pallets, equipment_type, temperature_min_f, temperature_max_f, hazmat, loaded_miles, empty_miles, internal_notes, customer_notes, customer_id, carrier_id, driver_id, truck_id, trailer_id, dispatcher_id, customers(name), carriers(name), drivers(name), trucks(unit_no), trailers(unit_no), load_stops(id,facility_id,stop_order,type,facility,address,appointment_from,appointment_to,actual_arrival,actual_departure), created_at";

function safeSearchTerm(value) {
  return String(value || "").trim().replace(/[,%()]/g, " ").replace(/\s+/g, " ");
}

async function matchingEntityIds(table, columns, search) {
  if (!search) return [];
  const filter = columns.map((column) => `${column}.ilike.%${search}%`).join(",");
  const { data, error } = await supabase.from(table).select("id").or(filter).limit(50);
  if (error) throw error;
  return (data || []).map((item) => item.id);
}

export async function listLoadsPage(options = {}) {
  const page = Math.max(1, Number(options.page) || 1);
  const pageSize = Math.min(100, Math.max(1, Number(options.pageSize) || 10));
  const sortColumns = new Set(["created_at", "load_no", "pickup_date", "delivery_date", "operational_status"]);
  const sortBy = sortColumns.has(options.sortBy) ? options.sortBy : "created_at";
  const ascending = options.sortDirection === "asc";
  const search = safeSearchTerm(options.search);
  const [customerIds, carrierIds, driverIds, truckIds] = await Promise.all([
    matchingEntityIds("customers", ["name"], search),
    matchingEntityIds("carriers", ["name"], search),
    matchingEntityIds("drivers", ["name"], search),
    matchingEntityIds("trucks", ["unit_no"], search),
  ]);
  const searchParts = ["load_no", "reference_no", "broker_reference", "commodity", "origin", "destination", "operational_status", "billing_status"]
    .map((column) => `${column}.ilike.%${search}%`);
  if (customerIds.length) searchParts.push(`customer_id.in.(${customerIds.join(",")})`);
  if (carrierIds.length) searchParts.push(`carrier_id.in.(${carrierIds.join(",")})`);
  if (driverIds.length) searchParts.push(`driver_id.in.(${driverIds.join(",")})`);
  if (truckIds.length) searchParts.push(`truck_id.in.(${truckIds.join(",")})`);

  let query = supabase.from("loads").select(loadListSelect, { count: "exact" });
  if (search) query = query.or(searchParts.join(","));
  if (options.status) query = query.or(`operational_status.eq.${options.status},status.eq.${options.status}`);
  if (options.customerId) query = query.eq("customer_id", options.customerId);
  if (options.driverId) query = query.eq("driver_id", options.driverId);
  if (options.dateFrom) query = query.gte("pickup_date", options.dateFrom);
  if (options.dateTo) query = query.lte("pickup_date", options.dateTo);
  if (options.view === "unassigned") query = query.or("driver_id.is.null,truck_id.is.null");
  if (options.view === "today") query = query.eq("pickup_date", options.today);
  if (options.view === "in_transit") query = query.eq("operational_status", "in_transit");
  if (options.view === "ready_to_bill") query = query.or("billing_status.eq.ready_to_bill,status.eq.ready_to_bill");
  const start = (page - 1) * pageSize;
  const { data, error, count } = await query.order(sortBy, { ascending, nullsFirst: false }).range(start, start + pageSize - 1);
  if (error) throw error;
  return { records: data || [], count: count || 0, page, pageSize };
}

export async function createLoad(load) {
  const { data, error } = await supabase
    .from("loads")
    .insert(load)
    .select("id, load_no, status, operational_status, document_status, billing_status, origin, destination, pickup_date, delivery_date, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function listCustomerRateAgreements() {
  const { data, error } = await supabase.from("customer_rate_agreements")
    .select("id,branch_id,customer_id,name,origin_pattern,destination_pattern,equipment_type,linehaul_rate,rate_per_mile,fuel_surcharge_percent,minimum_charge,effective_from,effective_to,active,notes,customers(name)")
    .order("created_at", { ascending: false });
  if (error) throw error;
  return data ?? [];
}

export async function createCustomerRateAgreement(agreement) {
  const { data, error } = await supabase.from("customer_rate_agreements").insert(agreement).select().single();
  if (error) throw error;
  return data;
}

export async function listFreightQuotes() {
  const { data, error } = await supabase.from("freight_quotes")
    .select("id,branch_id,quote_no,customer_id,rate_agreement_id,status,origin,destination,pickup_date,delivery_date,equipment_type,commodity,loaded_miles,linehaul,fuel_surcharge,accessorials,discount,customer_total,estimated_carrier_cost,estimated_margin,estimated_margin_percent,valid_until,customer_notes,internal_notes,sent_at,accepted_at,converted_load_id,created_at,customers(name),customer_rate_agreements(name),loads!freight_quotes_converted_load_id_fkey(load_no)")
    .order("created_at", { ascending: false });
  if (error) throw error;
  return data ?? [];
}

export async function createFreightQuote(quote) {
  const { data, error } = await supabase.from("freight_quotes").insert(quote).select().single();
  if (error) throw error;
  return data;
}

export async function updateFreightQuoteStatus(quoteId, status) {
  const { data, error } = await supabase.from("freight_quotes").update({ status }).eq("id", quoteId).select().single();
  if (error) throw error;
  return data;
}

export async function convertFreightQuoteToLoad(quoteId, loadNo) {
  const { data, error } = await supabase.rpc("convert_freight_quote_to_load", { target_quote_id: quoteId, target_load_no: loadNo });
  if (error) throw error;
  return data;
}

export async function listRecurringLoadTemplates() {
  const { data, error } = await supabase.from("recurring_load_templates")
    .select("id,branch_id,customer_id,name,frequency,interval_count,next_run_date,delivery_offset_days,origin,destination,equipment_type,commodity,loaded_miles,linehaul,fuel_surcharge,accessorials,internal_notes,customer_notes,active,last_generated_at,last_generated_load_id,customers(name),loads!recurring_load_templates_last_generated_load_id_fkey(load_no)")
    .order("next_run_date", { ascending: true });
  if (error) throw error;
  return data ?? [];
}

export async function createRecurringLoadTemplate(template) {
  const { data, error } = await supabase.from("recurring_load_templates").insert(template).select().single();
  if (error) throw error;
  return data;
}

export async function generateLoadFromRecurringTemplate(templateId, loadNo, pickupDate = null) {
  const { data, error } = await supabase.rpc("generate_load_from_recurring_template", { target_template_id: templateId, target_load_no: loadNo, target_pickup_date: pickupDate });
  if (error) throw error;
  return data;
}

export async function duplicateLoad({ sourceLoadId, loadNo, pickupDate, deliveryDate, copyStops }) {
  const { data, error } = await supabase.rpc("duplicate_tms_load", {
    source_load_id: sourceLoadId,
    target_load_no: loadNo,
    target_pickup_date: pickupDate,
    target_delivery_date: deliveryDate,
    copy_stops: copyStops,
  });

  if (error) throw error;
  return data;
}

export async function updateLoad(loadId, load) {
  const { data, error } = await supabase
    .from("loads")
    .update(load)
    .eq("id", loadId)
    .select("id, load_no, status, operational_status, document_status, billing_status, origin, destination, pickup_date, delivery_date, customer_id, carrier_id, driver_id, truck_id, trailer_id, created_at")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function updateLoadStatus(loadId, status) {
  const { data, error } = await supabase
    .from("loads")
    .update({ operational_status: status })
    .eq("id", loadId)
    .select("id, load_no, status, operational_status, document_status, billing_status")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function bulkUpdateLoadStatus(loadIds, status) {
  const { data, error } = await supabase.rpc("bulk_update_load_operational_status", { target_load_ids: loadIds, target_status: status });
  if (error) throw error;
  return data;
}

export async function driverUpdateLoadStatus(loadId, status) {
  const { data, error } = await supabase.rpc("driver_update_load_status", {
    target_load_id: loadId,
    target_status: status,
  });

  if (error) throw error;
  return data;
}

export async function updateLoadBillingStatus(loadId, billingStatus) {
  const { data, error } = await supabase
    .from("loads")
    .update({ billing_status: billingStatus })
    .eq("id", loadId)
    .select("id, load_no, status, operational_status, document_status, billing_status")
    .single();

  if (error) {
    throw error;
  }

  return data;
}

export async function updateLoadAssignment(loadId, assignment) {
  const { data, error } = await supabase
    .from("loads")
    .update(assignment)
    .eq("id", loadId)
    .select("id, load_no, status, operational_status, document_status, billing_status, driver_id, truck_id, trailer_id")
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
      "id, branch_id, load_no, status, operational_status, document_status, billing_status, origin, destination, pickup_date, delivery_date, notes, reference_no, broker_reference, commodity, weight_lbs, pieces, pallets, equipment_type, temperature_min_f, temperature_max_f, hazmat, loaded_miles, empty_miles, internal_notes, customer_notes, customer_id, carrier_id, driver_id, truck_id, trailer_id, customers(name), carriers(name), drivers(name), trucks(unit_no), trailers(unit_no), created_at"
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
      "id, load_id, facility_id, stop_order, type, facility, address, contact_name, contact_phone, appointment_from, appointment_to, actual_arrival, actual_departure, instructions, reference, created_at"
    )
    .eq("load_id", loadId)
    .order("stop_order", { ascending: true })
    .order("created_at", { ascending: true });

  if (error) {
    throw error;
  }

  return data ?? [];
}

export async function listDispatchDocumentStates() {
  const { data, error } = await supabase
    .from("documents")
    .select("entity_id,doc_type,status")
    .eq("entity_type", "load");

  if (error) throw error;
  return data ?? [];
}

export async function listLoadActivity(loadId) {
  const { data, error } = await supabase
    .from("audit_logs")
    .select("id,entity_type,entity_id,user_id,action,old_value,new_value,created_at")
    .in("entity_type", ["loads", "load_stops", "load_financials", "documents", "invoices"])
    .order("created_at", { ascending: false })
    .limit(500);

  if (error) throw error;

  return (data ?? []).filter((log) => {
    if (["loads", "load_financials"].includes(log.entity_type)) return log.entity_id === loadId;
    const oldValue = log.old_value || {};
    const newValue = log.new_value || {};
    if (log.entity_type === "documents") {
      return (oldValue.entity_type === "load" && oldValue.entity_id === loadId)
        || (newValue.entity_type === "load" && newValue.entity_id === loadId);
    }
    return oldValue.load_id === loadId || newValue.load_id === loadId;
  });
}

export async function listLoadMessages(loadId) {
  const { data, error } = await supabase
    .from("messages")
    .select("id,thread_type,thread_id,sender_id,body,attachment_id,created_at")
    .eq("thread_type", "load")
    .eq("thread_id", loadId)
    .order("created_at", { ascending: true });

  if (error) throw error;
  return data ?? [];
}

export async function createLoadMessage(loadId, senderId, body) {
  const { data, error } = await supabase
    .from("messages")
    .insert({ thread_type: "load", thread_id: loadId, sender_id: senderId, body })
    .select("id,thread_type,thread_id,sender_id,body,attachment_id,created_at")
    .single();

  if (error) throw error;
  return data;
}

export async function listDriverProblemReports(loadId) {
  const query = supabase
    .from("driver_problem_reports")
    .select("id,branch_id,load_id,stop_id,driver_id,reported_by,category,severity,description,status,created_at,updated_at")
    .order("created_at", { ascending: false });

  const { data, error } = loadId ? await query.eq("load_id", loadId) : await query.limit(50);
  if (error) throw error;
  return data ?? [];
}

export async function listPublicTrackingLinks(loadId) {
  const { data, error } = await supabase
    .from("public_tracking_links")
    .select("id,load_id,label,show_status,show_eta,show_location,show_driver,show_equipment,expires_at,revoked_at,created_by,created_at,updated_at")
    .eq("load_id", loadId)
    .order("created_at", { ascending: false });
  if (error) throw error;
  return data ?? [];
}

export async function listPublicTrackingAccessLogs(loadId) {
  const { data, error } = await supabase
    .from("public_tracking_access_logs")
    .select("id,tracking_link_id,load_id,access_status,user_agent,accessed_at")
    .eq("load_id", loadId)
    .order("accessed_at", { ascending: false })
    .limit(20);
  if (error) throw error;
  return data ?? [];
}

export async function createPublicTrackingLink({ loadId, label, expiresAt, privacy }) {
  const { data, error } = await supabase.rpc("create_public_tracking_link", {
    target_load_id: loadId,
    link_label: label,
    link_expires_at: expiresAt,
    show_status: privacy.show_status,
    show_eta: privacy.show_eta,
    show_location: privacy.show_location,
    show_driver: privacy.show_driver,
    show_equipment: privacy.show_equipment,
  });
  if (error) throw error;
  return Array.isArray(data) ? data[0] : data;
}

export async function revokePublicTrackingLink(linkId) {
  const { data, error } = await supabase.rpc("revoke_public_tracking_link", { link_id: linkId });
  if (error) throw error;
  return data;
}

export async function getPublicTrackingSnapshot(token) {
  const { data, error } = await supabase.rpc("get_public_tracking_snapshot", {
    raw_token: token,
    user_agent: navigator.userAgent,
  });
  if (error) throw error;
  return data;
}

export async function createDriverProblemReport(report) {
  const { data, error } = await supabase
    .from("driver_problem_reports")
    .insert(report)
    .select("id,branch_id,load_id,stop_id,driver_id,reported_by,category,severity,description,status,created_at")
    .single();

  if (error) throw error;
  return data;
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

export async function updateLoadStop(stopId, stop) {
  const { data, error } = await supabase
    .from("load_stops")
    .update(stop)
    .eq("id", stopId)
    .select(
      "id, load_id, stop_order, type, facility, address, contact_name, contact_phone, appointment_from, appointment_to, actual_arrival, actual_departure, instructions, reference, created_at"
    )
    .single();

  if (error) throw error;
  return data;
}

export async function moveLoadStop(stopId, direction) {
  const { error } = await supabase.rpc("move_tms_load_stop", { stop_id: stopId, direction });
  if (error) throw error;
}

export async function deleteLoadStop(stopId) {
  const { error } = await supabase.rpc("delete_tms_load_stop", { stop_id: stopId });
  if (error) throw error;
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

export async function listSavedReportViews() {
  const { data, error } = await supabase.from("saved_report_views").select("id,branch_id,name,dataset,columns,filters,group_by,sort_by,sort_direction,created_at,updated_at").order("updated_at",{ascending:false});
  if(error) throw error; return data ?? [];
}

export async function saveReportView(view) {
  const { data, error } = await supabase.from("saved_report_views").upsert(view,{onConflict:"branch_id,name"}).select().single();
  if(error) throw error; return data;
}

export async function deleteSavedReportView(id) {
  const { error } = await supabase.from("saved_report_views").delete().eq("id",id); if(error) throw error;
}

export async function listScheduledReports(){const {data,error}=await supabase.from("scheduled_reports").select("id,branch_id,report_view_id,name,frequency,recipients,export_format,next_run_at,last_run_at,processing_started_at,last_error,active,created_at,saved_report_views(name)").order("next_run_at");if(error)throw error;return data??[];}
export async function saveScheduledReport(report){const {data,error}=await supabase.from("scheduled_reports").upsert(report,{onConflict:"branch_id,name"}).select().single();if(error)throw error;return data;}
export async function deleteScheduledReport(id){const {error}=await supabase.from("scheduled_reports").delete().eq("id",id);if(error)throw error;}
export async function listReportDeliveryLogs(){const {data,error}=await supabase.from("report_delivery_logs").select("id,scheduled_report_id,status,export_format,row_count,file_name,error_message,created_at").order("created_at",{ascending:false}).limit(20);if(error)throw error;return data??[];}
export async function recordReportDelivery(log){const {data,error}=await supabase.from("report_delivery_logs").insert(log).select().single();if(error)throw error;return data;}
export async function updateScheduledReportRun(id,lastRunAt,nextRunAt){const {error}=await supabase.from("scheduled_reports").update({last_run_at:lastRunAt,next_run_at:nextRunAt}).eq("id",id);if(error)throw error;}

export async function listLoadProfitability() {
  const { data, error } = await supabase
    .from("load_financials")
    .select("load_id,linehaul,fsc,accessorials,deductions,carrier_cost,driver_pay,margin,loads!inner(id,branch_id,load_no,status,operational_status,document_status,billing_status,origin,destination,pickup_date,delivery_date,customer_id,carrier_id,driver_id,customers(name),carriers(name),drivers(name))")
    .order("updated_at", { ascending: false });
  if (error) throw error;
  return data ?? [];
}

export async function listLoadDocuments(loadId) {
  const { data, error } = await supabase
    .from("documents")
    .select("id, entity_type, entity_id, doc_type, file_url, storage_bucket, storage_path, file_name, mime_type, file_size, status, uploaded_by, ocr_text, created_at")
    .eq("entity_type", "load")
    .eq("entity_id", loadId)
    .order("created_at", { ascending: false });

  if (error) {
    throw error;
  }

  return attachDocumentLinks(data ?? []);
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
    .select("id, entity_type, entity_id, doc_type, file_url, storage_bucket, storage_path, file_name, mime_type, file_size, status, uploaded_by, ocr_text, created_at")
    .order("created_at", { ascending: false });

  if (error) {
    throw error;
  }

  return attachDocumentLinks(data ?? []);
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

async function attachDocumentLinks(documents) {
  const stored = documents.filter((document) => document.storage_path);
  if (!stored.length) return documents.map((document) => ({ ...document, signed_url: document.file_url || null }));

  const { data, error } = await supabase.storage
    .from(documentBucket)
    .createSignedUrls(stored.map((document) => document.storage_path), 300);
  if (error) throw error;

  const signedByPath = new Map((data || []).map((item) => [item.path, item.signedUrl]));
  return documents.map((document) => ({
    ...document,
    signed_url: document.storage_path ? signedByPath.get(document.storage_path) || null : document.file_url || null,
  }));
}

async function getEntityBranch(entityType, entityId) {
  const directTables = {
    load: "loads",
    driver: "drivers",
    truck: "trucks",
    trailer: "trailers",
    customer: "customers",
    carrier: "carriers",
    invoice: "invoices",
  };
  const table = directTables[entityType];

  if (table) {
    const { data, error } = await supabase.from(table).select("branch_id").eq("id", entityId).single();
    if (error) throw error;
    if (!data.branch_id) throw new Error("The selected entity does not have a branch");
    return data.branch_id;
  }

  const { data, error } = await supabase.rpc("tms_entity_branch", {
    target_type: entityType,
    target_id: entityId,
  });
  if (error) throw error;
  if (!data) throw new Error("Unable to resolve the entity branch");
  return data;
}

function safeFileName(name) {
  return name.normalize("NFKD").replace(/[^a-zA-Z0-9._-]+/g, "-").replace(/^-+|-+$/g, "") || "document";
}

export async function uploadTmsDocument({ entityType, entityId, docType, status, file, uploadedBy }) {
  if (!(file instanceof File) || !file.size) throw new Error("Choose a document to upload");
  if (!allowedDocumentTypes.has(file.type)) throw new Error("Only PDF, JPG and PNG files are allowed");
  if (file.size > maxDocumentSize) throw new Error("Document must be 15 MB or smaller");

  const branchId = await getEntityBranch(entityType, entityId);
  const documentId = crypto.randomUUID();
  const storagePath = `${branchId}/${entityType}/${entityId}/${documentId}/${safeFileName(file.name)}`;
  const { error: uploadError } = await supabase.storage.from(documentBucket).upload(storagePath, file, {
    contentType: file.type,
    upsert: false,
  });
  if (uploadError) throw uploadError;

  const { data, error } = await supabase
    .from("documents")
    .insert({
      id: documentId,
      entity_type: entityType,
      entity_id: entityId,
      doc_type: docType,
      status,
      uploaded_by: uploadedBy,
      storage_bucket: documentBucket,
      storage_path: storagePath,
      file_name: file.name,
      mime_type: file.type,
      file_size: file.size,
      file_url: null,
    })
    .select("id, entity_type, entity_id, doc_type, storage_bucket, storage_path, file_name, mime_type, file_size, status, uploaded_by, created_at")
    .single();

  if (error) {
    await supabase.storage.from(documentBucket).remove([storagePath]);
    throw error;
  }

  return data;
}

export async function deleteTmsDocument(documentId) {
  const { data: document, error: readError } = await supabase
    .from("documents")
    .select("id,storage_bucket,storage_path")
    .eq("id", documentId)
    .single();
  if (readError) throw readError;

  if (document.storage_path) {
    const { error: storageError } = await supabase.storage
      .from(document.storage_bucket || documentBucket)
      .remove([document.storage_path]);
    if (storageError) throw storageError;
  }

  const { error } = await supabase.from("documents").delete().eq("id", documentId);
  if (error) throw error;
}

export async function updateTmsDocumentStatus(documentId, status) {
  const { data, error } = await supabase
    .from("documents")
    .update({ status })
    .eq("id", documentId)
    .select("id, status")
    .single();

  if (error) throw error;
  return data;
}

export async function listInvoices() {
  const { data, error } = await supabase
    .from("invoices")
    .select("id, invoice_no, customer_id, load_id, amount, status, due_date, sent_at, created_at, line_items_snapshot, customers(name), loads(load_no), payments(id,amount,date,method,reference), invoice_credit_memos(id,credit_memo_no,amount,reason,status,approved_by,approved_at,applied_by,applied_at,voided_by,voided_at,created_at), invoice_line_items(id,item_type,description,customer_description,quantity,unit_amount,amount,receipt_required,customer_visible,approved,approval_status,approval_note,approved_by,approved_at,document_id,sort_order)")
    .order("created_at", { ascending: false });

  if (error) {
    throw error;
  }

  return data ?? [];
}

const invoiceListSelect = "id, branch_id, invoice_no, customer_id, load_id, amount, status, due_date, sent_at, created_at, line_items_snapshot, customers(name), loads(load_no), payments(id,amount,date,method,reference), invoice_credit_memos(id,credit_memo_no,amount,reason,status,approved_by,approved_at,applied_by,applied_at,voided_by,voided_at,created_at), invoice_line_items(id,item_type,description,customer_description,quantity,unit_amount,amount,receipt_required,customer_visible,approved,approval_status,approval_note,approved_by,approved_at,document_id,sort_order)";

export async function listInvoicesPage(options = {}) {
  const page = Math.max(1, Number(options.page) || 1);
  const pageSize = Math.min(100, Math.max(1, Number(options.pageSize) || 10));
  const sortColumns = new Set(["created_at", "invoice_no", "amount", "due_date", "status"]);
  const sortBy = sortColumns.has(options.sortBy) ? options.sortBy : "created_at";
  const ascending = options.sortDirection === "asc";
  const search = safeSearchTerm(options.search);
  const [customerIds, loadIds] = await Promise.all([
    matchingEntityIds("customers", ["name"], search),
    matchingEntityIds("loads", ["load_no"], search),
  ]);
  const searchParts = [`invoice_no.ilike.%${search}%`];
  if (customerIds.length) searchParts.push(`customer_id.in.(${customerIds.join(",")})`);
  if (loadIds.length) searchParts.push(`load_id.in.(${loadIds.join(",")})`);
  let query = supabase.from("invoices").select(invoiceListSelect, { count: "exact" });
  if (search) query = query.or(searchParts.join(","));
  if (options.status) query = query.eq("status", options.status);
  const start = (page - 1) * pageSize;
  const { data, error, count } = await query.order(sortBy, { ascending, nullsFirst: false }).range(start, start + pageSize - 1);
  if (error) throw error;
  return { records: data || [], count: count || 0, page, pageSize };
}

export async function listInvoiceBatches() {
  const { data, error } = await supabase
    .from("invoice_batches")
    .select("id,branch_id,batch_no,customer_id,billing_cycle,period_start,period_end,due_date,status,invoice_count,total_amount,notes,sent_at,created_at,customers(name,billing_address,billing_email),invoice_batch_items(id,invoice_id,amount_snapshot,invoices(invoice_no,status,due_date,amount))")
    .order("created_at", { ascending: false });
  if (error) throw error;
  return data ?? [];
}

export async function createInvoiceBatch(batch, invoiceIds) {
  const { data, error } = await supabase.rpc("create_invoice_batch", {
    batch_payload: batch,
    invoice_ids: invoiceIds,
  });
  if (error) throw error;
  return data;
}

export async function listInvoiceCollectionNotes() {
  const { data, error } = await supabase
    .from("invoice_collection_notes")
    .select("id,branch_id,invoice_id,invoice_batch_id,note_type,note,follow_up_date,status,created_at,created_by")
    .order("created_at", { ascending: false });
  if (error) throw error;
  return data ?? [];
}

export async function createInvoiceCollectionNote(note) {
  const payload = {
    invoice_id: note.invoice_id || null,
    invoice_batch_id: note.invoice_batch_id || null,
    note_type: note.note_type || "internal",
    note: note.note,
    follow_up_date: note.follow_up_date || null,
    status: note.status || "open",
  };
  const { data, error } = await supabase
    .from("invoice_collection_notes")
    .insert(payload)
    .select("id,invoice_id,invoice_batch_id,note_type,note,follow_up_date,status,created_at")
    .single();
  if (error) throw error;
  return data;
}

export async function createPayment(payment) {
  const { data, error } = await supabase.from("payments").insert(payment).select("id,invoice_id,amount,date,method,reference").single();
  if (error) throw error;
  return data;
}

export async function getInvoiceDetails(invoiceId) {
  const { data, error } = await supabase
    .from("invoices")
    .select("id, invoice_no, customer_id, load_id, amount, status, due_date, sent_at, created_at, line_items_snapshot, customers(name,billing_address), loads(load_no,status,operational_status,document_status,billing_status,origin,destination,pickup_date,delivery_date), payments(id,amount,date,method,reference), invoice_credit_memos(id,credit_memo_no,amount,reason,status,approved_by,approved_at,applied_by,applied_at,voided_by,voided_at,created_at), invoice_line_items(id,item_type,description,customer_description,quantity,unit_amount,amount,receipt_required,customer_visible,approved,approval_status,approval_note,approved_by,approved_at,document_id,sort_order)")
    .eq("id", invoiceId)
    .single();

  if (error) throw error;
  return data;
}

export async function createInvoiceWithLineItems(invoice, lineItems) {
  const { data, error } = await supabase
    .rpc("create_invoice_with_line_items", {
      invoice_payload: invoice,
      line_items_payload: lineItems,
    });

  if (error) throw error;
  return data;
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

export async function bulkUpdateInvoiceStatus(invoiceIds, status) {
  const { data, error } = await supabase.rpc("bulk_update_invoice_status", { target_invoice_ids: invoiceIds, target_status: status });
  if (error) throw error;
  return data;
}

export async function updateInvoiceLineItemApproval(lineItemId, approvalStatus, approvalNote = null) {
  const { data, error } = await supabase
    .from("invoice_line_items")
    .update({
      approval_status: approvalStatus,
      approval_note: approvalNote?.trim() || null,
    })
    .eq("id", lineItemId)
    .select("id,item_type,approval_status,approval_note,approved,approved_by,approved_at,receipt_required,document_id")
    .single();

  if (error) throw error;
  return data;
}

export async function attachInvoiceLineItemReceipt(lineItemId, documentId) {
  const { data, error } = await supabase
    .from("invoice_line_items")
    .update({ document_id: documentId })
    .eq("id", lineItemId)
    .select("id,invoice_id,item_type,receipt_required,document_id")
    .single();

  if (error) throw error;
  return data;
}

export async function createInvoiceCreditMemo(invoiceId, amount, reason) {
  const { data, error } = await supabase.rpc("create_invoice_credit_memo", {
    target_invoice_id: invoiceId,
    target_amount: amount,
    target_reason: reason,
  });
  if (error) throw error;
  return data;
}

export async function updateInvoiceCreditMemoStatus(creditMemoId, status) {
  const { data, error } = await supabase.rpc("update_invoice_credit_memo_status", {
    target_credit_memo_id: creditMemoId,
    target_status: status,
  });
  if (error) throw error;
  return data;
}

export async function listSettlementCandidates() {
  const { data, error } = await supabase
    .from("loads")
    .select("id,load_no,status,operational_status,document_status,billing_status,origin,destination,pickup_date,delivery_date,carrier_id,driver_id,carriers(name),drivers(name),load_financials(carrier_cost,driver_pay),settlement_items(id,payee_type,active)")
    .eq("operational_status", "delivered")
    .order("delivery_date", { ascending: true });
  if (error) throw error;
  return data ?? [];
}

export async function listSettlements() {
  const { data, error } = await supabase
    .from("settlements")
    .select("id,settlement_no,payee_type,carrier_id,driver_id,status,settlement_date,total_amount,notes,approved_at,paid_at,created_at,carriers(name),drivers(name),settlement_items(id,load_id,amount,description,active,loads(load_no,origin,destination,pickup_date,delivery_date))")
    .order("created_at", { ascending: false });
  if (error) throw error;
  return data ?? [];
}

export async function listSettlementPayBatches() {
  const { data, error } = await supabase
    .from("settlement_pay_batches")
    .select("id,batch_no,payee_type,carrier_id,driver_id,status,period_start,period_end,pay_date,settlement_count,total_amount,notes,approved_at,paid_at,created_at,carriers(name),drivers(name),settlement_pay_batch_items(id,settlement_id,amount_snapshot,settlements(settlement_no,status,total_amount,settlement_date,payee_type,settlement_items(id,load_id,amount,loads(load_no,origin,destination,pickup_date,delivery_date))))")
    .order("created_at", { ascending: false });
  if (error) throw error;
  return data ?? [];
}

export async function createSettlementPayBatch(batch, settlementIds) {
  const { data, error } = await supabase.rpc("create_settlement_pay_batch", {
    batch_payload: batch,
    settlement_ids: settlementIds,
  });
  if (error) throw error;
  return data;
}

export async function createLoadSettlement(loadId, payeeType, notes = null) {
  const { data, error } = await supabase.rpc("create_load_settlement", {
    target_load_id: loadId,
    target_payee_type: payeeType,
    target_notes: notes,
  });
  if (error) throw error;
  return data;
}

export async function updateSettlementStatus(settlementId, status) {
  const { data, error } = await supabase
    .from("settlements")
    .update({ status })
    .eq("id", settlementId)
    .select("id,settlement_no,status,total_amount,approved_at,paid_at")
    .single();
  if (error) throw error;
  return data;
}

export async function updateSettlementPayBatchStatus(batchId, status) {
  const { data, error } = await supabase
    .from("settlement_pay_batches")
    .update({ status })
    .eq("id", batchId)
    .select("id,batch_no,status,total_amount,approved_at,paid_at")
    .single();
  if (error) throw error;
  return data;
}
