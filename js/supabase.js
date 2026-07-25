import { createClient } from "../vendor/supabase/supabase.js";
import { appConfig } from "./config.js";

export const supabase = createClient(
  appConfig.supabaseUrl,
  appConfig.supabasePublishableKey
);
