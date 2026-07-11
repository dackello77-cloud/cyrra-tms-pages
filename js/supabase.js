import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const SUPABASE_URL = "https://dapegyncqmaezcpgrqxd.supabase.co";
const SUPABASE_ANON_KEY = "sb_publishable_fQUD4c2uFtFH3d0nU10mjg_M-lC7PWr";

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

