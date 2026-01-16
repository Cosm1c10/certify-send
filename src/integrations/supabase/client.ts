import { createClient, SupabaseClient } from '@supabase/supabase-js';

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

function createSupabaseClient(): SupabaseClient {
  if (!supabaseUrl || !supabaseAnonKey ||
      supabaseUrl === 'your_supabase_project_url_here' ||
      supabaseAnonKey === 'your_supabase_anon_key_here') {
    console.warn(
      'Supabase not configured. Please update .env.local with your Supabase credentials.'
    );
    // Return a client with placeholder values - will fail gracefully on API calls
    return createClient('https://placeholder.supabase.co', 'placeholder-key');
  }
  return createClient(supabaseUrl, supabaseAnonKey);
}

export const supabase = createSupabaseClient();
