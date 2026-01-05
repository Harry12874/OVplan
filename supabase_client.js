const SUPABASE_URL = "https://kdgsxghuevpmsnkhimfb.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtkZ3N4Z2h1ZXZwbXNua2hpbWZiIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Njc1MTU0OTcsImV4cCI6MjA4MzA5MTQ5N30.nLu4GvXRMFZDgwoSETaGfJ_0phYVmw9Qvi6rlWIg8zE";

const supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

function getSession() {
  return supabase.auth.getSession();
}

function signIn(email, password) {
  return supabase.auth.signInWithPassword({ email, password });
}

function signOut() {
  return supabase.auth.signOut();
}

export { supabase, getSession, signIn, signOut };
