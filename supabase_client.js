const SUPABASE_URL = "https://kdgsxghuevpmsnkhimfb.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtkZ3N4Z2h1ZXZwbXNua2hpbWZiIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Njc1MTU0OTcsImV4cCI6MjA4MzA5MTQ5N30.nLu4GvXRMFZDgwoSETaGfJ_0phYVmw9Qvi6rlWIg8zE";

const supabaseAvailable = Boolean(
  typeof window !== "undefined" &&
    window.supabase?.createClient &&
    SUPABASE_URL &&
    SUPABASE_ANON_KEY
);
const supabaseInitError = supabaseAvailable
  ? null
  : new Error(
      "Supabase client unavailable. Ensure @supabase/supabase-js loads and SUPABASE_URL/SUPABASE_ANON_KEY are set."
    );

function createSupabaseStub(error) {
  const stubError = error || new Error("Supabase client unavailable.");
  const withError = async (data = null) => ({ data, error: stubError });
  const stubTable = {
    select: () => withError([]),
    upsert: () => withError(),
    delete: () => ({
      eq: () => withError(),
    }),
  };
  return {
    from: () => stubTable,
    auth: {
      getSession: () => withError({ session: null }),
      signInWithPassword: () => withError(),
      signOut: () => withError(),
      onAuthStateChange: () => ({
        data: { subscription: { unsubscribe: () => {} } },
      }),
    },
  };
}

const supabase = supabaseAvailable
  ? window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY)
  : createSupabaseStub(supabaseInitError);

function getSession() {
  return supabase.auth.getSession();
}

function signIn(email, password) {
  return supabase.auth.signInWithPassword({ email, password });
}

function signOut() {
  return supabase.auth.signOut();
}

export { supabase, getSession, signIn, signOut, supabaseAvailable, supabaseInitError };
