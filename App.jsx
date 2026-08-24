import { useEffect, useState } from "react";
import { supabase } from "./supabaseClient";
import Login from "./Login";
import AccountingApp from "./AccountingApp";

export default function App() {
  const [session, setSession] = useState(undefined); // undefined = still checking

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => listener.subscription.unsubscribe();
  }, []);

  if (session === undefined) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "sans-serif", color: "#5B6B73" }}>
        กำลังโหลด...
      </div>
    );
  }

  if (!session) return <Login onLogin={setSession} />;

  return (
    <div>
      <div
        style={{
          display: "flex", justifyContent: "space-between", alignItems: "center",
          padding: "8px 16px", fontSize: 12, background: "#fff", borderBottom: "1px solid #E2E1D9",
        }}
      >
        <span style={{ color: "#5B6B73" }}>เข้าสู่ระบบในชื่อ {session.user.email.split("@")[0]}</span>
        <button
          onClick={() => supabase.auth.signOut()}
          style={{ background: "none", border: "none", color: "#5B6B73", cursor: "pointer", fontSize: 12, textDecoration: "underline" }}
        >
          ออกจากระบบ
        </button>
      </div>
      <AccountingApp />
    </div>
  );
}
