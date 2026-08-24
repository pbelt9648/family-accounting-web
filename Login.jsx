import { useState } from "react";
import { supabase } from "./supabaseClient";

// Supabase Auth needs an email address internally. Since only invited family
// members will ever have an account (public sign-up is turned off — see
// README.md), we let people log in with a plain username and quietly turn it
// into "username@family.local" behind the scenes.
const EMAIL_DOMAIN = "family.local";

export default function Login({ onLogin }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  async function handleSubmit(e) {
    e.preventDefault();
    if (!username.trim() || !password) return;
    setError("");
    setLoading(true);
    const email = `${username.trim().toLowerCase()}@${EMAIL_DOMAIN}`;
    const { data, error: authError } = await supabase.auth.signInWithPassword({ email, password });
    setLoading(false);
    if (authError) {
      setError("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง");
      return;
    }
    onLogin(data.session);
  }

  return (
    <div
      style={{
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background: "#F4F5F0",
        fontFamily: "Inter, sans-serif",
        padding: 16,
      }}
    >
      <form
        onSubmit={handleSubmit}
        style={{ background: "#fff", padding: 32, borderRadius: 12, width: 320, border: "1px solid #E2E1D9" }}
      >
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 20, marginBottom: 4, color: "#1E2124" }}>
          เข้าสู่ระบบบัญชี
        </div>
        <p style={{ fontSize: 13, color: "#5B6B73", marginBottom: 20 }}>สำหรับสมาชิกที่ได้รับอนุญาตเท่านั้น</p>

        <label style={{ fontSize: 13, fontWeight: 500, color: "#1E2124" }}>ชื่อผู้ใช้</label>
        <input
          value={username}
          onChange={(e) => setUsername(e.target.value)}
          autoFocus
          style={inputStyle}
        />

        <label style={{ fontSize: 13, fontWeight: 500, color: "#1E2124" }}>รหัสผ่าน</label>
        <input
          type="password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          style={inputStyle}
        />

        {error && <p style={{ color: "#C1462F", fontSize: 13, marginBottom: 8 }}>{error}</p>}

        <button disabled={loading} style={btnStyle}>
          {loading ? "กำลังเข้าสู่ระบบ..." : "เข้าสู่ระบบ"}
        </button>
      </form>
    </div>
  );
}

const inputStyle = {
  width: "100%",
  padding: "8px 10px",
  margin: "6px 0 14px",
  border: "1px solid #E2E1D9",
  borderRadius: 7,
  fontSize: 14,
  boxSizing: "border-box",
};

const btnStyle = {
  width: "100%",
  padding: "10px",
  background: "#1E2124",
  color: "#fff",
  border: "none",
  borderRadius: 8,
  fontWeight: 600,
  fontSize: 14,
  cursor: "pointer",
};
