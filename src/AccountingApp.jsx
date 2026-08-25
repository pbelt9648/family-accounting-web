import { useState, useEffect, useMemo, useRef } from "react";
import {
  LayoutDashboard, FileText, Wallet, Package, Users, SlidersHorizontal,
  Plus, Pencil, Trash2, X, Search, CheckCircle2, AlertTriangle, TrendingUp,
  TrendingDown, ArrowRight, Send, FileCheck2, Ban, RefreshCw, Loader2,
  BarChart3, Receipt, Printer, Download, Landmark,
} from "lucide-react";
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
} from "recharts";
import * as XLSX from "xlsx";

/* ---------------------------------- utils --------------------------------- */

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const today = () => new Date().toISOString().slice(0, 10);
const thb = (n) =>
  new Intl.NumberFormat("th-TH", { style: "currency", currency: "THB", maximumFractionDigits: 2 }).format(
    Number.isFinite(n) ? n : 0
  );
const fmtDate = (d) =>
  d ? new Date(d).toLocaleDateString("th-TH", { year: "numeric", month: "short", day: "numeric" }) : "-";
function resizeImageToDataUrl(file, maxSize = 200, quality = 0.72) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = reject;
    reader.onload = () => {
      const img = new Image();
      img.onerror = reject;
      img.onload = () => {
        const scale = Math.min(1, maxSize / Math.max(img.width, img.height));
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(img.width * scale);
        canvas.height = Math.round(img.height * scale);
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

const monthKey = (d) => (d || today()).slice(0, 7);
const monthLabel = (key) => {
  const [y, m] = key.split("-");
  return new Date(Number(y), Number(m) - 1, 1).toLocaleDateString("th-TH", { month: "short", year: "2-digit" });
};

function thaiBahtText(amount) {
  const digits = ["ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"];
  const positions = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน"];
  function convertGroup(numStr) {
    let result = "";
    const len = numStr.length;
    for (let i = 0; i < len; i++) {
      const digit = parseInt(numStr[i], 10);
      const pos = len - i - 1;
      if (digit === 0) continue;
      if (pos === 0) {
        result += digit === 1 && len > 1 ? "เอ็ด" : digits[digit];
      } else if (pos === 1) {
        result += digit === 1 ? "สิบ" : digit === 2 ? "ยี่สิบ" : digits[digit] + "สิบ";
      } else {
        result += digits[digit] + positions[pos];
      }
    }
    return result;
  }
  let n = Math.abs(Math.round((Number(amount) || 0) * 100) / 100);
  const negative = Number(amount) < 0;
  let [intPart, decPart] = n.toFixed(2).split(".");
  intPart = intPart.replace(/^0+(?=\d)/, "");
  let bahtText;
  if (intPart === "0") {
    bahtText = "ศูนย์บาท";
  } else {
    let groups = [];
    let s = intPart;
    while (s.length > 6) { groups.unshift(s.slice(-6)); s = s.slice(0, -6); }
    groups.unshift(s);
    bahtText = groups.map((g) => convertGroup(g)).join("ล้าน") + "บาท";
  }
  const satangText = decPart === "00" ? "ถ้วน" : convertGroup(decPart.replace(/^0+(?=\d)/, "")) + "สตางค์";
  return (negative ? "ลบ" : "") + bahtText + satangText;
}

const DEFAULT_SETTINGS = {
  companyName: "กิจการของฉัน",
  address: "",
  phone: "",
  taxId: "",
  vatRate: 7,
  invoicePrefix: "INV",
  quotePrefix: "QT",
  billingPrefix: "BN",
  taxInvoicePrefix: "TI",
  receiptPrefix: "RC",
  debitNotePrefix: "DN",
  creditNotePrefix: "CN",
  sellerCode: "001",
  nextInvoiceNumber: 1,
  nextQuoteNumber: 1,
  nextBillingNumber: 1,
  nextTaxInvoiceNumber: 1,
  nextReceiptNumber: 1,
  nextDebitNoteNumber: 1,
  nextCreditNoteNumber: 1,
  lowStockDefault: 5,
};

const WHT_TYPES = [
  { code: "service", label: "ค่าจ้างทำของ/ค่าบริการ", rate: 3 },
  { code: "rent", label: "ค่าเช่า", rate: 5 },
  { code: "transport", label: "ค่าขนส่ง", rate: 1 },
  { code: "ads", label: "ค่าโฆษณา", rate: 2 },
  { code: "professional", label: "ค่าวิชาชีพอิสระ", rate: 3 },
  { code: "other", label: "อื่นๆ", rate: 3 },
];

const DOC_TYPE_LABEL = {
  quote: "ใบเสนอราคา",
  invoice: "ใบแจ้งหนี้",
  taxinvoice: "ใบกำกับภาษี",
  receipt: "ใบเสร็จรับเงิน",
  debitnote: "ใบเพิ่มหนี้",
  creditnote: "ใบลดหนี้",
};

const PAYMENT_METHOD_LABEL = { cash: "เงินสด", cheque: "เช็ค", transfer: "โอนเงิน", credit: "บัตรเครดิต" };

function docPaidAmount(doc) {
  return (doc.payments || []).reduce((s, p) => s + (Number(p.amount) || 0), 0);
}
function docBalanceDue(doc) {
  return Math.max(0, (Number(doc.total) || 0) - docPaidAmount(doc));
}

const EXPENSE_CATEGORIES = ["วัตถุดิบ/ชิ้นส่วน", "ค่าแรง", "ค่าขนส่ง", "ค่าน้ำ-ค่าไฟ", "ค่าเช่า", "ค่าซ่อมบำรุง", "การตลาด", "ค่าเสื่อมราคา", "อื่นๆ"];
const INCOME_CATEGORIES = ["ขายสินค้า/บริการ", "รายได้อื่น"];

const STORAGE_KEYS = ["settings", "customers", "products", "documents", "transactions", "billingNotes", "assets"];

async function loadAll() {
  const out = {};
  await Promise.all(
    STORAGE_KEYS.map(async (k) => {
      try {
        const r = await window.storage.get(k, false);
        out[k] = r && r.value ? JSON.parse(r.value) : null;
      } catch {
        out[k] = null;
      }
    })
  );
  return out;
}

/* ---------------------------------- shell ---------------------------------- */

const NAV = [
  { id: "dashboard", label: "แดชบอร์ด", icon: LayoutDashboard },
  { id: "documents", label: "เอกสารขาย", icon: FileText },
  { id: "billing", label: "ใบวางบิล", icon: Receipt },
  { id: "salesReport", label: "รายงานขาย", icon: BarChart3 },
  { id: "transactions", label: "รายรับ-รายจ่าย", icon: Wallet },
  { id: "inventory", label: "สต็อกสินค้า", icon: Package },
  { id: "assets", label: "สินทรัพย์ถาวร", icon: Landmark },
  { id: "customers", label: "ลูกค้า", icon: Users },
  { id: "settings", label: "ตั้งค่า", icon: SlidersHorizontal },
];

export default function AccountingApp() {
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState("dashboard");
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [customers, setCustomers] = useState([]);
  const [products, setProducts] = useState([]);
  const [documents, setDocuments] = useState([]);
  const [transactions, setTransactions] = useState([]);
  const [billingNotes, setBillingNotes] = useState([]);
  const [assets, setAssets] = useState([]);
  const [toast, setToast] = useState(null);
  const [confirmDialog, setConfirmDialog] = useState(null);
  const toastTimer = useRef(null);

  useEffect(() => {
    (async () => {
      const d = await loadAll();
      setSettings(d.settings ? { ...DEFAULT_SETTINGS, ...d.settings } : DEFAULT_SETTINGS);
      setCustomers(d.customers || []);
      setProducts(d.products || []);
      setDocuments(d.documents || []);
      setTransactions(d.transactions || []);
      setBillingNotes(d.billingNotes || []);
      setAssets(d.assets || []);
      setLoading(false);
    })();
  }, []);

  function showToast(msg, type = "success") {
    setToast({ msg, type });
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToast(null), 2600);
  }

  function confirmAction(message, action) {
    setConfirmDialog({ message, action });
  }

  async function persist(key, value) {
    try {
      await window.storage.set(key, JSON.stringify(value), false);
    } catch {
      showToast("บันทึกข้อมูลไม่สำเร็จ ลองใหม่อีกครั้ง", "error");
    }
  }

  const updateSettings = (next) => { setSettings(next); persist("settings", next); };
  const updateCustomers = (next) => { setCustomers(next); persist("customers", next); };
  const updateProducts = (next) => { setProducts(next); persist("products", next); };
  const updateDocuments = (next) => { setDocuments(next); persist("documents", next); };
  const updateTransactions = (next) => { setTransactions(next); persist("transactions", next); };
  const updateBillingNotes = (next) => { setBillingNotes(next); persist("billingNotes", next); };
  const updateAssets = (next) => { setAssets(next); persist("assets", next); };

  function resetAll() {
    confirmAction("ล้างข้อมูลทั้งหมดถาวร? การกระทำนี้ยกเลิกไม่ได้", async () => {
      for (const k of STORAGE_KEYS) {
        try { await window.storage.delete(k, false); } catch {}
      }
      setSettings(DEFAULT_SETTINGS);
      setCustomers([]); setProducts([]); setDocuments([]); setTransactions([]); setBillingNotes([]); setAssets([]);
      showToast("ล้างข้อมูลเรียบร้อย");
    });
  }

  const netBalance = useMemo(
    () => transactions.reduce((s, t) => s + (t.type === "income" ? t.amount : -t.amount), 0),
    [transactions]
  );

  if (loading) {
    return (
      <Shell settings={settings}>
        <div className="flex items-center justify-center" style={{ height: 400 }}>
          <div className="flex flex-col items-center gap-3" style={{ color: "var(--steel)" }}>
            <Loader2 className="animate-spin" size={28} />
            <span style={{ fontFamily: "var(--font-body)" }}>กำลังโหลดข้อมูล...</span>
          </div>
        </div>
      </Shell>
    );
  }

  return (
    <Shell
      settings={settings} tab={tab} setTab={setTab} netBalance={netBalance} toast={toast}
      confirmDialog={confirmDialog}
      onConfirmYes={() => { const a = confirmDialog?.action; setConfirmDialog(null); a && a(); }}
      onConfirmNo={() => setConfirmDialog(null)}
    >
      {tab === "dashboard" && (
        <Dashboard documents={documents} transactions={transactions} products={products} customers={customers} settings={settings} />
      )}
      {tab === "documents" && (
        <DocumentsTab
          documents={documents} customers={customers} products={products} settings={settings}
          updateDocuments={updateDocuments} updateSettings={updateSettings}
          updateTransactions={updateTransactions} transactions={transactions}
          updateProducts={updateProducts} showToast={showToast} confirmAction={confirmAction}
        />
      )}
      {tab === "billing" && (
        <BillingTab
          documents={documents} customers={customers} settings={settings} updateSettings={updateSettings}
          billingNotes={billingNotes} updateBillingNotes={updateBillingNotes} updateDocuments={updateDocuments}
          showToast={showToast} confirmAction={confirmAction}
        />
      )}
      {tab === "salesReport" && (
        <SalesReportTab documents={documents} products={products} transactions={transactions} showToast={showToast} />
      )}
      {tab === "transactions" && (
        <TransactionsTab transactions={transactions} updateTransactions={updateTransactions} settings={settings} showToast={showToast} confirmAction={confirmAction} />
      )}
      {tab === "inventory" && (
        <InventoryTab products={products} updateProducts={updateProducts} settings={settings} showToast={showToast} confirmAction={confirmAction} />
      )}
      {tab === "assets" && (
        <AssetsTab assets={assets} updateAssets={updateAssets} updateTransactions={updateTransactions} transactions={transactions} showToast={showToast} confirmAction={confirmAction} />
      )}
      {tab === "customers" && (
        <CustomersTab customers={customers} updateCustomers={updateCustomers} showToast={showToast} confirmAction={confirmAction} />
      )}
      {tab === "settings" && (
        <SettingsTab settings={settings} updateSettings={updateSettings} onReset={resetAll} showToast={showToast} />
      )}
    </Shell>
  );
}

function Shell({ children, settings, tab, setTab, netBalance, toast, confirmDialog, onConfirmYes, onConfirmNo }) {
  return (
    <div style={{ fontFamily: "var(--font-body)", background: "var(--paper)", color: "var(--ink)", minHeight: "100%" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
        :root{
          --paper:#F4F5F0; --surface:#FFFFFF; --ink:#1E2124; --ink-soft:#4B5A63;
          --steel:#5B6B73; --steel-light:#8B9AA1; --border:#E2E1D9;
          --amber:#D98A2B; --amber-dark:#B56F1C; --green:#2E7D52; --green-bg:#E7F3EC;
          --red:#C1462F; --red-bg:#FBEBE7;
          --font-display:'Space Grotesk',sans-serif; --font-body:'Inter',sans-serif; --font-mono:'IBM Plex Mono',monospace;
        }
        .stripe{ height:5px; background-image: repeating-linear-gradient(135deg, var(--amber) 0px, var(--amber) 9px, var(--ink) 9px, var(--ink) 18px); }
        .card{ background:var(--surface); border:1px solid var(--border); border-radius:10px; }
        .btn-primary{ background:var(--ink); color:#fff; font-weight:600; border-radius:8px; transition:background .15s; }
        .btn-primary:hover{ background:var(--amber-dark); }
        .btn-ghost{ background:transparent; border:1px solid var(--border); border-radius:8px; color:var(--ink-soft); transition:.15s; }
        .btn-ghost:hover{ border-color:var(--steel); color:var(--ink); }
        .btn-danger{ color:var(--red); }
        input, select, textarea{ font-family:var(--font-body); border:1px solid var(--border); border-radius:7px; background:#fff; }
        input:focus, select:focus, textarea:focus{ outline:2px solid var(--amber); outline-offset:0; border-color:var(--amber); }
        table{ border-collapse:collapse; width:100%; }
        th{ text-align:left; font-size:11px; letter-spacing:.05em; text-transform:uppercase; color:var(--steel); font-weight:600; padding:10px 12px; border-bottom:1px solid var(--border); }
        td{ padding:12px; border-bottom:1px solid var(--border); font-size:14px; }
        tr:last-child td{ border-bottom:none; }
        .badge{ display:inline-flex; align-items:center; gap:4px; font-size:12px; font-weight:600; padding:3px 9px; border-radius:99px; }
        .mono{ font-family:var(--font-mono); }
        .navlink{ display:flex; align-items:center; gap:10px; padding:10px 14px; border-radius:8px; font-weight:500; font-size:14px; color:var(--ink-soft); cursor:pointer; white-space:nowrap; }
        .navlink:hover{ background:#EAEAE3; }
        .navlink.active{ background:var(--ink); color:#fff; }
        ::-webkit-scrollbar{ height:8px; width:8px; } ::-webkit-scrollbar-thumb{ background:var(--border); border-radius:8px; }
        @media print{
          body *{ visibility:hidden; }
          .print-area, .print-area *{ visibility:visible; }
          .print-area{ position:fixed; top:0; left:0; width:100%; }
          .no-print{ display:none !important; }
        }
      `}</style>

      <div className="flex flex-col md:flex-row" style={{ minHeight: "100%" }}>
        {/* Desktop sidebar */}
        <aside className="hidden md:flex md:flex-col md:w-60 shrink-0" style={{ borderRight: "1px solid var(--border)", background: "var(--surface)" }}>
          <div className="p-5">
            <div style={{ fontFamily: "var(--font-display)", fontWeight: 700, fontSize: 18 }}>{settings.companyName}</div>
            <div style={{ fontSize: 12, color: "var(--steel)" }}>ระบบบัญชีส่วนตัว</div>
          </div>
          <div className="stripe" />
          <nav className="p-3 flex flex-col gap-1 flex-1">
            {NAV.map((n) => (
              <div key={n.id} className={`navlink ${tab === n.id ? "active" : ""}`} onClick={() => setTab(n.id)}>
                <n.icon size={17} /> {n.label}
              </div>
            ))}
          </nav>
          <div className="p-4" style={{ borderTop: "1px solid var(--border)" }}>
            <div style={{ fontSize: 11, color: "var(--steel)" }}>ยอดคงเหลือสุทธิ</div>
            <div className="mono" style={{ fontWeight: 600, fontSize: 16, color: netBalance >= 0 ? "var(--green)" : "var(--red)" }}>
              {thb(netBalance)}
            </div>
          </div>
        </aside>

        {/* Mobile top bar */}
        <div className="md:hidden" style={{ background: "var(--surface)", borderBottom: "1px solid var(--border)" }}>
          <div className="flex items-center justify-between px-4 pt-4 pb-2">
            <div style={{ fontFamily: "var(--font-display)", fontWeight: 700, fontSize: 16 }}>{settings.companyName}</div>
            <div className="mono" style={{ fontSize: 13, fontWeight: 600, color: netBalance >= 0 ? "var(--green)" : "var(--red)" }}>{thb(netBalance)}</div>
          </div>
          <div className="stripe" />
          <nav className="flex gap-1 p-2 overflow-x-auto">
            {NAV.map((n) => (
              <div key={n.id} className={`navlink ${tab === n.id ? "active" : ""}`} onClick={() => setTab(n.id)}>
                <n.icon size={16} /> {n.label}
              </div>
            ))}
          </nav>
        </div>

        <main className="flex-1 p-4 md:p-8" style={{ maxWidth: 1100 }}>
          {children}
        </main>
      </div>

      {toast && (
        <div
          className="fixed left-1/2 flex items-center gap-2 shadow-lg"
          style={{
            bottom: 24, transform: "translateX(-50%)", background: toast.type === "error" ? "var(--red)" : "var(--ink)",
            color: "#fff", padding: "10px 18px", borderRadius: 8, fontSize: 14, fontWeight: 500, zIndex: 100,
          }}
        >
          {toast.type === "error" ? <AlertTriangle size={16} /> : <CheckCircle2 size={16} />} {toast.msg}
        </div>
      )}

      {confirmDialog && <ConfirmModal message={confirmDialog.message} onCancel={onConfirmNo} onConfirm={onConfirmYes} />}
    </div>
  );
}

function ConfirmModal({ message, onCancel, onConfirm }) {
  return (
    <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(20,20,18,0.55)", zIndex: 80 }}>
      <div className="card p-5" style={{ maxWidth: 380, width: "100%" }}>
        <div className="flex items-start gap-3 mb-5">
          <div className="flex items-center justify-center shrink-0" style={{ width: 34, height: 34, borderRadius: 8, background: "var(--red-bg)", color: "var(--red)" }}>
            <AlertTriangle size={17} />
          </div>
          <p style={{ fontSize: 14, paddingTop: 6 }}>{message}</p>
        </div>
        <div className="flex gap-2 justify-end">
          <button className="btn-ghost px-4 py-2 text-sm" onClick={onCancel}>ยกเลิก</button>
          <button className="px-4 py-2 text-sm" style={{ background: "var(--red)", color: "#fff", fontWeight: 600, borderRadius: 8 }} onClick={onConfirm}>ยืนยันลบ</button>
        </div>
      </div>
    </div>
  );
}

/* -------------------------------- shared bits ------------------------------- */

function SectionHeader({ title, subtitle, action }) {
  return (
    <div className="flex items-start justify-between mb-6 flex-wrap gap-3">
      <div>
        <h1 style={{ fontFamily: "var(--font-display)", fontWeight: 700, fontSize: 24 }}>{title}</h1>
        {subtitle && <p style={{ color: "var(--steel)", fontSize: 14, marginTop: 2 }}>{subtitle}</p>}
      </div>
      {action}
    </div>
  );
}

function Empty({ text, cta }) {
  return (
    <div className="card flex flex-col items-center justify-center text-center p-12" style={{ color: "var(--steel)" }}>
      <p style={{ fontSize: 14 }}>{text}</p>
      {cta}
    </div>
  );
}

function Modal({ title, onClose, children, wide }) {
  return (
    <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(20,20,18,0.5)", zIndex: 50 }}>
      <div className="card w-full flex flex-col" style={{ maxWidth: wide ? 720 : 460, maxHeight: "88vh" }}>
        <div className="flex items-center justify-between px-5 py-4" style={{ borderBottom: "1px solid var(--border)" }}>
          <h2 style={{ fontFamily: "var(--font-display)", fontWeight: 700, fontSize: 17 }}>{title}</h2>
          <button onClick={onClose} className="btn-ghost p-1"><X size={18} /></button>
        </div>
        <div className="p-5 overflow-y-auto">{children}</div>
      </div>
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="flex flex-col gap-1 mb-3" style={{ fontSize: 13 }}>
      <span style={{ fontWeight: 500, color: "var(--ink-soft)" }}>{label}</span>
      {children}
    </label>
  );
}

const inputCls = "px-3 py-2 text-sm w-full";

function StatCard({ label, value, icon: Icon, tone }) {
  const colors = {
    green: { bg: "var(--green-bg)", fg: "var(--green)" },
    red: { bg: "var(--red-bg)", fg: "var(--red)" },
    amber: { bg: "#FBF0DF", fg: "var(--amber-dark)" },
    ink: { bg: "#EAEAE3", fg: "var(--ink)" },
  }[tone || "ink"];
  return (
    <div className="card p-4 flex items-center gap-3">
      <div className="flex items-center justify-center shrink-0" style={{ width: 40, height: 40, borderRadius: 8, background: colors.bg, color: colors.fg }}>
        <Icon size={19} />
      </div>
      <div>
        <div style={{ fontSize: 12, color: "var(--steel)" }}>{label}</div>
        <div className="mono" style={{ fontWeight: 600, fontSize: 17 }}>{value}</div>
      </div>
    </div>
  );
}

function statusBadge(status) {
  const map = {
    draft: { bg: "#EAEAE3", fg: "var(--ink-soft)", label: "ร่าง" },
    sent: { bg: "#DEEAF6", fg: "#2A5F8A", label: "ส่งแล้ว" },
    issued: { bg: "#DEEAF6", fg: "#2A5F8A", label: "ออกบิลแล้ว" },
    accepted: { bg: "var(--green-bg)", fg: "var(--green)", label: "ลูกค้าตอบรับ" },
    completed: { bg: "var(--green-bg)", fg: "var(--green)", label: "ชำระแล้ว" },
    cancelled: { bg: "var(--red-bg)", fg: "var(--red)", label: "ยกเลิก" },
  };
  const s = map[status] || map.draft;
  return <span className="badge" style={{ background: s.bg, color: s.fg }}>{s.label}</span>;
}

const DOC_TYPE_COLOR = {
  quote: { bg: "#EAEAE3", fg: "var(--ink-soft)" },
  invoice: { bg: "#DEEAF6", fg: "#2A5F8A" },
  taxinvoice: { bg: "#E2E6FA", fg: "#2A4B9B" },
  receipt: { bg: "var(--green-bg)", fg: "var(--green)" },
  debitnote: { bg: "#FBF0DF", fg: "var(--amber-dark)" },
  creditnote: { bg: "var(--red-bg)", fg: "var(--red)" },
};

function docTypeBadge(docType) {
  const c = DOC_TYPE_COLOR[docType] || DOC_TYPE_COLOR.quote;
  return <span className="badge" style={{ background: c.bg, color: c.fg }}>{DOC_TYPE_LABEL[docType] || docType}</span>;
}

/* --------------------------------- dashboard -------------------------------- */

function Dashboard({ documents, transactions, products, customers, settings }) {
  const thisMonth = monthKey(today());
  const monthTx = transactions.filter((t) => monthKey(t.date) === thisMonth);
  const income = monthTx.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const expense = monthTx.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);
  const outstanding = documents
    .filter((d) => (d.docType === "invoice" || d.docType === "taxinvoice") && !["completed", "cancelled"].includes(d.status))
    .reduce((s, d) => s + docBalanceDue(d), 0);
  const lowStock = products.filter((p) => p.stock <= (p.lowStockThreshold ?? settings.lowStockDefault));

  const chartData = useMemo(() => {
    const months = [];
    const d = new Date();
    for (let i = 5; i >= 0; i--) {
      const dt = new Date(d.getFullYear(), d.getMonth() - i, 1);
      months.push(`${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}`);
    }
    return months.map((key) => {
      const txs = transactions.filter((t) => monthKey(t.date) === key);
      return {
        month: monthLabel(key),
        รายรับ: txs.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0),
        รายจ่าย: txs.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0),
      };
    });
  }, [transactions]);

  const topProducts = useMemo(() => {
    const rev = {};
    documents.filter((d) => d.docType === "invoice" && d.status === "completed").forEach((d) => {
      d.items.forEach((it) => {
        rev[it.name] = (rev[it.name] || 0) + it.qty * it.price;
      });
    });
    return Object.entries(rev).sort((a, b) => b[1] - a[1]).slice(0, 5);
  }, [documents]);

  return (
    <div>
      <SectionHeader title="แดชบอร์ด" subtitle={`ภาพรวมประจำเดือน ${monthLabel(thisMonth)}`} />
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
        <StatCard label="รายรับเดือนนี้" value={thb(income)} icon={TrendingUp} tone="green" />
        <StatCard label="รายจ่ายเดือนนี้" value={thb(expense)} icon={TrendingDown} tone="red" />
        <StatCard label="กำไรสุทธิเดือนนี้" value={thb(income - expense)} icon={Wallet} tone="amber" />
        <StatCard label="ค้างชำระ (ใบแจ้งหนี้)" value={thb(outstanding)} icon={FileText} tone="ink" />
      </div>

      <div className="grid md:grid-cols-3 gap-4">
        <div className="card p-4 md:col-span-2">
          <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 12 }}>รายรับ-รายจ่ายย้อนหลัง 6 เดือน</div>
          <div style={{ height: 260 }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chartData} barGap={4}>
                <CartesianGrid strokeDasharray="3 3" stroke="var(--border)" vertical={false} />
                <XAxis dataKey="month" tick={{ fontSize: 12, fill: "var(--steel)" }} axisLine={{ stroke: "var(--border)" }} tickLine={false} />
                <YAxis tick={{ fontSize: 11, fill: "var(--steel)" }} axisLine={false} tickLine={false} width={40} />
                <Tooltip formatter={(v) => thb(v)} contentStyle={{ fontSize: 12, borderRadius: 8, border: "1px solid var(--border)" }} />
                <Bar dataKey="รายรับ" fill="var(--green)" radius={[4, 4, 0, 0]} />
                <Bar dataKey="รายจ่าย" fill="var(--red)" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="card p-4">
          <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 12 }}>สต็อกใกล้หมด</div>
          {lowStock.length === 0 ? (
            <p style={{ fontSize: 13, color: "var(--steel)" }}>ไม่มีสินค้าใกล้หมดสต็อกในตอนนี้</p>
          ) : (
            <div className="flex flex-col gap-2">
              {lowStock.slice(0, 6).map((p) => (
                <div key={p.id} className="flex items-center justify-between" style={{ fontSize: 13 }}>
                  <span>{p.name}</span>
                  <span className="badge" style={{ background: "var(--red-bg)", color: "var(--red)" }}>{p.stock} {p.unit}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      {topProducts.length > 0 && (
        <div className="card p-4 mt-4">
          <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 12 }}>สินค้า/บริการขายดี (ตามยอดขาย)</div>
          <div className="flex flex-col gap-2">
            {topProducts.map(([name, rev]) => (
              <div key={name} className="flex items-center justify-between" style={{ fontSize: 13 }}>
                <span>{name}</span>
                <span className="mono" style={{ fontWeight: 600 }}>{thb(rev)}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {customers.length === 0 && documents.length === 0 && (
        <div className="mt-4">
          <Empty text="ยังไม่มีข้อมูล เริ่มจากเพิ่มลูกค้าและออกใบเสนอราคา/ใบแจ้งหนี้แรกของคุณได้ที่เมนู “เอกสารขาย”" />
        </div>
      )}
    </div>
  );
}

/* -------------------------------- customers -------------------------------- */

function CustomersTab({ customers, updateCustomers, showToast, confirmAction }) {
  const [modal, setModal] = useState(null); // {mode:'add'|'edit', data}
  const [q, setQ] = useState("");
  const filtered = customers.filter((c) => c.name.toLowerCase().includes(q.toLowerCase()));

  function save(form) {
    if (!form.name.trim()) return showToast("กรุณากรอกชื่อลูกค้า", "error");
    if (modal.mode === "add") {
      updateCustomers([{ ...form, id: uid() }, ...customers]);
      showToast("เพิ่มลูกค้าเรียบร้อย");
    } else {
      updateCustomers(customers.map((c) => (c.id === form.id ? form : c)));
      showToast("แก้ไขลูกค้าเรียบร้อย");
    }
    setModal(null);
  }

  function remove(id) {
    confirmAction("ลบลูกค้ารายนี้?", () => {
      updateCustomers(customers.filter((c) => c.id !== id));
      showToast("ลบลูกค้าแล้ว");
    });
  }

  return (
    <div>
      <SectionHeader
        title="ลูกค้า"
        subtitle={`ทั้งหมด ${customers.length} ราย`}
        action={
          <button className="btn-primary flex items-center gap-1.5 px-4 py-2 text-sm" onClick={() => setModal({ mode: "add", data: { name: "", contact: "", phone: "", address: "", taxId: "" } })}>
            <Plus size={16} /> เพิ่มลูกค้า
          </button>
        }
      />
      <div className="mb-4 flex items-center gap-2 card px-3 py-2" style={{ maxWidth: 320 }}>
        <Search size={15} color="var(--steel)" />
        <input className="text-sm flex-1" style={{ border: "none", outline: "none" }} placeholder="ค้นหาลูกค้า" value={q} onChange={(e) => setQ(e.target.value)} />
      </div>

      {filtered.length === 0 ? (
        <Empty text="ยังไม่มีลูกค้า เพิ่มลูกค้ารายแรกของคุณเพื่อเริ่มออกใบเสนอราคาหรือใบแจ้งหนี้" />
      ) : (
        <div className="card overflow-x-auto">
          <table>
            <thead><tr><th>ชื่อ</th><th>ผู้ติดต่อ</th><th>โทรศัพท์</th><th>เลขผู้เสียภาษี</th><th></th></tr></thead>
            <tbody>
              {filtered.map((c) => (
                <tr key={c.id}>
                  <td style={{ fontWeight: 500 }}>{c.name}</td>
                  <td>{c.contact || "-"}</td>
                  <td>{c.phone || "-"}</td>
                  <td className="mono">{c.taxId || "-"}</td>
                  <td>
                    <div className="flex gap-2 justify-end">
                      <button className="btn-ghost p-1.5" onClick={() => setModal({ mode: "edit", data: c })}><Pencil size={14} /></button>
                      <button className="btn-ghost btn-danger p-1.5" onClick={() => remove(c.id)}><Trash2 size={14} /></button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {modal && (
        <Modal title={modal.mode === "add" ? "เพิ่มลูกค้า" : "แก้ไขลูกค้า"} onClose={() => setModal(null)}>
          <CustomerForm data={modal.data} onSave={save} />
        </Modal>
      )}
    </div>
  );
}

function CustomerForm({ data, onSave }) {
  const [f, setF] = useState(data);
  return (
    <div>
      <Field label="ชื่อลูกค้า / บริษัท *"><input className={inputCls} value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} /></Field>
      <Field label="ชื่อผู้ติดต่อ"><input className={inputCls} value={f.contact} onChange={(e) => setF({ ...f, contact: e.target.value })} /></Field>
      <Field label="เบอร์โทรศัพท์"><input className={inputCls} value={f.phone} onChange={(e) => setF({ ...f, phone: e.target.value })} /></Field>
      <Field label="เลขประจำตัวผู้เสียภาษี"><input className={inputCls} value={f.taxId} onChange={(e) => setF({ ...f, taxId: e.target.value })} /></Field>
      <Field label="ที่อยู่"><textarea className={inputCls} rows={2} value={f.address} onChange={(e) => setF({ ...f, address: e.target.value })} /></Field>
      <button className="btn-primary w-full py-2.5 text-sm mt-2" onClick={() => onSave(f)}>บันทึก</button>
    </div>
  );
}

/* -------------------------------- inventory --------------------------------- */

function InventoryTab({ products, updateProducts, settings, showToast, confirmAction }) {
  const [modal, setModal] = useState(null);
  const [q, setQ] = useState("");
  const filtered = products.filter((p) => p.name.toLowerCase().includes(q.toLowerCase()));

  function save(form) {
    if (!form.name.trim()) return showToast("กรุณากรอกชื่อสินค้า", "error");
    const clean = { ...form, price: Number(form.price) || 0, stock: Number(form.stock) || 0, lowStockThreshold: form.lowStockThreshold === "" ? null : Number(form.lowStockThreshold) };
    if (modal.mode === "add") {
      updateProducts([{ ...clean, id: uid() }, ...products]);
      showToast("เพิ่มสินค้าเรียบร้อย");
    } else {
      updateProducts(products.map((p) => (p.id === clean.id ? clean : p)));
      showToast("แก้ไขสินค้าเรียบร้อย");
    }
    setModal(null);
  }

  function remove(id) {
    confirmAction("ลบสินค้ารายการนี้?", () => {
      updateProducts(products.filter((p) => p.id !== id));
      showToast("ลบสินค้าแล้ว");
    });
  }

  return (
    <div>
      <SectionHeader
        title="สต็อกสินค้า"
        subtitle={`ทั้งหมด ${products.length} รายการ`}
        action={
          <button className="btn-primary flex items-center gap-1.5 px-4 py-2 text-sm" onClick={() => setModal({ mode: "add", data: { name: "", sku: "", unit: "ชิ้น", price: "", stock: "", lowStockThreshold: "", imageUrl: "" } })}>
            <Plus size={16} /> เพิ่มสินค้า
          </button>
        }
      />
      <div className="mb-4 flex items-center gap-2 card px-3 py-2" style={{ maxWidth: 320 }}>
        <Search size={15} color="var(--steel)" />
        <input className="text-sm flex-1" style={{ border: "none", outline: "none" }} placeholder="ค้นหาสินค้า" value={q} onChange={(e) => setQ(e.target.value)} />
      </div>

      {filtered.length === 0 ? (
        <Empty text="ยังไม่มีสินค้าในสต็อก เพิ่มสินค้าเพื่อใช้อ้างอิงในใบเสนอราคา/ใบแจ้งหนี้และติดตามจำนวนคงเหลือ" />
      ) : (
        <div className="card overflow-x-auto">
          <table>
            <thead><tr><th style={{ width: 48 }}></th><th>สินค้า</th><th>SKU</th><th>ราคา/หน่วย</th><th>คงเหลือ</th><th></th></tr></thead>
            <tbody>
              {filtered.map((p) => {
                const low = p.stock <= (p.lowStockThreshold ?? settings.lowStockDefault);
                return (
                  <tr key={p.id}>
                    <td>
                      {p.imageUrl ? (
                        <img src={p.imageUrl} alt="" style={{ width: 32, height: 32, objectFit: "cover", borderRadius: 6 }} />
                      ) : (
                        <div style={{ width: 32, height: 32, borderRadius: 6, background: "var(--paper)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                          <Package size={14} color="var(--steel-light)" />
                        </div>
                      )}
                    </td>
                    <td style={{ fontWeight: 500 }}>{p.name}</td>
                    <td className="mono">{p.sku || "-"}</td>
                    <td className="mono">{thb(p.price)}</td>
                    <td>
                      <span className="badge" style={{ background: low ? "var(--red-bg)" : "var(--green-bg)", color: low ? "var(--red)" : "var(--green)" }}>
                        {p.stock} {p.unit}
                      </span>
                    </td>
                    <td>
                      <div className="flex gap-2 justify-end">
                        <button className="btn-ghost p-1.5" onClick={() => setModal({ mode: "edit", data: { ...p, price: String(p.price), stock: String(p.stock), lowStockThreshold: p.lowStockThreshold ?? "" } })}><Pencil size={14} /></button>
                        <button className="btn-ghost btn-danger p-1.5" onClick={() => remove(p.id)}><Trash2 size={14} /></button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {modal && (
        <Modal title={modal.mode === "add" ? "เพิ่มสินค้า" : "แก้ไขสินค้า"} onClose={() => setModal(null)}>
          <ProductForm data={modal.data} onSave={save} />
        </Modal>
      )}
    </div>
  );
}

function ProductForm({ data, onSave }) {
  const [f, setF] = useState(data);
  const [imgLoading, setImgLoading] = useState(false);

  async function handleImage(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    setImgLoading(true);
    try {
      const dataUrl = await resizeImageToDataUrl(file);
      setF((cur) => ({ ...cur, imageUrl: dataUrl }));
    } finally {
      setImgLoading(false);
    }
  }

  return (
    <div>
      <div className="flex items-center gap-3 mb-3">
        {f.imageUrl ? (
          <img src={f.imageUrl} alt="" style={{ width: 56, height: 56, objectFit: "cover", borderRadius: 8 }} />
        ) : (
          <div style={{ width: 56, height: 56, borderRadius: 8, background: "var(--paper)", display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Package size={22} color="var(--steel-light)" />
          </div>
        )}
        <div>
          <label className="btn-ghost px-3 py-1.5 text-sm" style={{ cursor: "pointer", display: "inline-block" }}>
            {imgLoading ? "กำลังโหลด..." : "อัปโหลดรูปสินค้า"}
            <input type="file" accept="image/*" style={{ display: "none" }} onChange={handleImage} />
          </label>
          {f.imageUrl && (
            <button className="btn-ghost btn-danger px-3 py-1.5 text-sm ml-2" onClick={() => setF({ ...f, imageUrl: "" })}>ลบรูป</button>
          )}
        </div>
      </div>
      <Field label="ชื่อสินค้า/บริการ *"><input className={inputCls} value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} /></Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="SKU / รหัส"><input className={inputCls} value={f.sku} onChange={(e) => setF({ ...f, sku: e.target.value })} /></Field>
        <Field label="หน่วยนับ"><input className={inputCls} value={f.unit} onChange={(e) => setF({ ...f, unit: e.target.value })} /></Field>
        <Field label="ราคาต่อหน่วย (บาท)"><input type="number" step="0.01" className={inputCls} value={f.price} onChange={(e) => setF({ ...f, price: e.target.value })} /></Field>
        <Field label="จำนวนคงเหลือ"><input type="number" className={inputCls} value={f.stock} onChange={(e) => setF({ ...f, stock: e.target.value })} /></Field>
      </div>
      <Field label="แจ้งเตือนเมื่อคงเหลือต่ำกว่า (ค่าว่าง = ใช้ค่าเริ่มต้น)"><input type="number" className={inputCls} value={f.lowStockThreshold} onChange={(e) => setF({ ...f, lowStockThreshold: e.target.value })} /></Field>
      <button className="btn-primary w-full py-2.5 text-sm mt-2" onClick={() => onSave(f)}>บันทึก</button>
    </div>
  );
}

/* ---------------------------------- assets ----------------------------------- */

function monthsBetween(from, to) {
  const f = new Date(from);
  const t = new Date(to);
  return Math.max(0, (t.getFullYear() - f.getFullYear()) * 12 + (t.getMonth() - f.getMonth()));
}

function assetDepInfo(a) {
  const lifeMonths = Math.max(1, Number(a.usefulLifeYears) || 1) * 12;
  const cost = Number(a.cost) || 0;
  const salvage = Number(a.salvageValue) || 0;
  const depreciable = Math.max(0, cost - salvage);
  const monthlyDep = depreciable / lifeMonths;
  const monthsElapsed = Math.min(monthsBetween(a.purchaseDate, today()) + 1, lifeMonths);
  const accumulated = a.disposed ? depreciable : Math.min(monthlyDep * monthsElapsed, depreciable);
  const bookValue = cost - accumulated;
  return { monthlyDep, accumulated, bookValue, isFullyDepreciated: monthsElapsed >= lifeMonths };
}

const ASSET_CATEGORIES = ["เครื่องจักร/อุปกรณ์", "ยานพาหนะ", "เครื่องใช้สำนักงาน", "คอมพิวเตอร์/IT", "อาคาร/สิ่งปลูกสร้าง", "อื่นๆ"];

function AssetsTab({ assets, updateAssets, transactions, updateTransactions, showToast, confirmAction }) {
  const [modal, setModal] = useState(null);
  const activeAssets = assets.filter((a) => !a.disposed);
  const totalBookValue = activeAssets.reduce((s, a) => s + assetDepInfo(a).bookValue, 0);
  const totalCost = activeAssets.reduce((s, a) => s + (Number(a.cost) || 0), 0);
  const monthlyDepTotal = activeAssets.reduce((s, a) => (assetDepInfo(a).isFullyDepreciated ? s : s + assetDepInfo(a).monthlyDep), 0);

  function save(form) {
    if (!form.name.trim()) return showToast("กรุณากรอกชื่อสินทรัพย์", "error");
    const clean = { ...form, cost: Number(form.cost) || 0, salvageValue: Number(form.salvageValue) || 0, usefulLifeYears: Number(form.usefulLifeYears) || 1 };
    if (modal.mode === "add") {
      updateAssets([{ ...clean, id: uid() }, ...assets]);
      showToast("เพิ่มสินทรัพย์เรียบร้อย");
    } else {
      updateAssets(assets.map((a) => (a.id === clean.id ? clean : a)));
      showToast("แก้ไขสินทรัพย์เรียบร้อย");
    }
    setModal(null);
  }

  function remove(id) {
    confirmAction("ลบสินทรัพย์รายการนี้?", () => {
      updateAssets(assets.filter((a) => a.id !== id));
      showToast("ลบสินทรัพย์แล้ว");
    });
  }

  function postMonthlyDepreciation() {
    const mKey = monthKey(today());
    const marker = `depr-${mKey}`;
    if (transactions.some((t) => t.sourceDocId === marker)) return showToast("บันทึกค่าเสื่อมราคาของเดือนนี้ไปแล้ว", "error");
    if (monthlyDepTotal <= 0) return showToast("ไม่มีค่าเสื่อมราคาที่ต้องบันทึกในเดือนนี้", "error");
    const tx = { id: uid(), date: today(), type: "expense", category: "ค่าเสื่อมราคา", amount: monthlyDepTotal, note: `ค่าเสื่อมราคาประจำเดือน ${monthLabel(mKey)}`, sourceDocId: marker };
    updateTransactions([tx, ...transactions]);
    showToast("บันทึกค่าเสื่อมราคาประจำเดือนเรียบร้อย");
  }

  return (
    <div>
      <SectionHeader
        title="สินทรัพย์ถาวร"
        subtitle="ทะเบียนสินทรัพย์และค่าเสื่อมราคาแบบเส้นตรง (Straight-line)"
        action={
          <div className="flex gap-2 flex-wrap">
            <button className="btn-ghost flex items-center gap-1.5 px-4 py-2 text-sm" onClick={postMonthlyDepreciation}>
              <Wallet size={16} /> บันทึกค่าเสื่อมราคาประจำเดือน
            </button>
            <button className="btn-primary flex items-center gap-1.5 px-4 py-2 text-sm" onClick={() => setModal({ mode: "add", data: { name: "", category: ASSET_CATEGORIES[0], purchaseDate: today(), cost: "", salvageValue: "0", usefulLifeYears: "5", note: "" } })}>
              <Plus size={16} /> เพิ่มสินทรัพย์
            </button>
          </div>
        }
      />

      <div className="grid grid-cols-2 md:grid-cols-3 gap-3 mb-6">
        <StatCard label="มูลค่าทุนรวม" value={thb(totalCost)} icon={Landmark} tone="ink" />
        <StatCard label="มูลค่าตามบัญชีรวม" value={thb(totalBookValue)} icon={TrendingDown} tone="amber" />
        <StatCard label="ค่าเสื่อมราคา/เดือน" value={thb(monthlyDepTotal)} icon={Wallet} tone="red" />
      </div>

      {assets.length === 0 ? (
        <Empty text="ยังไม่มีสินทรัพย์ถาวร เพิ่มรายการ เช่น เครื่องจักร รถยนต์ อุปกรณ์สำนักงาน เพื่อคำนวณค่าเสื่อมราคาอัตโนมัติ" />
      ) : (
        <div className="card overflow-x-auto">
          <table>
            <thead><tr><th>ชื่อสินทรัพย์</th><th>หมวดหมู่</th><th>วันที่ซื้อ</th><th>ราคาทุน</th><th>ค่าเสื่อมสะสม</th><th>มูลค่าตามบัญชี</th><th></th></tr></thead>
            <tbody>
              {assets.map((a) => {
                const info = assetDepInfo(a);
                return (
                  <tr key={a.id}>
                    <td style={{ fontWeight: 500 }}>{a.name}{a.disposed && <span className="badge ml-2" style={{ background: "var(--red-bg)", color: "var(--red)" }}>จำหน่ายแล้ว</span>}</td>
                    <td>{a.category}</td>
                    <td>{fmtDate(a.purchaseDate)}</td>
                    <td className="mono">{thb(a.cost)}</td>
                    <td className="mono">{thb(info.accumulated)}</td>
                    <td className="mono" style={{ fontWeight: 600 }}>{thb(info.bookValue)}</td>
                    <td>
                      <div className="flex gap-2 justify-end">
                        <button className="btn-ghost p-1.5" onClick={() => setModal({ mode: "edit", data: { ...a, cost: String(a.cost), salvageValue: String(a.salvageValue), usefulLifeYears: String(a.usefulLifeYears) } })}><Pencil size={14} /></button>
                        <button className="btn-ghost btn-danger p-1.5" onClick={() => remove(a.id)}><Trash2 size={14} /></button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {modal && (
        <Modal title={modal.mode === "add" ? "เพิ่มสินทรัพย์" : "แก้ไขสินทรัพย์"} onClose={() => setModal(null)}>
          <AssetForm data={modal.data} onSave={save} />
        </Modal>
      )}
    </div>
  );
}

function AssetForm({ data, onSave }) {
  const [f, setF] = useState(data);
  return (
    <div>
      <Field label="ชื่อสินทรัพย์ *"><input className={inputCls} value={f.name} onChange={(e) => setF({ ...f, name: e.target.value })} /></Field>
      <Field label="หมวดหมู่">
        <select className={inputCls} value={f.category} onChange={(e) => setF({ ...f, category: e.target.value })}>
          {ASSET_CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>
      </Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="วันที่ซื้อ"><input type="date" className={inputCls} value={f.purchaseDate} onChange={(e) => setF({ ...f, purchaseDate: e.target.value })} /></Field>
        <Field label="ราคาทุน (บาท) *"><input type="number" step="0.01" className={inputCls} value={f.cost} onChange={(e) => setF({ ...f, cost: e.target.value })} /></Field>
        <Field label="มูลค่าซาก (บาท)"><input type="number" step="0.01" className={inputCls} value={f.salvageValue} onChange={(e) => setF({ ...f, salvageValue: e.target.value })} /></Field>
        <Field label="อายุการใช้งาน (ปี)"><input type="number" step="1" className={inputCls} value={f.usefulLifeYears} onChange={(e) => setF({ ...f, usefulLifeYears: e.target.value })} /></Field>
      </div>
      <Field label="หมายเหตุ"><textarea className={inputCls} rows={2} value={f.note} onChange={(e) => setF({ ...f, note: e.target.value })} /></Field>
      <label className="flex items-center gap-2 mb-3" style={{ fontSize: 13 }}>
        <input type="checkbox" checked={f.disposed || false} onChange={(e) => setF({ ...f, disposed: e.target.checked })} /> จำหน่าย/เลิกใช้งานแล้ว
      </label>
      <button className="btn-primary w-full py-2.5 text-sm mt-2" onClick={() => onSave(f)}>บันทึก</button>
    </div>
  );
}

/* ------------------------------- transactions -------------------------------- */

function TransactionsTab({ transactions, updateTransactions, settings, showToast, confirmAction }) {
  const [modal, setModal] = useState(null);
  const [certTx, setCertTx] = useState(null);
  const [filter, setFilter] = useState("all");

  const filtered = transactions
    .filter((t) => filter === "all" || t.type === filter)
    .sort((a, b) => (a.date < b.date ? 1 : -1));

  const income = transactions.filter((t) => t.type === "income").reduce((s, t) => s + t.amount, 0);
  const expense = transactions.filter((t) => t.type === "expense").reduce((s, t) => s + t.amount, 0);

  function save(form) {
    const amount = Number(form.amount);
    if (!amount || amount <= 0) return showToast("กรุณากรอกจำนวนเงินให้ถูกต้อง", "error");
    if (modal.mode === "add") {
      updateTransactions([{ ...form, id: uid(), amount }, ...transactions]);
      showToast("บันทึกรายการเรียบร้อย");
    } else {
      updateTransactions(transactions.map((t) => (t.id === form.id ? { ...form, amount } : t)));
      showToast("แก้ไขรายการเรียบร้อย");
    }
    setModal(null);
  }

  function remove(id) {
    confirmAction("ลบรายการนี้?", () => {
      updateTransactions(transactions.filter((t) => t.id !== id));
      showToast("ลบรายการแล้ว");
    });
  }

  return (
    <div>
      <SectionHeader
        title="รายรับ-รายจ่าย"
        subtitle={`รายรับรวม ${thb(income)} · รายจ่ายรวม ${thb(expense)}`}
        action={
          <button className="btn-primary flex items-center gap-1.5 px-4 py-2 text-sm" onClick={() => setModal({ mode: "add", data: { date: today(), type: "income", category: INCOME_CATEGORIES[0], amount: "", note: "" } })}>
            <Plus size={16} /> เพิ่มรายการ
          </button>
        }
      />

      <div className="flex gap-2 mb-4">
        {[["all", "ทั้งหมด"], ["income", "รายรับ"], ["expense", "รายจ่าย"]].map(([k, l]) => (
          <button key={k} onClick={() => setFilter(k)} className="btn-ghost px-3 py-1.5 text-sm" style={filter === k ? { background: "var(--ink)", color: "#fff", borderColor: "var(--ink)" } : {}}>
            {l}
          </button>
        ))}
      </div>

      {filtered.length === 0 ? (
        <Empty text="ยังไม่มีรายการรายรับ-รายจ่าย เพิ่มรายการแรกเพื่อเริ่มติดตามกระแสเงินสด" />
      ) : (
        <div className="card overflow-x-auto">
          <table>
            <thead><tr><th>วันที่</th><th>ประเภท</th><th>หมวดหมู่</th><th>บันทึก</th><th>จำนวนเงิน</th><th></th></tr></thead>
            <tbody>
              {filtered.map((t) => (
                <tr key={t.id}>
                  <td>{fmtDate(t.date)}</td>
                  <td>
                    <span className="badge" style={{ background: t.type === "income" ? "var(--green-bg)" : "var(--red-bg)", color: t.type === "income" ? "var(--green)" : "var(--red)" }}>
                      {t.type === "income" ? "รายรับ" : "รายจ่าย"}
                    </span>
                  </td>
                  <td>{t.category}</td>
                  <td style={{ color: "var(--steel)" }}>{t.note || "-"}</td>
                  <td className="mono" style={{ fontWeight: 600, color: t.type === "income" ? "var(--green)" : "var(--red)" }}>
                    {t.type === "income" ? "+" : "-"}{thb(t.amount)}
                  </td>
                  <td>
                    <div className="flex gap-2 justify-end">
                      {t.hasWht && <button className="btn-ghost p-1.5" title="พิมพ์หนังสือรับรองหัก ณ ที่จ่าย (50 ทวิ)" onClick={() => setCertTx(t)}><FileText size={14} /></button>}
                      {!t.sourceDocId && <button className="btn-ghost p-1.5" onClick={() => setModal({ mode: "edit", data: { ...t, amount: String(t.amount) } })}><Pencil size={14} /></button>}
                      <button className="btn-ghost btn-danger p-1.5" onClick={() => remove(t.id)}><Trash2 size={14} /></button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {modal && (
        <Modal title={modal.mode === "add" ? "เพิ่มรายการ" : "แก้ไขรายการ"} onClose={() => setModal(null)}>
          <TransactionForm data={modal.data} onSave={save} />
        </Modal>
      )}

      {certTx && <WHTCertificatePrintView tx={certTx} settings={settings} onClose={() => setCertTx(null)} />}
    </div>
  );
}

function TransactionForm({ data, onSave }) {
  const [f, setF] = useState({ hasWht: false, whtType: "service", whtRate: 3, whtBase: "", payeeName: "", payeeTaxId: "", payeeAddress: "", hasInputVat: false, inputVat: "", ...data });
  const cats = f.type === "income" ? INCOME_CATEGORIES : EXPENSE_CATEGORIES;
  const whtAmount = f.hasWht ? (Number(f.whtBase || 0) * Number(f.whtRate || 0)) / 100 : 0;

  function pickWhtType(code) {
    const t = WHT_TYPES.find((w) => w.code === code);
    setF({ ...f, whtType: code, whtRate: t ? t.rate : f.whtRate });
  }

  function handleSave() {
    onSave({
      ...f,
      ...(f.hasWht
        ? { whtRate: Number(f.whtRate) || 0, whtBase: Number(f.whtBase) || 0, whtAmount }
        : { hasWht: false, whtType: undefined, whtRate: undefined, whtBase: undefined, whtAmount: undefined, payeeName: undefined, payeeTaxId: undefined, payeeAddress: undefined }),
      ...(f.hasInputVat ? { inputVat: Number(f.inputVat) || 0 } : { hasInputVat: false, inputVat: undefined }),
    });
  }

  return (
    <div>
      <div className="grid grid-cols-2 gap-3">
        <Field label="ประเภท">
          <select className={inputCls} value={f.type} onChange={(e) => setF({ ...f, type: e.target.value, category: e.target.value === "income" ? INCOME_CATEGORIES[0] : EXPENSE_CATEGORIES[0] })}>
            <option value="income">รายรับ</option>
            <option value="expense">รายจ่าย</option>
          </select>
        </Field>
        <Field label="วันที่"><input type="date" className={inputCls} value={f.date} onChange={(e) => setF({ ...f, date: e.target.value })} /></Field>
      </div>
      <Field label="หมวดหมู่">
        <select className={inputCls} value={f.category} onChange={(e) => setF({ ...f, category: e.target.value })}>
          {cats.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>
      </Field>
      <Field label="จำนวนเงิน (บาท) *"><input type="number" step="0.01" className={inputCls} value={f.amount} onChange={(e) => setF({ ...f, amount: e.target.value })} /></Field>
      <Field label="บันทึกเพิ่มเติม"><input className={inputCls} value={f.note} onChange={(e) => setF({ ...f, note: e.target.value })} /></Field>

      {f.type === "expense" && (
        <>
          <label className="flex items-center gap-2 mb-3" style={{ fontSize: 13 }}>
            <input type="checkbox" checked={f.hasInputVat} onChange={(e) => setF({ ...f, hasInputVat: e.target.checked })} /> มีภาษีซื้อ (ใบกำกับภาษีจากผู้ขาย)
          </label>
          {f.hasInputVat && (
            <Field label="ภาษีซื้อ (บาท)"><input type="number" step="0.01" className={inputCls} value={f.inputVat} onChange={(e) => setF({ ...f, inputVat: e.target.value })} /></Field>
          )}

          <label className="flex items-center gap-2 mb-3" style={{ fontSize: 13 }}>
            <input type="checkbox" checked={f.hasWht} onChange={(e) => setF({ ...f, hasWht: e.target.checked })} /> มีการหักภาษี ณ ที่จ่าย
          </label>
          {f.hasWht && (
            <div className="card p-3 mb-3" style={{ background: "var(--paper)" }}>
              <div className="grid grid-cols-2 gap-3">
                <Field label="ประเภทเงินได้">
                  <select className={inputCls} value={f.whtType} onChange={(e) => pickWhtType(e.target.value)}>
                    {WHT_TYPES.map((w) => <option key={w.code} value={w.code}>{w.label} ({w.rate}%)</option>)}
                  </select>
                </Field>
                <Field label="อัตราภาษี (%)"><input type="number" step="0.01" className={inputCls} value={f.whtRate} onChange={(e) => setF({ ...f, whtRate: e.target.value })} /></Field>
              </div>
              <Field label="ยอดเงินได้ที่จ่าย (ก่อนหักภาษี)"><input type="number" step="0.01" className={inputCls} value={f.whtBase} onChange={(e) => setF({ ...f, whtBase: e.target.value })} /></Field>
              <div className="flex justify-between mb-3" style={{ fontSize: 13, fontWeight: 600 }}>
                <span>ภาษีที่หักไว้</span><span className="mono">{thb(whtAmount)}</span>
              </div>
              <Field label="ชื่อผู้ถูกหักภาษี (ผู้รับเงิน)"><input className={inputCls} value={f.payeeName} onChange={(e) => setF({ ...f, payeeName: e.target.value })} /></Field>
              <Field label="เลขประจำตัวผู้เสียภาษีผู้ถูกหักภาษี"><input className={inputCls} value={f.payeeTaxId} onChange={(e) => setF({ ...f, payeeTaxId: e.target.value })} /></Field>
              <Field label="ที่อยู่ผู้ถูกหักภาษี"><textarea className={inputCls} rows={2} value={f.payeeAddress} onChange={(e) => setF({ ...f, payeeAddress: e.target.value })} /></Field>
            </div>
          )}
        </>
      )}

      <button className="btn-primary w-full py-2.5 text-sm mt-2" onClick={handleSave}>บันทึก</button>
    </div>
  );
}

function WHTCertificatePrintView({ tx, settings, onClose }) {
  const whtTypeLabel = WHT_TYPES.find((w) => w.code === tx.whtType)?.label || "อื่นๆ";
  return (
    <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(20,20,18,0.55)", zIndex: 60 }}>
      <div className="card w-full flex flex-col" style={{ maxWidth: 640, maxHeight: "92vh" }}>
        <div className="flex items-center justify-between px-5 py-4 no-print" style={{ borderBottom: "1px solid var(--border)" }}>
          <h2 style={{ fontFamily: "var(--font-display)", fontWeight: 700, fontSize: 17 }}>หนังสือรับรองการหักภาษี ณ ที่จ่าย</h2>
          <div className="flex gap-2">
            <button className="btn-primary flex items-center gap-1.5 px-3 py-1.5 text-sm" onClick={() => window.print()}><Printer size={14} /> พิมพ์ / บันทึก PDF</button>
            <button onClick={onClose} className="btn-ghost p-1"><X size={18} /></button>
          </div>
        </div>
        <div className="p-8 overflow-y-auto print-area">
          <div className="text-center mb-6">
            <div style={{ fontFamily: "var(--font-display)", fontWeight: 700, fontSize: 18 }}>หนังสือรับรองการหักภาษี ณ ที่จ่าย</div>
            <div style={{ fontSize: 12, color: "var(--steel)" }}>ตามมาตรา 50 ทวิ แห่งประมวลรัษฎากร</div>
          </div>

          <div className="grid grid-cols-2 gap-4 mb-4" style={{ fontSize: 13 }}>
            <div className="card p-3">
              <div style={{ color: "var(--steel)", fontSize: 11, textTransform: "uppercase", marginBottom: 4 }}>ผู้จ่ายเงิน</div>
              <div style={{ fontWeight: 600 }}>{settings.companyName}</div>
              <div style={{ color: "var(--steel)", whiteSpace: "pre-line" }}>{settings.address}</div>
              {settings.taxId && <div style={{ color: "var(--steel)" }}>เลขประจำตัวผู้เสียภาษี {settings.taxId}</div>}
            </div>
            <div className="card p-3">
              <div style={{ color: "var(--steel)", fontSize: 11, textTransform: "uppercase", marginBottom: 4 }}>ผู้ถูกหักภาษี ณ ที่จ่าย</div>
              <div style={{ fontWeight: 600 }}>{tx.payeeName || "-"}</div>
              <div style={{ color: "var(--steel)", whiteSpace: "pre-line" }}>{tx.payeeAddress}</div>
              {tx.payeeTaxId && <div style={{ color: "var(--steel)" }}>เลขประจำตัวผู้เสียภาษี {tx.payeeTaxId}</div>}
            </div>
          </div>

          <table style={{ marginBottom: 12 }}>
            <thead><tr><th>วันที่จ่ายเงิน</th><th>ประเภทเงินได้</th><th style={{ textAlign: "right" }}>จำนวนเงินที่จ่าย</th><th style={{ textAlign: "right" }}>อัตรา</th><th style={{ textAlign: "right" }}>ภาษีที่หักไว้</th></tr></thead>
            <tbody>
              <tr>
                <td>{fmtDate(tx.date)}</td>
                <td>{whtTypeLabel}</td>
                <td className="mono" style={{ textAlign: "right" }}>{thb(tx.whtBase)}</td>
                <td className="mono" style={{ textAlign: "right" }}>{tx.whtRate}%</td>
                <td className="mono" style={{ textAlign: "right" }}>{thb(tx.whtAmount)}</td>
              </tr>
            </tbody>
          </table>

          <div className="flex justify-end mb-6">
            <div style={{ minWidth: 220, fontSize: 13 }}>
              <div className="flex justify-between" style={{ fontWeight: 700, fontSize: 15, borderTop: "1px solid var(--border)", paddingTop: 8 }}>
                <span>รวมภาษีที่หักไว้</span><span className="mono">{thb(tx.whtAmount)}</span>
              </div>
            </div>
          </div>

          <div className="grid grid-cols-2 gap-8 mt-10" style={{ fontSize: 12 }}>
            <div className="text-center">
              <div style={{ borderTop: "1px solid var(--ink)", paddingTop: 6, marginTop: 40 }}>ผู้จ่ายเงิน (ลงชื่อ)</div>
              <div style={{ color: "var(--steel)", marginTop: 4 }}>วันที่ {fmtDate(tx.date)}</div>
            </div>
            <div className="text-center">
              <div style={{ borderTop: "1px solid var(--ink)", paddingTop: 6, marginTop: 40 }}>ตราประทับ (ถ้ามี)</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* -------------------------------- documents ---------------------------------- */

function DocumentsTab({ documents, customers, products, settings, updateDocuments, updateSettings, updateTransactions, transactions, updateProducts, showToast, confirmAction }) {
  const [modal, setModal] = useState(null);
  const [viewDoc, setViewDoc] = useState(null);
  const [paymentDoc, setPaymentDoc] = useState(null);
  const [typeFilter, setTypeFilter] = useState("all");
  const customerName = (id) => customers.find((c) => c.id === id)?.name || "ไม่ระบุ";
  const hasWorkflowActions = (d) =>
    (d.docType === "quote" && d.status === "draft") ||
    (d.docType === "quote" && d.status === "sent") ||
    (d.docType === "quote" && d.status === "accepted" && !d.convertedInvoiceId) ||
    (d.docType === "quote" && d.status === "accepted" && !d.convertedTaxInvoiceId) ||
    (d.docType === "invoice" && d.status === "draft");

  const filtered = documents.filter((d) => typeFilter === "all" || d.docType === typeFilter).sort((a, b) => (a.date < b.date ? 1 : -1));

  function genNumber(docType) {
    const map = {
      invoice: ["invoicePrefix", "nextInvoiceNumber"],
      quote: ["quotePrefix", "nextQuoteNumber"],
      taxinvoice: ["taxInvoicePrefix", "nextTaxInvoiceNumber"],
      receipt: ["receiptPrefix", "nextReceiptNumber"],
      debitnote: ["debitNotePrefix", "nextDebitNoteNumber"],
      creditnote: ["creditNotePrefix", "nextCreditNoteNumber"],
    };
    const [prefixField, numField] = map[docType];
    const prefix = settings[prefixField];
    const n = settings[numField];
    const year = new Date().getFullYear();
    updateSettings({ ...settings, [numField]: n + 1 });
    return `${prefix}-${year}-${String(n).padStart(4, "0")}`;
  }

  function saveDoc(form, isNew) {
    if (!form.customerId) return showToast("กรุณาเลือกลูกค้า", "error");
    if (form.items.length === 0) return showToast("กรุณาเพิ่มรายการสินค้า/บริการอย่างน้อย 1 รายการ", "error");
    if ((form.docType === "debitnote" || form.docType === "creditnote")) {
      if (!form.sourceDocId) return showToast("กรุณาเลือกเอกสารต้นทางที่จะอ้างอิง", "error");
      if (!form.adjustReason?.trim()) return showToast("กรุณาระบุเหตุผลที่ออกเอกสาร", "error");
    }
    if (isNew) {
      const docNumber = genNumber(form.docType);
      const doc = { ...form, id: uid(), docNumber, status: "draft", settled: false };
      updateDocuments([doc, ...documents]);
      showToast(`สร้าง${DOC_TYPE_LABEL[form.docType]}เรียบร้อย`);
    } else {
      updateDocuments(documents.map((d) => (d.id === form.id ? form : d)));
      showToast("แก้ไขเอกสารเรียบร้อย");
    }
    setModal(null);
  }

  function setStatus(doc, status) {
    updateDocuments(documents.map((d) => (d.id === doc.id ? { ...d, status } : d)));
    showToast("อัปเดตสถานะแล้ว");
  }

  function recordPayment(doc, { amount, method, note, date }) {
    if (doc.settled) return;
    const amt = Number(amount);
    const balance = docBalanceDue(doc);
    if (!amt || amt <= 0) return showToast("กรุณากรอกจำนวนเงินให้ถูกต้อง", "error");
    if (amt > balance + 0.01) return showToast("จำนวนเงินเกินยอดคงเหลือที่ต้องชำระ", "error");

    const paymentDate = date || today();
    const payment = { id: uid(), date: paymentDate, amount: amt, method: method || "", note: note || "" };
    const tx = { id: uid(), date: paymentDate, type: "income", category: "ขายสินค้า/บริการ", amount: amt, note: `รับชำระตาม ${doc.docNumber}`, sourceDocId: doc.id };
    updateTransactions([tx, ...transactions]);

    const newPayments = [...(doc.payments || []), payment];
    const isNowSettled = docBalanceDue({ ...doc, payments: newPayments }) <= 0.01;
    let updatedDoc = { ...doc, payments: newPayments };

    if (isNowSettled) {
      updatedDoc = { ...updatedDoc, status: "completed", settled: true, paidDate: paymentDate };
      if (doc.docType === "invoice") {
        const updatedProducts = products.map((p) => {
          const item = doc.items.find((it) => it.productId === p.id);
          return item ? { ...p, stock: p.stock - item.qty } : p;
        });
        updateProducts(updatedProducts);
      }
    }
    updateDocuments(documents.map((d) => (d.id === doc.id ? updatedDoc : d)));
    showToast(isNowSettled ? "บันทึกการชำระเงินครบถ้วนเรียบร้อย" : "บันทึกรับชำระบางส่วนเรียบร้อย");
    setPaymentDoc(null);
  }

  function convertQuote(quote, targetType) {
    const docNumber = genNumber(targetType);
    const converted = { ...quote, id: uid(), docType: targetType, docNumber, status: "draft", settled: false, quoteRef: quote.id, date: today() };
    const flagField = targetType === "invoice" ? "convertedInvoiceId" : "convertedTaxInvoiceId";
    updateDocuments([converted, ...documents.map((d) => (d.id === quote.id ? { ...d, [flagField]: converted.id } : d))]);
    showToast(`แปลงเป็น${DOC_TYPE_LABEL[targetType]}เรียบร้อย`);
  }

  function remove(id) {
    confirmAction("ลบเอกสารนี้? (จะไม่ลบรายการรายรับ-รายจ่ายที่เกิดขึ้นแล้ว)", () => {
      updateDocuments(documents.filter((d) => d.id !== id));
      showToast("ลบเอกสารแล้ว");
    });
  }

  return (
    <div>
      <SectionHeader
        title="เอกสารขาย"
        subtitle="ใบเสนอราคาและใบแจ้งหนี้ทั้งหมด"
        action={
          <div className="flex gap-2 flex-wrap">
            <button className="btn-ghost flex items-center gap-1.5 px-4 py-2 text-sm" onClick={() => setModal({ mode: "add", docType: "quote" })}>
              <Plus size={16} /> ใบเสนอราคา
            </button>
            <button className="btn-ghost flex items-center gap-1.5 px-4 py-2 text-sm" onClick={() => setModal({ mode: "add", docType: "invoice" })}>
              <Plus size={16} /> ใบแจ้งหนี้
            </button>
            <button className="btn-ghost flex items-center gap-1.5 px-4 py-2 text-sm" onClick={() => setModal({ mode: "add", docType: "taxinvoice" })}>
              <Plus size={16} /> ใบกำกับภาษี
            </button>
            <button className="btn-ghost flex items-center gap-1.5 px-4 py-2 text-sm" onClick={() => setModal({ mode: "add", docType: "receipt" })}>
              <Plus size={16} /> ใบเสร็จรับเงิน
            </button>
            <button className="btn-ghost flex items-center gap-1.5 px-4 py-2 text-sm" onClick={() => setModal({ mode: "add", docType: "debitnote" })}>
              <Plus size={16} /> ใบเพิ่มหนี้
            </button>
            <button className="btn-primary flex items-center gap-1.5 px-4 py-2 text-sm" onClick={() => setModal({ mode: "add", docType: "creditnote" })}>
              <Plus size={16} /> ใบลดหนี้
            </button>
          </div>
        }
      />

      <div className="flex gap-2 mb-4 flex-wrap">
        {[["all", "ทั้งหมด"], ["quote", "ใบเสนอราคา"], ["invoice", "ใบแจ้งหนี้"], ["taxinvoice", "ใบกำกับภาษี"], ["receipt", "ใบเสร็จรับเงิน"], ["debitnote", "ใบเพิ่มหนี้"], ["creditnote", "ใบลดหนี้"]].map(([k, l]) => (
          <button key={k} onClick={() => setTypeFilter(k)} className="btn-ghost px-3 py-1.5 text-sm" style={typeFilter === k ? { background: "var(--ink)", color: "#fff", borderColor: "var(--ink)" } : {}}>
            {l}
          </button>
        ))}
      </div>

      {filtered.length === 0 ? (
        <Empty text="ยังไม่มีเอกสาร สร้างใบเสนอราคาหรือใบแจ้งหนี้ฉบับแรกของคุณได้เลย" />
      ) : (
        <div className="card overflow-x-auto">
          <table>
            <thead><tr><th>เลขที่</th><th>ประเภท</th><th>ลูกค้า</th><th>วันที่</th><th>ยอดรวม</th><th>สถานะ</th><th></th></tr></thead>
            <tbody>
              {filtered.map((d) => (
                <tr key={d.id}>
                  <td className="mono" style={{ fontWeight: 500 }}>{d.docNumber}</td>
                  <td>{docTypeBadge(d.docType)}</td>
                  <td>{customerName(d.customerId)}</td>
                  <td>{fmtDate(d.date)}</td>
                  <td className="mono" style={{ fontWeight: 600 }}>
                    {thb(d.total)}
                    {(d.docType === "invoice" || d.docType === "taxinvoice") && docPaidAmount(d) > 0 && d.status !== "completed" && (
                      <div style={{ fontSize: 11, fontWeight: 500, color: "var(--amber-dark)" }}>คงเหลือ {thb(docBalanceDue(d))}</div>
                    )}
                  </td>
                  <td>{statusBadge(d.status)}</td>
                  <td>
                    <div className="flex items-center justify-end flex-wrap" style={{ gap: 2 }}>
                      <button className="btn-ghost p-1.5" title="ดู/พิมพ์เอกสาร" onClick={() => setViewDoc(d)}><Printer size={14} color="var(--steel)" /></button>

                      {hasWorkflowActions(d) && (
                        <span className="flex items-center" style={{ gap: 2, borderLeft: "1px solid var(--border)", marginLeft: 4, paddingLeft: 4 }}>
                          {d.docType === "quote" && d.status === "draft" && (
                            <button className="btn-ghost p-1.5" title="ทำเครื่องหมายว่าส่งแล้ว" onClick={() => setStatus(d, "sent")}><Send size={14} color="#2A5F8A" /></button>
                          )}
                          {d.docType === "quote" && d.status === "sent" && (
                            <button className="btn-ghost p-1.5" title="ลูกค้าตอบรับ" onClick={() => setStatus(d, "accepted")}><CheckCircle2 size={14} color="var(--green)" /></button>
                          )}
                          {d.docType === "quote" && d.status === "accepted" && !d.convertedInvoiceId && (
                            <button className="btn-ghost p-1.5" title="แปลงเป็นใบแจ้งหนี้" onClick={() => convertQuote(d, "invoice")}><ArrowRight size={14} color="#2A5F8A" /></button>
                          )}
                          {d.docType === "quote" && d.status === "accepted" && !d.convertedTaxInvoiceId && (
                            <button className="btn-ghost p-1.5" title="แปลงเป็นใบกำกับภาษี" onClick={() => convertQuote(d, "taxinvoice")}><FileCheck2 size={14} color="#2A4B9B" /></button>
                          )}
                          {d.docType === "invoice" && d.status === "draft" && (
                            <button className="btn-ghost p-1.5" title="ออกบิล" onClick={() => setStatus(d, "issued")}><FileCheck2 size={14} color="#2A5F8A" /></button>
                          )}
                        </span>
                      )}

                      {(d.docType === "invoice" || d.docType === "taxinvoice") && !d.settled && d.status !== "cancelled" && (
                        <span className="flex items-center" style={{ borderLeft: "1px solid var(--border)", marginLeft: 4, paddingLeft: 4 }}>
                          <button className="btn-ghost p-1.5" title="รับชำระเงิน" onClick={() => setPaymentDoc(d)}><CheckCircle2 size={14} color="var(--green)" /></button>
                        </span>
                      )}

                      <span className="flex items-center" style={{ gap: 2, borderLeft: "1px solid var(--border)", marginLeft: 4, paddingLeft: 4 }}>
                        {["draft", "sent", "issued"].includes(d.status) && (
                          <button className="btn-ghost p-1.5" title="ยกเลิก" onClick={() => setStatus(d, "cancelled")}><Ban size={14} color="var(--amber-dark)" /></button>
                        )}
                        {d.status === "draft" && !(d.payments?.length > 0) && (
                          <button className="btn-ghost p-1.5" title="แก้ไข" onClick={() => setModal({ mode: "edit", data: d })}><Pencil size={14} color="var(--steel)" /></button>
                        )}
                        <button className="btn-ghost btn-danger p-1.5" title="ลบ" onClick={() => remove(d.id)}><Trash2 size={14} /></button>
                      </span>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {modal && (
        <Modal title={modal.mode === "add" ? `สร้าง${DOC_TYPE_LABEL[modal.docType]}` : "แก้ไขเอกสาร"} onClose={() => setModal(null)} wide>
          <DocumentForm
            existing={modal.data}
            docType={modal.docType || modal.data?.docType}
            customers={customers}
            products={products}
            documents={documents}
            settings={settings}
            onSave={(form) => saveDoc(form, modal.mode === "add")}
          />
        </Modal>
      )}

      {viewDoc && ["taxinvoice", "receipt", "debitnote", "creditnote"].includes(viewDoc.docType) && (
        <TaxReceiptPrintView doc={viewDoc} customers={customers} documents={documents} settings={settings} onClose={() => setViewDoc(null)} />
      )}
      {viewDoc && (viewDoc.docType === "quote" || viewDoc.docType === "invoice") && (
        <DocumentPrintView doc={viewDoc} customers={customers} settings={settings} onClose={() => setViewDoc(null)} />
      )}

      {paymentDoc && (
        <Modal title={`รับชำระเงิน — ${paymentDoc.docNumber}`} onClose={() => setPaymentDoc(null)}>
          <PaymentForm doc={paymentDoc} onSave={(p) => recordPayment(paymentDoc, p)} />
        </Modal>
      )}
    </div>
  );
}

function PaymentForm({ doc, onSave }) {
  const balance = docBalanceDue(doc);
  const [amount, setAmount] = useState(String(balance));
  const [method, setMethod] = useState("cash");
  const [date, setDate] = useState(today());
  const [note, setNote] = useState("");

  return (
    <div>
      <div className="card p-3 mb-3 flex flex-col gap-1" style={{ background: "var(--paper)", fontSize: 13 }}>
        <div className="flex justify-between"><span>ยอดรวมเอกสาร</span><span className="mono">{thb(doc.total)}</span></div>
        <div className="flex justify-between"><span>ชำระแล้ว</span><span className="mono">{thb(docPaidAmount(doc))}</span></div>
        <div className="flex justify-between" style={{ fontWeight: 700 }}><span>คงเหลือ</span><span className="mono">{thb(balance)}</span></div>
      </div>
      <Field label="จำนวนเงินที่รับชำระ (บาท) *"><input type="number" step="0.01" className={inputCls} value={amount} onChange={(e) => setAmount(e.target.value)} /></Field>
      <div className="grid grid-cols-2 gap-3">
        <Field label="วิธีชำระเงิน">
          <select className={inputCls} value={method} onChange={(e) => setMethod(e.target.value)}>
            {Object.entries(PAYMENT_METHOD_LABEL).map(([k, l]) => <option key={k} value={k}>{l}</option>)}
          </select>
        </Field>
        <Field label="วันที่รับชำระ"><input type="date" className={inputCls} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
      </div>
      <Field label="บันทึกเพิ่มเติม"><input className={inputCls} value={note} onChange={(e) => setNote(e.target.value)} /></Field>
      <button className="btn-primary w-full py-2.5 text-sm mt-2" onClick={() => onSave({ amount, method, date, note })}>บันทึกรับชำระเงิน</button>
    </div>
  );
}

function DocumentForm({ existing, docType, customers, products, documents, settings, onSave }) {
  const [customerId, setCustomerId] = useState(existing?.customerId || "");
  const [date, setDate] = useState(existing?.date || today());
  const [dueDate, setDueDate] = useState(existing?.dueDate || "");
  const [includeVat, setIncludeVat] = useState(existing?.includeVat ?? true);
  const [note, setNote] = useState(existing?.note || "");
  const [items, setItems] = useState(existing?.items || []);
  const effectiveType = docType || existing?.docType || "quote";
  const isAdjustNote = effectiveType === "debitnote" || effectiveType === "creditnote";

  const [refNote, setRefNote] = useState(existing?.refNote || "");
  const [jobName, setJobName] = useState(existing?.jobName || "");
  const [contactName, setContactName] = useState(existing?.contactName || "");
  const [contactPhone, setContactPhone] = useState(existing?.contactPhone || "");
  const [paymentMethod, setPaymentMethod] = useState(existing?.paymentMethod || "");
  const [bankName, setBankName] = useState(existing?.bankName || "");
  const [paymentRefNumber, setPaymentRefNumber] = useState(existing?.paymentRefNumber || "");
  const [paymentDate, setPaymentDate] = useState(existing?.paymentDate || "");
  const [sourceDocId, setSourceDocId] = useState(existing?.sourceDocId || "");
  const [adjustReason, setAdjustReason] = useState(existing?.adjustReason || "");
  const [isSimplified, setIsSimplified] = useState(existing?.isSimplified || false);
  const canBeSimplified = effectiveType === "taxinvoice" || effectiveType === "receipt";

  const sourceCandidates = (documents || []).filter(
    (d) => (d.docType === "taxinvoice" || d.docType === "invoice") && (!customerId || d.customerId === customerId)
  );

  function addItem() {
    setItems([...items, { id: uid(), productId: "", name: "", sku: "", imageUrl: "", unit: "", qty: 1, price: 0, discount: 0 }]);
  }
  function updateItem(id, patch) {
    setItems(items.map((it) => (it.id === id ? { ...it, ...patch } : it)));
  }
  function removeItem(id) {
    setItems(items.filter((it) => it.id !== id));
  }
  function pickProduct(id, productId) {
    const p = products.find((x) => x.id === productId);
    updateItem(id, { productId, name: p ? p.name : "", price: p ? p.price : 0, unit: p ? p.unit : "", sku: p ? p.sku : "", imageUrl: p ? p.imageUrl : "" });
  }

  const subtotal = items.reduce((s, it) => s + (Number(it.qty || 0) * Number(it.price || 0) - Number(it.discount || 0)), 0);
  const vatAmount = includeVat ? subtotal * (settings.vatRate / 100) : 0;
  const total = subtotal + vatAmount;

  function handleSave() {
    onSave({
      ...(existing || {}),
      docType: effectiveType,
      customerId, date, dueDate, includeVat, note, items,
      subtotal, vatAmount, total,
      refNote, jobName, contactName, contactPhone,
      ...(effectiveType === "receipt" ? { paymentMethod, bankName, paymentRefNumber, paymentDate } : {}),
      ...(isAdjustNote ? { sourceDocId, adjustReason } : {}),
      ...(canBeSimplified ? { isSimplified } : {}),
    });
  }

  return (
    <div>
      <div className="grid grid-cols-2 gap-3">
        <Field label="ลูกค้า *">
          <select className={inputCls} value={customerId} onChange={(e) => setCustomerId(e.target.value)}>
            <option value="">เลือกลูกค้า</option>
            {customers.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
        </Field>
        <Field label="วันที่"><input type="date" className={inputCls} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
      </div>
      <Field label={effectiveType === "invoice" ? "กำหนดชำระ (ถ้ามี)" : "ยืนราคาถึงวันที่ (ถ้ามี)"}>
        <input type="date" className={inputCls} value={dueDate} onChange={(e) => setDueDate(e.target.value)} />
      </Field>
      {customers.length === 0 && <p style={{ fontSize: 12, color: "var(--red)", marginTop: -8, marginBottom: 10 }}>ยังไม่มีลูกค้าในระบบ กรุณาเพิ่มลูกค้าก่อนที่เมนู "ลูกค้า"</p>}

      <div className="grid grid-cols-2 gap-3">
        <Field label="อ้างอิง (เลขที่ใบเสนอราคา / PO)"><input className={inputCls} value={refNote} onChange={(e) => setRefNote(e.target.value)} /></Field>
        <Field label="ชื่องาน"><input className={inputCls} value={jobName} onChange={(e) => setJobName(e.target.value)} /></Field>
        <Field label="ผู้ติดต่อ"><input className={inputCls} value={contactName} onChange={(e) => setContactName(e.target.value)} /></Field>
        <Field label="เบอร์โทร"><input className={inputCls} value={contactPhone} onChange={(e) => setContactPhone(e.target.value)} /></Field>
      </div>

      {canBeSimplified && (
        <label className="flex items-center gap-2 mb-3" style={{ fontSize: 13 }}>
          <input type="checkbox" checked={isSimplified} onChange={(e) => setIsSimplified(e.target.checked)} />
          ออกเป็นใบกำกับภาษีอย่างย่อ (ไม่แสดงชื่อ-ที่อยู่ผู้ซื้อบนเอกสาร เหมาะกับการขายหน้าร้าน)
        </label>
      )}

      {isAdjustNote && (
        <div className="card p-3 mb-3" style={{ background: "var(--paper)" }}>
          <Field label="อ้างอิงใบกำกับภาษี/ใบแจ้งหนี้ต้นทาง *">
            <select className={inputCls} value={sourceDocId} onChange={(e) => setSourceDocId(e.target.value)}>
              <option value="">เลือกเอกสารต้นทาง</option>
              {sourceCandidates.map((d) => <option key={d.id} value={d.id}>{d.docNumber} · {DOC_TYPE_LABEL[d.docType]} · {thb(d.total)}</option>)}
            </select>
          </Field>
          {sourceCandidates.length === 0 && <p style={{ fontSize: 12, color: "var(--steel)" }}>ยังไม่มีใบกำกับภาษี/ใบแจ้งหนี้ของลูกค้ารายนี้ที่จะอ้างอิงได้</p>}
          <Field label={`เหตุผลที่ออก${DOC_TYPE_LABEL[effectiveType]} *`}>
            <textarea className={inputCls} rows={2} placeholder={effectiveType === "debitnote" ? "เช่น เรียกเก็บเงินเพิ่มเนื่องจากคำนวณราคาผิดพลาด" : "เช่น รับคืนสินค้า / ลดราคาหลังส่งมอบ"} value={adjustReason} onChange={(e) => setAdjustReason(e.target.value)} />
          </Field>
        </div>
      )}

      <div className="flex items-center justify-between mb-2 mt-3">
        <span style={{ fontWeight: 600, fontSize: 13 }}>รายการสินค้า/บริการ</span>
        <button className="btn-ghost text-xs px-2 py-1 flex items-center gap-1" onClick={addItem}><Plus size={13} /> เพิ่มรายการ</button>
      </div>

      <div className="flex flex-col gap-2 mb-3">
        {items.map((it) => (
          <div key={it.id} className="flex gap-2 items-start card p-2" style={{ background: "var(--paper)" }}>
            {it.imageUrl ? (
              <img src={it.imageUrl} alt="" style={{ width: 32, height: 32, objectFit: "cover", borderRadius: 6, flexShrink: 0 }} />
            ) : (
              <div style={{ width: 32, height: 32, borderRadius: 6, background: "var(--surface)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <Package size={14} color="var(--steel-light)" />
              </div>
            )}
            <select className="text-sm" style={{ flex: 2, padding: "6px 8px" }} value={it.productId} onChange={(e) => pickProduct(it.id, e.target.value)}>
              <option value="">กำหนดเอง / เลือกสินค้า</option>
              {products.map((p) => <option key={p.id} value={p.id}>{p.name}</option>)}
            </select>
            {!it.productId && (
              <input className="text-sm" style={{ flex: 2, padding: "6px 8px" }} placeholder="ชื่อรายการ" value={it.name} onChange={(e) => updateItem(it.id, { name: e.target.value })} />
            )}
            <input type="number" className="text-sm mono" style={{ width: 55, padding: "6px 8px" }} placeholder="จำนวน" value={it.qty} onChange={(e) => updateItem(it.id, { qty: e.target.value })} />
            <input className="text-sm" style={{ width: 60, padding: "6px 8px" }} placeholder="หน่วย" value={it.unit} onChange={(e) => updateItem(it.id, { unit: e.target.value })} />
            <input type="number" step="0.01" className="text-sm mono" style={{ width: 90, padding: "6px 8px" }} placeholder="ราคา" value={it.price} onChange={(e) => updateItem(it.id, { price: e.target.value })} />
            <input type="number" step="0.01" className="text-sm mono" style={{ width: 80, padding: "6px 8px" }} placeholder="ส่วนลด" value={it.discount || 0} onChange={(e) => updateItem(it.id, { discount: e.target.value })} />
            <div className="mono text-sm" style={{ width: 90, padding: "6px 0", textAlign: "right", fontWeight: 600 }}>{thb(Number(it.qty || 0) * Number(it.price || 0) - Number(it.discount || 0))}</div>
            <button className="btn-ghost btn-danger p-1.5" onClick={() => removeItem(it.id)}><Trash2 size={13} /></button>
          </div>
        ))}
        {items.length === 0 && <p style={{ fontSize: 13, color: "var(--steel)" }}>ยังไม่มีรายการ กด “เพิ่มรายการ” เพื่อเริ่มต้น</p>}
      </div>

      <label className="flex items-center gap-2 mb-3" style={{ fontSize: 13 }}>
        <input type="checkbox" checked={includeVat} onChange={(e) => setIncludeVat(e.target.checked)} /> คิดภาษีมูลค่าเพิ่ม {settings.vatRate}%
      </label>

      <div className="card p-3 flex flex-col gap-1" style={{ background: "var(--paper)", fontSize: 13 }}>
        <div className="flex justify-between"><span>ยอดรวมก่อนภาษี</span><span className="mono">{thb(subtotal)}</span></div>
        {includeVat && <div className="flex justify-between"><span>ภาษีมูลค่าเพิ่ม ({settings.vatRate}%)</span><span className="mono">{thb(vatAmount)}</span></div>}
        <div className="flex justify-between" style={{ fontWeight: 700, fontSize: 15, paddingTop: 4, borderTop: "1px solid var(--border)" }}><span>ยอดรวมสุทธิ</span><span className="mono">{thb(total)}</span></div>
      </div>

      {effectiveType === "receipt" && (
        <div className="card p-3 mb-3" style={{ background: "var(--paper)" }}>
          <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 8 }}>การรับชำระเงิน</div>
          <Field label="วิธีชำระเงิน">
            <select className={inputCls} value={paymentMethod} onChange={(e) => setPaymentMethod(e.target.value)}>
              <option value="">ไม่ระบุ</option>
              <option value="cash">เงินสด</option>
              <option value="cheque">เช็ค</option>
              <option value="transfer">โอนเงิน</option>
              <option value="credit">บัตรเครดิต</option>
            </select>
          </Field>
          {(paymentMethod === "cheque" || paymentMethod === "transfer") && (
            <div className="grid grid-cols-2 gap-3">
              <Field label="ธนาคาร"><input className={inputCls} value={bankName} onChange={(e) => setBankName(e.target.value)} /></Field>
              <Field label="เลขที่"><input className={inputCls} value={paymentRefNumber} onChange={(e) => setPaymentRefNumber(e.target.value)} /></Field>
              <Field label="วันที่ชำระ"><input type="date" className={inputCls} value={paymentDate} onChange={(e) => setPaymentDate(e.target.value)} /></Field>
            </div>
          )}
        </div>
      )}

      <Field label="หมายเหตุ"><textarea className={inputCls} rows={2} value={note} onChange={(e) => setNote(e.target.value)} /></Field>

      <button className="btn-primary w-full py-2.5 text-sm mt-2" onClick={handleSave}>บันทึกเอกสาร</button>
    </div>
  );
}

function DocumentPrintView({ doc, customers, settings, onClose }) {
  const customer = customers.find((c) => c.id === doc.customerId);
  const isInvoice = doc.docType === "invoice";
  const title = isInvoice ? "ใบแจ้งหนี้ / INVOICE" : "ใบเสนอราคา / QUOTATION";

  return (
    <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(20,20,18,0.55)", zIndex: 60 }}>
      <div className="card w-full flex flex-col" style={{ maxWidth: 760, maxHeight: "92vh" }}>
        <div className="flex items-center justify-between px-5 py-4 no-print" style={{ borderBottom: "1px solid var(--border)" }}>
          <h2 style={{ fontFamily: "var(--font-display)", fontWeight: 700, fontSize: 17 }}>{doc.docNumber}</h2>
          <div className="flex gap-2">
            <button className="btn-primary flex items-center gap-1.5 px-3 py-1.5 text-sm" onClick={() => window.print()}><Printer size={14} /> พิมพ์ / บันทึก PDF</button>
            <button onClick={onClose} className="btn-ghost p-1"><X size={18} /></button>
          </div>
        </div>

        <div className="p-8 overflow-y-auto print-area">
          <div className="stripe" style={{ marginBottom: 20, marginLeft: -32, marginRight: -32, marginTop: -32 }} />

          <div className="flex justify-between items-start mb-8 mt-2">
            <div>
              <div style={{ fontFamily: "var(--font-display)", fontWeight: 700, fontSize: 22 }}>{settings.companyName}</div>
              <div style={{ fontSize: 12, color: "var(--steel)", whiteSpace: "pre-line", marginTop: 4 }}>{settings.address}</div>
              <div style={{ fontSize: 12, color: "var(--steel)" }}>
                {settings.phone && `โทร. ${settings.phone}`} {settings.taxId && `· เลขผู้เสียภาษี ${settings.taxId}`}
              </div>
            </div>
            <div className="text-right">
              <div style={{ fontFamily: "var(--font-display)", fontWeight: 700, fontSize: 18, color: "var(--ink)" }}>{title}</div>
              <table style={{ marginTop: 8, fontSize: 12 }}>
                <tbody>
                  <tr><td style={{ padding: "2px 8px 2px 0", color: "var(--steel)", border: "none" }}>เลขที่</td><td className="mono" style={{ padding: "2px 0", border: "none", fontWeight: 600 }}>{doc.docNumber}</td></tr>
                  <tr><td style={{ padding: "2px 8px 2px 0", color: "var(--steel)", border: "none" }}>ผู้ขาย</td><td style={{ padding: "2px 0", border: "none" }}>{settings.sellerCode}</td></tr>
                  {doc.refNote && <tr><td style={{ padding: "2px 8px 2px 0", color: "var(--steel)", border: "none" }}>อ้างอิง</td><td style={{ padding: "2px 0", border: "none" }}>{doc.refNote}</td></tr>}
                </tbody>
              </table>
            </div>
          </div>

          <div className="flex justify-between items-start mb-6 gap-6">
            <div style={{ fontSize: 13 }}>
              <div style={{ color: "var(--steel)", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 2 }}>ลูกค้า / Customer</div>
              <div style={{ fontWeight: 600 }}>{customer?.name || "ไม่ระบุ"}</div>
              {customer?.address && <div style={{ color: "var(--steel)" }}>{customer.address}</div>}
              {customer?.phone && <div style={{ color: "var(--steel)" }}>โทร. {customer.phone}</div>}
              {customer?.taxId && <div style={{ color: "var(--steel)" }}>เลขผู้เสียภาษี {customer.taxId}</div>}
            </div>
            <div className="card p-3" style={{ background: "var(--paper)", fontSize: 12, minWidth: 190 }}>
              <div className="flex justify-between mb-1"><span style={{ color: "var(--steel)" }}>วันที่ออกเอกสาร</span><span>{fmtDate(doc.date)}</span></div>
              {doc.dueDate && (
                <div className="flex justify-between"><span style={{ color: "var(--steel)" }}>{isInvoice ? "กำหนดชำระ" : "ยืนราคาถึง"}</span><span>{fmtDate(doc.dueDate)}</span></div>
              )}
            </div>
          </div>

          {(doc.jobName || doc.contactName || doc.contactPhone) && (
            <div className="card p-3 mb-6 grid grid-cols-3 gap-2" style={{ fontSize: 12 }}>
              <div><div style={{ color: "var(--steel)", fontSize: 11 }}>ชื่องาน</div>{doc.jobName || "-"}</div>
              <div><div style={{ color: "var(--steel)", fontSize: 11 }}>ผู้ติดต่อ</div>{doc.contactName || "-"}</div>
              <div><div style={{ color: "var(--steel)", fontSize: 11 }}>เบอร์โทร</div>{doc.contactPhone || "-"}</div>
            </div>
          )}

          <table style={{ marginBottom: 6 }}>
            <thead>
              <tr>
                <th style={{ width: 36 }}>ลำดับ</th>
                <th style={{ width: 44 }}>รูป</th>
                <th>รายการ</th>
                <th style={{ textAlign: "right" }}>จำนวน</th>
                <th>หน่วย</th>
                <th style={{ textAlign: "right" }}>ราคา/หน่วย</th>
                <th style={{ textAlign: "right" }}>ส่วนลด</th>
                <th style={{ textAlign: "right" }}>จำนวนเงิน</th>
              </tr>
            </thead>
            <tbody>
              {doc.items.map((it, i) => (
                <tr key={it.id}>
                  <td>{i + 1}</td>
                  <td>
                    {it.imageUrl ? (
                      <img src={it.imageUrl} alt="" style={{ width: 36, height: 36, objectFit: "cover", borderRadius: 4 }} />
                    ) : (
                      <div style={{ width: 36, height: 36, borderRadius: 4, background: "var(--paper)" }} />
                    )}
                  </td>
                  <td>
                    {it.sku && <div style={{ fontSize: 11, color: "var(--steel)" }}>รหัสสินค้า: {it.sku}</div>}
                    {it.name}
                  </td>
                  <td className="mono" style={{ textAlign: "right" }}>{it.qty}</td>
                  <td>{it.unit || "-"}</td>
                  <td className="mono" style={{ textAlign: "right" }}>{thb(it.price)}</td>
                  <td className="mono" style={{ textAlign: "right" }}>{it.discount ? thb(it.discount) : "-"}</td>
                  <td className="mono" style={{ textAlign: "right" }}>{thb(Number(it.qty) * Number(it.price) - Number(it.discount || 0))}</td>
                </tr>
              ))}
            </tbody>
          </table>

          <div className="flex justify-end mb-2">
            <div style={{ minWidth: 230, fontSize: 13 }}>
              <div className="flex justify-between" style={{ padding: "5px 0" }}><span>ยอดรวมก่อนภาษี</span><span className="mono">{thb(doc.subtotal)}</span></div>
              {doc.includeVat && (
                <div className="flex justify-between" style={{ padding: "5px 0" }}><span>ภาษีมูลค่าเพิ่ม ({settings.vatRate}%)</span><span className="mono">{thb(doc.vatAmount)}</span></div>
              )}
              <div className="flex justify-between" style={{ fontWeight: 700, fontSize: 16, borderTop: "1px solid var(--border)", padding: "8px 0 4px" }}>
                <span>ยอดรวมสุทธิ</span><span className="mono">{thb(doc.total)}</span>
              </div>
            </div>
          </div>
          <div className="text-right mb-8" style={{ fontSize: 12, color: "var(--steel)" }}>({thaiBahtText(doc.total)})</div>

          {doc.payments?.length > 0 && (
            <div className="card p-3 mb-6" style={{ fontSize: 12, background: "var(--paper)" }}>
              <div style={{ color: "var(--steel)", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>ประวัติการรับชำระ</div>
              {doc.payments.map((p) => (
                <div key={p.id} className="flex justify-between" style={{ padding: "2px 0" }}>
                  <span>{fmtDate(p.date)} · {PAYMENT_METHOD_LABEL[p.method] || p.method || "-"}</span>
                  <span className="mono">{thb(p.amount)}</span>
                </div>
              ))}
              <div className="flex justify-between" style={{ fontWeight: 700, paddingTop: 6, marginTop: 4, borderTop: "1px solid var(--border)" }}>
                <span>คงเหลือ</span><span className="mono">{thb(docBalanceDue(doc))}</span>
              </div>
            </div>
          )}

          {doc.note && (
            <div className="mb-8" style={{ fontSize: 12 }}>
              <div style={{ color: "var(--steel)", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 2 }}>หมายเหตุ</div>
              <div>{doc.note}</div>
            </div>
          )}

          <div className="grid grid-cols-2 gap-8 mt-10" style={{ fontSize: 12 }}>
            <div className="text-center">
              <div style={{ borderTop: "1px solid var(--ink)", paddingTop: 6, marginTop: 40 }}>ผู้จัดทำเอกสาร</div>
              <div style={{ color: "var(--steel)", marginTop: 4 }}>วันที่ ..... /..... /.....</div>
            </div>
            <div className="text-center">
              <div style={{ borderTop: "1px solid var(--ink)", paddingTop: 6, marginTop: 40 }}>{isInvoice ? "ผู้อนุมัติชำระเงิน" : "ผู้สั่งซื้อ / อนุมัติ"}</div>
              <div style={{ color: "var(--steel)", marginTop: 4 }}>วันที่ ..... /..... /.....</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function TaxReceiptPrintView({ doc, customers, documents, settings, onClose }) {
  const customer = customers.find((c) => c.id === doc.customerId);
  const isReceipt = doc.docType === "receipt";
  const isDebitNote = doc.docType === "debitnote";
  const isCreditNote = doc.docType === "creditnote";
  const title = (doc.isSimplified ? "ใบกำกับภาษีอย่างย่อ" : DOC_TYPE_LABEL[doc.docType]) || "ใบกำกับภาษี";
  const cornerColor = isReceipt ? "var(--green)" : isDebitNote ? "var(--amber-dark)" : isCreditNote ? "var(--red)" : "#2A4B9B";
  const paymentLabel = { cash: "เงินสด", cheque: "เช็ค", transfer: "โอนเงิน", credit: "บัตรเครดิต" };
  const sourceDoc = (documents || []).find((d) => d.id === doc.sourceDocId);

  return (
    <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(20,20,18,0.55)", zIndex: 60 }}>
      <div className="card w-full flex flex-col" style={{ maxWidth: 760, maxHeight: "92vh" }}>
        <div className="flex items-center justify-between px-5 py-4 no-print" style={{ borderBottom: "1px solid var(--border)" }}>
          <h2 style={{ fontFamily: "var(--font-display)", fontWeight: 700, fontSize: 17 }}>{doc.docNumber}</h2>
          <div className="flex gap-2">
            <button className="btn-primary flex items-center gap-1.5 px-3 py-1.5 text-sm" onClick={() => window.print()}><Printer size={14} /> พิมพ์ / บันทึก PDF</button>
            <button onClick={onClose} className="btn-ghost p-1"><X size={18} /></button>
          </div>
        </div>

        <div className="p-8 overflow-y-auto print-area" style={{ position: "relative" }}>
          <div
            style={{
              position: "absolute", top: 0, right: 0, width: 0, height: 0,
              borderStyle: "solid", borderWidth: "0 64px 64px 0", borderColor: `transparent ${cornerColor} transparent transparent`,
            }}
          >
            <span style={{ position: "absolute", top: -54, right: -30, color: "#fff", fontSize: 12, fontWeight: 700 }}>1</span>
          </div>

          <div className="flex justify-between items-start mb-6">
            <div>
              <div style={{ fontFamily: "var(--font-display)", fontWeight: 700, fontSize: 22 }}>{settings.companyName}</div>
              <div style={{ fontSize: 12, color: "var(--steel)", whiteSpace: "pre-line", marginTop: 4 }}>{settings.address}</div>
              <div style={{ fontSize: 12, color: "var(--steel)" }}>
                {settings.taxId && `เลขประจำตัวผู้เสียภาษี ${settings.taxId}`}
              </div>
              <div style={{ fontSize: 12, color: "var(--steel)" }}>{settings.phone && `โทร. ${settings.phone}`}</div>
            </div>
            <div className="text-right">
              <div style={{ fontFamily: "var(--font-display)", fontWeight: 700, fontSize: 20 }}>{title}</div>
              <div style={{ fontSize: 11, color: "var(--steel)" }}>ต้นฉบับ (เอกสารออกเป็นชุด)</div>
              <table style={{ marginTop: 8, fontSize: 12 }}>
                <tbody>
                  <tr><td style={{ padding: "2px 8px 2px 0", color: "var(--steel)", border: "none" }}>เลขที่</td><td className="mono" style={{ padding: "2px 0", border: "none", fontWeight: 600 }}>{doc.docNumber}</td></tr>
                  <tr><td style={{ padding: "2px 8px 2px 0", color: "var(--steel)", border: "none" }}>วันที่</td><td style={{ padding: "2px 0", border: "none" }}>{fmtDate(doc.date)}</td></tr>
                  <tr><td style={{ padding: "2px 8px 2px 0", color: "var(--steel)", border: "none" }}>ผู้ขาย</td><td style={{ padding: "2px 0", border: "none" }}>{settings.sellerCode}</td></tr>
                  {doc.refNote && <tr><td style={{ padding: "2px 8px 2px 0", color: "var(--steel)", border: "none" }}>อ้างอิง</td><td style={{ padding: "2px 0", border: "none" }}>{doc.refNote}</td></tr>}
                </tbody>
              </table>
            </div>
          </div>

          {!doc.isSimplified && (
            <div className="card p-3 mb-4" style={{ fontSize: 13 }}>
              <div style={{ color: "var(--steel)", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 4 }}>ลูกค้า</div>
              <div style={{ fontWeight: 600 }}>{customer?.name || "ไม่ระบุ"}</div>
              {customer?.address && <div style={{ color: "var(--steel)" }}>{customer.address}</div>}
              {customer?.taxId && <div style={{ color: "var(--steel)" }}>เลขประจำตัวผู้เสียภาษี {customer.taxId}</div>}
              {(doc.jobName || doc.contactName || doc.contactPhone) && (
                <div className="grid grid-cols-3 gap-2 mt-3 pt-2" style={{ borderTop: "1px solid var(--border)" }}>
                  <div><div style={{ color: "var(--steel)", fontSize: 11 }}>ชื่องาน</div>{doc.jobName || "-"}</div>
                  <div><div style={{ color: "var(--steel)", fontSize: 11 }}>ผู้ติดต่อ</div>{doc.contactName || "-"}</div>
                  <div><div style={{ color: "var(--steel)", fontSize: 11 }}>เบอร์โทร</div>{doc.contactPhone || "-"}</div>
                </div>
              )}
            </div>
          )}

          {(isDebitNote || isCreditNote) && (
            <div className="card p-3 mb-4" style={{ fontSize: 13, background: "var(--paper)" }}>
              <div className="flex justify-between flex-wrap gap-2">
                <span><span style={{ color: "var(--steel)" }}>อ้างอิง{DOC_TYPE_LABEL[sourceDoc?.docType] || "เอกสาร"}เลขที่</span> <span className="mono" style={{ fontWeight: 600 }}>{sourceDoc?.docNumber || "-"}</span></span>
                {sourceDoc && <span style={{ color: "var(--steel)" }}>วันที่ {fmtDate(sourceDoc.date)}</span>}
              </div>
              {doc.adjustReason && <div className="mt-2"><span style={{ color: "var(--steel)" }}>เหตุผล:</span> {doc.adjustReason}</div>}
            </div>
          )}

          <table style={{ marginBottom: 6 }}>
            <thead>
              <tr>
                <th style={{ width: 32 }}>#</th>
                <th style={{ width: 44 }}>รูป</th>
                <th>รายละเอียด</th>
                <th style={{ textAlign: "right" }}>จำนวน</th>
                <th>หน่วย</th>
                <th style={{ textAlign: "right" }}>ราคาต่อหน่วย</th>
                <th style={{ textAlign: "right" }}>ส่วนลด</th>
                <th style={{ textAlign: "right" }}>มูลค่า</th>
              </tr>
            </thead>
            <tbody>
              {doc.items.map((it, i) => (
                <tr key={it.id}>
                  <td>{i + 1}</td>
                  <td>
                    {it.imageUrl ? (
                      <img src={it.imageUrl} alt="" style={{ width: 36, height: 36, objectFit: "cover", borderRadius: 4 }} />
                    ) : (
                      <div style={{ width: 36, height: 36, borderRadius: 4, background: "var(--surface)" }} />
                    )}
                  </td>
                  <td>
                    {it.sku && <div style={{ fontSize: 11, color: "var(--steel)" }}>รหัสสินค้า: {it.sku}</div>}
                    {it.name}
                  </td>
                  <td className="mono" style={{ textAlign: "right" }}>{it.qty}</td>
                  <td>{it.unit || "-"}</td>
                  <td className="mono" style={{ textAlign: "right" }}>{thb(it.price)}</td>
                  <td className="mono" style={{ textAlign: "right" }}>{it.discount ? thb(it.discount) : "-"}</td>
                  <td className="mono" style={{ textAlign: "right" }}>{thb(Number(it.qty) * Number(it.price) - Number(it.discount || 0))}</td>
                </tr>
              ))}
            </tbody>
          </table>

          <div className="flex justify-end mb-2">
            <div style={{ minWidth: 230, fontSize: 13 }}>
              <div className="flex justify-between" style={{ padding: "5px 0" }}><span>รวมเป็นเงิน</span><span className="mono">{thb(doc.subtotal)}</span></div>
              {doc.includeVat && (
                <div className="flex justify-between" style={{ padding: "5px 0" }}><span>ภาษีมูลค่าเพิ่ม ({settings.vatRate}%)</span><span className="mono">{thb(doc.vatAmount)}</span></div>
              )}
              <div className="flex justify-between" style={{ fontWeight: 700, fontSize: 16, borderTop: "1px solid var(--border)", padding: "8px 0 4px" }}>
                <span>{isDebitNote || isCreditNote ? `จำนวนเงินที่${isDebitNote ? "เพิ่ม" : "ลด"}` : "จำนวนเงินรวมทั้งสิ้น"}</span><span className="mono">{thb(doc.total)}</span>
              </div>
            </div>
          </div>
          <div className="text-right mb-6" style={{ fontSize: 12, color: "var(--steel)" }}>({thaiBahtText(doc.total)})</div>

          {isReceipt && (
            <div className="card p-3 mb-6" style={{ fontSize: 13 }}>
              <div className="flex items-center gap-4 flex-wrap mb-2">
                <span style={{ color: "var(--steel)" }}>การชำระเงิน:</span>
                {["cash", "cheque", "transfer", "credit"].map((m) => (
                  <label key={m} className="flex items-center gap-1.5">
                    <input type="checkbox" readOnly checked={doc.paymentMethod === m} /> {paymentLabel[m]}
                  </label>
                ))}
              </div>
              {(doc.bankName || doc.paymentRefNumber || doc.paymentDate) && (
                <div className="flex gap-6 flex-wrap" style={{ color: "var(--steel)", fontSize: 12 }}>
                  {doc.bankName && <span>ธนาคาร {doc.bankName}</span>}
                  {doc.paymentRefNumber && <span>เลขที่ {doc.paymentRefNumber}</span>}
                  {doc.paymentDate && <span>วันที่ {fmtDate(doc.paymentDate)}</span>}
                </div>
              )}
            </div>
          )}

          {doc.docType === "taxinvoice" && doc.payments?.length > 0 && (
            <div className="card p-3 mb-6" style={{ fontSize: 12, background: "var(--paper)" }}>
              <div style={{ color: "var(--steel)", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 6 }}>ประวัติการรับชำระ</div>
              {doc.payments.map((p) => (
                <div key={p.id} className="flex justify-between" style={{ padding: "2px 0" }}>
                  <span>{fmtDate(p.date)} · {PAYMENT_METHOD_LABEL[p.method] || p.method || "-"}</span>
                  <span className="mono">{thb(p.amount)}</span>
                </div>
              ))}
              <div className="flex justify-between" style={{ fontWeight: 700, paddingTop: 6, marginTop: 4, borderTop: "1px solid var(--border)" }}>
                <span>คงเหลือ</span><span className="mono">{thb(docBalanceDue(doc))}</span>
              </div>
            </div>
          )}

          {doc.note && (
            <div className="mb-8" style={{ fontSize: 12 }}>
              <div style={{ color: "var(--steel)", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 2 }}>หมายเหตุ</div>
              <div>{doc.note}</div>
            </div>
          )}

          <div className="grid grid-cols-2 gap-8 mt-10" style={{ fontSize: 12 }}>
            <div className="text-center">
              <div style={{ borderTop: "1px solid var(--ink)", paddingTop: 6, marginTop: 40 }}>{isReceipt ? "ผู้จ่ายเงิน" : "ผู้รับสินค้า / บริการ"}</div>
              <div style={{ color: "var(--steel)", marginTop: 4 }}>วันที่ ..... /..... /.....</div>
            </div>
            <div className="text-center">
              <div style={{ borderTop: "1px solid var(--ink)", paddingTop: 6, marginTop: 40 }}>{isReceipt ? "ผู้รับเงิน" : "ผู้อนุมัติ"}</div>
              <div style={{ color: "var(--steel)", marginTop: 4 }}>วันที่ ..... /..... /.....</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------- sales report -------------------------------- */

function SalesReportTab({ documents, products, transactions, showToast }) {
  const availableMonths = useMemo(() => {
    const set = new Set([monthKey(today())]);
    documents.forEach((d) => { if (d.docType === "invoice") set.add(monthKey(d.date)); });
    return Array.from(set).sort().reverse();
  }, [documents]);

  const [selectedMonth, setSelectedMonth] = useState(availableMonths[0] || monthKey(today()));

  const outputVatDocs = useMemo(
    () => documents.filter((d) => ["taxinvoice", "debitnote", "creditnote"].includes(d.docType) && d.status !== "cancelled" && monthKey(d.date) === selectedMonth),
    [documents, selectedMonth]
  );
  const outputVat = outputVatDocs.reduce((s, d) => {
    const v = d.vatAmount || 0;
    return s + (d.docType === "creditnote" ? -v : v);
  }, 0);
  const inputVatTx = useMemo(
    () => transactions.filter((t) => t.type === "expense" && t.hasInputVat && monthKey(t.date) === selectedMonth),
    [transactions, selectedMonth]
  );
  const inputVat = inputVatTx.reduce((s, t) => s + (Number(t.inputVat) || 0), 0);
  const vatPayable = outputVat - inputVat;

  const arAging = useMemo(() => {
    const buckets = { current: [], d30: [], d60: [], d90: [] };
    const now = new Date(today());
    documents
      .filter((d) => (d.docType === "invoice" || d.docType === "taxinvoice") && !["completed", "cancelled"].includes(d.status))
      .forEach((d) => {
        const due = new Date(d.dueDate || d.date);
        const daysOverdue = Math.floor((now - due) / (1000 * 60 * 60 * 24));
        if (daysOverdue <= 0) buckets.current.push(d);
        else if (daysOverdue <= 30) buckets.d30.push(d);
        else if (daysOverdue <= 60) buckets.d60.push(d);
        else buckets.d90.push(d);
      });
    return buckets;
  }, [documents]);
  const arTotal = (arr) => arr.reduce((s, d) => s + docBalanceDue(d), 0);
  const arGrandTotal = arTotal(arAging.current) + arTotal(arAging.d30) + arTotal(arAging.d60) + arTotal(arAging.d90);

  const monthInvoices = useMemo(
    () =>
      documents
        .filter((d) => d.docType === "invoice" && ["issued", "completed"].includes(d.status) && monthKey(d.date) === selectedMonth)
        .sort((a, b) => (a.date < b.date ? 1 : -1)),
    [documents, selectedMonth]
  );

  const totalSales = monthInvoices.reduce((s, d) => s + d.total, 0);
  const totalVat = monthInvoices.reduce((s, d) => s + (d.vatAmount || 0), 0);
  const totalSubtotal = monthInvoices.reduce((s, d) => s + (d.subtotal || 0), 0);
  const avgSale = monthInvoices.length ? totalSales / monthInvoices.length : 0;

  const productBreakdown = useMemo(() => {
    const map = {};
    monthInvoices.forEach((d) => {
      d.items.forEach((it) => {
        if (!map[it.name]) map[it.name] = { qty: 0, revenue: 0 };
        map[it.name].qty += Number(it.qty) || 0;
        map[it.name].revenue += (Number(it.qty) || 0) * (Number(it.price) || 0);
      });
    });
    return Object.entries(map).sort((a, b) => b[1].revenue - a[1].revenue);
  }, [monthInvoices]);

  function exportExcel() {
    if (monthInvoices.length === 0) return showToast("ไม่มีข้อมูลยอดขายในเดือนนี้", "error");
    const wsSummary = XLSX.utils.aoa_to_sheet([
      ["รายงานยอดขายประจำเดือน", monthLabel(selectedMonth)],
      [],
      ["จำนวนใบแจ้งหนี้", monthInvoices.length],
      ["ยอดขายก่อนภาษี", totalSubtotal],
      ["ภาษีมูลค่าเพิ่มรวม", totalVat],
      ["ยอดขายรวมสุทธิ", totalSales],
      ["ยอดขายเฉลี่ยต่อบิล", avgSale],
    ]);
    const wsInvoices = XLSX.utils.json_to_sheet(
      monthInvoices.map((d) => ({
        เลขที่เอกสาร: d.docNumber,
        วันที่: d.date,
        ยอดก่อนภาษี: d.subtotal,
        ภาษี: d.vatAmount,
        ยอดรวม: d.total,
        สถานะ: d.status === "completed" ? "ชำระแล้ว" : "ออกบิลแล้ว",
      }))
    );
    const wsProducts = XLSX.utils.json_to_sheet(
      productBreakdown.map(([name, v]) => ({ "สินค้า/บริการ": name, จำนวนที่ขาย: v.qty, ยอดขาย: v.revenue }))
    );
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsSummary, "สรุป");
    XLSX.utils.book_append_sheet(wb, wsInvoices, "รายการใบแจ้งหนี้");
    XLSX.utils.book_append_sheet(wb, wsProducts, "สินค้าขายดี");
    XLSX.writeFile(wb, `รายงานขาย-${selectedMonth}.xlsx`);
    showToast("ดาวน์โหลดรายงาน Excel เรียบร้อย");
  }

  return (
    <div>
      <SectionHeader
        title="รายงานขายประจำเดือน"
        subtitle="สรุปยอดขายจากใบแจ้งหนี้ที่ออกบิลแล้ว"
        action={
          <div className="flex gap-2 items-center flex-wrap">
            <select className={inputCls} style={{ width: 170 }} value={selectedMonth} onChange={(e) => setSelectedMonth(e.target.value)}>
              {availableMonths.map((m) => <option key={m} value={m}>{monthLabel(m)}</option>)}
            </select>
            <button className="btn-primary flex items-center gap-1.5 px-4 py-2 text-sm" onClick={exportExcel}>
              <Download size={16} /> Export Excel
            </button>
          </div>
        }
      />
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
        <StatCard label="ยอดขายรวม" value={thb(totalSales)} icon={TrendingUp} tone="green" />
        <StatCard label="จำนวนใบแจ้งหนี้" value={monthInvoices.length} icon={FileText} tone="ink" />
        <StatCard label="ภาษีมูลค่าเพิ่มรวม" value={thb(totalVat)} icon={Wallet} tone="amber" />
        <StatCard label="ยอดขายเฉลี่ย/บิล" value={thb(avgSale)} icon={BarChart3} tone="ink" />
      </div>

      {monthInvoices.length === 0 ? (
        <Empty text="ยังไม่มีใบแจ้งหนี้ในเดือนที่เลือก" />
      ) : (
        <>
          <div className="card overflow-x-auto mb-4">
            <table>
              <thead><tr><th>เลขที่</th><th>วันที่</th><th>ยอดก่อนภาษี</th><th>ภาษี</th><th>ยอดรวม</th><th>สถานะ</th></tr></thead>
              <tbody>
                {monthInvoices.map((d) => (
                  <tr key={d.id}>
                    <td className="mono">{d.docNumber}</td>
                    <td>{fmtDate(d.date)}</td>
                    <td className="mono">{thb(d.subtotal)}</td>
                    <td className="mono">{thb(d.vatAmount)}</td>
                    <td className="mono" style={{ fontWeight: 600 }}>{thb(d.total)}</td>
                    <td>{statusBadge(d.status)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div className="card p-4">
            <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 12 }}>ยอดขายแยกตามสินค้า/บริการ</div>
            <div className="flex flex-col gap-2">
              {productBreakdown.map(([name, v]) => (
                <div key={name} className="flex items-center justify-between" style={{ fontSize: 13 }}>
                  <span>{name} <span style={{ color: "var(--steel)" }}>x{v.qty}</span></span>
                  <span className="mono" style={{ fontWeight: 600 }}>{thb(v.revenue)}</span>
                </div>
              ))}
            </div>
          </div>
        </>
      )}

      <div className="card p-4 mt-4">
        <div className="flex items-center justify-between mb-3 flex-wrap gap-2">
          <div style={{ fontWeight: 600, fontSize: 14 }}>รายงานภาษีมูลค่าเพิ่ม (แบบ ภ.พ.30) — {monthLabel(selectedMonth)}</div>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-2" style={{ fontSize: 13 }}>
          <div className="flex justify-between p-2" style={{ background: "var(--paper)", borderRadius: 8 }}>
            <span>ภาษีขาย (จากใบกำกับภาษี/ใบเพิ่มหนี้/ใบลดหนี้)</span><span className="mono" style={{ fontWeight: 600 }}>{thb(outputVat)}</span>
          </div>
          <div className="flex justify-between p-2" style={{ background: "var(--paper)", borderRadius: 8 }}>
            <span>ภาษีซื้อ (จากรายจ่ายที่มีภาษีซื้อ)</span><span className="mono" style={{ fontWeight: 600 }}>{thb(inputVat)}</span>
          </div>
          <div className="flex justify-between p-2" style={{ background: vatPayable >= 0 ? "var(--red-bg)" : "var(--green-bg)", borderRadius: 8 }}>
            <span>{vatPayable >= 0 ? "ภาษีที่ต้องชำระ" : "ภาษีที่ขอคืน/ยกไป"}</span>
            <span className="mono" style={{ fontWeight: 700, color: vatPayable >= 0 ? "var(--red)" : "var(--green)" }}>{thb(Math.abs(vatPayable))}</span>
          </div>
        </div>
        {outputVatDocs.length === 0 && inputVatTx.length === 0 && (
          <p style={{ fontSize: 12, color: "var(--steel)", marginTop: 8 }}>ยังไม่มีใบกำกับภาษีหรือรายจ่ายที่มีภาษีซื้อในเดือนนี้</p>
        )}
      </div>

      <div className="card p-4 mt-4">
        <div style={{ fontWeight: 600, fontSize: 14, marginBottom: 12 }}>อายุลูกหนี้ค้างชำระ (AR Aging)</div>
        {arGrandTotal === 0 ? (
          <p style={{ fontSize: 13, color: "var(--steel)" }}>ไม่มีใบแจ้งหนี้/ใบกำกับภาษีค้างชำระในขณะนี้</p>
        ) : (
          <div className="grid grid-cols-2 md:grid-cols-4 gap-2" style={{ fontSize: 13 }}>
            {[
              ["ยังไม่ครบกำหนด", arAging.current, "ink"],
              ["เกิน 1-30 วัน", arAging.d30, "amber"],
              ["เกิน 31-60 วัน", arAging.d60, "red"],
              ["เกิน 60 วันขึ้นไป", arAging.d90, "red"],
            ].map(([label, arr, tone]) => (
              <div key={label} className="p-3" style={{ background: "var(--paper)", borderRadius: 8 }}>
                <div style={{ color: "var(--steel)", fontSize: 11, marginBottom: 4 }}>{label}</div>
                <div className="mono" style={{ fontWeight: 700, fontSize: 15, color: tone === "red" ? "var(--red)" : tone === "amber" ? "var(--amber-dark)" : "var(--ink)" }}>{thb(arTotal(arr))}</div>
                <div style={{ color: "var(--steel)", fontSize: 11 }}>{arr.length} รายการ</div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

/* --------------------------------- billing note -------------------------------- */

function BillingTab({ documents, customers, settings, updateSettings, billingNotes, updateBillingNotes, updateDocuments, showToast, confirmAction }) {
  const [modal, setModal] = useState(false);
  const [viewNote, setViewNote] = useState(null);
  const customerName = (id) => customers.find((c) => c.id === id)?.name || "ไม่ระบุ";

  function genBillingNumber() {
    const n = settings.nextBillingNumber;
    const year = new Date().getFullYear();
    updateSettings({ ...settings, nextBillingNumber: n + 1 });
    return `${settings.billingPrefix}-${year}-${String(n).padStart(4, "0")}`;
  }

  function createNote(form) {
    if (!form.customerId) return showToast("กรุณาเลือกลูกค้า", "error");
    if (form.invoiceIds.length === 0) return showToast("กรุณาเลือกใบแจ้งหนี้อย่างน้อย 1 รายการ", "error");
    const selectedInvoices = documents.filter((d) => form.invoiceIds.includes(d.id));
    const total = selectedInvoices.reduce((s, d) => s + d.total, 0);
    const docNumber = genBillingNumber();
    const note = { id: uid(), docNumber, customerId: form.customerId, date: form.date, dueDate: form.dueDate, note: form.note, invoiceIds: form.invoiceIds, total };
    updateBillingNotes([note, ...billingNotes]);
    updateDocuments(documents.map((d) => (form.invoiceIds.includes(d.id) ? { ...d, billingNoteId: note.id } : d)));
    showToast("สร้างใบวางบิลเรียบร้อย");
    setModal(false);
  }

  function removeNote(note) {
    confirmAction("ลบใบวางบิลนี้? ใบแจ้งหนี้ที่เกี่ยวข้องจะกลับไปเป็นสถานะยังไม่วางบิล", () => {
      updateBillingNotes(billingNotes.filter((n) => n.id !== note.id));
      updateDocuments(documents.map((d) => (note.invoiceIds.includes(d.id) ? { ...d, billingNoteId: null } : d)));
      showToast("ลบใบวางบิลแล้ว");
    });
  }

  return (
    <div>
      <SectionHeader
        title="ใบวางบิล"
        subtitle="รวมใบแจ้งหนี้ค้างชำระของลูกค้าแต่ละรายเพื่อส่งเรียกเก็บเงิน"
        action={
          <button className="btn-primary flex items-center gap-1.5 px-4 py-2 text-sm" onClick={() => setModal(true)}>
            <Plus size={16} /> สร้างใบวางบิล
          </button>
        }
      />
      {billingNotes.length === 0 ? (
        <Empty text="ยังไม่มีใบวางบิล สร้างใบวางบิลเพื่อรวมใบแจ้งหนี้ค้างชำระของลูกค้าแล้วส่งเรียกเก็บเงินได้ในคราวเดียว" />
      ) : (
        <div className="card overflow-x-auto">
          <table>
            <thead><tr><th>เลขที่</th><th>ลูกค้า</th><th>วันที่</th><th>จำนวนบิล</th><th>ยอดรวม</th><th></th></tr></thead>
            <tbody>
              {[...billingNotes].sort((a, b) => (a.date < b.date ? 1 : -1)).map((n) => (
                <tr key={n.id}>
                  <td className="mono" style={{ fontWeight: 500 }}>{n.docNumber}</td>
                  <td>{customerName(n.customerId)}</td>
                  <td>{fmtDate(n.date)}</td>
                  <td>{n.invoiceIds.length}</td>
                  <td className="mono" style={{ fontWeight: 600 }}>{thb(n.total)}</td>
                  <td>
                    <div className="flex gap-2 justify-end">
                      <button className="btn-ghost p-1.5" title="ดู/พิมพ์" onClick={() => setViewNote(n)}><Printer size={14} /></button>
                      <button className="btn-ghost btn-danger p-1.5" onClick={() => removeNote(n)}><Trash2 size={14} /></button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {modal && (
        <Modal title="สร้างใบวางบิล" onClose={() => setModal(false)} wide>
          <BillingForm documents={documents} customers={customers} onSave={createNote} />
        </Modal>
      )}

      {viewNote && (
        <BillingPrintView note={viewNote} documents={documents} customers={customers} settings={settings} onClose={() => setViewNote(null)} />
      )}
    </div>
  );
}

function BillingForm({ documents, customers, onSave }) {
  const [customerId, setCustomerId] = useState("");
  const [date, setDate] = useState(today());
  const [dueDate, setDueDate] = useState("");
  const [note, setNote] = useState("");
  const [selected, setSelected] = useState([]);

  const outstanding = documents.filter((d) => d.docType === "invoice" && d.customerId === customerId && d.status === "issued" && !d.billingNoteId);

  function toggle(id) {
    setSelected((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
  }

  const total = documents.filter((d) => selected.includes(d.id)).reduce((s, d) => s + d.total, 0);

  return (
    <div>
      <div className="grid grid-cols-2 gap-3">
        <Field label="ลูกค้า *">
          <select className={inputCls} value={customerId} onChange={(e) => { setCustomerId(e.target.value); setSelected([]); }}>
            <option value="">เลือกลูกค้า</option>
            {customers.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
          </select>
        </Field>
        <Field label="วันที่ออกใบวางบิล"><input type="date" className={inputCls} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
      </div>
      <Field label="กำหนดชำระ (ถ้ามี)"><input type="date" className={inputCls} value={dueDate} onChange={(e) => setDueDate(e.target.value)} /></Field>

      {customerId && (
        outstanding.length === 0 ? (
          <p style={{ fontSize: 13, color: "var(--steel)", marginBottom: 12 }}>ลูกค้ารายนี้ไม่มีใบแจ้งหนี้ค้างชำระที่ยังไม่ได้วางบิล</p>
        ) : (
          <div className="flex flex-col gap-2 mb-3">
            {outstanding.map((d) => (
              <label key={d.id} className="card flex items-center justify-between p-2" style={{ fontSize: 13, cursor: "pointer" }}>
                <span className="flex items-center gap-2">
                  <input type="checkbox" checked={selected.includes(d.id)} onChange={() => toggle(d.id)} />
                  <span className="mono">{d.docNumber}</span> · {fmtDate(d.date)}
                </span>
                <span className="mono" style={{ fontWeight: 600 }}>{thb(d.total)}</span>
              </label>
            ))}
          </div>
        )
      )}

      <Field label="หมายเหตุ"><textarea className={inputCls} rows={2} value={note} onChange={(e) => setNote(e.target.value)} /></Field>

      <div className="card p-3 flex justify-between items-center mb-3" style={{ background: "var(--paper)" }}>
        <span style={{ fontWeight: 600, fontSize: 13 }}>ยอดรวมที่เรียกเก็บ</span>
        <span className="mono" style={{ fontWeight: 700, fontSize: 16 }}>{thb(total)}</span>
      </div>

      <button className="btn-primary w-full py-2.5 text-sm" onClick={() => onSave({ customerId, date, dueDate, note, invoiceIds: selected })}>สร้างใบวางบิล</button>
    </div>
  );
}

function BillingPrintView({ note, documents, customers, settings, onClose }) {
  const customer = customers.find((c) => c.id === note.customerId);
  const invoices = documents.filter((d) => note.invoiceIds.includes(d.id));

  return (
    <div className="fixed inset-0 flex items-center justify-center p-4" style={{ background: "rgba(20,20,18,0.5)", zIndex: 60 }}>
      <div className="card w-full flex flex-col" style={{ maxWidth: 640, maxHeight: "90vh" }}>
        <div className="flex items-center justify-between px-5 py-4 no-print" style={{ borderBottom: "1px solid var(--border)" }}>
          <h2 style={{ fontFamily: "var(--font-display)", fontWeight: 700, fontSize: 17 }}>ใบวางบิล {note.docNumber}</h2>
          <div className="flex gap-2">
            <button className="btn-primary flex items-center gap-1.5 px-3 py-1.5 text-sm" onClick={() => window.print()}><Printer size={14} /> พิมพ์</button>
            <button onClick={onClose} className="btn-ghost p-1"><X size={18} /></button>
          </div>
        </div>
        <div className="p-8 overflow-y-auto print-area" id="billing-print-area">
          <div className="flex justify-between items-start mb-6">
            <div>
              <div style={{ fontFamily: "var(--font-display)", fontWeight: 700, fontSize: 20 }}>{settings.companyName}</div>
              <div style={{ fontSize: 12, color: "var(--steel)", whiteSpace: "pre-line" }}>{settings.address}</div>
              <div style={{ fontSize: 12, color: "var(--steel)" }}>{settings.phone} {settings.taxId && `· เลขผู้เสียภาษี ${settings.taxId}`}</div>
            </div>
            <div className="text-right">
              <div style={{ fontFamily: "var(--font-display)", fontWeight: 700, fontSize: 18 }}>ใบวางบิล</div>
              <div className="mono" style={{ fontSize: 13 }}>{note.docNumber}</div>
              <div style={{ fontSize: 12, color: "var(--steel)" }}>วันที่ {fmtDate(note.date)}</div>
              {note.dueDate && <div style={{ fontSize: 12, color: "var(--steel)" }}>กำหนดชำระ {fmtDate(note.dueDate)}</div>}
            </div>
          </div>

          <div className="mb-4" style={{ fontSize: 13 }}>
            <div style={{ color: "var(--steel)", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em" }}>เรียน</div>
            <div style={{ fontWeight: 600 }}>{customer?.name || "ไม่ระบุ"}</div>
            {customer?.address && <div style={{ color: "var(--steel)" }}>{customer.address}</div>}
            {customer?.taxId && <div style={{ color: "var(--steel)" }}>เลขผู้เสียภาษี {customer.taxId}</div>}
          </div>

          <table style={{ marginBottom: 16 }}>
            <thead><tr><th>เลขที่ใบแจ้งหนี้</th><th>วันที่</th><th style={{ textAlign: "right" }}>จำนวนเงิน</th></tr></thead>
            <tbody>
              {invoices.map((d) => (
                <tr key={d.id}>
                  <td className="mono">{d.docNumber}</td>
                  <td>{fmtDate(d.date)}</td>
                  <td className="mono" style={{ textAlign: "right" }}>{thb(d.total)}</td>
                </tr>
              ))}
            </tbody>
          </table>

          <div className="flex justify-end mb-2">
            <div style={{ minWidth: 220 }}>
              <div className="flex justify-between" style={{ fontWeight: 700, fontSize: 16, borderTop: "1px solid var(--border)", paddingTop: 8 }}>
                <span>ยอดรวมที่เรียกเก็บ</span><span className="mono">{thb(note.total)}</span>
              </div>
            </div>
          </div>
          <div className="text-right mb-6" style={{ fontSize: 12, color: "var(--steel)" }}>({thaiBahtText(note.total)})</div>

          {note.note && <div style={{ fontSize: 12, color: "var(--steel)" }}>หมายเหตุ: {note.note}</div>}
        </div>
      </div>
    </div>
  );
}

/* --------------------------------- settings ----------------------------------- */

function SettingsTab({ settings, updateSettings, onReset, showToast }) {
  const [f, setF] = useState(settings);

  function save() {
    updateSettings({
      ...f,
      vatRate: Number(f.vatRate) || 0,
      nextInvoiceNumber: Number(f.nextInvoiceNumber) || 1,
      nextQuoteNumber: Number(f.nextQuoteNumber) || 1,
      nextBillingNumber: Number(f.nextBillingNumber) || 1,
      nextTaxInvoiceNumber: Number(f.nextTaxInvoiceNumber) || 1,
      nextReceiptNumber: Number(f.nextReceiptNumber) || 1,
      nextDebitNoteNumber: Number(f.nextDebitNoteNumber) || 1,
      nextCreditNoteNumber: Number(f.nextCreditNoteNumber) || 1,
      lowStockDefault: Number(f.lowStockDefault) || 0,
    });
    showToast("บันทึกการตั้งค่าเรียบร้อย");
  }

  return (
    <div>
      <SectionHeader title="ตั้งค่า" subtitle="ข้อมูลกิจการและค่าตั้งต้นของระบบ" />

      <div className="card p-5 mb-4" style={{ maxWidth: 560 }}>
        <div style={{ fontWeight: 600, marginBottom: 12 }}>ข้อมูลกิจการ</div>
        <Field label="ชื่อกิจการ"><input className={inputCls} value={f.companyName} onChange={(e) => setF({ ...f, companyName: e.target.value })} /></Field>
        <Field label="ที่อยู่"><textarea className={inputCls} rows={2} value={f.address} onChange={(e) => setF({ ...f, address: e.target.value })} /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="เบอร์โทรศัพท์"><input className={inputCls} value={f.phone} onChange={(e) => setF({ ...f, phone: e.target.value })} /></Field>
          <Field label="เลขประจำตัวผู้เสียภาษี"><input className={inputCls} value={f.taxId} onChange={(e) => setF({ ...f, taxId: e.target.value })} /></Field>
          <Field label="รหัสผู้ขาย (แสดงบนใบกำกับภาษี/ใบเสร็จ)"><input className={inputCls} value={f.sellerCode} onChange={(e) => setF({ ...f, sellerCode: e.target.value })} /></Field>
        </div>
      </div>

      <div className="card p-5 mb-4" style={{ maxWidth: 560 }}>
        <div style={{ fontWeight: 600, marginBottom: 12 }}>เอกสารและภาษี</div>
        <div className="grid grid-cols-2 gap-3">
          <Field label="อัตราภาษีมูลค่าเพิ่ม (%)"><input type="number" className={inputCls} value={f.vatRate} onChange={(e) => setF({ ...f, vatRate: e.target.value })} /></Field>
          <Field label="แจ้งเตือนสต็อกต่ำกว่า (ค่าเริ่มต้น)"><input type="number" className={inputCls} value={f.lowStockDefault} onChange={(e) => setF({ ...f, lowStockDefault: e.target.value })} /></Field>
          <Field label="คำนำหน้าใบเสนอราคา"><input className={inputCls} value={f.quotePrefix} onChange={(e) => setF({ ...f, quotePrefix: e.target.value })} /></Field>
          <Field label="เลขที่ใบเสนอราคาถัดไป"><input type="number" className={inputCls} value={f.nextQuoteNumber} onChange={(e) => setF({ ...f, nextQuoteNumber: e.target.value })} /></Field>
          <Field label="คำนำหน้าใบแจ้งหนี้"><input className={inputCls} value={f.invoicePrefix} onChange={(e) => setF({ ...f, invoicePrefix: e.target.value })} /></Field>
          <Field label="เลขที่ใบแจ้งหนี้ถัดไป"><input type="number" className={inputCls} value={f.nextInvoiceNumber} onChange={(e) => setF({ ...f, nextInvoiceNumber: e.target.value })} /></Field>
          <Field label="คำนำหน้าใบวางบิล"><input className={inputCls} value={f.billingPrefix} onChange={(e) => setF({ ...f, billingPrefix: e.target.value })} /></Field>
          <Field label="เลขที่ใบวางบิลถัดไป"><input type="number" className={inputCls} value={f.nextBillingNumber} onChange={(e) => setF({ ...f, nextBillingNumber: e.target.value })} /></Field>
          <Field label="คำนำหน้าใบกำกับภาษี"><input className={inputCls} value={f.taxInvoicePrefix} onChange={(e) => setF({ ...f, taxInvoicePrefix: e.target.value })} /></Field>
          <Field label="เลขที่ใบกำกับภาษีถัดไป"><input type="number" className={inputCls} value={f.nextTaxInvoiceNumber} onChange={(e) => setF({ ...f, nextTaxInvoiceNumber: e.target.value })} /></Field>
          <Field label="คำนำหน้าใบเสร็จรับเงิน"><input className={inputCls} value={f.receiptPrefix} onChange={(e) => setF({ ...f, receiptPrefix: e.target.value })} /></Field>
          <Field label="เลขที่ใบเสร็จรับเงินถัดไป"><input type="number" className={inputCls} value={f.nextReceiptNumber} onChange={(e) => setF({ ...f, nextReceiptNumber: e.target.value })} /></Field>
          <Field label="คำนำหน้าใบเพิ่มหนี้"><input className={inputCls} value={f.debitNotePrefix} onChange={(e) => setF({ ...f, debitNotePrefix: e.target.value })} /></Field>
          <Field label="เลขที่ใบเพิ่มหนี้ถัดไป"><input type="number" className={inputCls} value={f.nextDebitNoteNumber} onChange={(e) => setF({ ...f, nextDebitNoteNumber: e.target.value })} /></Field>
          <Field label="คำนำหน้าใบลดหนี้"><input className={inputCls} value={f.creditNotePrefix} onChange={(e) => setF({ ...f, creditNotePrefix: e.target.value })} /></Field>
          <Field label="เลขที่ใบลดหนี้ถัดไป"><input type="number" className={inputCls} value={f.nextCreditNoteNumber} onChange={(e) => setF({ ...f, nextCreditNoteNumber: e.target.value })} /></Field>
        </div>
        <button className="btn-primary px-5 py-2.5 text-sm mt-1" onClick={save}>บันทึกการตั้งค่า</button>
      </div>

      <div className="card p-5" style={{ maxWidth: 560, borderColor: "var(--red-bg)" }}>
        <div style={{ fontWeight: 600, marginBottom: 6, color: "var(--red)" }}>โซนอันตราย</div>
        <p style={{ fontSize: 13, color: "var(--steel)", marginBottom: 12 }}>ล้างข้อมูลลูกค้า สินค้า เอกสาร และรายรับ-รายจ่ายทั้งหมดอย่างถาวร</p>
        <button className="btn-ghost btn-danger flex items-center gap-1.5 px-4 py-2 text-sm" onClick={onReset}><RefreshCw size={14} /> ล้างข้อมูลทั้งหมด</button>
      </div>
    </div>
  );
}
