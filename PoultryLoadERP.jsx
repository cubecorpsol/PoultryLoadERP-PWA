import { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  Home, PlusCircle, ClipboardList, BarChart3, Settings as SettingsIcon,
  X, Trash2, Pencil, Search, Download, Upload, LogOut, ChevronRight,
  ChevronLeft, Check, Plus, Building2, Truck, Package, Calendar, Clock,
  ArrowLeft, KeyRound, DatabaseBackup, AlertTriangle, User, Lock, MapPin,
  Phone, FileSpreadsheet, Printer, Filter, Bird, Weight, Feather, Sun,
} from "lucide-react";

/* ---------------------------------------------------------------------- */
/* Tokens & helpers                                                       */
/* ---------------------------------------------------------------------- */

const C = {
  primary: "#2E7D32",
  primaryDark: "#1B5E20",
  primaryLight: "#E8F5E9",
  egg: "#FFFDF5",
  ink: "#173A1B",
  muted: "#5C7A60",
  amber: "#D98E2B",
  amberLight: "#FBF0DD",
  danger: "#C0392B",
  border: "#DCEEDD",
};

const STORAGE_KEY = "poultry_erp_data_v1";
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const round2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;

function nowLocalInput() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function isoFromLocalInput(v) {
  if (!v) return new Date().toISOString();
  const d = new Date(v);
  return d.toISOString();
}
function fmtDate(iso) {
  return new Date(iso).toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
}
function fmtTime(iso) {
  return new Date(iso).toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" });
}
function isSameDay(iso, d) {
  const a = new Date(iso);
  return a.getFullYear() === d.getFullYear() && a.getMonth() === d.getMonth() && a.getDate() === d.getDate();
}
function startOfDay(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function endOfDay(d) { const x = new Date(d); x.setHours(23, 59, 59, 999); return x; }

function calcTotals(boxes) {
  let birds = 0, ew = 0, lw = 0;
  boxes.forEach((b) => {
    ew += parseFloat(b.emptyWeight) || 0;
    lw += parseFloat(b.loadedWeight) || 0;
    birds += parseInt(b.birdCount) || 0;
  });
  const net = lw - ew;
  return {
    boxCount: boxes.length,
    birds,
    emptyWeight: round2(ew),
    loadedWeight: round2(lw),
    netWeight: round2(net),
    avgBirdWeight: birds > 0 ? round2(net / birds) : 0,
  };
}

function seedData() {
  const farms = [
    { id: "f1", name: "Green Valley Poultry", village: "Kondampatti", owner: "Murugan R", mobile: "9894512300" },
    { id: "f2", name: "Sunrise Broilers", village: "Anaikatti", owner: "Selvi K", mobile: "" },
    { id: "f3", name: "Om Sakthi Farm", village: "Thondamuthur", owner: "Karthik P", mobile: "9843221100" },
  ];
  const companies = ["Company 1", "Company 2"];

  const mkBoxes = (rows) => rows.map((r) => ({ emptyWeight: r[0], loadedWeight: r[1], birdCount: r[2] }));

  const today = new Date();
  const yesterday = new Date(today); yesterday.setDate(yesterday.getDate() - 1);

  const l1boxes = mkBoxes([[4.5, 32.1, 24], [4.5, 33.4, 25], [4.6, 31.8, 23], [4.5, 34.0, 26], [4.6, 32.6, 24]]);
  const l2boxes = mkBoxes([[4.4, 29.8, 22], [4.5, 30.5, 23], [4.5, 31.2, 24]]);
  const l3boxes = mkBoxes([[4.6, 33.9, 25], [4.5, 32.2, 24], [4.6, 34.5, 26], [4.5, 31.9, 23]]);

  const loadings = [
    {
      id: uid(), farmId: "f1", farmName: "Green Valley Poultry", vehicleNumber: "TN 38 AT 4521",
      dateTime: new Date(today.setHours(8, 45)).toISOString(), company: "Company 1",
      boxes: l1boxes, totals: calcTotals(l1boxes),
    },
    {
      id: uid(), farmId: "f2", farmName: "Sunrise Broilers", vehicleNumber: "TN 37 BQ 9012",
      dateTime: new Date(today.setHours(11, 10)).toISOString(), company: "Company 2",
      boxes: l2boxes, totals: calcTotals(l2boxes),
    },
    {
      id: uid(), farmId: "f3", farmName: "Om Sakthi Farm", vehicleNumber: "TN 38 CJ 7788",
      dateTime: new Date(yesterday.setHours(9, 20)).toISOString(), company: "Company 1",
      boxes: l3boxes, totals: calcTotals(l3boxes),
    },
  ];

  return { farms, companies, loadings, settings: { password: "12345" } };
}

/* ---------------------------------------------------------------------- */
/* Small UI atoms                                                         */
/* ---------------------------------------------------------------------- */

function GlobalStyle() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@500;600;700;800&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@500;600;700&display=swap');
      .pw-root{ font-family:'Inter',sans-serif; background:${C.egg}; color:${C.ink}; }
      .pw-display{ font-family:'Poppins',sans-serif; }
      .pw-mono{ font-family:'JetBrains Mono',monospace; }
      .pw-card{ background:#fff; border:1px solid ${C.border}; border-radius:18px; }
      .pw-btn-primary{ background:linear-gradient(135deg,${C.primary},${C.primaryDark}); color:#fff; border-radius:16px; font-weight:600; transition:transform .15s ease, box-shadow .15s ease; box-shadow:0 8px 20px -8px rgba(46,125,50,.55); }
      .pw-btn-primary:active{ transform:scale(.97); }
      .pw-btn-primary:disabled{ opacity:.45; box-shadow:none; }
      .pw-btn-soft{ background:${C.primaryLight}; color:${C.primaryDark}; border-radius:16px; font-weight:600; }
      .pw-readout{ background:#0F2A15; color:#8FE3A0; border-radius:14px; font-family:'JetBrains Mono',monospace; }
      .pw-input{ background:#fff; border:1.5px solid ${C.border}; border-radius:14px; padding:12px 14px; font-size:15px; width:100%; }
      .pw-input:focus{ outline:none; border-color:${C.primary}; box-shadow:0 0 0 3px ${C.primaryLight}; }
      .pw-input.pl-10{ padding-left:2.5rem; }
      .pw-input.pw-input-sm{ padding:8px 10px; font-size:13px; }
      .pw-input.pw-input-compact{ padding:10px 10px; font-size:12.5px; }
      .pw-label{ font-size:12.5px; font-weight:600; color:${C.muted}; letter-spacing:.02em; margin-bottom:6px; display:block; }
      .pw-label.pw-label-sm{ font-size:10.5px; margin-bottom:4px; }
      .pw-modal-sheet{ border-bottom-left-radius:0; border-bottom-right-radius:0; }
      @media (min-width:640px){ .pw-modal-sheet{ border-bottom-left-radius:18px; border-bottom-right-radius:18px; } }
      .rb-28{ border-bottom-left-radius:28px; border-bottom-right-radius:28px; }
      .rb-32{ border-bottom-left-radius:32px; border-bottom-right-radius:32px; }
      .mh-85{ max-height:85vh; }
      .mh-90{ max-height:90vh; }
      .t-10{ font-size:10px; } .t-10-5{ font-size:10.5px; } .t-11{ font-size:11px; } .t-11-5{ font-size:11.5px; }
      .t-12{ font-size:12px; } .t-12-5{ font-size:12.5px; } .t-13{ font-size:13px; } .t-13-5{ font-size:13.5px; }
      .t-14{ font-size:14px; } .t-15{ font-size:15px; } .t-16{ font-size:16px; } .t-17{ font-size:17px; }
      .t-18{ font-size:18px; } .t-22{ font-size:22px; }
      .pw-ticket{ border:2px dashed ${C.primary}; border-radius:18px; background:#fff; position:relative; }
      .pw-ticket::before, .pw-ticket::after{ content:''; position:absolute; width:16px; height:16px; background:${C.egg}; border-radius:50%; top:50%; transform:translateY(-50%); }
      .pw-ticket::before{ left:-9px; } .pw-ticket::after{ right:-9px; }
      @keyframes pwScaleIn{ from{opacity:0; transform:scale(.55);} to{opacity:1; transform:scale(1);} }
      .pw-scale-in{ animation:pwScaleIn .4s cubic-bezier(.34,1.56,.64,1); }
      @keyframes pwFadeUp{ from{opacity:0; transform:translateY(12px);} to{opacity:1; transform:translateY(0);} }
      .pw-fade-up{ animation:pwFadeUp .28s ease-out; }
      @keyframes pwPulse{ 0%,100%{opacity:1;} 50%{opacity:.55;} }
      .pw-pulse{ animation:pwPulse 1.6s ease-in-out infinite; }
      .pw-scrollbar-none::-webkit-scrollbar{ display:none; }
      @media print { .no-print{ display:none !important; } .print-area{ padding:0 !important; } }
    `}</style>
  );
}

function IconBadge({ Icon, bg, color, size = 40 }) {
  return (
    <div style={{ width: size, height: size, background: bg, color }} className="rounded-2xl flex items-center justify-center shrink-0">
      <Icon size={size * 0.5} />
    </div>
  );
}

function Toast({ toast }) {
  if (!toast) return null;
  const bad = toast.type === "error";
  return (
    <div
      className="fixed left-1/2 -translate-x-1/2 z-50 pw-fade-up px-4 py-3 rounded-2xl shadow-lg text-sm font-medium flex items-center gap-2"
      style={{ top: 18, background: bad ? "#FDECEA" : C.primaryDark, color: bad ? C.danger : "#fff" }}
    >
      {bad ? <AlertTriangle size={16} /> : <Check size={16} />}
      {toast.msg}
    </div>
  );
}

function EmptyState({ Icon, title, sub }) {
  return (
    <div className="flex flex-col items-center justify-center text-center py-14 px-6">
      <div className="w-16 h-16 rounded-full flex items-center justify-center mb-4" style={{ background: C.primaryLight, color: C.primary }}>
        <Icon size={28} />
      </div>
      <p className="pw-display font-semibold t-15" style={{ color: C.ink }}>{title}</p>
      {sub && <p className="text-sm mt-1 max-w-xs" style={{ color: C.muted }}>{sub}</p>}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Login                                                                  */
/* ---------------------------------------------------------------------- */

function LoginScreen({ onLogin, correctPassword }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [err, setErr] = useState("");

  function attemptLogin() {
    if (username.trim().toLowerCase() === "admin" && password.trim() === correctPassword) {
      setErr("");
      onLogin();
    } else {
      setErr("Incorrect username or password.");
    }
  }
  function onKeyDown(e) {
    if (e.key === "Enter") attemptLogin();
  }

  return (
    <div className="pw-root min-h-screen flex flex-col justify-center px-6 py-10 relative overflow-hidden">
      <GlobalStyle />
      <div className="absolute -top-20 -right-24 w-72 h-72 rounded-full" style={{ background: C.primaryLight, opacity: 0.7 }} />
      <div className="absolute -bottom-24 -left-16 w-64 h-64 rounded-full" style={{ background: C.amberLight, opacity: 0.6 }} />

      <div className="relative z-10 max-w-sm w-full mx-auto">
        <div className="flex flex-col items-center mb-8">
          <svg width="120" height="96" viewBox="0 0 120 96" className="mb-4">
            <circle cx="90" cy="20" r="14" fill={C.amber} opacity="0.85" />
            <path d="M10 80 L30 40 L50 80 Z" fill={C.primary} />
            <rect x="16" y="55" width="28" height="25" fill={C.primaryDark} />
            <rect x="24" y="64" width="10" height="16" fill={C.egg} />
            <ellipse cx="72" cy="66" rx="20" ry="16" fill={C.ink} />
            <circle cx="88" cy="58" r="7" fill={C.ink} />
            <path d="M95 56 L104 59 L95 62 Z" fill={C.amber} />
            <circle cx="91" cy="56.5" r="1.4" fill="#fff" />
            <path d="M78 50 Q84 42 90 50" stroke={C.amber} strokeWidth="3" fill="none" strokeLinecap="round" />
            <ellipse cx="60" cy="85" rx="52" ry="4" fill={C.primary} opacity="0.12" />
          </svg>
          <h1 className="pw-display t-22 font-bold" style={{ color: C.ink }}>PoultryLoad ERP</h1>
          <p className="t-13 mt-1" style={{ color: C.muted }}>Farm loading, weighed and tracked</p>
        </div>

        <div className="pw-card p-6 shadow-sm space-y-4">
          <div>
            <label className="pw-label">Username</label>
            <div className="relative">
              <User size={17} className="absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none" style={{ color: C.muted }} />
              <input
                className="pw-input pl-10"
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                onKeyDown={onKeyDown}
                placeholder="Enter username"
                autoCapitalize="none"
              />
            </div>
          </div>
          <div>
            <label className="pw-label">Password</label>
            <div className="relative">
              <Lock size={17} className="absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none" style={{ color: C.muted }} />
              <input
                type="password"
                className="pw-input pl-10"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                onKeyDown={onKeyDown}
                placeholder="Enter password"
              />
            </div>
          </div>
          {err && (
            <div className="t-13 px-3 py-2 rounded-xl flex items-center gap-2" style={{ background: "#FDECEA", color: C.danger }}>
              <AlertTriangle size={14} /> {err}
            </div>
          )}
          <button type="button" onClick={attemptLogin} className="pw-btn-primary w-full py-3.5 t-15">Log In</button>
          <p className="text-center t-12 pt-1" style={{ color: C.muted }}>Demo credentials — admin / 12345</p>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Shell (header + bottom nav)                                            */
/* ---------------------------------------------------------------------- */

function BottomNav({ screen, setScreen }) {
  const items = [
    { key: "dashboard", icon: Home, label: "Dashboard" },
    { key: "newLoading", icon: PlusCircle, label: "New", isFab: true },
    { key: "history", icon: ClipboardList, label: "History" },
    { key: "reports", icon: BarChart3, label: "Reports" },
    { key: "settings", icon: SettingsIcon, label: "Settings" },
  ];
  return (
    <div className="no-print fixed bottom-0 left-0 right-0 z-40 px-3 pt-2" style={{ background: "rgba(255,253,245,0.92)", backdropFilter: "blur(8px)", borderTop: `1px solid ${C.border}`, paddingBottom: "max(10px, env(safe-area-inset-bottom))" }}>
      <div className="max-w-md mx-auto flex items-end justify-between">
        {items.map((it) => {
          const active = screen === it.key;
          if (it.isFab) {
            return (
              <button key={it.key} onClick={() => setScreen("newLoading")} className="flex flex-col items-center -mt-7">
                <div className="w-14 h-14 rounded-full flex items-center justify-center shadow-lg" style={{ background: `linear-gradient(135deg, ${C.primary}, ${C.primaryDark})` }}>
                  <Plus size={26} color="#fff" />
                </div>
                <span className="t-10-5 font-semibold mt-1" style={{ color: C.primaryDark }}>New</span>
              </button>
            );
          }
          return (
            <button key={it.key} onClick={() => setScreen(it.key)} className="flex flex-col items-center gap-1 py-1.5 px-2 flex-1">
              <it.icon size={21} color={active ? C.primary : C.muted} strokeWidth={active ? 2.4 : 2} />
              <span className="t-10-5 font-medium" style={{ color: active ? C.primary : C.muted }}>{it.label}</span>
            </button>
          );
        })}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Dashboard                                                              */
/* ---------------------------------------------------------------------- */

function Dashboard({ data, setScreen, openLoading }) {
  const today = new Date();
  const todays = data.loadings.filter((l) => isSameDay(l.dateTime, today));
  const totalBirds = todays.reduce((s, l) => s + l.totals.birds, 0);
  const totalWeight = round2(todays.reduce((s, l) => s + l.totals.netWeight, 0));
  const totalBoxes = todays.reduce((s, l) => s + l.totals.boxCount, 0);
  const totalVehicles = new Set(todays.map((l) => l.vehicleNumber)).size;

  const byCompany = {};
  data.companies.forEach((c) => (byCompany[c] = 0));
  todays.forEach((l) => { byCompany[l.company] = round2((byCompany[l.company] || 0) + l.totals.netWeight); });

  const recent = [...data.loadings].sort((a, b) => new Date(b.dateTime) - new Date(a.dateTime)).slice(0, 4);

  const cards = [
    { icon: Bird, label: "Today's Total Birds", value: totalBirds, bg: C.primaryLight, color: C.primaryDark },
    { icon: Weight, label: "Today's Total Weight", value: `${totalWeight} kg`, bg: C.amberLight, color: C.amber },
    { icon: Package, label: "Total Boxes", value: totalBoxes, bg: C.primaryLight, color: C.primaryDark },
    { icon: Truck, label: "Today's Vehicles", value: totalVehicles, bg: C.amberLight, color: C.amber },
  ];

  return (
    <div className="pb-28">
      <div className="px-5 pt-6 pb-8 rb-32" style={{ background: `linear-gradient(135deg, ${C.primary}, ${C.primaryDark})` }}>
        <p className="t-13" style={{ color: "#CDEBD0" }}>{today.toLocaleDateString("en-IN", { weekday: "long", day: "numeric", month: "long" })}</p>
        <h1 className="pw-display text-2xl font-bold text-white mt-1">Good day, Admin</h1>
        <p className="t-13 mt-1" style={{ color: "#CDEBD0" }}>Here's how today's loading is going</p>
      </div>

      <div className="px-5 -mt-6 grid grid-cols-2 gap-3">
        {cards.map((c, i) => (
          <div key={i} className="pw-card p-4 shadow-sm pw-fade-up" style={{ animationDelay: `${i * 40}ms` }}>
            <IconBadge Icon={c.icon} bg={c.bg} color={c.color} />
            <p className="pw-mono text-xl font-bold mt-3" style={{ color: C.ink }}>{c.value}</p>
            <p className="t-12 mt-0.5" style={{ color: C.muted }}>{c.label}</p>
          </div>
        ))}
      </div>

      <div className="px-5 mt-3 grid grid-cols-2 gap-3">
        {data.companies.slice(0, 2).map((c, i) => (
          <div key={c} className="pw-card p-4 shadow-sm">
            <IconBadge Icon={Building2} bg={C.primaryLight} color={C.primaryDark} size={34} />
            <p className="pw-mono text-lg font-bold mt-2.5" style={{ color: C.ink }}>{byCompany[c] || 0} kg</p>
            <p className="t-12 mt-0.5" style={{ color: C.muted }}>{c} Load</p>
          </div>
        ))}
      </div>

      <div className="px-5 mt-5">
        <button onClick={() => setScreen("newLoading")} className="pw-btn-primary w-full py-4 flex items-center justify-center gap-2 t-15">
          <PlusCircle size={19} /> New Loading
        </button>
      </div>

      <div className="px-5 mt-7">
        <div className="flex items-center justify-between mb-3">
          <h2 className="pw-display font-semibold t-15" style={{ color: C.ink }}>Recent Loadings</h2>
          <button onClick={() => setScreen("history")} className="t-12-5 font-semibold flex items-center gap-0.5" style={{ color: C.primary }}>
            View all <ChevronRight size={14} />
          </button>
        </div>
        {recent.length === 0 ? (
          <EmptyState Icon={Feather} title="No loadings yet" sub="Tap New Loading to record your first truck of the day." />
        ) : (
          <div className="space-y-2.5">
            {recent.map((l) => (
              <button key={l.id} onClick={() => openLoading(l)} className="pw-card w-full p-3.5 flex items-center gap-3 text-left shadow-sm">
                <IconBadge Icon={Truck} bg={C.primaryLight} color={C.primaryDark} size={38} />
                <div className="flex-1 min-w-0">
                  <p className="font-semibold t-13-5 truncate" style={{ color: C.ink }}>{l.farmName}</p>
                  <p className="t-12" style={{ color: C.muted }}>{l.vehicleNumber} · {fmtTime(l.dateTime)}</p>
                </div>
                <div className="text-right">
                  <p className="pw-mono font-bold t-13-5" style={{ color: C.primaryDark }}>{l.totals.netWeight}kg</p>
                  <p className="t-11" style={{ color: C.muted }}>{l.totals.birds} birds</p>
                </div>
              </button>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Farm modal (add/edit)                                                  */
/* ---------------------------------------------------------------------- */

function FarmModal({ initial, onClose, onSave }) {
  const [name, setName] = useState(initial?.name || "");
  const [village, setVillage] = useState(initial?.village || "");
  const [owner, setOwner] = useState(initial?.owner || "");
  const [mobile, setMobile] = useState(initial?.mobile || "");
  const valid = name.trim() && village.trim() && owner.trim();

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center">
      <div className="absolute inset-0 bg-black/40" onClick={onClose} />
      <div className="relative pw-card w-full sm:max-w-sm pw-modal-sheet p-5 pb-7 pw-fade-up mh-90 overflow-y-auto">
        <div className="flex items-center justify-between mb-4">
          <h3 className="pw-display font-semibold t-16">{initial ? "Edit Farm" : "Add Farm"}</h3>
          <button onClick={onClose}><X size={20} color={C.muted} /></button>
        </div>
        <div className="space-y-3.5">
          <div>
            <label className="pw-label">Farm Name</label>
            <input className="pw-input" value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Green Valley Poultry" />
          </div>
          <div>
            <label className="pw-label">Village</label>
            <input className="pw-input" value={village} onChange={(e) => setVillage(e.target.value)} placeholder="e.g. Kondampatti" />
          </div>
          <div>
            <label className="pw-label">Owner Name</label>
            <input className="pw-input" value={owner} onChange={(e) => setOwner(e.target.value)} placeholder="e.g. Murugan R" />
          </div>
          <div>
            <label className="pw-label">Mobile Number (optional)</label>
            <input className="pw-input" value={mobile} onChange={(e) => setMobile(e.target.value.replace(/[^0-9]/g, ""))} placeholder="10-digit number" maxLength={10} />
          </div>
        </div>
        <button
          disabled={!valid}
          onClick={() => onSave({ id: initial?.id || uid(), name: name.trim(), village: village.trim(), owner: owner.trim(), mobile: mobile.trim() })}
          className="pw-btn-primary w-full py-3.5 mt-5 t-15"
        >
          Save Farm
        </button>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* New Loading                                                            */
/* ---------------------------------------------------------------------- */

function NewLoadingScreen({ data, persist, setScreen, showToast }) {
  const [step, setStep] = useState(1);
  const [farmId, setFarmId] = useState("");
  const [vehicle, setVehicle] = useState("");
  const [dt, setDt] = useState(nowLocalInput());
  const [company, setCompany] = useState(data.companies[0] || "");
  const [boxes, setBoxes] = useState([{ id: uid(), emptyWeight: "", loadedWeight: "", birdCount: "" }]);
  const [showFarmModal, setShowFarmModal] = useState(false);
  const [saving, setSaving] = useState(false);
  const [success, setSuccess] = useState(false);

  const farm = data.farms.find((f) => f.id === farmId);
  const totals = useMemo(() => calcTotals(boxes), [boxes]);

  function updateBox(id, field, value) {
    setBoxes((bs) => bs.map((b) => (b.id === id ? { ...b, [field]: value } : b)));
  }
  function addBox() {
    setBoxes((bs) => [...bs, { id: uid(), emptyWeight: "", loadedWeight: "", birdCount: "" }]);
  }
  function removeBox(id) {
    setBoxes((bs) => (bs.length > 1 ? bs.filter((b) => b.id !== id) : bs));
  }

  async function saveFarmQuick(newFarm) {
    const nextFarms = [...data.farms, newFarm];
    await persist({ ...data, farms: nextFarms });
    setFarmId(newFarm.id);
    setShowFarmModal(false);
    showToast("Farm added");
  }

  function goStep2() {
    if (!farmId) return showToast("Please select a farm", "error");
    if (!vehicle.trim()) return showToast("Please enter vehicle number", "error");
    setStep(2);
  }

  async function save() {
    const validBoxes = boxes.filter((b) => (parseFloat(b.loadedWeight) || 0) > 0 || (parseInt(b.birdCount) || 0) > 0);
    if (validBoxes.length === 0) return showToast("Add at least one box with weight or bird count", "error");
    setSaving(true);
    const loading = {
      id: uid(),
      farmId,
      farmName: farm.name,
      vehicleNumber: vehicle.trim().toUpperCase(),
      dateTime: isoFromLocalInput(dt),
      company,
      boxes: validBoxes.map((b) => ({ emptyWeight: b.emptyWeight, loadedWeight: b.loadedWeight, birdCount: b.birdCount })),
      totals: calcTotals(validBoxes),
    };
    await persist({ ...data, loadings: [loading, ...data.loadings] });
    setSaving(false);
    setSuccess(true);
    setTimeout(() => {
      setSuccess(false);
      setStep(1); setFarmId(""); setVehicle(""); setDt(nowLocalInput()); setCompany(data.companies[0] || "");
      setBoxes([{ id: uid(), emptyWeight: "", loadedWeight: "", birdCount: "" }]);
      setScreen("dashboard");
    }, 1400);
  }

  return (
    <div className="pb-40 relative">
      <div className="px-5 pt-6 pb-5 rb-28" style={{ background: `linear-gradient(135deg, ${C.primary}, ${C.primaryDark})` }}>
        <div className="flex items-center gap-3">
          {step === 2 && (
            <button onClick={() => setStep(1)} className="text-white"><ArrowLeft size={20} /></button>
          )}
          <div>
            <h1 className="pw-display text-lg font-bold text-white">New Loading</h1>
            <p className="t-12" style={{ color: "#CDEBD0" }}>Step {step} of 2 — {step === 1 ? "Trip details" : "Box entries"}</p>
          </div>
        </div>
        <div className="flex gap-1.5 mt-3">
          <div className="h-1.5 flex-1 rounded-full bg-white/90" />
          <div className="h-1.5 flex-1 rounded-full" style={{ background: step === 2 ? "#fff" : "rgba(255,255,255,0.35)" }} />
        </div>
      </div>

      {step === 1 && (
        <div className="px-5 mt-5 space-y-4 pw-fade-up">
          <div>
            <label className="pw-label">Select Farm</label>
            <select className="pw-input" value={farmId} onChange={(e) => setFarmId(e.target.value)}>
              <option value="">Choose a farm…</option>
              {data.farms.map((f) => <option key={f.id} value={f.id}>{f.name} — {f.village}</option>)}
            </select>
            <button onClick={() => setShowFarmModal(true)} className="mt-2 t-13 font-semibold flex items-center gap-1" style={{ color: C.primary }}>
              <Plus size={15} /> Add New Farm
            </button>
          </div>

          <div>
            <label className="pw-label">Vehicle Number</label>
            <input className="pw-input pw-mono uppercase" value={vehicle} onChange={(e) => setVehicle(e.target.value.toUpperCase())} placeholder="TN 38 AT 4521" />
          </div>

          <div>
            <label className="pw-label">Date &amp; Time</label>
            <div className="relative">
              <Clock size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2" style={{ color: C.muted }} />
              <input type="datetime-local" className="pw-input pl-10" value={dt} onChange={(e) => setDt(e.target.value)} />
            </div>
          </div>

          <div>
            <label className="pw-label">Taken By</label>
            <div className="grid grid-cols-2 gap-2.5">
              {data.companies.map((c) => (
                <button
                  key={c}
                  onClick={() => setCompany(c)}
                  className="rounded-2xl py-3.5 px-3 t-13-5 font-semibold border-2 flex items-center justify-center gap-2"
                  style={{ borderColor: company === c ? C.primary : C.border, background: company === c ? C.primaryLight : "#fff", color: company === c ? C.primaryDark : C.muted }}
                >
                  <span className="w-4 h-4 rounded-full border-2 flex items-center justify-center" style={{ borderColor: company === c ? C.primary : C.border }}>
                    {company === c && <span className="w-2 h-2 rounded-full" style={{ background: C.primary }} />}
                  </span>
                  {c}
                </button>
              ))}
            </div>
          </div>

          <button onClick={goStep2} className="pw-btn-primary w-full py-4 mt-2 flex items-center justify-center gap-2 t-15">
            Continue to Boxes <ChevronRight size={18} />
          </button>
        </div>
      )}

      {step === 2 && (
        <div className="px-5 mt-5 pw-fade-up">
          <div className="space-y-3">
            {boxes.map((b, i) => {
              const net = round2((parseFloat(b.loadedWeight) || 0) - (parseFloat(b.emptyWeight) || 0));
              return (
                <div key={b.id} className="pw-card p-3.5 shadow-sm">
                  <div className="flex items-center justify-between mb-2.5">
                    <span className="t-12-5 font-bold" style={{ color: C.primaryDark }}>Box #{i + 1}</span>
                    {boxes.length > 1 && (
                      <button onClick={() => removeBox(b.id)}><Trash2 size={15} color={C.danger} /></button>
                    )}
                  </div>
                  <div className="grid grid-cols-3 gap-2">
                    <div>
                      <label className="pw-label pw-label-sm">Empty (kg)</label>
                      <input inputMode="decimal" className="pw-input pw-input-sm pw-mono" value={b.emptyWeight} onChange={(e) => updateBox(b.id, "emptyWeight", e.target.value)} placeholder="0" />
                    </div>
                    <div>
                      <label className="pw-label pw-label-sm">Loaded (kg)</label>
                      <input inputMode="decimal" className="pw-input pw-input-sm pw-mono" value={b.loadedWeight} onChange={(e) => updateBox(b.id, "loadedWeight", e.target.value)} placeholder="0" />
                    </div>
                    <div>
                      <label className="pw-label pw-label-sm">Birds</label>
                      <input inputMode="numeric" className="pw-input pw-input-sm pw-mono" value={b.birdCount} onChange={(e) => updateBox(b.id, "birdCount", e.target.value)} placeholder="0" />
                    </div>
                  </div>
                  <div className="mt-2.5 pw-readout px-3 py-1.5 flex items-center justify-between">
                    <span className="t-11 opacity-80">NET WEIGHT</span>
                    <span className="font-bold t-14">{net} kg</span>
                  </div>
                </div>
              );
            })}
          </div>

          <button onClick={addBox} className="pw-btn-soft w-full py-3.5 mt-3.5 flex items-center justify-center gap-2 t-14">
            <Plus size={17} /> Add Box
          </button>

          <div className="pw-ticket p-4 mt-5">
            <p className="pw-display font-semibold t-13 mb-3 text-center" style={{ color: C.primaryDark }}>LIVE SUMMARY</p>
            <div className="grid grid-cols-2 gap-y-2.5 t-13">
              <span style={{ color: C.muted }}>Total Boxes</span><span className="text-right pw-mono font-semibold">{totals.boxCount}</span>
              <span style={{ color: C.muted }}>Total Birds</span><span className="text-right pw-mono font-semibold">{totals.birds}</span>
              <span style={{ color: C.muted }}>Empty Weight</span><span className="text-right pw-mono font-semibold">{totals.emptyWeight} kg</span>
              <span style={{ color: C.muted }}>Loaded Weight</span><span className="text-right pw-mono font-semibold">{totals.loadedWeight} kg</span>
              <span style={{ color: C.muted }}>Avg Bird Weight</span><span className="text-right pw-mono font-semibold">{totals.avgBirdWeight} kg</span>
            </div>
            <div className="mt-3 pw-readout px-3 py-2.5 flex items-center justify-between">
              <span className="t-12 opacity-80">TOTAL NET WEIGHT</span>
              <span className="font-bold t-18">{totals.netWeight} kg</span>
            </div>
          </div>

          <button disabled={saving} onClick={save} className="pw-btn-primary w-full py-4 mt-5 mb-4 t-15">
            {saving ? "Saving…" : "SAVE LOADING"}
          </button>
        </div>
      )}

      {showFarmModal && <FarmModal onClose={() => setShowFarmModal(false)} onSave={saveFarmQuick} />}

      {success && (
        <div className="fixed inset-0 z-50 flex items-center justify-center" style={{ background: "rgba(23,58,27,0.55)" }}>
          <div className="pw-card px-10 py-9 flex flex-col items-center pw-scale-in">
            <div className="w-16 h-16 rounded-full flex items-center justify-center mb-4" style={{ background: C.primaryLight }}>
              <Check size={30} color={C.primary} strokeWidth={3} />
            </div>
            <p className="pw-display font-bold t-16" style={{ color: C.ink }}>Loading Saved Successfully</p>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Loading Detail Modal                                                   */
/* ---------------------------------------------------------------------- */

function LoadingDetailModal({ loading, onClose }) {
  if (!loading) return null;
  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center">
      <div className="absolute inset-0 bg-black/40" onClick={onClose} />
      <div className="relative pw-card w-full sm:max-w-sm pw-modal-sheet p-5 pw-fade-up mh-85 overflow-y-auto">
        <div className="flex items-center justify-between mb-1">
          <h3 className="pw-display font-semibold t-16">{loading.farmName}</h3>
          <button onClick={onClose}><X size={20} color={C.muted} /></button>
        </div>
        <p className="t-12-5" style={{ color: C.muted }}>{fmtDate(loading.dateTime)} · {fmtTime(loading.dateTime)}</p>

        <div className="grid grid-cols-2 gap-2.5 mt-4 t-13">
          <div className="pw-card p-3"><span className="block t-11" style={{ color: C.muted }}>Vehicle</span><span className="pw-mono font-semibold">{loading.vehicleNumber}</span></div>
          <div className="pw-card p-3"><span className="block t-11" style={{ color: C.muted }}>Company</span><span className="font-semibold">{loading.company}</span></div>
        </div>

        <p className="pw-label mt-4 mb-2">Box Entries ({loading.boxes.length})</p>
        <div className="space-y-1.5 max-h-40 overflow-y-auto pw-scrollbar-none">
          {loading.boxes.map((b, i) => (
            <div key={i} className="flex items-center justify-between t-12-5 px-3 py-2 rounded-xl" style={{ background: C.primaryLight }}>
              <span className="font-semibold">Box #{i + 1}</span>
              <span className="pw-mono" style={{ color: C.muted }}>{b.emptyWeight}→{b.loadedWeight}kg</span>
              <span className="pw-mono font-semibold">{round2((parseFloat(b.loadedWeight) || 0) - (parseFloat(b.emptyWeight) || 0))}kg</span>
              <span className="pw-mono">{b.birdCount} birds</span>
            </div>
          ))}
        </div>

        <div className="pw-readout px-4 py-3 mt-4 flex items-center justify-between">
          <span className="t-12 opacity-80">TOTAL NET WEIGHT</span>
          <span className="font-bold t-17">{loading.totals.netWeight} kg</span>
        </div>
        <div className="grid grid-cols-3 gap-2 mt-2.5 text-center">
          <div className="pw-card p-2.5"><p className="pw-mono font-bold t-14">{loading.totals.birds}</p><p className="t-10-5" style={{ color: C.muted }}>Birds</p></div>
          <div className="pw-card p-2.5"><p className="pw-mono font-bold t-14">{loading.totals.boxCount}</p><p className="t-10-5" style={{ color: C.muted }}>Boxes</p></div>
          <div className="pw-card p-2.5"><p className="pw-mono font-bold t-14">{loading.totals.avgBirdWeight}</p><p className="t-10-5" style={{ color: C.muted }}>Avg kg/bird</p></div>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* History                                                                */
/* ---------------------------------------------------------------------- */

function HistoryScreen({ data, openLoading }) {
  const [text, setText] = useState("");
  const [company, setCompany] = useState("");
  const [date, setDate] = useState("");

  const filtered = data.loadings
    .filter((l) => {
      if (text && !(`${l.farmName} ${l.vehicleNumber}`.toLowerCase().includes(text.toLowerCase()))) return false;
      if (company && l.company !== company) return false;
      if (date && !isSameDay(l.dateTime, new Date(date))) return false;
      return true;
    })
    .sort((a, b) => new Date(b.dateTime) - new Date(a.dateTime));

  return (
    <div className="pb-28 px-5 pt-6">
      <h1 className="pw-display text-xl font-bold" style={{ color: C.ink }}>History</h1>
      <p className="t-13 mt-0.5" style={{ color: C.muted }}>{filtered.length} loading{filtered.length !== 1 ? "s" : ""} found</p>

      <div className="mt-4 space-y-2.5">
        <div className="relative">
          <Search size={16} className="absolute left-3.5 top-1/2 -translate-y-1/2" style={{ color: C.muted }} />
          <input className="pw-input pl-10" placeholder="Search farm or vehicle…" value={text} onChange={(e) => setText(e.target.value)} />
        </div>
        <div className="grid grid-cols-2 gap-2.5">
          <select className="pw-input" value={company} onChange={(e) => setCompany(e.target.value)}>
            <option value="">All Companies</option>
            {data.companies.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
          <input type="date" className="pw-input" value={date} onChange={(e) => setDate(e.target.value)} />
        </div>
      </div>

      <div className="mt-5 space-y-2.5">
        {filtered.length === 0 ? (
          <EmptyState Icon={ClipboardList} title="No loadings found" sub="Try adjusting your filters, or record a new loading." />
        ) : (
          filtered.map((l) => (
            <button key={l.id} onClick={() => openLoading(l)} className="pw-card w-full p-4 text-left shadow-sm block">
              <div className="flex items-center justify-between">
                <span className="font-semibold t-14" style={{ color: C.ink }}>{l.farmName}</span>
                <span className="t-10-5 font-semibold px-2 py-0.5 rounded-full" style={{ background: C.amberLight, color: C.amber }}>{l.company}</span>
              </div>
              <p className="t-12 mt-0.5" style={{ color: C.muted }}>{fmtDate(l.dateTime)} · {fmtTime(l.dateTime)} · {l.vehicleNumber}</p>
              <div className="flex items-center gap-4 mt-2.5 pw-mono t-12-5">
                <span><b>{l.totals.birds}</b> birds</span>
                <span><b>{l.totals.netWeight}</b> kg</span>
                <span><b>{l.totals.boxCount}</b> boxes</span>
              </div>
            </button>
          ))
        )}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Reports                                                                */
/* ---------------------------------------------------------------------- */

function ReportsScreen({ data, showToast }) {
  const [preset, setPreset] = useState("today");
  const [customFrom, setCustomFrom] = useState("");
  const [customTo, setCustomTo] = useState("");
  const [farm, setFarm] = useState("");
  const [vehicle, setVehicle] = useState("");
  const [company, setCompany] = useState("");

  const range = useMemo(() => {
    const now = new Date();
    if (preset === "today") return [startOfDay(now), endOfDay(now)];
    if (preset === "yesterday") { const y = new Date(now); y.setDate(y.getDate() - 1); return [startOfDay(y), endOfDay(y)]; }
    if (preset === "week") { const s = new Date(now); s.setDate(s.getDate() - 6); return [startOfDay(s), endOfDay(now)]; }
    if (preset === "month") { const s = new Date(now); s.setDate(s.getDate() - 29); return [startOfDay(s), endOfDay(now)]; }
    if (preset === "custom") { if (!customFrom || !customTo) return [null, null]; return [startOfDay(new Date(customFrom)), endOfDay(new Date(customTo))]; }
    return [null, null];
  }, [preset, customFrom, customTo]);

  const filtered = data.loadings.filter((l) => {
    const d = new Date(l.dateTime);
    if (range[0] && (d < range[0] || d > range[1])) return false;
    if (farm && l.farmId !== farm) return false;
    if (vehicle && !l.vehicleNumber.toLowerCase().includes(vehicle.toLowerCase())) return false;
    if (company && l.company !== company) return false;
    return true;
  });

  const totalBirds = filtered.reduce((s, l) => s + l.totals.birds, 0);
  const totalWeight = round2(filtered.reduce((s, l) => s + l.totals.netWeight, 0));
  const avg = totalBirds > 0 ? round2(totalWeight / totalBirds) : 0;
  const trips = filtered.length;

  function exportExcel() {
    if (filtered.length === 0) return showToast("No data to export", "error");
    const rows = filtered.map((l) => ({
      Date: fmtDate(l.dateTime), Time: fmtTime(l.dateTime), Farm: l.farmName, Vehicle: l.vehicleNumber,
      Company: l.company, Boxes: l.totals.boxCount, Birds: l.totals.birds,
      "Empty Weight (kg)": l.totals.emptyWeight, "Loaded Weight (kg)": l.totals.loadedWeight,
      "Net Weight (kg)": l.totals.netWeight, "Avg Bird Weight (kg)": l.totals.avgBirdWeight,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/octet-stream" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `poultry-report-${preset}.xlsx`;
    a.click();
    showToast("Excel report downloaded");
  }

  function exportPdf() { window.print(); }

  const presets = [
    { key: "today", label: "Today" }, { key: "yesterday", label: "Yesterday" },
    { key: "week", label: "Weekly" }, { key: "month", label: "Monthly" }, { key: "custom", label: "Custom" },
  ];

  return (
    <div className="pb-28 px-5 pt-6 print-area">
      <h1 className="pw-display text-xl font-bold" style={{ color: C.ink }}>Reports</h1>
      <p className="t-13 mt-0.5" style={{ color: C.muted }}>Analyze loading performance</p>

      <div className="flex gap-2 mt-4 overflow-x-auto pw-scrollbar-none no-print">
        {presets.map((p) => (
          <button key={p.key} onClick={() => setPreset(p.key)} className="px-4 py-2 rounded-full t-13 font-semibold whitespace-nowrap shrink-0"
            style={{ background: preset === p.key ? C.primary : "#fff", color: preset === p.key ? "#fff" : C.muted, border: `1.5px solid ${preset === p.key ? C.primary : C.border}` }}>
            {p.label}
          </button>
        ))}
      </div>

      {preset === "custom" && (
        <div className="grid grid-cols-2 gap-2.5 mt-3 no-print">
          <input type="date" className="pw-input" value={customFrom} onChange={(e) => setCustomFrom(e.target.value)} />
          <input type="date" className="pw-input" value={customTo} onChange={(e) => setCustomTo(e.target.value)} />
        </div>
      )}

      <div className="grid grid-cols-3 gap-2.5 mt-3 no-print">
        <select className="pw-input pw-input-compact" value={farm} onChange={(e) => setFarm(e.target.value)}>
          <option value="">All Farms</option>
          {data.farms.map((f) => <option key={f.id} value={f.id}>{f.name}</option>)}
        </select>
        <input className="pw-input pw-input-compact" placeholder="Vehicle" value={vehicle} onChange={(e) => setVehicle(e.target.value)} />
        <select className="pw-input pw-input-compact" value={company} onChange={(e) => setCompany(e.target.value)}>
          <option value="">All Companies</option>
          {data.companies.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>
      </div>

      <div className="pw-ticket p-4 mt-5">
        <p className="pw-display font-semibold t-13 mb-3 text-center" style={{ color: C.primaryDark }}>REPORT SUMMARY</p>
        <div className="grid grid-cols-2 gap-3">
          <div className="text-center"><p className="pw-mono text-xl font-bold" style={{ color: C.ink }}>{totalBirds}</p><p className="t-11-5" style={{ color: C.muted }}>Total Birds</p></div>
          <div className="text-center"><p className="pw-mono text-xl font-bold" style={{ color: C.ink }}>{totalWeight}kg</p><p className="t-11-5" style={{ color: C.muted }}>Total Weight</p></div>
          <div className="text-center"><p className="pw-mono text-xl font-bold" style={{ color: C.ink }}>{avg}kg</p><p className="t-11-5" style={{ color: C.muted }}>Avg Bird Weight</p></div>
          <div className="text-center"><p className="pw-mono text-xl font-bold" style={{ color: C.ink }}>{trips}</p><p className="t-11-5" style={{ color: C.muted }}>Trips</p></div>
        </div>
      </div>

      <div className="grid grid-cols-2 gap-2.5 mt-4 no-print">
        <button onClick={exportPdf} className="pw-btn-soft py-3 flex items-center justify-center gap-2 t-13-5"><Printer size={16} /> Print / PDF</button>
        <button onClick={exportExcel} className="pw-btn-soft py-3 flex items-center justify-center gap-2 t-13-5"><FileSpreadsheet size={16} /> Excel</button>
      </div>

      <p className="pw-label mt-6 mb-2">Trips in range ({filtered.length})</p>
      <div className="space-y-2">
        {filtered.length === 0 ? (
          <EmptyState Icon={BarChart3} title="No data for this range" sub="Adjust the filters above to see report data." />
        ) : filtered.slice(0, 30).map((l) => (
          <div key={l.id} className="pw-card p-3 flex items-center justify-between t-12-5">
            <div><p className="font-semibold">{l.farmName}</p><p style={{ color: C.muted }}>{fmtDate(l.dateTime)} · {l.vehicleNumber}</p></div>
            <p className="pw-mono font-semibold">{l.totals.netWeight}kg</p>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Settings                                                               */
/* ---------------------------------------------------------------------- */

function SettingsScreen({ data, persist, showToast, onLogout }) {
  const [view, setView] = useState("main");
  const [editingFarm, setEditingFarm] = useState(null);
  const [showFarmModal, setShowFarmModal] = useState(false);
  const [newCompany, setNewCompany] = useState("");
  const fileRef = useRef(null);

  const [curPass, setCurPass] = useState("");
  const [newPass, setNewPass] = useState("");
  const [confPass, setConfPass] = useState("");

  async function saveFarm(f) {
    const exists = data.farms.some((x) => x.id === f.id);
    const nextFarms = exists ? data.farms.map((x) => (x.id === f.id ? f : x)) : [...data.farms, f];
    await persist({ ...data, farms: nextFarms });
    setShowFarmModal(false); setEditingFarm(null);
    showToast(exists ? "Farm updated" : "Farm added");
  }
  async function deleteFarm(id) {
    await persist({ ...data, farms: data.farms.filter((f) => f.id !== id) });
    showToast("Farm removed");
  }
  async function addCompany() {
    if (!newCompany.trim()) return;
    if (data.companies.includes(newCompany.trim())) return showToast("Company already exists", "error");
    await persist({ ...data, companies: [...data.companies, newCompany.trim()] });
    setNewCompany("");
    showToast("Company added");
  }
  async function removeCompany(c) {
    if (data.companies.length <= 1) return showToast("At least one company is required", "error");
    await persist({ ...data, companies: data.companies.filter((x) => x !== c) });
    showToast("Company removed");
  }
  async function changePassword() {
    if (curPass !== data.settings.password) return showToast("Current password is incorrect", "error");
    if (newPass.length < 4) return showToast("New password must be at least 4 characters", "error");
    if (newPass !== confPass) return showToast("Passwords do not match", "error");
    await persist({ ...data, settings: { ...data.settings, password: newPass } });
    setCurPass(""); setNewPass(""); setConfPass("");
    setView("main");
    showToast("Password updated");
  }
  function backup() {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `poultry-erp-backup-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    showToast("Backup downloaded");
  }
  function restore(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async () => {
      try {
        const parsed = JSON.parse(reader.result);
        if (!parsed.farms || !parsed.loadings || !parsed.companies) throw new Error("bad shape");
        await persist(parsed);
        showToast("Database restored");
      } catch {
        showToast("Invalid backup file", "error");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  }

  const Row = ({ Icon, label, sub, onClick, danger }) => (
    <button onClick={onClick} className="pw-card w-full p-4 flex items-center gap-3 text-left shadow-sm">
      <IconBadge Icon={Icon} bg={danger ? "#FDECEA" : C.primaryLight} color={danger ? C.danger : C.primaryDark} size={38} />
      <div className="flex-1">
        <p className="font-semibold t-13-5" style={{ color: danger ? C.danger : C.ink }}>{label}</p>
        {sub && <p className="t-11-5" style={{ color: C.muted }}>{sub}</p>}
      </div>
      <ChevronRight size={17} color={C.muted} />
    </button>
  );

  if (view === "farms") {
    return (
      <div className="pb-28 px-5 pt-6">
        <div className="flex items-center gap-2 mb-1">
          <button onClick={() => setView("main")}><ArrowLeft size={19} color={C.ink} /></button>
          <h1 className="pw-display text-xl font-bold" style={{ color: C.ink }}>Manage Farms</h1>
        </div>
        <button onClick={() => { setEditingFarm(null); setShowFarmModal(true); }} className="pw-btn-primary w-full py-3.5 mt-4 flex items-center justify-center gap-2 t-14">
          <Plus size={17} /> Add Farm
        </button>
        <div className="mt-4 space-y-2.5">
          {data.farms.length === 0 ? (
            <EmptyState Icon={Building2} title="No farms yet" sub="Add a farm to start recording loadings." />
          ) : data.farms.map((f) => (
            <div key={f.id} className="pw-card p-4 shadow-sm">
              <div className="flex items-start justify-between">
                <div>
                  <p className="font-semibold t-14" style={{ color: C.ink }}>{f.name}</p>
                  <p className="t-12 flex items-center gap-1 mt-0.5" style={{ color: C.muted }}><MapPin size={12} /> {f.village}</p>
                  <p className="t-12 flex items-center gap-1 mt-0.5" style={{ color: C.muted }}><User size={12} /> {f.owner}</p>
                  {f.mobile && <p className="t-12 flex items-center gap-1 mt-0.5" style={{ color: C.muted }}><Phone size={12} /> {f.mobile}</p>}
                </div>
                <div className="flex gap-2">
                  <button onClick={() => { setEditingFarm(f); setShowFarmModal(true); }} className="w-8 h-8 rounded-full flex items-center justify-center" style={{ background: C.primaryLight }}><Pencil size={14} color={C.primaryDark} /></button>
                  <button onClick={() => deleteFarm(f.id)} className="w-8 h-8 rounded-full flex items-center justify-center" style={{ background: "#FDECEA" }}><Trash2 size={14} color={C.danger} /></button>
                </div>
              </div>
            </div>
          ))}
        </div>
        {showFarmModal && <FarmModal initial={editingFarm} onClose={() => { setShowFarmModal(false); setEditingFarm(null); }} onSave={saveFarm} />}
      </div>
    );
  }

  if (view === "companies") {
    return (
      <div className="pb-28 px-5 pt-6">
        <div className="flex items-center gap-2 mb-4">
          <button onClick={() => setView("main")}><ArrowLeft size={19} color={C.ink} /></button>
          <h1 className="pw-display text-xl font-bold" style={{ color: C.ink }}>Manage Companies</h1>
        </div>
        <div className="flex gap-2 mb-4">
          <input className="pw-input" placeholder="New company name" value={newCompany} onChange={(e) => setNewCompany(e.target.value)} />
          <button onClick={addCompany} className="pw-btn-primary px-5"><Plus size={18} /></button>
        </div>
        <div className="space-y-2.5">
          {data.companies.map((c) => (
            <div key={c} className="pw-card p-4 flex items-center justify-between shadow-sm">
              <span className="font-semibold t-14">{c}</span>
              <button onClick={() => removeCompany(c)}><Trash2 size={16} color={C.danger} /></button>
            </div>
          ))}
        </div>
      </div>
    );
  }

  if (view === "password") {
    return (
      <div className="pb-28 px-5 pt-6">
        <div className="flex items-center gap-2 mb-4">
          <button onClick={() => setView("main")}><ArrowLeft size={19} color={C.ink} /></button>
          <h1 className="pw-display text-xl font-bold" style={{ color: C.ink }}>Change Password</h1>
        </div>
        <div className="space-y-3.5">
          <div><label className="pw-label">Current Password</label><input type="password" className="pw-input" value={curPass} onChange={(e) => setCurPass(e.target.value)} /></div>
          <div><label className="pw-label">New Password</label><input type="password" className="pw-input" value={newPass} onChange={(e) => setNewPass(e.target.value)} /></div>
          <div><label className="pw-label">Confirm New Password</label><input type="password" className="pw-input" value={confPass} onChange={(e) => setConfPass(e.target.value)} /></div>
        </div>
        <button onClick={changePassword} className="pw-btn-primary w-full py-3.5 mt-5 t-14">Update Password</button>
      </div>
    );
  }

  return (
    <div className="pb-28 px-5 pt-6">
      <h1 className="pw-display text-xl font-bold" style={{ color: C.ink }}>Settings</h1>
      <p className="t-13 mt-0.5 mb-4" style={{ color: C.muted }}>Manage your farm ERP</p>

      <p className="pw-label mb-2">General</p>
      <div className="space-y-2.5">
        <Row Icon={Building2} label="Manage Farms" sub={`${data.farms.length} farms registered`} onClick={() => setView("farms")} />
        <Row Icon={Truck} label="Manage Companies" sub={`${data.companies.length} companies`} onClick={() => setView("companies")} />
        <Row Icon={KeyRound} label="Change Password" onClick={() => setView("password")} />
      </div>

      <p className="pw-label mt-6 mb-2">Data</p>
      <div className="space-y-2.5">
        <Row Icon={DatabaseBackup} label="Backup Database" sub="Download all data as a JSON file" onClick={backup} />
        <Row Icon={Upload} label="Restore Database" sub="Load data from a backup file" onClick={() => fileRef.current?.click()} />
        <input ref={fileRef} type="file" accept="application/json" className="hidden" onChange={restore} />
      </div>

      <p className="pw-label mt-6 mb-2">Account</p>
      <Row Icon={LogOut} label="Logout" danger onClick={onLogout} />

      <p className="text-center t-11-5 mt-8" style={{ color: C.muted }}>PoultryLoad ERP · installable PWA · works offline on your device</p>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* App                                                                    */
/* ---------------------------------------------------------------------- */

export default function App() {
  const [ready, setReady] = useState(false);
  const [authed, setAuthed] = useState(false);
  const [screen, setScreen] = useState("dashboard");
  const [data, setData] = useState(null);
  const [toast, setToast] = useState(null);
  const [selectedLoading, setSelectedLoading] = useState(null);

  useEffect(() => { load(); }, []);

  async function load() {
    let parsed = null;
    try {
      const res = await window.storage.get(STORAGE_KEY, false);
      if (res && res.value) parsed = JSON.parse(res.value);
    } catch (e) { parsed = null; }
    if (!parsed) {
      parsed = seedData();
      try { await window.storage.set(STORAGE_KEY, JSON.stringify(parsed), false); } catch (e) {}
    }
    setData(parsed);
    setReady(true);
  }

  async function persist(next) {
    setData(next);
    try {
      const res = await window.storage.set(STORAGE_KEY, JSON.stringify(next), false);
      if (!res) showToast("Could not save changes", "error");
    } catch (e) {
      showToast("Storage error — changes may not persist", "error");
    }
  }

  function showToast(msg, type = "success") {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 2400);
  }

  if (!ready || !data) {
    return (
      <div className="pw-root min-h-screen flex items-center justify-center">
        <GlobalStyle />
        <div className="flex flex-col items-center pw-pulse">
          <div className="w-12 h-12 rounded-2xl flex items-center justify-center mb-3" style={{ background: C.primaryLight }}>
            <Bird size={22} color={C.primary} />
          </div>
          <p className="t-13 font-medium" style={{ color: C.muted }}>Loading PoultryLoad ERP…</p>
        </div>
      </div>
    );
  }

  if (!authed) {
    return <LoginScreen onLogin={() => setAuthed(true)} correctPassword={data.settings.password} />;
  }

  return (
    <div className="pw-root min-h-screen">
      <GlobalStyle />
      <Toast toast={toast} />
      <div className="max-w-md mx-auto relative">
        {screen === "dashboard" && <Dashboard data={data} setScreen={setScreen} openLoading={setSelectedLoading} />}
        {screen === "newLoading" && <NewLoadingScreen data={data} persist={persist} setScreen={setScreen} showToast={showToast} />}
        {screen === "history" && <HistoryScreen data={data} openLoading={setSelectedLoading} />}
        {screen === "reports" && <ReportsScreen data={data} showToast={showToast} />}
        {screen === "settings" && (
          <SettingsScreen data={data} persist={persist} showToast={showToast} onLogout={() => { setAuthed(false); setScreen("dashboard"); }} />
        )}
        <BottomNav screen={screen} setScreen={setScreen} />
      </div>
      {selectedLoading && <LoadingDetailModal loading={selectedLoading} onClose={() => setSelectedLoading(null)} />}
    </div>
  );
}
