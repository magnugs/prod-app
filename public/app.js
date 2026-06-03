import React, { useState, useEffect, useMemo, useRef } from "react";
import { Upload, Search, Plus, Edit2, Trash2, Download, X, AlertCircle, CheckCircle2, Clock, Users, Activity, FileJson, ChevronDown, ChevronUp, FileSpreadsheet, Play, Square, RefreshCw } from "lucide-react";
import * as XLSX from "xlsx";
const TOKEN_KEY = "bo_token_v1";
const ROLE_KEY = "bo_role_v1";
function getToken() {
  try {
    return localStorage.getItem(TOKEN_KEY);
  } catch {
    return null;
  }
}
function setToken(t) {
  try {
    t ? localStorage.setItem(TOKEN_KEY, t) : localStorage.removeItem(TOKEN_KEY);
  } catch {
  }
}
function getRole() {
  try {
    return localStorage.getItem(ROLE_KEY);
  } catch {
    return null;
  }
}
function setRole(r) {
  try {
    r ? localStorage.setItem(ROLE_KEY, r) : localStorage.removeItem(ROLE_KEY);
  } catch {
  }
}
class AuthError extends Error {
  constructor(message) {
    super(message);
    this.name = "AuthError";
    this.status = 401;
  }
}
class ApiError extends Error {
  constructor(message, status) {
    super(message);
    this.name = "ApiError";
    this.status = status;
  }
}
async function apiFetch(path, opts = {}) {
  const token = getToken();
  const headers = { "content-type": "application/json", ...opts.headers || {} };
  if (token) headers.authorization = `Bearer ${token}`;
  const res = await fetch(path, { ...opts, headers });
  if (res.status === 401) {
    setToken(null);
    setRole(null);
    window.dispatchEvent(new CustomEvent("auth-required"));
    throw new AuthError("P\xE5logging utl\xF8pt");
  }
  if (!res.ok) {
    let msg = `HTTP ${res.status}`;
    try {
      const body = await res.json();
      if (body?.error) msg = body.error;
    } catch {
    }
    throw new ApiError(msg, res.status);
  }
  if (res.status === 204) return null;
  return await res.json();
}
async function apiLogin(password, asAdmin = false) {
  const res = await fetch("/api/login", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({ password, asAdmin })
  });
  if (!res.ok) {
    let msg = "Feil passord";
    try {
      const body = await res.json();
      if (body?.error) msg = body.error;
    } catch {
    }
    throw new ApiError(msg, res.status);
  }
  const data = await res.json();
  setToken(data.token);
  setRole(data.role);
  return data;
}
async function apiCheckSession() {
  return await apiFetch("/api/me");
}
async function loadRegistrations() {
  return await apiFetch("/api/registrations");
}
async function saveRegistration(reg) {
  return await apiFetch("/api/registrations", { method: "POST", body: JSON.stringify(reg) });
}
async function deleteRegistration(id, adminPassword) {
  await apiLogin(adminPassword, true);
  return await apiFetch(`/api/registrations/${encodeURIComponent(id)}`, { method: "DELETE" });
}
async function loadActiveSessions() {
  return await apiFetch("/api/active");
}
async function saveActiveSession(session) {
  return await apiFetch("/api/active", { method: "POST", body: JSON.stringify(session) });
}
async function deleteActiveSession(boNumber, line) {
  const q = new URLSearchParams({ bo: boNumber, line: String(line) });
  return await apiFetch(`/api/active?${q}`, { method: "DELETE" });
}
async function loadBatchOrders() {
  return await apiFetch("/api/batch-orders");
}
function cleanDate(iso) {
  if (!iso || typeof iso !== "string") return "";
  if (iso.startsWith("1900-")) return "";
  return iso;
}
function mapBatchOrders(rawData) {
  const records = Array.isArray(rawData) ? rawData : rawData?.value || [];
  if (!Array.isArray(records) || records.length === 0) return [];
  return records.filter((d) => d && d.BatchOrderNumber).map((d) => ({
    BatchOrderNumber: d.BatchOrderNumber,
    BatchOrderName: d.BatchOrderName || "",
    DataAreaId: d.dataAreaId || "",
    ItemNumber: d.ItemNumber || "",
    ItemBatchNumber: d.ItemBatchNumber || "",
    Warehouse: d.PlannedReceiptWarehouseId || "",
    Site: d.ProductionSiteId || "",
    Status: d.BatchOrderStatus || "Unknown",
    RemainderStatus: d.BatchOrderRemainderStatus || "",
    ScheduledDate: cleanDate(d.ScheduledDate),
    ScheduledStartDate: cleanDate(d.ScheduledStartDate),
    ScheduledEndDate: cleanDate(d.ScheduledEndDate),
    StartedDate: cleanDate(d.StartedDate),
    EndedDate: cleanDate(d.EndedDate),
    DeliveryDate: cleanDate(d.DeliveryDate),
    ScheduledQuantity: d.ScheduledQuantity ?? 0,
    StartedQuantity: d.StartedQuantity ?? 0,
    EstimatedQuantity: d.EstimatedQuantity ?? 0,
    Dimension: d.DefaultLedgerDimensionDisplayValue || "",
    ProductionPool: d.ProductionPoolId || "",
    Priority: d.BatchOrderPriority ?? 0
  }));
}
function fmtDate(iso) {
  if (!iso) return "\u2014";
  try {
    const d = new Date(iso);
    return d.toLocaleDateString("nb-NO", { year: "numeric", month: "2-digit", day: "2-digit" });
  } catch {
    return iso;
  }
}
function fmtDateTime(iso) {
  if (!iso) return "\u2014";
  try {
    const d = new Date(iso);
    return d.toLocaleString("nb-NO", { year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" });
  } catch {
    return iso;
  }
}
function durationMinutes(start, end) {
  if (!start || !end) return null;
  const s = new Date(start).getTime();
  const e = new Date(end).getTime();
  if (isNaN(s) || isNaN(e) || e <= s) return null;
  return Math.round((e - s) / 6e4);
}
function personHours(start, end, people) {
  const mins = durationMinutes(start, end);
  if (mins == null || !people) return null;
  return +(mins * people / 60).toFixed(2);
}
function fmtMinutes(min) {
  if (min == null) return "\u2014";
  const h = Math.floor(min / 60);
  const m = min % 60;
  return h > 0 ? `${h}t ${m}m` : `${m}m`;
}
function isoLocalFromMs(ms) {
  const d = new Date(ms);
  d.setSeconds(0, 0);
  const tz = d.getTimezoneOffset() * 6e4;
  return new Date(d - tz).toISOString().slice(0, 16);
}
function isoLocalNow() {
  return isoLocalFromMs(Date.now());
}
function toCSV(rows, headers) {
  const escape = (v) => {
    if (v == null) return "";
    const s = String(v);
    if (s.includes(",") || s.includes('"') || s.includes("\n")) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  };
  const lines = [headers.map((h) => escape(h.label)).join(",")];
  for (const r of rows) {
    lines.push(headers.map((h) => escape(typeof h.get === "function" ? h.get(r) : r[h.key])).join(","));
  }
  return lines.join("\n");
}
function download(filename, content, type = "text/csv;charset=utf-8") {
  const blob = new Blob(["\uFEFF" + content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
function LoginScreen({ onSuccess }) {
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const handleSubmit = async (e) => {
    e?.preventDefault?.();
    if (!password) return;
    setBusy(true);
    setErr("");
    try {
      await apiLogin(password, false);
      onSuccess();
    } catch (e2) {
      setErr(e2?.message || "P\xE5logging feilet");
      setBusy(false);
    }
  };
  return /* @__PURE__ */ React.createElement("div", { className: "min-h-screen bg-stone-950 flex items-center justify-center p-6" }, /* @__PURE__ */ React.createElement("form", { onSubmit: handleSubmit, className: "w-full max-w-sm" }, /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-3 mb-8 justify-center" }, /* @__PURE__ */ React.createElement("div", { className: "w-10 h-10 bg-amber-400 flex items-center justify-center text-stone-950 font-bold font-mono text-xl" }, "B"), /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono tracking-[0.3em] text-amber-400" }, "BATCH ORDER TRACKER")), /* @__PURE__ */ React.createElement("div", { className: "bg-stone-900/40 border border-stone-800 p-6 sm:p-8" }, /* @__PURE__ */ React.createElement("label", { className: "block text-[10px] font-mono tracking-wider text-stone-400 mb-3" }, "PASSORD"), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "password",
      value: password,
      onChange: (e) => setPassword(e.target.value),
      autoFocus: true,
      disabled: busy,
      className: "w-full bg-stone-900 border border-stone-700 px-4 py-3 text-stone-100 font-mono text-base focus:border-amber-400 focus:outline-none disabled:opacity-50",
      placeholder: "\u2022\u2022\u2022\u2022\u2022\u2022\u2022\u2022"
    }
  ), err && /* @__PURE__ */ React.createElement("div", { className: "mt-4 p-3 border border-rose-500/40 bg-rose-500/10 text-rose-200 text-xs font-mono flex gap-2" }, /* @__PURE__ */ React.createElement(AlertCircle, { className: "w-4 h-4 flex-shrink-0 mt-0.5" }), /* @__PURE__ */ React.createElement("span", null, err)), /* @__PURE__ */ React.createElement(
    "button",
    {
      type: "submit",
      disabled: busy || !password,
      className: "w-full mt-5 py-3 bg-amber-400 hover:bg-amber-300 disabled:opacity-30 disabled:cursor-not-allowed text-stone-950 font-mono font-bold text-sm tracking-wider"
    },
    busy ? "LOGGER INN ..." : "LOGG INN"
  ))));
}
function AdminPasswordModal({ onConfirm, onCancel }) {
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const handleSubmit = async (e) => {
    e?.preventDefault?.();
    if (!password) return;
    setBusy(true);
    setErr("");
    try {
      await onConfirm(password);
    } catch (e2) {
      setErr(e2?.message || "Feil admin-passord");
      setBusy(false);
    }
  };
  return /* @__PURE__ */ React.createElement("div", { className: "fixed inset-0 z-[60] bg-black/80 backdrop-blur-sm flex items-center justify-center p-4" }, /* @__PURE__ */ React.createElement("form", { onSubmit: handleSubmit, className: "bg-stone-950 border border-rose-500/40 max-w-md w-full" }, /* @__PURE__ */ React.createElement("div", { className: "px-6 py-4 border-b border-stone-800 flex items-center justify-between" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono tracking-[0.3em] text-rose-400" }, "ADMIN-OPERASJON"), /* @__PURE__ */ React.createElement("div", { className: "font-mono text-stone-100 text-lg mt-1" }, "Slett registrering")), /* @__PURE__ */ React.createElement("button", { type: "button", onClick: onCancel, className: "text-stone-500 hover:text-stone-200 p-1" }, /* @__PURE__ */ React.createElement(X, { className: "w-5 h-5" }))), /* @__PURE__ */ React.createElement("div", { className: "px-6 py-5" }, /* @__PURE__ */ React.createElement("div", { className: "flex gap-3 mb-4 p-3 bg-rose-500/5 border border-rose-500/20 text-xs text-stone-300" }, /* @__PURE__ */ React.createElement(AlertCircle, { className: "w-4 h-4 text-rose-400 flex-shrink-0 mt-0.5" }), /* @__PURE__ */ React.createElement("div", null, "Denne handlingen kan ikke angres. Registreringen fjernes permanent fra databasen.")), /* @__PURE__ */ React.createElement("label", { className: "block text-[10px] font-mono tracking-wider text-stone-400 mb-2" }, "ADMIN-PASSORD"), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "password",
      value: password,
      onChange: (e) => setPassword(e.target.value),
      autoFocus: true,
      disabled: busy,
      className: "w-full bg-stone-900 border border-stone-700 px-4 py-3 text-stone-100 font-mono text-sm focus:border-rose-400 focus:outline-none disabled:opacity-50",
      placeholder: "\u2022\u2022\u2022\u2022\u2022\u2022\u2022\u2022"
    }
  ), err && /* @__PURE__ */ React.createElement("div", { className: "mt-3 p-2 border border-rose-500/40 bg-rose-500/10 text-rose-200 text-xs font-mono" }, err)), /* @__PURE__ */ React.createElement("div", { className: "px-6 py-4 border-t border-stone-800 flex items-center justify-end gap-2" }, /* @__PURE__ */ React.createElement(
    "button",
    {
      type: "button",
      onClick: onCancel,
      className: "px-5 py-2 border border-stone-700 hover:border-stone-500 text-stone-300 text-sm font-mono"
    },
    "AVBRYT"
  ), /* @__PURE__ */ React.createElement(
    "button",
    {
      type: "submit",
      disabled: busy || !password,
      className: "px-5 py-2 bg-rose-500 hover:bg-rose-400 disabled:opacity-30 text-stone-950 font-mono font-bold text-sm inline-flex items-center gap-2"
    },
    /* @__PURE__ */ React.createElement(Trash2, { className: "w-4 h-4" }),
    busy ? "SLETTER ..." : "BEKREFT SLETT"
  ))));
}
function StatusBadge({ status }) {
  const styles = {
    Created: "bg-stone-400/10 text-stone-300 border-stone-400/30",
    StartedUp: "bg-amber-400/15 text-amber-300 border-amber-400/40",
    ReportedFinished: "bg-blue-400/10 text-blue-300 border-blue-400/30",
    Completed: "bg-emerald-400/10 text-emerald-300 border-emerald-400/25",
    CostEstimated: "bg-violet-400/10 text-violet-300 border-violet-400/25",
    Unknown: "bg-stone-500/10 text-stone-400 border-stone-500/25"
  };
  const label = {
    Created: "OPPRETTET",
    StartedUp: "STARTET",
    ReportedFinished: "FERDIGSTILT",
    Completed: "FULLF\xD8RT",
    CostEstimated: "ESTIMERT",
    Unknown: "UKJENT"
  }[status] || (status || "\u2014").toUpperCase();
  return /* @__PURE__ */ React.createElement("span", { className: `inline-flex items-center px-2 py-0.5 text-[10px] font-mono font-bold tracking-wider border ${styles[status] || styles.Unknown}` }, label);
}
function LineChip({ line }) {
  const colors = {
    1: "bg-rose-400/15 text-rose-200 border-rose-400/40",
    2: "bg-amber-400/15 text-amber-200 border-amber-400/40",
    3: "bg-emerald-400/15 text-emerald-200 border-emerald-400/40",
    4: "bg-sky-400/15 text-sky-200 border-sky-400/40",
    5: "bg-violet-400/15 text-violet-200 border-violet-400/40"
  };
  return /* @__PURE__ */ React.createElement("span", { className: `inline-flex items-center justify-center w-7 h-7 text-xs font-mono font-bold border ${colors[line]}` }, "L", line);
}
function UploadPanel({ onLoad, onDemo }) {
  const inputRef = useRef(null);
  const [err, setErr] = useState("");
  const [dragging, setDragging] = useState(false);
  const handleFile = async (file) => {
    setErr("");
    try {
      const text = await file.text();
      const json = JSON.parse(text);
      const orders = mapBatchOrders(json);
      if (orders.length === 0) {
        setErr("Ingen batch orders funnet i filen. Forventer JSON-array eller { value: [...] } med BatchOrderNumber-felt.");
        return;
      }
      onLoad(orders, file.name);
    } catch (e) {
      setErr("Kunne ikke lese filen: " + e.message);
    }
  };
  return /* @__PURE__ */ React.createElement("div", { className: "max-w-3xl mx-auto pt-16 px-6" }, /* @__PURE__ */ React.createElement("div", { className: "mb-10" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono tracking-[0.3em] text-amber-400 mb-3" }, "SYSTEM // 01"), /* @__PURE__ */ React.createElement("h2", { className: "text-3xl font-bold text-stone-100 mb-2", style: { fontFamily: "'JetBrains Mono', monospace" } }, "Last inn batch orders"), /* @__PURE__ */ React.createElement("p", { className: "text-stone-400 text-sm leading-relaxed max-w-xl" }, "Last opp en JSON-fil hentet fra Dynamics 365 F&O OData-endepunktet, eller en hvilken som helst JSON-array med batch order-linjer. Linjer aggregeres automatisk per ", /* @__PURE__ */ React.createElement("code", { className: "text-amber-300 bg-stone-900/60 px-1.5 py-0.5 text-xs" }, "BatchOrderNumber"), ".")), /* @__PURE__ */ React.createElement(
    "div",
    {
      onDragOver: (e) => {
        e.preventDefault();
        setDragging(true);
      },
      onDragLeave: () => setDragging(false),
      onDrop: (e) => {
        e.preventDefault();
        setDragging(false);
        if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
      },
      onClick: () => inputRef.current?.click(),
      className: `border-2 border-dashed p-12 cursor-pointer transition-all ${dragging ? "border-amber-400 bg-amber-400/5" : "border-stone-700 hover:border-stone-500 bg-stone-900/30"}`
    },
    /* @__PURE__ */ React.createElement("div", { className: "flex flex-col items-center text-center" }, /* @__PURE__ */ React.createElement(Upload, { className: "w-10 h-10 text-stone-500 mb-4", strokeWidth: 1.5 }), /* @__PURE__ */ React.createElement("div", { className: "text-stone-200 font-medium mb-1" }, "Slipp JSON-fil her, eller klikk for \xE5 velge"), /* @__PURE__ */ React.createElement("div", { className: "text-xs text-stone-500 font-mono" }, ".json \u2014 OData format eller array")),
    /* @__PURE__ */ React.createElement(
      "input",
      {
        ref: inputRef,
        type: "file",
        accept: ".json,application/json",
        className: "hidden",
        onChange: (e) => e.target.files[0] && handleFile(e.target.files[0])
      }
    )
  ), err && /* @__PURE__ */ React.createElement("div", { className: "mt-4 p-3 border border-rose-500/40 bg-rose-500/10 text-rose-200 text-sm flex gap-2" }, /* @__PURE__ */ React.createElement(AlertCircle, { className: "w-4 h-4 flex-shrink-0 mt-0.5" }), /* @__PURE__ */ React.createElement("span", null, err)), /* @__PURE__ */ React.createElement("div", { className: "mt-6 flex items-center gap-4" }, /* @__PURE__ */ React.createElement("div", { className: "flex-1 h-px bg-stone-800" }), /* @__PURE__ */ React.createElement("span", { className: "text-xs font-mono text-stone-600 tracking-wider" }, "ELLER"), /* @__PURE__ */ React.createElement("div", { className: "flex-1 h-px bg-stone-800" })), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: onDemo,
      className: "mt-6 w-full py-3 border border-stone-700 hover:border-amber-400/50 hover:bg-amber-400/5 text-stone-300 hover:text-amber-200 font-mono text-sm tracking-wider transition-colors"
    },
    "BRUK DEMO-DATA (5 batch orders)"
  ), /* @__PURE__ */ React.createElement("div", { className: "mt-10 p-4 bg-stone-900/40 border border-stone-800 text-xs text-stone-400 font-mono leading-relaxed" }, /* @__PURE__ */ React.createElement("div", { className: "text-stone-300 mb-2 font-bold" }, "FORVENTET FORMAT"), /* @__PURE__ */ React.createElement("pre", { className: "text-stone-500 overflow-x-auto" }, `[
  {
    "BatchOrderNumber": "66-000255",
    "BatchOrderName": "Roser Premium 12x50cm",
    "ItemNumber": "300009",
    "ScheduledDate": "2025-06-14T12:00:00Z",
    "ScheduledQuantity": 64,
    "BatchOrderStatus": "Completed",
    "ProductionSiteId": "BBS",
    "PlannedReceiptWarehouseId": "Inbound",
    ...
  },
  ...
]`)));
}
function NumberStepper({ label, value, onChange, min = 0, max = 999 }) {
  const dec = () => onChange(Math.max(min, value - 1));
  const inc = () => onChange(Math.min(max, value + 1));
  return /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { className: "block text-[10px] font-mono tracking-wider text-stone-400 mb-2" }, label), /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-3" }, /* @__PURE__ */ React.createElement(
    "button",
    {
      type: "button",
      onClick: dec,
      className: "w-12 h-12 sm:w-11 sm:h-11 border border-stone-700 hover:border-amber-400/50 active:bg-amber-400/10 text-stone-300 font-mono text-xl flex items-center justify-center select-none",
      "aria-label": `Senk ${label}`
    },
    "\u2212"
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "number",
      min,
      max,
      inputMode: "numeric",
      value,
      onChange: (e) => {
        const v = +e.target.value;
        if (!isNaN(v)) onChange(Math.max(min, Math.min(max, v)));
      },
      className: "flex-1 min-w-0 bg-stone-900 border border-stone-700 px-3 py-3 sm:py-2 text-stone-100 font-mono text-center text-lg focus:border-amber-400 focus:outline-none"
    }
  ), /* @__PURE__ */ React.createElement(
    "button",
    {
      type: "button",
      onClick: inc,
      className: "w-12 h-12 sm:w-11 sm:h-11 border border-stone-700 hover:border-amber-400/50 active:bg-amber-400/10 text-stone-300 font-mono text-xl flex items-center justify-center select-none",
      "aria-label": `\xD8k ${label}`
    },
    "+"
  )));
}
function RegistrationForm({ batchOrder, existing, activeSession, activeSessions, onSave, onCancel, onDelete, onStart, onStop, onCancelActive }) {
  const initialPhase = existing ? "review" : activeSession ? "running" : "setup";
  const busyLines = (activeSessions || []).map((s) => s.line);
  const firstFreeLine = [1, 2, 3, 4, 5].find((n) => !busyLines.includes(n)) || 1;
  const [phase, setPhase] = useState(initialPhase);
  const [line, setLine] = useState(existing?.line || activeSession?.line || firstFreeLine);
  const [people, setPeople] = useState(existing?.people || activeSession?.people || 1);
  const [buckets, setBuckets] = useState(existing?.buckets ?? activeSession?.buckets ?? 0);
  const [start, setStart] = useState(existing?.start || activeSession?.start || "");
  const [end, setEnd] = useState(existing?.end || "");
  const [notes, setNotes] = useState(existing?.notes || "");
  const [startMs, setStartMs] = useState(() => {
    if (activeSession?.startMs) return activeSession.startMs;
    if (activeSession?.start) return new Date(activeSession.start).getTime();
    if (existing?.start) return new Date(existing.start).getTime();
    return null;
  });
  const [now, setNow] = useState(Date.now());
  useEffect(() => {
    if (phase !== "running") return;
    const t = setInterval(() => setNow(Date.now()), 1e3);
    return () => clearInterval(t);
  }, [phase]);
  useEffect(() => {
    if (phase !== "running" || !startMs) return;
    onStart({
      batchOrderNumber: batchOrder.BatchOrderNumber,
      line: Number(line),
      people: Number(people),
      buckets: Number(buckets),
      start,
      startMs
    });
  }, [buckets, people]);
  const dur = durationMinutes(start, end);
  const mh = personHours(start, end, people);
  const liveElapsedSec = phase === "running" && startMs ? Math.max(0, Math.floor((now - startMs) / 1e3)) : 0;
  const liveHours = Math.floor(liveElapsedSec / 3600);
  const liveMinutes = Math.floor(liveElapsedSec % 3600 / 60);
  const liveSeconds = liveElapsedSec % 60;
  const livePersonHours = (liveElapsedSec / 3600 * people).toFixed(2);
  const handleStart = () => {
    if (busyLines.includes(Number(line))) {
      alert(`Linje L${line} har allerede en p\xE5g\xE5ende registrering. Velg en annen linje.`);
      return;
    }
    const nowMs = Date.now();
    const startTime = isoLocalFromMs(nowMs);
    setStartMs(nowMs);
    setStart(startTime);
    onStart({
      batchOrderNumber: batchOrder.BatchOrderNumber,
      line: Number(line),
      people: Number(people),
      buckets: Number(buckets),
      start: startTime,
      startMs: nowMs
    });
    setPhase("running");
  };
  const handleStop = () => {
    const endTime = isoLocalNow();
    setEnd(endTime);
    setPhase("review");
  };
  const handleSubmit = () => {
    if (!start) return alert("Starttidspunkt mangler");
    if (end && new Date(end) <= new Date(start)) return alert("Slutt m\xE5 v\xE6re etter start");
    onSave({
      id: existing?.id || `reg_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
      batchOrderNumber: batchOrder.BatchOrderNumber,
      line: Number(line),
      people: Number(people),
      buckets: Number(buckets),
      start,
      end: end || null,
      notes,
      createdAt: existing?.createdAt || (/* @__PURE__ */ new Date()).toISOString(),
      updatedAt: (/* @__PURE__ */ new Date()).toISOString()
    });
  };
  const handleCloseRequest = () => {
    if (phase === "running") {
      if (!confirm("Lukke uten \xE5 avslutte? Registreringen forblir aktiv og kan gjen\xE5pnes senere.")) return;
    }
    onCancel();
  };
  const handleCancelActive = () => {
    if (!confirm("Forkaste p\xE5g\xE5ende registrering? Tidsbruk vil ikke bli lagret.")) return;
    onCancelActive(batchOrder.BatchOrderNumber, Number(line));
  };
  const headerLabel = phase === "running" ? "P\xC5G\xC5R" : phase === "review" && !existing ? "FULLF\xD8R REGISTRERING" : existing ? "REDIGER REGISTRERING" : "NY REGISTRERING";
  return /* @__PURE__ */ React.createElement("div", { className: "fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-start justify-center p-4 overflow-y-auto" }, /* @__PURE__ */ React.createElement("div", { className: "bg-stone-950 border border-stone-700 max-w-2xl w-full mt-12 mb-12" }, /* @__PURE__ */ React.createElement("div", { className: "px-6 py-4 border-b border-stone-800 flex items-center justify-between" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: `text-[10px] font-mono tracking-[0.3em] ${phase === "running" ? "text-emerald-400 animate-pulse" : "text-amber-400"}` }, phase === "running" && "\u25CF ", headerLabel), /* @__PURE__ */ React.createElement("div", { className: "font-mono text-stone-100 text-lg mt-1" }, batchOrder.BatchOrderNumber)), /* @__PURE__ */ React.createElement("button", { onClick: handleCloseRequest, className: "text-stone-500 hover:text-stone-200 p-1" }, /* @__PURE__ */ React.createElement(X, { className: "w-5 h-5" }))), /* @__PURE__ */ React.createElement("div", { className: "px-6 py-5 grid grid-cols-2 gap-4 border-b border-stone-800 bg-stone-900/30 text-xs" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 font-mono mb-1" }, "VARE"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-200 font-mono" }, batchOrder.ItemNumber || "\u2014"), batchOrder.BatchOrderName && /* @__PURE__ */ React.createElement("div", { className: "text-stone-400 mt-0.5 uppercase tracking-wide", style: { fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif" } }, batchOrder.BatchOrderName)), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 font-mono mb-1" }, "PLANLAGT"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-200 font-mono" }, fmtDate(batchOrder.ScheduledDate))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 font-mono mb-1" }, "MENGDE"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-200 font-mono" }, batchOrder.ScheduledQuantity || "\u2014")), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 font-mono mb-1" }, "SITE / LAGER"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-200 font-mono" }, batchOrder.Site, " / ", batchOrder.Warehouse)), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 font-mono mb-1" }, "STATUS"), /* @__PURE__ */ React.createElement(StatusBadge, { status: batchOrder.Status }))), phase === "setup" && /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("div", { className: "px-6 py-6 space-y-5" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { className: "block text-[10px] font-mono tracking-wider text-stone-400 mb-2" }, "PRODUKSJONSLINJE"), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-5 gap-2" }, [1, 2, 3, 4, 5].map((n) => {
    const isBusy = busyLines.includes(n);
    const isSelected = line === n;
    return /* @__PURE__ */ React.createElement(
      "button",
      {
        key: n,
        type: "button",
        onClick: () => !isBusy && setLine(n),
        disabled: isBusy,
        title: isBusy ? `L${n} har allerede en p\xE5g\xE5ende registrering` : void 0,
        className: `py-3 sm:py-3 min-h-[48px] font-mono font-bold border transition-all relative ${isBusy ? "bg-stone-900 text-stone-600 border-stone-800 cursor-not-allowed" : isSelected ? "bg-amber-400 text-stone-950 border-amber-400" : "bg-stone-900 text-stone-400 border-stone-700 hover:border-stone-500 active:bg-stone-800"}`
      },
      "L",
      n,
      isBusy && /* @__PURE__ */ React.createElement("span", { className: "absolute top-1 right-1 w-1.5 h-1.5 bg-emerald-400 rounded-full animate-pulse" })
    );
  })), busyLines.length > 0 && /* @__PURE__ */ React.createElement("div", { className: "mt-2 text-[10px] font-mono text-stone-500" }, "Linje", busyLines.length > 1 ? "r" : "", " ", busyLines.map((l) => `L${l}`).join(", "), " p\xE5g\xE5r \u2014 kan ikke startes igjen")), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-2 gap-4" }, /* @__PURE__ */ React.createElement(NumberStepper, { label: "ANTALL PERSONER", value: people, onChange: setPeople, min: 1 }), /* @__PURE__ */ React.createElement(NumberStepper, { label: "ANTALL B\xD8TTER", value: buckets, onChange: setBuckets, min: 0 }))), /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 py-4 border-t border-stone-800 flex items-center justify-between gap-3" }, /* @__PURE__ */ React.createElement("button", { onClick: onCancel, className: "px-4 sm:px-5 py-3 sm:py-2 border border-stone-700 hover:border-stone-500 text-stone-300 text-sm font-mono" }, "AVBRYT"), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: handleStart,
      className: "inline-flex items-center gap-2 px-6 py-3 bg-emerald-500 hover:bg-emerald-400 text-stone-950 font-mono font-bold text-sm tracking-wider"
    },
    /* @__PURE__ */ React.createElement(Play, { className: "w-4 h-4 fill-current" }),
    "START PRODUKSJON"
  ))), phase === "running" && /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 py-6 sm:py-8 flex flex-col items-center border-b border-stone-800 bg-stone-900/20" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono tracking-[0.3em] text-emerald-400 mb-3" }, "PRODUKSJON P\xC5G\xC5R \xB7 L", line), /* @__PURE__ */ React.createElement("div", { className: "font-mono text-stone-100 text-5xl sm:text-6xl tabular-nums", style: { fontVariantNumeric: "tabular-nums" } }, String(liveHours).padStart(2, "0"), ":", String(liveMinutes).padStart(2, "0"), ":", String(liveSeconds).padStart(2, "0")), /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 font-mono text-xs mt-3" }, "Startet ", startMs ? fmtDateTime(new Date(startMs).toISOString()) : "\u2014"), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-3 gap-2 sm:gap-3 mt-6 w-full max-w-md text-center" }, /* @__PURE__ */ React.createElement("div", { className: "bg-stone-900/60 border border-stone-800 px-2 sm:px-3 py-2" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "PERSONER"), /* @__PURE__ */ React.createElement("div", { className: "text-amber-300 font-mono text-lg mt-1" }, people)), /* @__PURE__ */ React.createElement("div", { className: "bg-stone-900/60 border border-stone-800 px-2 sm:px-3 py-2" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "B\xD8TTER"), /* @__PURE__ */ React.createElement("div", { className: "text-amber-300 font-mono text-lg mt-1" }, buckets)), /* @__PURE__ */ React.createElement("div", { className: "bg-stone-900/60 border border-stone-800 px-2 sm:px-3 py-2" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "PERSONTIMER"), /* @__PURE__ */ React.createElement("div", { className: "text-amber-300 font-mono text-lg mt-1" }, livePersonHours, "h"))), /* @__PURE__ */ React.createElement("div", { className: "w-full max-w-md mt-5" }, /* @__PURE__ */ React.createElement(NumberStepper, { label: "OPPDATER B\xD8TTER", value: buckets, onChange: setBuckets, min: 0 }))), /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 py-4 border-t border-stone-800 flex items-center justify-between gap-3 flex-wrap" }, /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: handleCancelActive,
      className: "px-3 py-2 text-rose-400 hover:text-rose-200 hover:bg-rose-500/10 text-sm font-mono inline-flex items-center gap-2"
    },
    /* @__PURE__ */ React.createElement(Trash2, { className: "w-4 h-4" }),
    " FORKAST"
  ), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: handleStop,
      className: "inline-flex items-center gap-2 px-5 sm:px-6 py-3 bg-rose-500 hover:bg-rose-400 text-stone-950 font-mono font-bold text-sm tracking-wider"
    },
    /* @__PURE__ */ React.createElement(Square, { className: "w-4 h-4 fill-current" }),
    "AVSLUTT"
  ))), phase === "review" && /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("div", { className: "px-6 py-6 space-y-5" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { className: "block text-[10px] font-mono tracking-wider text-stone-400 mb-2" }, "PRODUKSJONSLINJE"), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-5 gap-2" }, [1, 2, 3, 4, 5].map((n) => /* @__PURE__ */ React.createElement(
    "button",
    {
      key: n,
      type: "button",
      onClick: () => setLine(n),
      className: `py-3 min-h-[48px] font-mono font-bold border transition-all ${line === n ? "bg-amber-400 text-stone-950 border-amber-400" : "bg-stone-900 text-stone-400 border-stone-700 hover:border-stone-500 active:bg-stone-800"}`
    },
    "L",
    n
  )))), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-2 gap-4" }, /* @__PURE__ */ React.createElement(NumberStepper, { label: "ANTALL PERSONER", value: people, onChange: setPeople, min: 1 }), /* @__PURE__ */ React.createElement(NumberStepper, { label: "ANTALL B\xD8TTER", value: buckets, onChange: setBuckets, min: 0 })), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-1 sm:grid-cols-2 gap-4" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { className: "block text-[10px] font-mono tracking-wider text-stone-400 mb-2" }, "START"), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "datetime-local",
      value: start,
      onChange: (e) => setStart(e.target.value),
      className: "w-full bg-stone-900 border border-stone-700 px-3 py-3 sm:py-2 text-stone-100 font-mono text-sm focus:border-amber-400 focus:outline-none"
    }
  )), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { className: "block text-[10px] font-mono tracking-wider text-stone-400 mb-2" }, "SLUTT", existing ? " (valgfri)" : ""), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "datetime-local",
      value: end,
      onChange: (e) => setEnd(e.target.value),
      className: "w-full bg-stone-900 border border-stone-700 px-3 py-3 sm:py-2 text-stone-100 font-mono text-sm focus:border-amber-400 focus:outline-none"
    }
  ))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { className: "block text-[10px] font-mono tracking-wider text-stone-400 mb-2" }, "NOTAT (valgfritt)"), /* @__PURE__ */ React.createElement(
    "textarea",
    {
      value: notes,
      onChange: (e) => setNotes(e.target.value),
      rows: 2,
      className: "w-full bg-stone-900 border border-stone-700 px-3 py-2 text-stone-200 text-sm focus:border-amber-400 focus:outline-none resize-none",
      placeholder: "Avvik, kommentarer ..."
    }
  )), dur != null && /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-2 sm:grid-cols-4 gap-3 pt-2" }, /* @__PURE__ */ React.createElement("div", { className: "bg-stone-900/60 border border-stone-800 px-4 py-3" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "VARIGHET"), /* @__PURE__ */ React.createElement("div", { className: "text-amber-300 font-mono text-lg mt-1" }, fmtMinutes(dur))), /* @__PURE__ */ React.createElement("div", { className: "bg-stone-900/60 border border-stone-800 px-4 py-3" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "PERSONTIMER"), /* @__PURE__ */ React.createElement("div", { className: "text-amber-300 font-mono text-lg mt-1" }, mh, "h")), /* @__PURE__ */ React.createElement("div", { className: "bg-stone-900/60 border border-stone-800 px-4 py-3" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "B\xD8TTER"), /* @__PURE__ */ React.createElement("div", { className: "text-amber-300 font-mono text-lg mt-1" }, buckets)), /* @__PURE__ */ React.createElement("div", { className: "bg-stone-900/60 border border-stone-800 px-4 py-3" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "B\xD8TTER/TIME"), /* @__PURE__ */ React.createElement("div", { className: "text-amber-300 font-mono text-lg mt-1" }, mh > 0 ? (buckets / mh).toFixed(1) : "\u2014")))), /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 py-4 border-t border-stone-800 flex items-center justify-between gap-3 flex-wrap" }, /* @__PURE__ */ React.createElement("div", null, existing && /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => {
        if (confirm("Slette registreringen permanent?")) onDelete(existing.id);
      },
      className: "px-3 py-2 text-rose-400 hover:text-rose-200 hover:bg-rose-500/10 text-sm font-mono inline-flex items-center gap-2"
    },
    /* @__PURE__ */ React.createElement(Trash2, { className: "w-4 h-4" }),
    " SLETT"
  )), /* @__PURE__ */ React.createElement("div", { className: "flex gap-2" }, /* @__PURE__ */ React.createElement("button", { onClick: onCancel, className: "px-5 py-2 border border-stone-700 hover:border-stone-500 text-stone-300 text-sm font-mono" }, "AVBRYT"), /* @__PURE__ */ React.createElement("button", { onClick: handleSubmit, className: "px-5 py-2 bg-amber-400 hover:bg-amber-300 text-stone-950 font-mono font-bold text-sm" }, existing ? "OPPDATER" : "LAGRE"))))));
}
function BatchOrderDetailView({ batchOrder, registrations, activeSessions, onEdit, onNew, onClose, onResumeActive }) {
  const regs = useMemo(() => {
    return registrations.filter((r) => r.batchOrderNumber === batchOrder.BatchOrderNumber).map((r) => ({
      ...r,
      durationMin: durationMinutes(r.start, r.end),
      personHours: personHours(r.start, r.end, r.people)
    })).sort((a, b) => (b.start || "").localeCompare(a.start || ""));
  }, [registrations, batchOrder]);
  const stats = useMemo(() => {
    const closed = regs.filter((r) => r.end);
    const totalMin = closed.reduce((s, r) => s + (r.durationMin || 0), 0);
    const totalMh = closed.reduce((s, r) => s + (r.personHours || 0), 0);
    const totalBuckets = regs.reduce((s, r) => s + (Number(r.buckets) || 0), 0);
    const byLine = {};
    for (const r of closed) {
      if (!byLine[r.line]) byLine[r.line] = { hours: 0, buckets: 0 };
      byLine[r.line].hours += r.personHours || 0;
      byLine[r.line].buckets += Number(r.buckets) || 0;
    }
    return {
      total: regs.length,
      closed: closed.length,
      open: regs.length - closed.length,
      totalMin,
      totalMh: totalMh.toFixed(1),
      totalBuckets,
      bucketsPerHour: totalMh > 0 ? (totalBuckets / totalMh).toFixed(1) : null,
      byLine
    };
  }, [regs]);
  const hasActive = activeSessions && activeSessions.length > 0;
  return /* @__PURE__ */ React.createElement("div", { className: "fixed inset-0 z-40 bg-stone-950 overflow-y-auto" }, /* @__PURE__ */ React.createElement("div", { className: "min-h-full flex flex-col" }, /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 md:px-8 py-4 sm:py-5 border-b border-stone-800 flex items-start justify-between gap-3 sticky top-0 bg-stone-950 z-10" }, /* @__PURE__ */ React.createElement("div", { className: "flex-1 min-w-0" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono tracking-[0.3em] text-amber-400" }, "BATCH ORDER"), /* @__PURE__ */ React.createElement("div", { className: "flex items-baseline gap-3 mt-1 flex-wrap" }, /* @__PURE__ */ React.createElement("h2", { className: "text-xl sm:text-2xl text-stone-100 font-mono" }, batchOrder.BatchOrderNumber), /* @__PURE__ */ React.createElement(StatusBadge, { status: batchOrder.Status })), batchOrder.BatchOrderName && /* @__PURE__ */ React.createElement("div", { className: "text-stone-300 mt-1 uppercase tracking-wide text-sm" }, batchOrder.BatchOrderName)), /* @__PURE__ */ React.createElement("button", { onClick: onClose, className: "text-stone-400 hover:text-stone-100 p-2 flex-shrink-0" }, /* @__PURE__ */ React.createElement(X, { className: "w-6 h-6" }))), /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 md:px-8 py-4 sm:py-5 border-b border-stone-800 grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-3 sm:gap-4 text-xs bg-stone-900/30" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 font-mono mb-1" }, "VARE"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-200 font-mono" }, batchOrder.ItemNumber || "\u2014")), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 font-mono mb-1" }, "PLANLAGT"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-200 font-mono" }, fmtDate(batchOrder.ScheduledDate))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 font-mono mb-1" }, "MENGDE"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-200 font-mono" }, batchOrder.ScheduledQuantity || "\u2014")), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 font-mono mb-1" }, "SITE"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-200 font-mono" }, batchOrder.Site || "\u2014")), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 font-mono mb-1" }, "LAGER"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-200 font-mono" }, batchOrder.Warehouse || "\u2014"))), /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 md:px-8 py-4 sm:py-5 border-b border-stone-800 grid grid-cols-2 md:grid-cols-5 gap-2 sm:gap-3" }, /* @__PURE__ */ React.createElement("div", { className: "border border-stone-800 px-3 sm:px-4 py-3 bg-stone-900/30" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "TOTALT"), /* @__PURE__ */ React.createElement("div", { className: "text-xl sm:text-2xl font-mono text-stone-100 mt-1" }, stats.total)), /* @__PURE__ */ React.createElement("div", { className: "border border-stone-800 px-3 sm:px-4 py-3 bg-stone-900/30" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "P\xC5G\xC5R"), /* @__PURE__ */ React.createElement("div", { className: "text-xl sm:text-2xl font-mono text-emerald-300 mt-1" }, activeSessions?.length || 0)), /* @__PURE__ */ React.createElement("div", { className: "border border-stone-800 px-3 sm:px-4 py-3 bg-stone-900/30" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "VARIGHET"), /* @__PURE__ */ React.createElement("div", { className: "text-xl sm:text-2xl font-mono text-stone-100 mt-1" }, fmtMinutes(stats.totalMin))), /* @__PURE__ */ React.createElement("div", { className: "border border-stone-800 px-3 sm:px-4 py-3 bg-stone-900/30" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "PERSONTIMER"), /* @__PURE__ */ React.createElement("div", { className: "text-xl sm:text-2xl font-mono text-amber-300 mt-1" }, stats.totalMh, "h")), /* @__PURE__ */ React.createElement("div", { className: "border border-stone-800 px-3 sm:px-4 py-3 bg-stone-900/30" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "B\xD8TTER"), /* @__PURE__ */ React.createElement("div", { className: "text-xl sm:text-2xl font-mono text-amber-300 mt-1" }, stats.totalBuckets), stats.bucketsPerHour && /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 mt-0.5" }, stats.bucketsPerHour, "/time"))), Object.keys(stats.byLine).length > 0 && /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 md:px-8 py-4 border-b border-stone-800 flex items-center gap-3 flex-wrap" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono tracking-wider text-stone-500" }, "FORDELING"), [1, 2, 3, 4, 5].map((n) => stats.byLine[n] && /* @__PURE__ */ React.createElement("div", { key: n, className: "flex items-center gap-2 text-xs font-mono" }, /* @__PURE__ */ React.createElement(LineChip, { line: n }), /* @__PURE__ */ React.createElement("span", { className: "text-stone-300" }, stats.byLine[n].hours.toFixed(1), "h"), /* @__PURE__ */ React.createElement("span", { className: "text-stone-500" }, "\xB7"), /* @__PURE__ */ React.createElement("span", { className: "text-stone-400" }, stats.byLine[n].buckets, " b\xF8tter")))), hasActive && /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 md:px-8 py-4 border-b border-stone-800 bg-emerald-500/5" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono tracking-wider text-emerald-400 mb-3" }, "P\xC5G\xC5ENDE REGISTRERING", activeSessions.length > 1 ? "ER" : "", " (", activeSessions.length, ")"), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3" }, activeSessions.sort((a, b) => a.line - b.line).map((s) => /* @__PURE__ */ React.createElement("div", { key: s.line, className: "bg-stone-900/40 border border-emerald-500/30 p-3 flex items-center gap-3" }, /* @__PURE__ */ React.createElement(LineChip, { line: s.line }), /* @__PURE__ */ React.createElement("div", { className: "flex-1 min-w-0" }, /* @__PURE__ */ React.createElement("div", { className: "font-mono text-stone-200 text-sm" }, s.people, " pers \xB7 ", s.buckets ?? 0, " b\xF8tter"), /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 mt-0.5" }, "Startet ", fmtDateTime(s.start))), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => onResumeActive(s.line),
      className: "inline-flex items-center gap-1.5 px-3 py-2 bg-emerald-500 hover:bg-emerald-400 text-stone-950 font-mono font-bold text-xs tracking-wider flex-shrink-0"
    },
    /* @__PURE__ */ React.createElement(Square, { className: "w-3 h-3 fill-current" }),
    "\xC5PNE"
  ))))), /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 md:px-8 py-4 border-b border-stone-800 flex items-center justify-between" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono tracking-wider text-stone-500" }, "ALLE REGISTRERINGER"), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: onNew,
      className: "inline-flex items-center gap-2 px-4 py-2 bg-amber-400 hover:bg-amber-300 text-stone-950 font-mono font-bold text-xs tracking-wider"
    },
    /* @__PURE__ */ React.createElement(Plus, { className: "w-3.5 h-3.5" }),
    "NY REGISTRERING"
  )), /* @__PURE__ */ React.createElement("div", { className: "flex-1" }, regs.length === 0 ? /* @__PURE__ */ React.createElement("div", { className: "px-8 py-16 text-center text-stone-500" }, /* @__PURE__ */ React.createElement(Activity, { className: "w-12 h-12 mx-auto mb-3", strokeWidth: 1 }), /* @__PURE__ */ React.createElement("div", { className: "font-mono text-sm" }, "Ingen registreringer enn\xE5"), /* @__PURE__ */ React.createElement("div", { className: "text-xs mt-1" }, "Trykk NY REGISTRERING for \xE5 starte")) : /* @__PURE__ */ React.createElement("div", { className: "divide-y divide-stone-900" }, regs.map((r) => /* @__PURE__ */ React.createElement(
    "div",
    {
      key: r.id,
      className: "px-4 sm:px-6 md:px-8 py-4 hover:bg-stone-900/40 group flex items-start sm:items-center gap-3 sm:gap-4 flex-col sm:flex-row"
    },
    /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-3 w-full sm:w-auto" }, /* @__PURE__ */ React.createElement(LineChip, { line: r.line }), /* @__PURE__ */ React.createElement("div", { className: "flex-1 sm:hidden" }, /* @__PURE__ */ React.createElement("div", { className: "text-stone-200 font-mono text-sm" }, fmtDateTime(r.start)), /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 font-mono text-xs mt-0.5" }, r.end ? `\u2192 ${fmtDateTime(r.end)}` : /* @__PURE__ */ React.createElement("span", { className: "text-amber-400" }, "p\xE5g\xE5r"))), /* @__PURE__ */ React.createElement(
      "button",
      {
        onClick: () => onEdit(r),
        className: "sm:hidden inline-flex items-center gap-1.5 px-3 py-2 border border-stone-700 hover:border-amber-400/50 text-stone-300 hover:text-amber-200 text-xs font-mono tracking-wider flex-shrink-0"
      },
      /* @__PURE__ */ React.createElement(Edit2, { className: "w-3.5 h-3.5" }),
      "REDIGER"
    )),
    /* @__PURE__ */ React.createElement("div", { className: "flex-1 min-w-0 grid grid-cols-2 sm:grid-cols-5 gap-3 items-center w-full" }, /* @__PURE__ */ React.createElement("div", { className: "hidden sm:block" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "START"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-200 font-mono text-sm mt-0.5" }, fmtDateTime(r.start))), /* @__PURE__ */ React.createElement("div", { className: "hidden sm:block" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "SLUTT"), /* @__PURE__ */ React.createElement("div", { className: "font-mono text-sm mt-0.5" }, r.end ? /* @__PURE__ */ React.createElement("span", { className: "text-stone-200" }, fmtDateTime(r.end)) : /* @__PURE__ */ React.createElement("span", { className: "text-amber-400" }, "\u2014 p\xE5g\xE5r"))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "VARIGHET"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-200 font-mono text-sm mt-0.5" }, fmtMinutes(r.durationMin))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "PERS \xB7 B\xD8TTER"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-200 font-mono text-sm mt-0.5 inline-flex items-center gap-2" }, /* @__PURE__ */ React.createElement("span", { className: "inline-flex items-center gap-1" }, /* @__PURE__ */ React.createElement(Users, { className: "w-3 h-3 text-stone-500" }), r.people), /* @__PURE__ */ React.createElement("span", { className: "text-stone-600" }, "\xB7"), /* @__PURE__ */ React.createElement("span", null, r.buckets ?? 0))), /* @__PURE__ */ React.createElement("div", { className: "col-span-2 sm:col-span-1" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "PERSONTIMER"), /* @__PURE__ */ React.createElement("div", { className: "text-amber-300 font-mono text-sm mt-0.5" }, r.personHours != null ? `${r.personHours}h` : "\u2014"))),
    r.notes && /* @__PURE__ */ React.createElement("div", { className: "w-full sm:w-48 text-xs text-stone-400 italic", title: r.notes }, /* @__PURE__ */ React.createElement("div", { className: "truncate" }, r.notes)),
    /* @__PURE__ */ React.createElement(
      "button",
      {
        onClick: () => onEdit(r),
        className: "hidden sm:inline-flex items-center gap-1.5 px-3 py-2 border border-stone-700 hover:border-amber-400/50 hover:bg-amber-400/5 text-stone-300 hover:text-amber-200 text-xs font-mono tracking-wider flex-shrink-0 opacity-60 group-hover:opacity-100 transition-all"
      },
      /* @__PURE__ */ React.createElement(Edit2, { className: "w-3.5 h-3.5" }),
      "REDIGER"
    )
  ))))));
}
function RegistrationsPanel({ registrations, onEdit, onClose, batchOrders }) {
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [selectedLines, setSelectedLines] = useState(/* @__PURE__ */ new Set([1, 2, 3, 4, 5]));
  const enriched = useMemo(() => {
    const boMap = new Map(batchOrders.map((b) => [b.BatchOrderNumber, b]));
    return registrations.map((r) => ({
      ...r,
      bo: boMap.get(r.batchOrderNumber),
      durationMin: durationMinutes(r.start, r.end),
      personHours: personHours(r.start, r.end, r.people)
    }));
  }, [registrations, batchOrders]);
  const filtered = useMemo(() => {
    const fromTs = dateFrom ? (/* @__PURE__ */ new Date(dateFrom + "T00:00:00")).getTime() : null;
    const toTs = dateTo ? (/* @__PURE__ */ new Date(dateTo + "T23:59:59")).getTime() : null;
    return enriched.filter((r) => {
      if (!selectedLines.has(r.line)) return false;
      if (fromTs != null || toTs != null) {
        if (!r.start) return false;
        const ts = new Date(r.start).getTime();
        if (fromTs != null && ts < fromTs) return false;
        if (toTs != null && ts > toTs) return false;
      }
      return true;
    });
  }, [enriched, dateFrom, dateTo, selectedLines]);
  const hasActiveFilters = dateFrom || dateTo || selectedLines.size < 5;
  const clearFilters = () => {
    setDateFrom("");
    setDateTo("");
    setSelectedLines(/* @__PURE__ */ new Set([1, 2, 3, 4, 5]));
  };
  const toggleLine = (n) => {
    setSelectedLines((prev) => {
      const next = new Set(prev);
      if (next.has(n)) next.delete(n);
      else next.add(n);
      if (next.size === 0) return /* @__PURE__ */ new Set([1, 2, 3, 4, 5]);
      return next;
    });
  };
  const stats = useMemo(() => {
    const closed = filtered.filter((r) => r.end);
    const open = filtered.filter((r) => !r.end);
    const totalMh = closed.reduce((s, r) => s + (r.personHours || 0), 0);
    const totalBuckets = filtered.reduce((s, r) => s + (Number(r.buckets) || 0), 0);
    const byLine = {};
    for (const r of closed) {
      if (!byLine[r.line]) byLine[r.line] = { hours: 0, buckets: 0 };
      byLine[r.line].hours += r.personHours || 0;
      byLine[r.line].buckets += Number(r.buckets) || 0;
    }
    return { closed: closed.length, open: open.length, totalMh: totalMh.toFixed(1), totalBuckets, byLine };
  }, [filtered]);
  const exportHeaders = [
    { key: "id", label: "RegistrationId" },
    { key: "batchOrderNumber", label: "BatchOrderNumber" },
    { label: "BatchOrderName", get: (r) => r.bo?.BatchOrderName || "" },
    { key: "line", label: "Line" },
    { key: "people", label: "People" },
    { label: "Buckets", get: (r) => r.buckets ?? 0 },
    { key: "start", label: "StartTime" },
    { key: "end", label: "EndTime" },
    { label: "DurationMinutes", get: (r) => r.durationMin ?? "" },
    { label: "PersonHours", get: (r) => r.personHours ?? "" },
    { label: "BucketsPerPersonHour", get: (r) => {
      const b = Number(r.buckets) || 0;
      return r.personHours && r.personHours > 0 ? +(b / r.personHours).toFixed(2) : "";
    } },
    { label: "ItemNumber", get: (r) => r.bo?.ItemNumber || "" },
    { label: "Site", get: (r) => r.bo?.Site || "" },
    { label: "Warehouse", get: (r) => r.bo?.Warehouse || "" },
    { label: "ScheduledDate", get: (r) => r.bo?.ScheduledDate || "" },
    { key: "notes", label: "Notes" },
    { key: "createdAt", label: "CreatedAt" },
    { key: "updatedAt", label: "UpdatedAt" }
  ];
  const handleExportCsv = () => {
    download(
      `bo_work_info_${(/* @__PURE__ */ new Date()).toISOString().slice(0, 10)}.csv`,
      toCSV(filtered, exportHeaders)
    );
  };
  const handleExportXlsx = () => {
    const rows = filtered.map((r) => {
      const obj = {};
      for (const h of exportHeaders) {
        obj[h.label] = typeof h.get === "function" ? h.get(r) : r[h.key];
      }
      return obj;
    });
    const wsData = XLSX.utils.json_to_sheet(rows, { header: exportHeaders.map((h) => h.label) });
    wsData["!cols"] = [
      { wch: 22 },
      { wch: 16 },
      { wch: 28 },
      { wch: 6 },
      { wch: 6 },
      { wch: 8 },
      { wch: 18 },
      { wch: 18 },
      { wch: 8 },
      { wch: 10 },
      { wch: 12 },
      { wch: 14 },
      { wch: 8 },
      { wch: 12 },
      { wch: 20 },
      { wch: 30 },
      { wch: 22 },
      { wch: 22 }
    ];
    const byLine = {};
    for (const r of filtered.filter((x) => x.end)) {
      const k = r.line;
      if (!byLine[k]) byLine[k] = { Line: k, Registrations: 0, TotalMinutes: 0, PersonHours: 0, Buckets: 0, UniqueBatchOrders: /* @__PURE__ */ new Set() };
      byLine[k].Registrations++;
      byLine[k].TotalMinutes += r.durationMin || 0;
      byLine[k].PersonHours += r.personHours || 0;
      byLine[k].Buckets += Number(r.buckets) || 0;
      byLine[k].UniqueBatchOrders.add(r.batchOrderNumber);
    }
    const summary = Object.values(byLine).sort((a, b) => a.Line - b.Line).map((s) => ({
      Line: s.Line,
      Registrations: s.Registrations,
      UniqueBatchOrders: s.UniqueBatchOrders.size,
      TotalMinutes: s.TotalMinutes,
      PersonHours: +s.PersonHours.toFixed(2),
      Buckets: s.Buckets,
      BucketsPerPersonHour: s.PersonHours > 0 ? +(s.Buckets / s.PersonHours).toFixed(2) : 0,
      AvgPersonHoursPerRegistration: +(s.PersonHours / s.Registrations).toFixed(2)
    }));
    const wsSummary = XLSX.utils.json_to_sheet(summary);
    wsSummary["!cols"] = [{ wch: 6 }, { wch: 14 }, { wch: 18 }, { wch: 14 }, { wch: 12 }, { wch: 10 }, { wch: 22 }, { wch: 26 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsData, "Registreringer");
    XLSX.utils.book_append_sheet(wb, wsSummary, "Sammendrag pr linje");
    XLSX.writeFile(wb, `bo_work_info_${(/* @__PURE__ */ new Date()).toISOString().slice(0, 10)}.xlsx`);
  };
  return /* @__PURE__ */ React.createElement("div", { className: "fixed inset-0 z-40 bg-stone-950 overflow-y-auto" }, /* @__PURE__ */ React.createElement("div", { className: "min-h-full flex flex-col" }, /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 md:px-8 py-4 sm:py-5 border-b border-stone-800 flex items-center justify-between sticky top-0 bg-stone-950 z-10" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono tracking-[0.3em] text-amber-400" }, "BO_WORK_INFO"), /* @__PURE__ */ React.createElement("h2", { className: "text-xl sm:text-2xl text-stone-100 font-mono mt-1" }, "Registreringer")), /* @__PURE__ */ React.createElement("button", { onClick: onClose, className: "text-stone-400 hover:text-stone-100 p-2" }, /* @__PURE__ */ React.createElement(X, { className: "w-6 h-6" }))), /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 md:px-8 py-3 sm:py-4 border-b border-stone-800 flex flex-wrap items-center gap-2 sm:gap-3" }, /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-2 flex-wrap" }, /* @__PURE__ */ React.createElement("span", { className: "text-[10px] font-mono tracking-wider text-stone-500" }, "LINJER"), /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-1" }, [1, 2, 3, 4, 5].map((n) => {
    const active = selectedLines.has(n);
    return /* @__PURE__ */ React.createElement(
      "button",
      {
        key: n,
        onClick: () => toggleLine(n),
        className: `w-9 h-9 sm:w-8 sm:h-8 font-mono font-bold text-xs border transition-all ${active ? "bg-amber-400 text-stone-950 border-amber-400" : "bg-stone-900 text-stone-500 border-stone-700 hover:border-stone-500"}`
      },
      "L",
      n
    );
  }))), /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-2 flex-wrap" }, /* @__PURE__ */ React.createElement("span", { className: "text-[10px] font-mono tracking-wider text-stone-500" }, "PERIODE"), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "date",
      value: dateFrom,
      onChange: (e) => setDateFrom(e.target.value),
      className: "bg-stone-900 border border-stone-800 px-2 py-2 sm:py-1.5 text-xs text-stone-200 font-mono focus:border-amber-400 focus:outline-none"
    }
  ), /* @__PURE__ */ React.createElement("span", { className: "text-stone-600 text-xs font-mono" }, "\u2192"), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "date",
      value: dateTo,
      onChange: (e) => setDateTo(e.target.value),
      className: "bg-stone-900 border border-stone-800 px-2 py-2 sm:py-1.5 text-xs text-stone-200 font-mono focus:border-amber-400 focus:outline-none"
    }
  )), /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-1 flex-wrap" }, (() => {
    const toISO = (d) => {
      const tz = d.getTimezoneOffset() * 6e4;
      return new Date(d - tz).toISOString().slice(0, 10);
    };
    const today = /* @__PURE__ */ new Date();
    today.setHours(0, 0, 0, 0);
    const dayAt = (offset) => {
      const d = new Date(today);
      d.setDate(d.getDate() + offset);
      return toISO(d);
    };
    const startOfWeek = new Date(today);
    const dow = (today.getDay() + 6) % 7;
    startOfWeek.setDate(today.getDate() - dow);
    const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    const quickRanges = [
      { label: "I DAG", from: dayAt(0), to: dayAt(0) },
      { label: "I G\xC5R", from: dayAt(-1), to: dayAt(-1) },
      { label: "DENNE UKEN", from: toISO(startOfWeek), to: dayAt(0) },
      { label: "DENNE M\xC5NED", from: toISO(startOfMonth), to: dayAt(0) }
    ];
    return quickRanges.map((r) => {
      const active = dateFrom === r.from && dateTo === r.to;
      return /* @__PURE__ */ React.createElement(
        "button",
        {
          key: r.label,
          onClick: () => {
            setDateFrom(r.from);
            setDateTo(r.to);
          },
          className: `px-2 py-2 sm:py-1.5 text-[10px] font-mono tracking-wider border ${active ? "bg-amber-400 text-stone-950 border-amber-400" : "border-stone-800 text-stone-400 hover:border-stone-600 hover:text-stone-200"}`
        },
        r.label
      );
    });
  })()), hasActiveFilters && /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: clearFilters,
      className: "inline-flex items-center gap-1.5 px-2.5 py-2 sm:py-1.5 text-[10px] font-mono tracking-wider text-rose-300 hover:text-rose-200 hover:bg-rose-500/10 border border-rose-500/30"
    },
    /* @__PURE__ */ React.createElement(X, { className: "w-3 h-3" }),
    "NULLSTILL"
  ), /* @__PURE__ */ React.createElement("div", { className: "ml-auto text-xs font-mono text-stone-500 whitespace-nowrap" }, filtered.length, " / ", registrations.length)), /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 md:px-8 py-4 sm:py-5 border-b border-stone-800 grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-2 sm:gap-3" }, /* @__PURE__ */ React.createElement("div", { className: "border border-stone-800 px-3 sm:px-4 py-3 bg-stone-900/30" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "TOTALT"), /* @__PURE__ */ React.createElement("div", { className: "text-xl sm:text-2xl font-mono text-stone-100 mt-1" }, filtered.length)), /* @__PURE__ */ React.createElement("div", { className: "border border-stone-800 px-3 sm:px-4 py-3 bg-stone-900/30" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "P\xC5G\xC5R"), /* @__PURE__ */ React.createElement("div", { className: "text-xl sm:text-2xl font-mono text-amber-300 mt-1" }, stats.open)), /* @__PURE__ */ React.createElement("div", { className: "border border-stone-800 px-3 sm:px-4 py-3 bg-stone-900/30" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "FULLF\xD8RT"), /* @__PURE__ */ React.createElement("div", { className: "text-xl sm:text-2xl font-mono text-emerald-300 mt-1" }, stats.closed)), /* @__PURE__ */ React.createElement("div", { className: "border border-stone-800 px-3 sm:px-4 py-3 bg-stone-900/30" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "PERSONTIMER"), /* @__PURE__ */ React.createElement("div", { className: "text-xl sm:text-2xl font-mono text-stone-100 mt-1" }, stats.totalMh, "h")), /* @__PURE__ */ React.createElement("div", { className: "border border-stone-800 px-3 sm:px-4 py-3 bg-stone-900/30 col-span-2 sm:col-span-1" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono text-stone-500 tracking-wider" }, "B\xD8TTER"), /* @__PURE__ */ React.createElement("div", { className: "text-xl sm:text-2xl font-mono text-amber-300 mt-1" }, stats.totalBuckets))), Object.keys(stats.byLine).length > 0 && /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 md:px-8 py-4 border-b border-stone-800 flex items-center gap-3 flex-wrap" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono tracking-wider text-stone-500" }, "FORDELING PR LINJE"), [1, 2, 3, 4, 5].map((n) => stats.byLine[n] && /* @__PURE__ */ React.createElement("div", { key: n, className: "flex items-center gap-2 text-xs font-mono" }, /* @__PURE__ */ React.createElement(LineChip, { line: n }), /* @__PURE__ */ React.createElement("span", { className: "text-stone-300" }, stats.byLine[n].hours.toFixed(1), "h"), /* @__PURE__ */ React.createElement("span", { className: "text-stone-500" }, "\xB7"), /* @__PURE__ */ React.createElement("span", { className: "text-stone-400" }, stats.byLine[n].buckets, " b\xF8tter")))), /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 md:px-8 py-3 border-b border-stone-800 flex items-center justify-end gap-2" }, /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: handleExportCsv,
      disabled: registrations.length === 0,
      className: "inline-flex items-center gap-2 px-3 py-1.5 border border-stone-700 hover:border-amber-400/50 hover:bg-amber-400/5 text-stone-300 hover:text-amber-200 text-xs font-mono tracking-wider disabled:opacity-30 disabled:hover:border-stone-700 disabled:hover:bg-transparent"
    },
    /* @__PURE__ */ React.createElement(Download, { className: "w-3.5 h-3.5" }),
    "EKSPORT CSV"
  ), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: handleExportXlsx,
      disabled: registrations.length === 0,
      className: "inline-flex items-center gap-2 px-3 py-1.5 border border-emerald-500/40 hover:border-emerald-400 hover:bg-emerald-500/10 text-emerald-300 hover:text-emerald-200 text-xs font-mono tracking-wider disabled:opacity-30 disabled:hover:border-emerald-500/40 disabled:hover:bg-transparent"
    },
    /* @__PURE__ */ React.createElement(FileSpreadsheet, { className: "w-3.5 h-3.5" }),
    "EKSPORT EXCEL"
  )), /* @__PURE__ */ React.createElement("div", { className: "flex-1 overflow-auto" }, registrations.length === 0 ? /* @__PURE__ */ React.createElement("div", { className: "flex flex-col items-center justify-center py-20 text-stone-500" }, /* @__PURE__ */ React.createElement(Activity, { className: "w-12 h-12 mb-3", strokeWidth: 1 }), /* @__PURE__ */ React.createElement("div", { className: "font-mono text-sm" }, "Ingen registreringer enn\xE5")) : filtered.length === 0 ? /* @__PURE__ */ React.createElement("div", { className: "flex flex-col items-center justify-center py-20 text-stone-500" }, /* @__PURE__ */ React.createElement(Activity, { className: "w-12 h-12 mb-3", strokeWidth: 1 }), /* @__PURE__ */ React.createElement("div", { className: "font-mono text-sm" }, "Ingen registreringer matcher filtrene"), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: clearFilters,
      className: "mt-4 text-xs font-mono tracking-wider text-amber-400 hover:text-amber-300 border border-amber-400/30 hover:border-amber-400/60 px-3 py-1.5"
    },
    "NULLSTILL FILTRE"
  )) : /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("div", { className: "md:hidden divide-y divide-stone-900" }, filtered.sort((a, b) => (b.start || "").localeCompare(a.start || "")).map((r) => /* @__PURE__ */ React.createElement("div", { key: r.id, className: "p-4 flex items-start gap-3", onClick: () => onEdit(r), role: "button" }, /* @__PURE__ */ React.createElement(LineChip, { line: r.line }), /* @__PURE__ */ React.createElement("div", { className: "flex-1 min-w-0" }, /* @__PURE__ */ React.createElement("div", { className: "flex items-center justify-between gap-2 mb-1" }, /* @__PURE__ */ React.createElement("div", { className: "font-mono text-stone-200 text-sm" }, r.batchOrderNumber), r.end ? /* @__PURE__ */ React.createElement("span", { className: "inline-flex items-center gap-1 text-emerald-400 text-[10px] font-mono" }, /* @__PURE__ */ React.createElement(CheckCircle2, { className: "w-3 h-3" }), "OK") : /* @__PURE__ */ React.createElement("span", { className: "inline-flex items-center gap-1 text-amber-400 text-[10px] font-mono" }, /* @__PURE__ */ React.createElement(Clock, { className: "w-3 h-3" }), "\xC5PEN")), /* @__PURE__ */ React.createElement("div", { className: "text-stone-400 text-xs font-mono" }, fmtDateTime(r.start)), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-4 gap-2 mt-2 text-xs font-mono" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] text-stone-500" }, "PERS"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-300" }, r.people)), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] text-stone-500" }, "B\xD8TTER"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-300" }, r.buckets ?? 0)), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] text-stone-500" }, "VARIGHET"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-300" }, fmtMinutes(r.durationMin))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] text-stone-500" }, "PT"), /* @__PURE__ */ React.createElement("div", { className: "text-amber-300" }, r.personHours != null ? `${r.personHours}h` : "\u2014")))), /* @__PURE__ */ React.createElement(Edit2, { className: "w-4 h-4 text-stone-500 flex-shrink-0 mt-1" })))), /* @__PURE__ */ React.createElement("table", { className: "hidden md:table w-full text-sm" }, /* @__PURE__ */ React.createElement("thead", { className: "bg-stone-900 sticky top-0" }, /* @__PURE__ */ React.createElement("tr", { className: "text-left text-[10px] font-mono tracking-wider text-stone-500 border-b border-stone-800" }, /* @__PURE__ */ React.createElement("th", { className: "px-8 py-3" }, "BATCH ORDER"), /* @__PURE__ */ React.createElement("th", { className: "py-3" }, "LINJE"), /* @__PURE__ */ React.createElement("th", { className: "py-3" }, "PERS"), /* @__PURE__ */ React.createElement("th", { className: "py-3" }, "B\xD8TTER"), /* @__PURE__ */ React.createElement("th", { className: "py-3" }, "START"), /* @__PURE__ */ React.createElement("th", { className: "py-3" }, "SLUTT"), /* @__PURE__ */ React.createElement("th", { className: "py-3" }, "VARIGHET"), /* @__PURE__ */ React.createElement("th", { className: "py-3" }, "PT"), /* @__PURE__ */ React.createElement("th", { className: "py-3" }, "STATUS"), /* @__PURE__ */ React.createElement("th", { className: "py-3 pr-8" }))), /* @__PURE__ */ React.createElement("tbody", null, filtered.sort((a, b) => (b.start || "").localeCompare(a.start || "")).map((r) => /* @__PURE__ */ React.createElement("tr", { key: r.id, className: "border-b border-stone-900 hover:bg-stone-900/40" }, /* @__PURE__ */ React.createElement("td", { className: "px-8 py-3 font-mono text-stone-200" }, r.batchOrderNumber), /* @__PURE__ */ React.createElement("td", { className: "py-3" }, /* @__PURE__ */ React.createElement(LineChip, { line: r.line })), /* @__PURE__ */ React.createElement("td", { className: "py-3 font-mono text-stone-300" }, /* @__PURE__ */ React.createElement("span", { className: "inline-flex items-center gap-1" }, /* @__PURE__ */ React.createElement(Users, { className: "w-3 h-3 text-stone-500" }), r.people)), /* @__PURE__ */ React.createElement("td", { className: "py-3 font-mono text-stone-300" }, r.buckets ?? 0), /* @__PURE__ */ React.createElement("td", { className: "py-3 font-mono text-stone-300 text-xs" }, fmtDateTime(r.start)), /* @__PURE__ */ React.createElement("td", { className: "py-3 font-mono text-stone-300 text-xs" }, r.end ? fmtDateTime(r.end) : /* @__PURE__ */ React.createElement("span", { className: "text-amber-400" }, "\u2014 p\xE5g\xE5r")), /* @__PURE__ */ React.createElement("td", { className: "py-3 font-mono text-stone-300" }, fmtMinutes(r.durationMin)), /* @__PURE__ */ React.createElement("td", { className: "py-3 font-mono text-amber-300" }, r.personHours != null ? `${r.personHours}h` : "\u2014"), /* @__PURE__ */ React.createElement("td", { className: "py-3" }, r.end ? /* @__PURE__ */ React.createElement("span", { className: "inline-flex items-center gap-1 text-emerald-400 text-xs font-mono" }, /* @__PURE__ */ React.createElement(CheckCircle2, { className: "w-3 h-3" }), "OK") : /* @__PURE__ */ React.createElement("span", { className: "inline-flex items-center gap-1 text-amber-400 text-xs font-mono" }, /* @__PURE__ */ React.createElement(Clock, { className: "w-3 h-3" }), "\xC5PEN")), /* @__PURE__ */ React.createElement("td", { className: "py-3 pr-8 text-right" }, /* @__PURE__ */ React.createElement("button", { onClick: () => onEdit(r), className: "text-stone-400 hover:text-amber-300 p-1" }, /* @__PURE__ */ React.createElement(Edit2, { className: "w-4 h-4" })))))))))));
}
const DEMO_DATA = [
  {
    BatchOrderNumber: "66-00017905",
    BatchOrderName: "Roser Premium 12x50cm",
    DataAreaId: "66",
    ItemNumber: "300009",
    ItemBatchNumber: "",
    Warehouse: "Inbound",
    Site: "BBS",
    Status: "Created",
    ScheduledDate: "2026-06-09T12:00:00Z",
    ScheduledQuantity: 64,
    Dimension: "1008"
  },
  {
    BatchOrderNumber: "66-00017904",
    BatchOrderName: "Roser 10x50cm",
    DataAreaId: "66",
    ItemNumber: "300004",
    ItemBatchNumber: "",
    Warehouse: "Inbound",
    Site: "BBS",
    Status: "StartedUp",
    ScheduledDate: "2026-06-02T12:00:00Z",
    ScheduledQuantity: 1e3,
    Dimension: "1008"
  },
  {
    BatchOrderNumber: "66-00017902",
    BatchOrderName: "Tulipaner 7PK Gul",
    DataAreaId: "66",
    ItemNumber: "302501",
    ItemBatchNumber: "",
    Warehouse: "BBT",
    Site: "BBS",
    Status: "StartedUp",
    ScheduledDate: "2026-05-27T12:00:00Z",
    ScheduledQuantity: 240,
    Dimension: "1008"
  },
  {
    BatchOrderNumber: "66-00017800",
    BatchOrderName: "Peoner 5PK NG",
    DataAreaId: "66",
    ItemNumber: "302327",
    ItemBatchNumber: "210611-660000064",
    Warehouse: "BBT",
    Site: "BBS",
    Status: "ReportedFinished",
    ScheduledDate: "2026-05-18T12:00:00Z",
    ScheduledQuantity: 120,
    Dimension: "1008"
  },
  {
    BatchOrderNumber: "66-00017750",
    BatchOrderName: "Gerbera Mix 3PK",
    DataAreaId: "66",
    ItemNumber: "302055",
    ItemBatchNumber: "210611-660000031",
    Warehouse: "Outbound",
    Site: "BBS",
    Status: "Completed",
    ScheduledDate: "2026-05-12T12:00:00Z",
    ScheduledQuantity: 480,
    Dimension: "1008"
  }
];
function App() {
  const [orders, setOrders] = useState([]);
  const [sourceName, setSourceName] = useState("");
  const [ordersUpdatedAt, setOrdersUpdatedAt] = useState(null);
  const [refreshing, setRefreshing] = useState(false);
  const [registrations, setRegistrations] = useState([]);
  const [activeSessions, setActiveSessions] = useState([]);
  const [loading, setLoading] = useState(true);
  const [authState, setAuthState] = useState("checking");
  const [apiError, setApiError] = useState(null);
  const [pendingDeleteId, setPendingDeleteId] = useState(null);
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("active");
  const [warehouseFilter, setWarehouseFilter] = useState("all");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [showOnlyWithReg, setShowOnlyWithReg] = useState(false);
  const [selectedOrder, setSelectedOrder] = useState(null);
  const [editingReg, setEditingReg] = useState(null);
  const [activeSessionContext, setActiveSessionContext] = useState(null);
  const [detailOrder, setDetailOrder] = useState(null);
  const [showPanel, setShowPanel] = useState(false);
  const [sortBy, setSortBy] = useState("date");
  const [sortDir, setSortDir] = useState("desc");
  const navIgnoreRef = useRef(false);
  const pushView = (view) => {
    navIgnoreRef.current = true;
    window.history.pushState(view, "");
    applyView(view);
    navIgnoreRef.current = false;
  };
  const applyView = (view) => {
    if (!view) {
      setDetailOrder(null);
      setSelectedOrder(null);
      setEditingReg(null);
      setActiveSessionContext(null);
      setShowPanel(false);
      return;
    }
    if (view.type === "detail") {
      const bo = orders.find((o) => o.BatchOrderNumber === view.batchOrderNumber);
      setDetailOrder(bo || null);
      setSelectedOrder(null);
      setEditingReg(null);
      setActiveSessionContext(null);
      setShowPanel(false);
    } else if (view.type === "panel") {
      setShowPanel(true);
      setSelectedOrder(null);
      setEditingReg(null);
      setActiveSessionContext(null);
    } else if (view.type === "form") {
      const bo = orders.find((o) => o.BatchOrderNumber === view.batchOrderNumber);
      if (!bo) return;
      setSelectedOrder(bo);
      setEditingReg(view.editingRegId ? registrations.find((r) => r.id === view.editingRegId) : null);
      setActiveSessionContext(view.activeLine ? { batchOrderNumber: view.batchOrderNumber, line: view.activeLine } : null);
    }
  };
  useEffect(() => {
    const onPopState = (e) => {
      if (navIgnoreRef.current) return;
      applyView(e.state);
    };
    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
  }, [orders, registrations]);
  const popView = () => {
    window.history.back();
  };
  useEffect(() => {
    let cancelled = false;
    async function init() {
      const onAuth = () => {
        setAuthState("needs-login");
        setRegistrations([]);
        setActiveSessions([]);
      };
      window.addEventListener("auth-required", onAuth);
      if (window.history.state === null) {
        window.history.replaceState(null, "");
      }
      const token = getToken();
      if (!token) {
        if (!cancelled) {
          setAuthState("needs-login");
          setLoading(false);
        }
        return () => window.removeEventListener("auth-required", onAuth);
      }
      try {
        await apiCheckSession();
        if (cancelled) return;
        setAuthState("logged-in");
        const [regs, sessions] = await Promise.all([loadRegistrations(), loadActiveSessions()]);
        if (cancelled) return;
        setRegistrations(regs);
        setActiveSessions(sessions);
        await refreshOrdersFromBackend(cancelled);
      } catch (e) {
        if (!cancelled) setAuthState("needs-login");
      } finally {
        if (!cancelled) setLoading(false);
      }
      return () => window.removeEventListener("auth-required", onAuth);
    }
    init();
    return () => {
      cancelled = true;
    };
  }, []);
  const refreshOrdersFromBackend = async (cancelled = false) => {
    try {
      const data = await loadBatchOrders();
      if (cancelled) return;
      if (data && Array.isArray(data.orders) && data.orders.length > 0) {
        const mapped = mapBatchOrders(data.orders);
        setOrders(mapped);
        setSourceName(data.source || "API");
        setOrdersUpdatedAt(data.updatedAt || null);
      }
    } catch (e) {
      console.warn("Kunne ikke hente batch orders fra backend:", e);
    }
  };
  const handleLoginSuccess = async () => {
    setAuthState("logged-in");
    setLoading(true);
    try {
      const [regs, sessions] = await Promise.all([loadRegistrations(), loadActiveSessions()]);
      setRegistrations(regs);
      setActiveSessions(sessions);
      await refreshOrdersFromBackend();
    } catch (e) {
      setApiError(e.message || "Kunne ikke laste data");
    } finally {
      setLoading(false);
    }
  };
  const handleLogout = () => {
    setToken(null);
    setRole(null);
    setAuthState("needs-login");
    setRegistrations([]);
    setActiveSessions([]);
    setOrders([]);
    setSourceName("");
    try {
      window.history.replaceState(null, "");
    } catch {
    }
  };
  const activeSessionsFor = (boNumber) => activeSessions.filter((s) => s.batchOrderNumber === boNumber);
  const activeSessionForLine = (boNumber, line) => activeSessions.find((s) => s.batchOrderNumber === boNumber && s.line === line) || null;
  const handleStartSession = async (session) => {
    setActiveSessions((prev) => {
      const others = prev.filter(
        (s) => !(s.batchOrderNumber === session.batchOrderNumber && s.line === session.line)
      );
      return [...others, session];
    });
    try {
      await saveActiveSession(session);
    } catch (e) {
      setApiError(e.message || "Kunne ikke lagre p\xE5g\xE5ende sesjon");
      try {
        setActiveSessions(await loadActiveSessions());
      } catch {
      }
    }
  };
  const handleStopSession = async () => {
  };
  const handleCancelActive = async (boNumber, line) => {
    setActiveSessions(
      (prev) => prev.filter((s) => !(s.batchOrderNumber === boNumber && s.line === line))
    );
    try {
      await deleteActiveSession(boNumber, line);
    } catch (e) {
      setApiError(e.message || "Kunne ikke forkaste sesjon");
      try {
        setActiveSessions(await loadActiveSessions());
      } catch {
      }
    }
    popView();
  };
  const handleSaveRegistration = async (reg) => {
    setRegistrations((prev) => {
      const idx = prev.findIndex((r) => r.id === reg.id);
      return idx >= 0 ? prev.map((r, i) => i === idx ? reg : r) : [...prev, reg];
    });
    try {
      const saved = await saveRegistration(reg);
      setRegistrations((prev) => prev.map((r) => r.id === saved.id ? saved : r));
    } catch (e) {
      setApiError(e.message || "Kunne ikke lagre registrering");
      try {
        setRegistrations(await loadRegistrations());
      } catch {
      }
      return;
    }
    const hadActive = activeSessions.some(
      (s) => s.batchOrderNumber === reg.batchOrderNumber && s.line === reg.line
    );
    if (hadActive) {
      setActiveSessions(
        (prev) => prev.filter((s) => !(s.batchOrderNumber === reg.batchOrderNumber && s.line === reg.line))
      );
      try {
        await deleteActiveSession(reg.batchOrderNumber, reg.line);
      } catch {
      }
    }
    popView();
  };
  const handleDeleteRegistration = (id) => {
    setPendingDeleteId(id);
  };
  const confirmDelete = async (adminPassword) => {
    if (!pendingDeleteId) return;
    try {
      await deleteRegistration(pendingDeleteId, adminPassword);
      setRegistrations((prev) => prev.filter((r) => r.id !== pendingDeleteId));
      setPendingDeleteId(null);
      popView();
    } catch (e) {
      throw e;
    }
  };
  const regCountByBO = useMemo(() => {
    const m = {};
    for (const r of registrations) m[r.batchOrderNumber] = (m[r.batchOrderNumber] || 0) + 1;
    return m;
  }, [registrations]);
  const warehouses = useMemo(() => {
    const set = /* @__PURE__ */ new Set();
    for (const o of orders) if (o.Warehouse) set.add(o.Warehouse);
    return Array.from(set).sort();
  }, [orders]);
  const filtered = useMemo(() => {
    const q = search.toLowerCase().trim();
    const fromTs = dateFrom ? (/* @__PURE__ */ new Date(dateFrom + "T00:00:00")).getTime() : null;
    const toTs = dateTo ? (/* @__PURE__ */ new Date(dateTo + "T23:59:59")).getTime() : null;
    let list = orders.filter((o) => {
      if (statusFilter === "active") {
        if (o.Status === "Completed") return false;
      } else if (statusFilter !== "all" && o.Status !== statusFilter) {
        return false;
      }
      if (warehouseFilter !== "all" && o.Warehouse !== warehouseFilter) return false;
      if (showOnlyWithReg && !regCountByBO[o.BatchOrderNumber]) return false;
      if (fromTs != null || toTs != null) {
        if (!o.ScheduledDate) return false;
        const ts = new Date(o.ScheduledDate).getTime();
        if (fromTs != null && ts < fromTs) return false;
        if (toTs != null && ts > toTs) return false;
      }
      if (q) {
        const hay = `${o.BatchOrderNumber} ${o.BatchOrderName || ""} ${o.ItemNumber || ""} ${o.Site} ${o.Warehouse} ${o.ItemBatchNumber} ${o.Dimension}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
    list.sort((a, b) => {
      let av, bv;
      if (sortBy === "date") {
        av = a.ScheduledDate;
        bv = b.ScheduledDate;
      } else if (sortBy === "number") {
        av = a.BatchOrderNumber;
        bv = b.BatchOrderNumber;
      } else if (sortBy === "qty") {
        av = a.ScheduledQuantity;
        bv = b.ScheduledQuantity;
      } else {
        av = a.Status;
        bv = b.Status;
      }
      const cmp = av < bv ? -1 : av > bv ? 1 : 0;
      return sortDir === "asc" ? cmp : -cmp;
    });
    return list;
  }, [orders, search, statusFilter, warehouseFilter, dateFrom, dateTo, showOnlyWithReg, regCountByBO, sortBy, sortDir]);
  const hasActiveFilters = search || statusFilter !== "active" || warehouseFilter !== "all" || dateFrom || dateTo || showOnlyWithReg;
  const clearFilters = () => {
    setSearch("");
    setStatusFilter("active");
    setWarehouseFilter("all");
    setDateFrom("");
    setDateTo("");
    setShowOnlyWithReg(false);
  };
  const toggleSort = (col) => {
    if (sortBy === col) setSortDir(sortDir === "asc" ? "desc" : "asc");
    else {
      setSortBy(col);
      setSortDir("desc");
    }
  };
  const SortHeader = ({ col, children, className = "" }) => /* @__PURE__ */ React.createElement("th", { className: `py-3 cursor-pointer hover:text-stone-200 select-none ${className}`, onClick: () => toggleSort(col) }, /* @__PURE__ */ React.createElement("span", { className: "inline-flex items-center gap-1" }, children, sortBy === col && (sortDir === "asc" ? /* @__PURE__ */ React.createElement(ChevronUp, { className: "w-3 h-3" }) : /* @__PURE__ */ React.createElement(ChevronDown, { className: "w-3 h-3" }))));
  if (loading || authState === "checking") {
    return /* @__PURE__ */ React.createElement("div", { className: "min-h-screen bg-stone-950 flex items-center justify-center text-stone-500 font-mono text-sm" }, "Initialiserer ...");
  }
  if (authState === "needs-login") {
    return /* @__PURE__ */ React.createElement(LoginScreen, { onSuccess: handleLoginSuccess });
  }
  return /* @__PURE__ */ React.createElement("div", { className: "min-h-screen bg-stone-950 text-stone-100" }, /* @__PURE__ */ React.createElement("header", { className: "border-b border-stone-800 bg-stone-950 sticky top-0 z-30" }, /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 md:px-8 py-3 sm:py-4 flex items-center justify-between gap-3" }, /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-3 min-w-0" }, /* @__PURE__ */ React.createElement("div", { className: "w-8 h-8 bg-amber-400 flex items-center justify-center text-stone-950 font-bold font-mono flex-shrink-0" }, "B"), /* @__PURE__ */ React.createElement("div", { className: "min-w-0" }, /* @__PURE__ */ React.createElement("div", { className: "text-[10px] font-mono tracking-[0.3em] text-amber-400 truncate" }, "BATCH ORDER TRACKER"), /* @__PURE__ */ React.createElement("div", { className: "text-xs sm:text-sm text-stone-400 font-mono truncate" }, orders.length > 0 ? /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("span", { className: "hidden sm:inline" }, orders.length, " ORDERS // "), /* @__PURE__ */ React.createElement("span", { className: "sm:hidden" }, orders.length, "/"), registrations.length, /* @__PURE__ */ React.createElement("span", { className: "hidden sm:inline" }, " REGISTRERINGER"), activeSessions.length > 0 && /* @__PURE__ */ React.createElement("span", { className: "text-emerald-400 ml-1" }, "// ", activeSessions.length, " P\xC5G\xC5R")) : "INGEN DATA", sourceName && /* @__PURE__ */ React.createElement("span", { className: "text-stone-600 hidden md:inline" }, " \xB7 ", sourceName), ordersUpdatedAt && /* @__PURE__ */ React.createElement("span", { className: "text-stone-600 hidden md:inline" }, " \xB7 synket ", fmtDateTime(ordersUpdatedAt))))), /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-1.5 sm:gap-2 flex-shrink-0" }, orders.length > 0 && /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => pushView({ type: "panel" }),
      className: "inline-flex items-center gap-1.5 sm:gap-2 px-2 sm:px-3 py-2 border border-stone-700 hover:border-amber-400/50 hover:bg-amber-400/5 text-stone-300 hover:text-amber-200 text-xs font-mono tracking-wider",
      title: "Vis alle registreringer"
    },
    /* @__PURE__ */ React.createElement(Activity, { className: "w-4 h-4" }),
    /* @__PURE__ */ React.createElement("span", { className: "hidden sm:inline" }, "BO_WORK_INFO"),
    registrations.length > 0 && /* @__PURE__ */ React.createElement("span", { className: "bg-amber-400 text-stone-950 px-1.5 font-bold" }, registrations.length)
  ), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: async () => {
        setRefreshing(true);
        try {
          await refreshOrdersFromBackend();
        } finally {
          setRefreshing(false);
        }
      },
      disabled: refreshing,
      className: "px-2 sm:px-3 py-2 border border-stone-700 hover:border-stone-500 text-stone-400 hover:text-stone-200 text-xs font-mono tracking-wider disabled:opacity-50",
      title: ordersUpdatedAt ? `Sist oppdatert: ${fmtDateTime(ordersUpdatedAt)}` : "Hent batch orders fra backend"
    },
    /* @__PURE__ */ React.createElement(RefreshCw, { className: `w-4 h-4 sm:hidden ${refreshing ? "animate-spin" : ""}` }),
    /* @__PURE__ */ React.createElement("span", { className: "hidden sm:inline" }, refreshing ? "OPPDATERER ..." : "OPPDATER")
  ), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => {
        setOrders([]);
        setSourceName("");
      },
      className: "hidden md:inline px-3 py-2 border border-stone-700 hover:border-stone-500 text-stone-400 hover:text-stone-200 text-xs font-mono tracking-wider",
      title: "Bytt til manuell filopplasting"
    },
    "FIL"
  )), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: handleLogout,
      className: "px-2 sm:px-3 py-2 border border-stone-700 hover:border-rose-400/50 hover:bg-rose-400/5 text-stone-400 hover:text-rose-300 text-xs font-mono tracking-wider",
      title: "Logg ut"
    },
    /* @__PURE__ */ React.createElement(X, { className: "w-4 h-4 sm:hidden" }),
    /* @__PURE__ */ React.createElement("span", { className: "hidden sm:inline" }, "LOGG UT")
  )))), orders.length === 0 ? /* @__PURE__ */ React.createElement(
    UploadPanel,
    {
      onLoad: (o, name) => {
        setOrders(o);
        setSourceName(name);
      },
      onDemo: () => {
        setOrders(DEMO_DATA);
        setSourceName("DEMO");
      }
    }
  ) : /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("div", { className: "border-b border-stone-800 bg-stone-950" }, /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 md:px-8 pt-3 sm:pt-4 pb-3 flex flex-wrap items-center gap-2 sm:gap-3" }, /* @__PURE__ */ React.createElement("div", { className: "relative w-full sm:flex-1 sm:min-w-[240px] sm:max-w-md" }, /* @__PURE__ */ React.createElement(Search, { className: "absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-stone-500" }), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "text",
      placeholder: "S\xF8k batch order, vare, lager...",
      value: search,
      onChange: (e) => setSearch(e.target.value),
      className: "w-full bg-stone-900 border border-stone-800 pl-10 pr-3 py-2.5 sm:py-2 text-sm text-stone-200 placeholder-stone-600 focus:border-amber-400 focus:outline-none font-mono"
    }
  )), /* @__PURE__ */ React.createElement("div", { className: "flex items-center border border-stone-800 flex-wrap" }, [
    { v: "all", l: "ALLE" },
    { v: "active", l: "AKTIVE" },
    { v: "Completed", l: "FULLF\xD8RT" }
  ].map((s) => /* @__PURE__ */ React.createElement(
    "button",
    {
      key: s.v,
      onClick: () => setStatusFilter(s.v),
      className: `px-2.5 sm:px-3 py-2 sm:py-1.5 text-[10px] font-mono tracking-wider ${statusFilter === s.v ? "bg-amber-400 text-stone-950" : "text-stone-400 hover:text-stone-200"}`
    },
    s.l
  )), /* @__PURE__ */ React.createElement(
    "select",
    {
      value: ["all", "active", "Completed"].includes(statusFilter) ? "" : statusFilter,
      onChange: (e) => setStatusFilter(e.target.value || "all"),
      className: "bg-stone-900 border-l border-stone-800 px-2 py-2 sm:py-1.5 text-[10px] text-stone-300 font-mono tracking-wider focus:outline-none"
    },
    /* @__PURE__ */ React.createElement("option", { value: "" }, "SPESIFIKK ..."),
    /* @__PURE__ */ React.createElement("option", { value: "Created" }, "OPPRETTET"),
    /* @__PURE__ */ React.createElement("option", { value: "StartedUp" }, "STARTET"),
    /* @__PURE__ */ React.createElement("option", { value: "ReportedFinished" }, "FERDIGSTILT"),
    /* @__PURE__ */ React.createElement("option", { value: "Completed" }, "FULLF\xD8RT"),
    /* @__PURE__ */ React.createElement("option", { value: "CostEstimated" }, "ESTIMERT")
  )), /* @__PURE__ */ React.createElement("label", { className: "hidden sm:flex items-center gap-2 text-xs font-mono text-stone-400 cursor-pointer hover:text-stone-200" }, /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "checkbox",
      checked: showOnlyWithReg,
      onChange: (e) => setShowOnlyWithReg(e.target.checked),
      className: "accent-amber-400"
    }
  ), "MED REGISTRERING"), /* @__PURE__ */ React.createElement("div", { className: "ml-auto text-xs font-mono text-stone-500 whitespace-nowrap" }, filtered.length, " / ", orders.length)), /* @__PURE__ */ React.createElement("div", { className: "px-4 sm:px-6 md:px-8 pb-4 flex flex-wrap items-center gap-2 sm:gap-3" }, /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-2 flex-1 sm:flex-initial min-w-[140px]" }, /* @__PURE__ */ React.createElement("span", { className: "text-[10px] font-mono tracking-wider text-stone-500 flex-shrink-0" }, "LAGER"), /* @__PURE__ */ React.createElement(
    "select",
    {
      value: warehouseFilter,
      onChange: (e) => setWarehouseFilter(e.target.value),
      className: "flex-1 sm:flex-initial bg-stone-900 border border-stone-800 px-3 py-2 sm:py-1.5 text-xs text-stone-200 font-mono focus:border-amber-400 focus:outline-none sm:min-w-[140px]"
    },
    /* @__PURE__ */ React.createElement("option", { value: "all" }, "Alle (", warehouses.length, ")"),
    warehouses.map((w) => /* @__PURE__ */ React.createElement("option", { key: w, value: w }, w))
  )), /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-2 flex-wrap" }, /* @__PURE__ */ React.createElement("span", { className: "text-[10px] font-mono tracking-wider text-stone-500" }, "PLANLAGT"), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "date",
      value: dateFrom,
      onChange: (e) => setDateFrom(e.target.value),
      className: "bg-stone-900 border border-stone-800 px-2 py-2 sm:py-1.5 text-xs text-stone-200 font-mono focus:border-amber-400 focus:outline-none"
    }
  ), /* @__PURE__ */ React.createElement("span", { className: "text-stone-600 text-xs font-mono" }, "\u2192"), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "date",
      value: dateTo,
      onChange: (e) => setDateTo(e.target.value),
      className: "bg-stone-900 border border-stone-800 px-2 py-2 sm:py-1.5 text-xs text-stone-200 font-mono focus:border-amber-400 focus:outline-none"
    }
  )), /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-1 flex-wrap" }, (() => {
    const toISO = (d) => {
      const tz = d.getTimezoneOffset() * 6e4;
      return new Date(d - tz).toISOString().slice(0, 10);
    };
    const today = /* @__PURE__ */ new Date();
    today.setHours(0, 0, 0, 0);
    const dayAt = (offset) => {
      const d = new Date(today);
      d.setDate(d.getDate() + offset);
      return toISO(d);
    };
    const quickRanges = [
      { label: "I G\xC5R", from: dayAt(-1), to: dayAt(-1) },
      { label: "I DAG", from: dayAt(0), to: dayAt(0) },
      { label: "I MORGEN", from: dayAt(1), to: dayAt(1) },
      { label: "+2 DAGER", from: dayAt(0), to: dayAt(2) }
    ];
    return quickRanges.map((r) => {
      const active = dateFrom === r.from && dateTo === r.to;
      return /* @__PURE__ */ React.createElement(
        "button",
        {
          key: r.label,
          onClick: () => {
            setDateFrom(r.from);
            setDateTo(r.to);
          },
          className: `px-2 py-2 sm:py-1.5 text-[10px] font-mono tracking-wider border ${active ? "bg-amber-400 text-stone-950 border-amber-400" : "border-stone-800 text-stone-400 hover:border-stone-600 hover:text-stone-200"}`
        },
        r.label
      );
    });
  })()), /* @__PURE__ */ React.createElement("label", { className: "sm:hidden flex items-center gap-2 text-xs font-mono text-stone-400 cursor-pointer w-full mt-1" }, /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "checkbox",
      checked: showOnlyWithReg,
      onChange: (e) => setShowOnlyWithReg(e.target.checked),
      className: "accent-amber-400"
    }
  ), "MED REGISTRERING"), hasActiveFilters && /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: clearFilters,
      className: "inline-flex items-center gap-1.5 px-2.5 py-2 sm:py-1.5 text-[10px] font-mono tracking-wider text-rose-300 hover:text-rose-200 hover:bg-rose-500/10 border border-rose-500/30"
    },
    /* @__PURE__ */ React.createElement(X, { className: "w-3 h-3" }),
    "NULLSTILL"
  ))), /* @__PURE__ */ React.createElement("div", { className: "md:hidden divide-y divide-stone-900" }, filtered.slice(0, 500).map((o) => {
    const regs = registrations.filter((r) => r.batchOrderNumber === o.BatchOrderNumber);
    const actives = activeSessionsFor(o.BatchOrderNumber);
    const hasActive = actives.length > 0;
    return /* @__PURE__ */ React.createElement("div", { key: o.BatchOrderNumber, className: `p-4 ${hasActive ? "bg-emerald-500/5" : ""}` }, /* @__PURE__ */ React.createElement("div", { className: "flex items-start justify-between gap-3 mb-2" }, /* @__PURE__ */ React.createElement("div", { className: "min-w-0 flex-1" }, /* @__PURE__ */ React.createElement(
      "button",
      {
        onClick: () => pushView({ type: "detail", batchOrderNumber: o.BatchOrderNumber }),
        className: "font-mono text-stone-100 text-base hover:text-amber-300 transition-colors flex items-center gap-2"
      },
      hasActive && /* @__PURE__ */ React.createElement("span", { className: "w-1.5 h-1.5 bg-emerald-400 rounded-full animate-pulse flex-shrink-0" }),
      o.BatchOrderNumber
    ), /* @__PURE__ */ React.createElement("div", { className: "font-mono text-stone-300 text-xs mt-1" }, o.ItemNumber || "\u2014"), o.BatchOrderName && /* @__PURE__ */ React.createElement("div", { className: "text-stone-400 text-xs mt-0.5 uppercase tracking-wide" }, o.BatchOrderName)), /* @__PURE__ */ React.createElement(StatusBadge, { status: o.Status })), /* @__PURE__ */ React.createElement("div", { className: "grid grid-cols-4 gap-2 text-xs font-mono mb-3" }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 text-[10px]" }, "MENGDE"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-300" }, o.ScheduledQuantity || "\u2014")), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 text-[10px]" }, "PLANLAGT"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-300" }, fmtDate(o.ScheduledDate))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 text-[10px]" }, "SITE"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-300" }, o.Site || "\u2014")), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { className: "text-stone-500 text-[10px]" }, "LAGER"), /* @__PURE__ */ React.createElement("div", { className: "text-stone-300" }, o.Warehouse || "\u2014"))), /* @__PURE__ */ React.createElement("div", { className: "flex items-center justify-between gap-2 flex-wrap" }, /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-1.5 flex-wrap" }, actives.sort((a, b) => a.line - b.line).map((s) => /* @__PURE__ */ React.createElement(
      "button",
      {
        key: s.line,
        onClick: () => pushView({ type: "form", batchOrderNumber: o.BatchOrderNumber, activeLine: s.line }),
        className: "inline-flex items-center gap-1 px-2 py-1 bg-emerald-500/15 border border-emerald-400/40 hover:bg-emerald-500/25 text-emerald-300 text-xs font-mono font-bold tracking-wider"
      },
      /* @__PURE__ */ React.createElement("span", { className: "w-1.5 h-1.5 bg-emerald-400 rounded-full animate-pulse" }),
      "L",
      s.line
    )), regs.length > 0 && /* @__PURE__ */ React.createElement(
      "button",
      {
        onClick: () => pushView({ type: "detail", batchOrderNumber: o.BatchOrderNumber }),
        className: "inline-flex items-center justify-center min-w-[24px] h-6 px-2 bg-amber-400/15 border border-amber-400/40 text-amber-300 text-xs font-mono font-bold"
      },
      regs.length
    )), /* @__PURE__ */ React.createElement(
      "button",
      {
        onClick: () => pushView({ type: "form", batchOrderNumber: o.BatchOrderNumber }),
        className: "inline-flex items-center gap-1.5 px-4 py-2 bg-stone-800 hover:bg-amber-400 text-stone-300 hover:text-stone-950 text-xs font-mono font-bold transition-colors"
      },
      /* @__PURE__ */ React.createElement(Plus, { className: "w-3.5 h-3.5" }),
      " NY"
    )));
  })), /* @__PURE__ */ React.createElement("div", { className: "hidden md:block overflow-x-auto" }, /* @__PURE__ */ React.createElement("table", { className: "w-full text-sm" }, /* @__PURE__ */ React.createElement("thead", { className: "bg-stone-900/60" }, /* @__PURE__ */ React.createElement("tr", { className: "text-left text-[10px] font-mono tracking-wider text-stone-500 border-b border-stone-800" }, /* @__PURE__ */ React.createElement(SortHeader, { col: "number", className: "pl-6 md:pl-8" }, "BATCH ORDER"), /* @__PURE__ */ React.createElement("th", { className: "py-3" }, "VARE"), /* @__PURE__ */ React.createElement("th", { className: "py-3" }, "SITE"), /* @__PURE__ */ React.createElement("th", { className: "py-3" }, "LAGER"), /* @__PURE__ */ React.createElement(SortHeader, { col: "qty" }, "MENGDE"), /* @__PURE__ */ React.createElement(SortHeader, { col: "date" }, "PLANLAGT"), /* @__PURE__ */ React.createElement(SortHeader, { col: "status" }, "STATUS"), /* @__PURE__ */ React.createElement("th", { className: "py-3" }, "REG"), /* @__PURE__ */ React.createElement("th", { className: "py-3 pr-6 md:pr-8 text-right" }))), /* @__PURE__ */ React.createElement("tbody", null, filtered.slice(0, 500).map((o) => {
    const regs = registrations.filter((r) => r.batchOrderNumber === o.BatchOrderNumber);
    const actives = activeSessionsFor(o.BatchOrderNumber);
    const hasActive = actives.length > 0;
    return /* @__PURE__ */ React.createElement("tr", { key: o.BatchOrderNumber, className: `border-b border-stone-900 hover:bg-stone-900/40 group ${hasActive ? "bg-emerald-500/5" : ""}` }, /* @__PURE__ */ React.createElement("td", { className: "pl-6 md:pl-8 py-3 font-mono text-stone-100" }, hasActive && /* @__PURE__ */ React.createElement("span", { className: "inline-block w-1.5 h-1.5 bg-emerald-400 rounded-full mr-2 align-middle animate-pulse", title: "P\xE5g\xE5ende registrering" }), /* @__PURE__ */ React.createElement(
      "button",
      {
        onClick: () => pushView({ type: "detail", batchOrderNumber: o.BatchOrderNumber }),
        className: "hover:text-amber-300 hover:underline underline-offset-4 transition-colors",
        title: "Vis alle registreringer"
      },
      o.BatchOrderNumber
    )), /* @__PURE__ */ React.createElement("td", { className: "py-3 text-xs" }, /* @__PURE__ */ React.createElement("div", { className: "font-mono text-stone-200" }, o.ItemNumber || "\u2014"), o.BatchOrderName && /* @__PURE__ */ React.createElement("div", { className: "text-stone-400 mt-0.5 uppercase tracking-wide", style: { fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif" } }, o.BatchOrderName)), /* @__PURE__ */ React.createElement("td", { className: "py-3 font-mono text-stone-300 text-xs" }, o.Site || "\u2014"), /* @__PURE__ */ React.createElement("td", { className: "py-3 font-mono text-stone-300 text-xs" }, o.Warehouse || "\u2014"), /* @__PURE__ */ React.createElement("td", { className: "py-3 font-mono text-stone-300 text-xs" }, o.ScheduledQuantity || "\u2014"), /* @__PURE__ */ React.createElement("td", { className: "py-3 font-mono text-stone-300 text-xs" }, fmtDate(o.ScheduledDate)), /* @__PURE__ */ React.createElement("td", { className: "py-3" }, /* @__PURE__ */ React.createElement(StatusBadge, { status: o.Status })), /* @__PURE__ */ React.createElement("td", { className: "py-3" }, hasActive ? /* @__PURE__ */ React.createElement("div", { className: "flex items-center gap-1 flex-wrap" }, actives.sort((a, b) => a.line - b.line).map((s) => /* @__PURE__ */ React.createElement(
      "button",
      {
        key: s.line,
        onClick: () => pushView({ type: "form", batchOrderNumber: o.BatchOrderNumber, activeLine: s.line }),
        className: "inline-flex items-center gap-1 px-1.5 py-0.5 bg-emerald-500/15 border border-emerald-400/40 hover:bg-emerald-500/25 text-emerald-300 text-[10px] font-mono font-bold tracking-wider",
        title: `\xC5pne p\xE5g\xE5ende L${s.line}`
      },
      /* @__PURE__ */ React.createElement("span", { className: "w-1 h-1 bg-emerald-400 rounded-full animate-pulse" }),
      "L",
      s.line
    )), regs.length > 0 && /* @__PURE__ */ React.createElement(
      "button",
      {
        onClick: () => pushView({ type: "detail", batchOrderNumber: o.BatchOrderNumber }),
        className: "inline-flex items-center justify-center min-w-[18px] h-5 px-1 bg-amber-400/15 border border-amber-400/40 text-amber-300 text-[10px] font-mono font-bold hover:bg-amber-400/25 ml-1",
        title: `${regs.length} fullf\xF8rte registreringer`
      },
      regs.length
    )) : regs.length > 0 ? /* @__PURE__ */ React.createElement(
      "button",
      {
        onClick: () => pushView({ type: "detail", batchOrderNumber: o.BatchOrderNumber }),
        className: "flex items-center gap-1.5",
        title: "Vis alle registreringer"
      },
      /* @__PURE__ */ React.createElement("span", { className: "inline-flex items-center justify-center min-w-[20px] h-5 px-1.5 bg-amber-400/15 border border-amber-400/40 text-amber-300 text-[10px] font-mono font-bold hover:bg-amber-400/25" }, regs.length),
      /* @__PURE__ */ React.createElement("div", { className: "flex gap-0.5" }, [...new Set(regs.map((r) => r.line))].sort().map((l) => /* @__PURE__ */ React.createElement("span", { key: l, className: "w-1.5 h-4 bg-stone-700", style: {
        backgroundColor: ["#f43f5e", "#f59e0b", "#10b981", "#0ea5e9", "#8b5cf6"][l - 1]
      }, title: `Linje ${l}` })))
    ) : /* @__PURE__ */ React.createElement("span", { className: "text-stone-700" }, "\u2014")), /* @__PURE__ */ React.createElement("td", { className: "py-3 pr-6 md:pr-8 text-right" }, /* @__PURE__ */ React.createElement(
      "button",
      {
        onClick: () => pushView({ type: "form", batchOrderNumber: o.BatchOrderNumber }),
        className: "inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-mono font-bold transition-colors bg-stone-800 hover:bg-amber-400 text-stone-300 hover:text-stone-950"
      },
      /* @__PURE__ */ React.createElement(Plus, { className: "w-3.5 h-3.5" }),
      " NY"
    )));
  }))), filtered.length > 500 && /* @__PURE__ */ React.createElement("div", { className: "px-6 md:px-8 py-4 text-xs font-mono text-stone-500 border-b border-stone-900" }, "Viser de f\xF8rste 500 av ", filtered.length, " treff. Bruk s\xF8k og filtre for \xE5 snevre inn."), filtered.length === 0 && /* @__PURE__ */ React.createElement("div", { className: "px-8 py-16 text-center text-stone-500 font-mono text-sm" }, "Ingen batch orders matcher filtrene."))), selectedOrder && /* @__PURE__ */ React.createElement(
    RegistrationForm,
    {
      batchOrder: selectedOrder,
      existing: editingReg,
      activeSession: !editingReg && activeSessionContext && activeSessionContext.batchOrderNumber === selectedOrder.BatchOrderNumber ? activeSessionForLine(selectedOrder.BatchOrderNumber, activeSessionContext.line) : null,
      activeSessions: activeSessionsFor(selectedOrder.BatchOrderNumber),
      onSave: handleSaveRegistration,
      onCancel: popView,
      onDelete: handleDeleteRegistration,
      onStart: handleStartSession,
      onStop: handleStopSession,
      onCancelActive: handleCancelActive
    }
  ), showPanel && /* @__PURE__ */ React.createElement(
    RegistrationsPanel,
    {
      registrations,
      batchOrders: orders,
      onEdit: (reg) => pushView({ type: "detail", batchOrderNumber: reg.batchOrderNumber }),
      onClose: popView
    }
  ), detailOrder && !selectedOrder && /* @__PURE__ */ React.createElement(
    BatchOrderDetailView,
    {
      batchOrder: detailOrder,
      registrations,
      activeSessions: activeSessionsFor(detailOrder.BatchOrderNumber),
      onEdit: (reg) => pushView({ type: "form", batchOrderNumber: detailOrder.BatchOrderNumber, editingRegId: reg.id }),
      onNew: () => pushView({ type: "form", batchOrderNumber: detailOrder.BatchOrderNumber }),
      onResumeActive: (line) => pushView({ type: "form", batchOrderNumber: detailOrder.BatchOrderNumber, activeLine: line }),
      onClose: popView
    }
  ), pendingDeleteId && /* @__PURE__ */ React.createElement(
    AdminPasswordModal,
    {
      onConfirm: confirmDelete,
      onCancel: () => setPendingDeleteId(null)
    }
  ), apiError && /* @__PURE__ */ React.createElement("div", { className: "fixed bottom-4 left-1/2 -translate-x-1/2 z-[70] max-w-md bg-rose-950 border border-rose-500/40 px-4 py-3 flex items-start gap-3 shadow-lg" }, /* @__PURE__ */ React.createElement(AlertCircle, { className: "w-4 h-4 text-rose-400 flex-shrink-0 mt-0.5" }), /* @__PURE__ */ React.createElement("div", { className: "text-rose-200 text-sm font-mono flex-1" }, apiError), /* @__PURE__ */ React.createElement("button", { onClick: () => setApiError(null), className: "text-rose-400 hover:text-rose-200" }, /* @__PURE__ */ React.createElement(X, { className: "w-4 h-4" }))));
}
export {
  App as default
};
