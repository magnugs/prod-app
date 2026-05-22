import React, { useState, useEffect, useMemo, useRef } from "react";
import { Upload, Search, Plus, Edit2, Trash2, Download, X, AlertCircle, CheckCircle2, Clock, Users, Activity, FileJson, ChevronDown, ChevronUp, FileSpreadsheet, Play, Square } from "lucide-react";
import * as XLSX from "xlsx";

// ============ STORAGE HELPERS ============
const STORAGE_KEY = "bo_work_info_v1";
const ACTIVE_KEY = "bo_work_info_active_v1";

async function loadRegistrations() {
  try {
    const res = await window.storage.get(STORAGE_KEY);
    return res ? JSON.parse(res.value) : [];
  } catch (e) {
    return [];
  }
}

async function saveRegistrations(list) {
  try {
    await window.storage.set(STORAGE_KEY, JSON.stringify(list));
    return true;
  } catch (e) {
    console.error("Save failed:", e);
    return false;
  }
}

async function loadActiveSessions() {
  try {
    const res = await window.storage.get(ACTIVE_KEY);
    return res ? JSON.parse(res.value) : [];
  } catch (e) {
    return [];
  }
}

async function saveActiveSessions(list) {
  try {
    await window.storage.set(ACTIVE_KEY, JSON.stringify(list));
    return true;
  } catch (e) {
    return false;
  }
}

// ============ DATA AGGREGATION ============
// Filtrerer ut placeholder-datoer (D365 bruker 1900-01-01 som "tom")
function cleanDate(iso) {
  if (!iso || typeof iso !== "string") return "";
  if (iso.startsWith("1900-")) return "";
  return iso;
}

function mapBatchOrders(rawData) {
  // Accept either { value: [...] } (OData format) or a plain array
  const records = Array.isArray(rawData) ? rawData : (rawData?.value || []);
  if (!Array.isArray(records) || records.length === 0) return [];

  // Hver record = én batch order header
  return records
    .filter((d) => d && d.BatchOrderNumber)
    .map((d) => ({
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
      Priority: d.BatchOrderPriority ?? 0,
    }));
}

// ============ UTILS ============
function fmtDate(iso) {
  if (!iso) return "—";
  try {
    const d = new Date(iso);
    return d.toLocaleDateString("nb-NO", { year: "numeric", month: "2-digit", day: "2-digit" });
  } catch { return iso; }
}

function fmtDateTime(iso) {
  if (!iso) return "—";
  try {
    const d = new Date(iso);
    return d.toLocaleString("nb-NO", { year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" });
  } catch { return iso; }
}

function durationMinutes(start, end) {
  if (!start || !end) return null;
  const s = new Date(start).getTime();
  const e = new Date(end).getTime();
  if (isNaN(s) || isNaN(e) || e <= s) return null;
  return Math.round((e - s) / 60000);
}

function personHours(start, end, people) {
  const mins = durationMinutes(start, end);
  if (mins == null || !people) return null;
  return +((mins * people) / 60).toFixed(2);
}

function fmtMinutes(min) {
  if (min == null) return "—";
  const h = Math.floor(min / 60);
  const m = min % 60;
  return h > 0 ? `${h}t ${m}m` : `${m}m`;
}

function isoLocalFromMs(ms) {
  const d = new Date(ms);
  d.setSeconds(0, 0);
  const tz = d.getTimezoneOffset() * 60000;
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
  const blob = new Blob(["\ufeff" + content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

// ============ COMPONENTS ============

function StatusBadge({ status }) {
  const styles = {
    Created: "bg-stone-400/10 text-stone-300 border-stone-400/30",
    StartedUp: "bg-amber-400/15 text-amber-300 border-amber-400/40",
    ReportedFinished: "bg-blue-400/10 text-blue-300 border-blue-400/30",
    Completed: "bg-emerald-400/10 text-emerald-300 border-emerald-400/25",
    CostEstimated: "bg-violet-400/10 text-violet-300 border-violet-400/25",
    Unknown: "bg-stone-500/10 text-stone-400 border-stone-500/25",
  };
  const label = {
    Created: "OPPRETTET",
    StartedUp: "STARTET",
    ReportedFinished: "FERDIGSTILT",
    Completed: "FULLFØRT",
    CostEstimated: "ESTIMERT",
    Unknown: "UKJENT",
  }[status] || (status || "—").toUpperCase();
  return (
    <span className={`inline-flex items-center px-2 py-0.5 text-[10px] font-mono font-bold tracking-wider border ${styles[status] || styles.Unknown}`}>
      {label}
    </span>
  );
}

function LineChip({ line }) {
  const colors = {
    1: "bg-rose-400/15 text-rose-200 border-rose-400/40",
    2: "bg-amber-400/15 text-amber-200 border-amber-400/40",
    3: "bg-emerald-400/15 text-emerald-200 border-emerald-400/40",
    4: "bg-sky-400/15 text-sky-200 border-sky-400/40",
    5: "bg-violet-400/15 text-violet-200 border-violet-400/40",
  };
  return (
    <span className={`inline-flex items-center justify-center w-7 h-7 text-xs font-mono font-bold border ${colors[line]}`}>
      L{line}
    </span>
  );
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

  return (
    <div className="max-w-3xl mx-auto pt-16 px-6">
      <div className="mb-10">
        <div className="text-[10px] font-mono tracking-[0.3em] text-amber-400 mb-3">SYSTEM // 01</div>
        <h2 className="text-3xl font-bold text-stone-100 mb-2" style={{ fontFamily: "'JetBrains Mono', monospace" }}>
          Last inn batch orders
        </h2>
        <p className="text-stone-400 text-sm leading-relaxed max-w-xl">
          Last opp en JSON-fil hentet fra Dynamics 365 F&amp;O OData-endepunktet, eller en hvilken som helst JSON-array
          med batch order-linjer. Linjer aggregeres automatisk per <code className="text-amber-300 bg-stone-900/60 px-1.5 py-0.5 text-xs">BatchOrderNumber</code>.
        </p>
      </div>

      <div
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => {
          e.preventDefault();
          setDragging(false);
          if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
        }}
        onClick={() => inputRef.current?.click()}
        className={`border-2 border-dashed p-12 cursor-pointer transition-all ${
          dragging ? "border-amber-400 bg-amber-400/5" : "border-stone-700 hover:border-stone-500 bg-stone-900/30"
        }`}
      >
        <div className="flex flex-col items-center text-center">
          <Upload className="w-10 h-10 text-stone-500 mb-4" strokeWidth={1.5} />
          <div className="text-stone-200 font-medium mb-1">Slipp JSON-fil her, eller klikk for å velge</div>
          <div className="text-xs text-stone-500 font-mono">.json — OData format eller array</div>
        </div>
        <input
          ref={inputRef}
          type="file"
          accept=".json,application/json"
          className="hidden"
          onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])}
        />
      </div>

      {err && (
        <div className="mt-4 p-3 border border-rose-500/40 bg-rose-500/10 text-rose-200 text-sm flex gap-2">
          <AlertCircle className="w-4 h-4 flex-shrink-0 mt-0.5" />
          <span>{err}</span>
        </div>
      )}

      <div className="mt-6 flex items-center gap-4">
        <div className="flex-1 h-px bg-stone-800" />
        <span className="text-xs font-mono text-stone-600 tracking-wider">ELLER</span>
        <div className="flex-1 h-px bg-stone-800" />
      </div>

      <button
        onClick={onDemo}
        className="mt-6 w-full py-3 border border-stone-700 hover:border-amber-400/50 hover:bg-amber-400/5 text-stone-300 hover:text-amber-200 font-mono text-sm tracking-wider transition-colors"
      >
        BRUK DEMO-DATA (5 batch orders)
      </button>

      <div className="mt-10 p-4 bg-stone-900/40 border border-stone-800 text-xs text-stone-400 font-mono leading-relaxed">
        <div className="text-stone-300 mb-2 font-bold">FORVENTET FORMAT</div>
        <pre className="text-stone-500 overflow-x-auto">{`[
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
]`}</pre>
      </div>
    </div>
  );
}

function RegistrationForm({ batchOrder, existing, activeSession, onSave, onCancel, onDelete, onStart, onStop, onCancelActive }) {
  // Tre faser:
  //   "setup"  – ny: velg linje + personer, så Start-knapp
  //   "running" – timer går; Avslutt-knapp
  //   "review" – juster tider/notat, lagre
  // Eksisterende -> rett til review
  // Aktiv (pågående) registrering -> running
  const initialPhase = existing
    ? "review"
    : activeSession
    ? "running"
    : "setup";

  const [phase, setPhase] = useState(initialPhase);
  const [line, setLine] = useState(existing?.line || activeSession?.line || 1);
  const [people, setPeople] = useState(existing?.people || activeSession?.people || 1);
  // start (string) fylles inn først når Start trykkes — eller fra eksisterende/aktiv sesjon
  const [start, setStart] = useState(existing?.start || activeSession?.start || "");
  const [end, setEnd] = useState(existing?.end || "");
  const [notes, setNotes] = useState(existing?.notes || "");

  // Presis starttidspunkt i ms — for live-telling uten sekundavrunding
  const [startMs, setStartMs] = useState(() => {
    if (activeSession?.startMs) return activeSession.startMs;
    if (activeSession?.start) return new Date(activeSession.start).getTime();
    if (existing?.start) return new Date(existing.start).getTime();
    return null;
  });

  // Tikkende klokke for running-fase
  const [now, setNow] = useState(Date.now());
  useEffect(() => {
    if (phase !== "running") return;
    const t = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(t);
  }, [phase]);

  const dur = durationMinutes(start, end);
  const mh = personHours(start, end, people);

  // Live elapsed-time — bruker ms-presisjon
  const liveElapsedSec = phase === "running" && startMs
    ? Math.max(0, Math.floor((now - startMs) / 1000))
    : 0;
  const liveHours = Math.floor(liveElapsedSec / 3600);
  const liveMinutes = Math.floor((liveElapsedSec % 3600) / 60);
  const liveSeconds = liveElapsedSec % 60;
  const livePersonHours = (liveElapsedSec / 3600 * people).toFixed(2);

  const handleStart = () => {
    const nowMs = Date.now();
    const startTime = isoLocalFromMs(nowMs);
    setStartMs(nowMs);
    setStart(startTime);
    onStart({
      batchOrderNumber: batchOrder.BatchOrderNumber,
      line: Number(line),
      people: Number(people),
      start: startTime,
      startMs: nowMs,
    });
    setPhase("running");
  };

  const handleStop = () => {
    const endTime = isoLocalNow();
    setEnd(endTime);
    setPhase("review");
    // Sesjonen ryddes ikke her — først ved lagring/forkasting.
    // Slik kan brukeren angre seg og fortsette uten å miste tidsbruk.
  };

  const handleSubmit = () => {
    if (!start) return alert("Starttidspunkt mangler");
    if (end && new Date(end) <= new Date(start)) return alert("Slutt må være etter start");
    onSave({
      id: existing?.id || `reg_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
      batchOrderNumber: batchOrder.BatchOrderNumber,
      line: Number(line),
      people: Number(people),
      start,
      end: end || null,
      notes,
      createdAt: existing?.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    });
  };

  const handleCloseRequest = () => {
    if (phase === "running") {
      if (!confirm("Lukke uten å avslutte? Registreringen forblir aktiv og kan gjenåpnes senere.")) return;
    }
    onCancel();
  };

  const handleCancelActive = () => {
    if (!confirm("Forkaste pågående registrering? Tidsbruk vil ikke bli lagret.")) return;
    onCancelActive();
  };

  const headerLabel = phase === "running"
    ? "PÅGÅR"
    : phase === "review" && !existing
    ? "FULLFØR REGISTRERING"
    : existing
    ? "REDIGER REGISTRERING"
    : "NY REGISTRERING";

  return (
    <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-start justify-center p-4 overflow-y-auto">
      <div className="bg-stone-950 border border-stone-700 max-w-2xl w-full mt-12 mb-12">
        <div className="px-6 py-4 border-b border-stone-800 flex items-center justify-between">
          <div>
            <div className={`text-[10px] font-mono tracking-[0.3em] ${phase === "running" ? "text-emerald-400 animate-pulse" : "text-amber-400"}`}>
              {phase === "running" && "● "}{headerLabel}
            </div>
            <div className="font-mono text-stone-100 text-lg mt-1">{batchOrder.BatchOrderNumber}</div>
          </div>
          <button onClick={handleCloseRequest} className="text-stone-500 hover:text-stone-200 p-1">
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="px-6 py-5 grid grid-cols-2 gap-4 border-b border-stone-800 bg-stone-900/30 text-xs">
          <div>
            <div className="text-stone-500 font-mono mb-1">VARE</div>
            <div className="text-stone-200 font-mono">{batchOrder.ItemNumber || "—"}</div>
            {batchOrder.BatchOrderName && (
              <div className="text-stone-400 mt-0.5 uppercase tracking-wide" style={{ fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif" }}>{batchOrder.BatchOrderName}</div>
            )}
          </div>
          <div>
            <div className="text-stone-500 font-mono mb-1">PLANLAGT</div>
            <div className="text-stone-200 font-mono">{fmtDate(batchOrder.ScheduledDate)}</div>
          </div>
          <div>
            <div className="text-stone-500 font-mono mb-1">MENGDE</div>
            <div className="text-stone-200 font-mono">{batchOrder.ScheduledQuantity || "—"}</div>
          </div>
          <div>
            <div className="text-stone-500 font-mono mb-1">SITE / LAGER</div>
            <div className="text-stone-200 font-mono">{batchOrder.Site} / {batchOrder.Warehouse}</div>
          </div>
          <div>
            <div className="text-stone-500 font-mono mb-1">STATUS</div>
            <StatusBadge status={batchOrder.Status} />
          </div>
        </div>

        {/* ====================== SETUP-fase ====================== */}
        {phase === "setup" && (
          <>
            <div className="px-6 py-6 space-y-5">
              <div>
                <label className="block text-[10px] font-mono tracking-wider text-stone-400 mb-2">PRODUKSJONSLINJE</label>
                <div className="grid grid-cols-5 gap-2">
                  {[1, 2, 3, 4, 5].map((n) => (
                    <button
                      key={n}
                      type="button"
                      onClick={() => setLine(n)}
                      className={`py-3 font-mono font-bold border transition-all ${
                        line === n
                          ? "bg-amber-400 text-stone-950 border-amber-400"
                          : "bg-stone-900 text-stone-400 border-stone-700 hover:border-stone-500"
                      }`}
                    >
                      L{n}
                    </button>
                  ))}
                </div>
              </div>
              <div>
                <label className="block text-[10px] font-mono tracking-wider text-stone-400 mb-2">ANTALL PERSONER</label>
                <div className="flex items-center gap-3">
                  <button
                    type="button"
                    onClick={() => setPeople(Math.max(1, people - 1))}
                    className="w-10 h-10 border border-stone-700 hover:border-amber-400/50 text-stone-300 font-mono text-lg"
                  >−</button>
                  <input
                    type="number"
                    min="1"
                    value={people}
                    onChange={(e) => setPeople(Math.max(1, +e.target.value || 1))}
                    className="flex-1 bg-stone-900 border border-stone-700 px-4 py-2 text-stone-100 font-mono text-center focus:border-amber-400 focus:outline-none"
                  />
                  <button
                    type="button"
                    onClick={() => setPeople(people + 1)}
                    className="w-10 h-10 border border-stone-700 hover:border-amber-400/50 text-stone-300 font-mono text-lg"
                  >+</button>
                </div>
              </div>
            </div>
            <div className="px-6 py-4 border-t border-stone-800 flex items-center justify-between gap-3">
              <button onClick={onCancel} className="px-5 py-2 border border-stone-700 hover:border-stone-500 text-stone-300 text-sm font-mono">
                AVBRYT
              </button>
              <button
                onClick={handleStart}
                className="inline-flex items-center gap-2 px-6 py-3 bg-emerald-500 hover:bg-emerald-400 text-stone-950 font-mono font-bold text-sm tracking-wider"
              >
                <Play className="w-4 h-4 fill-current" />
                START PRODUKSJON
              </button>
            </div>
          </>
        )}

        {/* ====================== RUNNING-fase ====================== */}
        {phase === "running" && (
          <>
            <div className="px-6 py-8 flex flex-col items-center border-b border-stone-800 bg-stone-900/20">
              <div className="text-[10px] font-mono tracking-[0.3em] text-emerald-400 mb-3">PRODUKSJON PÅGÅR</div>
              <div className="font-mono text-stone-100 text-6xl tabular-nums" style={{ fontVariantNumeric: "tabular-nums" }}>
                {String(liveHours).padStart(2, "0")}:{String(liveMinutes).padStart(2, "0")}:{String(liveSeconds).padStart(2, "0")}
              </div>
              <div className="text-stone-500 font-mono text-xs mt-3">
                Startet {startMs ? fmtDateTime(new Date(startMs).toISOString()) : "—"}
              </div>
              <div className="grid grid-cols-3 gap-3 mt-6 w-full max-w-md text-center">
                <div className="bg-stone-900/60 border border-stone-800 px-3 py-2">
                  <div className="text-[10px] font-mono text-stone-500 tracking-wider">LINJE</div>
                  <div className="text-amber-300 font-mono text-lg mt-1">L{line}</div>
                </div>
                <div className="bg-stone-900/60 border border-stone-800 px-3 py-2">
                  <div className="text-[10px] font-mono text-stone-500 tracking-wider">PERSONER</div>
                  <div className="text-amber-300 font-mono text-lg mt-1">{people}</div>
                </div>
                <div className="bg-stone-900/60 border border-stone-800 px-3 py-2">
                  <div className="text-[10px] font-mono text-stone-500 tracking-wider">PERSONTIMER</div>
                  <div className="text-amber-300 font-mono text-lg mt-1">{livePersonHours}h</div>
                </div>
              </div>
            </div>
            <div className="px-6 py-4 border-t border-stone-800 flex items-center justify-between gap-3">
              <button
                onClick={handleCancelActive}
                className="px-3 py-2 text-rose-400 hover:text-rose-200 hover:bg-rose-500/10 text-sm font-mono inline-flex items-center gap-2"
              >
                <Trash2 className="w-4 h-4" /> FORKAST
              </button>
              <button
                onClick={handleStop}
                className="inline-flex items-center gap-2 px-6 py-3 bg-rose-500 hover:bg-rose-400 text-stone-950 font-mono font-bold text-sm tracking-wider"
              >
                <Square className="w-4 h-4 fill-current" />
                AVSLUTT PRODUKSJON
              </button>
            </div>
          </>
        )}

        {/* ====================== REVIEW-fase ====================== */}
        {phase === "review" && (
          <>
            <div className="px-6 py-6 space-y-5">
              <div>
                <label className="block text-[10px] font-mono tracking-wider text-stone-400 mb-2">PRODUKSJONSLINJE</label>
                <div className="grid grid-cols-5 gap-2">
                  {[1, 2, 3, 4, 5].map((n) => (
                    <button
                      key={n}
                      type="button"
                      onClick={() => setLine(n)}
                      className={`py-3 font-mono font-bold border transition-all ${
                        line === n
                          ? "bg-amber-400 text-stone-950 border-amber-400"
                          : "bg-stone-900 text-stone-400 border-stone-700 hover:border-stone-500"
                      }`}
                    >
                      L{n}
                    </button>
                  ))}
                </div>
              </div>

              <div>
                <label className="block text-[10px] font-mono tracking-wider text-stone-400 mb-2">ANTALL PERSONER</label>
                <div className="flex items-center gap-3">
                  <button
                    type="button"
                    onClick={() => setPeople(Math.max(1, people - 1))}
                    className="w-10 h-10 border border-stone-700 hover:border-amber-400/50 text-stone-300 font-mono text-lg"
                  >−</button>
                  <input
                    type="number"
                    min="1"
                    value={people}
                    onChange={(e) => setPeople(Math.max(1, +e.target.value || 1))}
                    className="flex-1 bg-stone-900 border border-stone-700 px-4 py-2 text-stone-100 font-mono text-center focus:border-amber-400 focus:outline-none"
                  />
                  <button
                    type="button"
                    onClick={() => setPeople(people + 1)}
                    className="w-10 h-10 border border-stone-700 hover:border-amber-400/50 text-stone-300 font-mono text-lg"
                  >+</button>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-[10px] font-mono tracking-wider text-stone-400 mb-2">START</label>
                  <input
                    type="datetime-local"
                    value={start}
                    onChange={(e) => setStart(e.target.value)}
                    className="w-full bg-stone-900 border border-stone-700 px-3 py-2 text-stone-100 font-mono text-sm focus:border-amber-400 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="block text-[10px] font-mono tracking-wider text-stone-400 mb-2">SLUTT{existing ? " (valgfri)" : ""}</label>
                  <input
                    type="datetime-local"
                    value={end}
                    onChange={(e) => setEnd(e.target.value)}
                    className="w-full bg-stone-900 border border-stone-700 px-3 py-2 text-stone-100 font-mono text-sm focus:border-amber-400 focus:outline-none"
                  />
                </div>
              </div>

              <div>
                <label className="block text-[10px] font-mono tracking-wider text-stone-400 mb-2">NOTAT (valgfritt)</label>
                <textarea
                  value={notes}
                  onChange={(e) => setNotes(e.target.value)}
                  rows={2}
                  className="w-full bg-stone-900 border border-stone-700 px-3 py-2 text-stone-200 text-sm focus:border-amber-400 focus:outline-none resize-none"
                  placeholder="Avvik, kommentarer ..."
                />
              </div>

              {dur != null && (
                <div className="grid grid-cols-2 gap-3 pt-2">
                  <div className="bg-stone-900/60 border border-stone-800 px-4 py-3">
                    <div className="text-[10px] font-mono text-stone-500 tracking-wider">VARIGHET</div>
                    <div className="text-amber-300 font-mono text-lg mt-1">{fmtMinutes(dur)}</div>
                  </div>
                  <div className="bg-stone-900/60 border border-stone-800 px-4 py-3">
                    <div className="text-[10px] font-mono text-stone-500 tracking-wider">PERSONTIMER</div>
                    <div className="text-amber-300 font-mono text-lg mt-1">{mh}h</div>
                  </div>
                </div>
              )}
            </div>
            <div className="px-6 py-4 border-t border-stone-800 flex items-center justify-between gap-3">
              <div>
                {existing && (
                  <button
                    onClick={() => {
                      if (confirm("Slette registreringen permanent?")) onDelete(existing.id);
                    }}
                    className="px-3 py-2 text-rose-400 hover:text-rose-200 hover:bg-rose-500/10 text-sm font-mono inline-flex items-center gap-2"
                  >
                    <Trash2 className="w-4 h-4" /> SLETT
                  </button>
                )}
              </div>
              <div className="flex gap-2">
                <button onClick={onCancel} className="px-5 py-2 border border-stone-700 hover:border-stone-500 text-stone-300 text-sm font-mono">
                  AVBRYT
                </button>
                <button onClick={handleSubmit} className="px-5 py-2 bg-amber-400 hover:bg-amber-300 text-stone-950 font-mono font-bold text-sm">
                  {existing ? "OPPDATER" : "LAGRE"}
                </button>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function BatchOrderDetailView({ batchOrder, registrations, activeSession, onEdit, onNew, onClose, onResumeActive }) {
  const regs = useMemo(() => {
    return registrations
      .filter((r) => r.batchOrderNumber === batchOrder.BatchOrderNumber)
      .map((r) => ({
        ...r,
        durationMin: durationMinutes(r.start, r.end),
        personHours: personHours(r.start, r.end, r.people),
      }))
      .sort((a, b) => (b.start || "").localeCompare(a.start || ""));
  }, [registrations, batchOrder]);

  const stats = useMemo(() => {
    const closed = regs.filter((r) => r.end);
    const totalMin = closed.reduce((s, r) => s + (r.durationMin || 0), 0);
    const totalMh = closed.reduce((s, r) => s + (r.personHours || 0), 0);
    const byLine = {};
    for (const r of closed) {
      byLine[r.line] = (byLine[r.line] || 0) + (r.personHours || 0);
    }
    return {
      total: regs.length,
      closed: closed.length,
      open: regs.length - closed.length,
      totalMin,
      totalMh: totalMh.toFixed(1),
      byLine,
    };
  }, [regs]);

  return (
    <div className="fixed inset-0 z-40 bg-stone-950 overflow-y-auto">
      <div className="min-h-full flex flex-col">
        {/* HEADER */}
        <div className="px-6 md:px-8 py-5 border-b border-stone-800 flex items-start justify-between gap-4 sticky top-0 bg-stone-950 z-10">
          <div className="flex-1 min-w-0">
            <div className="text-[10px] font-mono tracking-[0.3em] text-amber-400">BATCH ORDER</div>
            <div className="flex items-baseline gap-4 mt-1 flex-wrap">
              <h2 className="text-2xl text-stone-100 font-mono">{batchOrder.BatchOrderNumber}</h2>
              <StatusBadge status={batchOrder.Status} />
            </div>
            {batchOrder.BatchOrderName && (
              <div className="text-stone-300 mt-1 uppercase tracking-wide text-sm">{batchOrder.BatchOrderName}</div>
            )}
          </div>
          <button onClick={onClose} className="text-stone-400 hover:text-stone-100 p-2 flex-shrink-0">
            <X className="w-6 h-6" />
          </button>
        </div>

        {/* META */}
        <div className="px-6 md:px-8 py-5 border-b border-stone-800 grid grid-cols-2 md:grid-cols-5 gap-4 text-xs bg-stone-900/30">
          <div>
            <div className="text-stone-500 font-mono mb-1">VARE</div>
            <div className="text-stone-200 font-mono">{batchOrder.ItemNumber || "—"}</div>
          </div>
          <div>
            <div className="text-stone-500 font-mono mb-1">PLANLAGT</div>
            <div className="text-stone-200 font-mono">{fmtDate(batchOrder.ScheduledDate)}</div>
          </div>
          <div>
            <div className="text-stone-500 font-mono mb-1">MENGDE</div>
            <div className="text-stone-200 font-mono">{batchOrder.ScheduledQuantity || "—"}</div>
          </div>
          <div>
            <div className="text-stone-500 font-mono mb-1">SITE</div>
            <div className="text-stone-200 font-mono">{batchOrder.Site || "—"}</div>
          </div>
          <div>
            <div className="text-stone-500 font-mono mb-1">LAGER</div>
            <div className="text-stone-200 font-mono">{batchOrder.Warehouse || "—"}</div>
          </div>
        </div>

        {/* STATS */}
        <div className="px-6 md:px-8 py-5 border-b border-stone-800 grid grid-cols-2 md:grid-cols-4 gap-3">
          <div className="border border-stone-800 px-4 py-3 bg-stone-900/30">
            <div className="text-[10px] font-mono text-stone-500 tracking-wider">TOTALT</div>
            <div className="text-2xl font-mono text-stone-100 mt-1">{stats.total}</div>
          </div>
          <div className="border border-stone-800 px-4 py-3 bg-stone-900/30">
            <div className="text-[10px] font-mono text-stone-500 tracking-wider">PÅGÅR</div>
            <div className="text-2xl font-mono text-amber-300 mt-1">{stats.open}</div>
          </div>
          <div className="border border-stone-800 px-4 py-3 bg-stone-900/30">
            <div className="text-[10px] font-mono text-stone-500 tracking-wider">VARIGHET</div>
            <div className="text-2xl font-mono text-stone-100 mt-1">{fmtMinutes(stats.totalMin)}</div>
          </div>
          <div className="border border-stone-800 px-4 py-3 bg-stone-900/30">
            <div className="text-[10px] font-mono text-stone-500 tracking-wider">PERSONTIMER</div>
            <div className="text-2xl font-mono text-amber-300 mt-1">{stats.totalMh}h</div>
          </div>
        </div>

        {/* LINJEFORDELING */}
        {Object.keys(stats.byLine).length > 0 && (
          <div className="px-6 md:px-8 py-4 border-b border-stone-800 flex items-center gap-3 flex-wrap">
            <div className="text-[10px] font-mono tracking-wider text-stone-500">FORDELING</div>
            {[1, 2, 3, 4, 5].map((n) => stats.byLine[n] && (
              <div key={n} className="flex items-center gap-2 text-xs font-mono">
                <LineChip line={n} />
                <span className="text-stone-300">{stats.byLine[n].toFixed(1)}h</span>
              </div>
            ))}
          </div>
        )}

        {/* AKTIV SESJON-BANNER */}
        {activeSession && (
          <div className="px-6 md:px-8 py-4 border-b border-stone-800 bg-emerald-500/5 flex items-center justify-between gap-3">
            <div className="flex items-center gap-3">
              <div className="w-2 h-2 bg-emerald-400 rounded-full animate-pulse" />
              <div>
                <div className="text-[10px] font-mono tracking-wider text-emerald-400">PÅGÅENDE REGISTRERING</div>
                <div className="text-sm text-stone-200 font-mono mt-0.5">
                  L{activeSession.line} · {activeSession.people} pers · startet {fmtDateTime(activeSession.start)}
                </div>
              </div>
            </div>
            <button
              onClick={onResumeActive}
              className="inline-flex items-center gap-2 px-4 py-2 bg-emerald-500 hover:bg-emerald-400 text-stone-950 font-mono font-bold text-xs tracking-wider"
            >
              <Square className="w-3.5 h-3.5 fill-current" />
              ÅPNE
            </button>
          </div>
        )}

        {/* ACTIONS */}
        <div className="px-6 md:px-8 py-4 border-b border-stone-800 flex items-center justify-between">
          <div className="text-[10px] font-mono tracking-wider text-stone-500">
            ALLE REGISTRERINGER
          </div>
          {!activeSession && (
            <button
              onClick={onNew}
              className="inline-flex items-center gap-2 px-4 py-2 bg-amber-400 hover:bg-amber-300 text-stone-950 font-mono font-bold text-xs tracking-wider"
            >
              <Plus className="w-3.5 h-3.5" />
              NY REGISTRERING
            </button>
          )}
        </div>

        {/* LISTE */}
        <div className="flex-1">
          {regs.length === 0 ? (
            <div className="px-8 py-16 text-center text-stone-500">
              <Activity className="w-12 h-12 mx-auto mb-3" strokeWidth={1} />
              <div className="font-mono text-sm">Ingen registreringer ennå</div>
              <div className="text-xs mt-1">Trykk NY REGISTRERING for å starte</div>
            </div>
          ) : (
            <div className="divide-y divide-stone-900">
              {regs.map((r) => (
                <div
                  key={r.id}
                  className="px-6 md:px-8 py-4 hover:bg-stone-900/40 group flex items-center gap-4"
                >
                  <LineChip line={r.line} />
                  <div className="flex-1 min-w-0 grid grid-cols-1 md:grid-cols-4 gap-3 items-center">
                    <div>
                      <div className="text-[10px] font-mono text-stone-500 tracking-wider">START</div>
                      <div className="text-stone-200 font-mono text-sm mt-0.5">{fmtDateTime(r.start)}</div>
                    </div>
                    <div>
                      <div className="text-[10px] font-mono text-stone-500 tracking-wider">SLUTT</div>
                      <div className="font-mono text-sm mt-0.5">
                        {r.end ? <span className="text-stone-200">{fmtDateTime(r.end)}</span> : <span className="text-amber-400">— pågår</span>}
                      </div>
                    </div>
                    <div className="flex gap-5">
                      <div>
                        <div className="text-[10px] font-mono text-stone-500 tracking-wider">VARIGHET</div>
                        <div className="text-stone-200 font-mono text-sm mt-0.5">{fmtMinutes(r.durationMin)}</div>
                      </div>
                      <div>
                        <div className="text-[10px] font-mono text-stone-500 tracking-wider">PERS</div>
                        <div className="text-stone-200 font-mono text-sm mt-0.5 inline-flex items-center gap-1">
                          <Users className="w-3 h-3 text-stone-500" />{r.people}
                        </div>
                      </div>
                      <div>
                        <div className="text-[10px] font-mono text-stone-500 tracking-wider">PERSONTIMER</div>
                        <div className="text-amber-300 font-mono text-sm mt-0.5">{r.personHours != null ? `${r.personHours}h` : "—"}</div>
                      </div>
                    </div>
                    <div>
                      {r.notes && (
                        <>
                          <div className="text-[10px] font-mono text-stone-500 tracking-wider">NOTAT</div>
                          <div className="text-stone-300 text-xs mt-0.5 truncate" title={r.notes}>{r.notes}</div>
                        </>
                      )}
                    </div>
                  </div>
                  <button
                    onClick={() => onEdit(r)}
                    className="opacity-60 group-hover:opacity-100 inline-flex items-center gap-1.5 px-3 py-2 border border-stone-700 hover:border-amber-400/50 hover:bg-amber-400/5 text-stone-300 hover:text-amber-200 text-xs font-mono tracking-wider flex-shrink-0 transition-all"
                  >
                    <Edit2 className="w-3.5 h-3.5" />
                    REDIGER
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function RegistrationsPanel({ registrations, onEdit, onClose, batchOrders }) {
  const enriched = useMemo(() => {
    const boMap = new Map(batchOrders.map((b) => [b.BatchOrderNumber, b]));
    return registrations.map((r) => ({
      ...r,
      bo: boMap.get(r.batchOrderNumber),
      durationMin: durationMinutes(r.start, r.end),
      personHours: personHours(r.start, r.end, r.people),
    }));
  }, [registrations, batchOrders]);

  const stats = useMemo(() => {
    const closed = enriched.filter((r) => r.end);
    const open = enriched.filter((r) => !r.end);
    const totalMh = closed.reduce((s, r) => s + (r.personHours || 0), 0);
    const byLine = {};
    for (const r of closed) {
      byLine[r.line] = (byLine[r.line] || 0) + (r.personHours || 0);
    }
    return { closed: closed.length, open: open.length, totalMh: totalMh.toFixed(1), byLine };
  }, [enriched]);

  const exportHeaders = [
    { key: "id", label: "RegistrationId" },
    { key: "batchOrderNumber", label: "BatchOrderNumber" },
    { label: "BatchOrderName", get: (r) => r.bo?.BatchOrderName || "" },
    { key: "line", label: "Line" },
    { key: "people", label: "People" },
    { key: "start", label: "StartTime" },
    { key: "end", label: "EndTime" },
    { label: "DurationMinutes", get: (r) => r.durationMin ?? "" },
    { label: "PersonHours", get: (r) => r.personHours ?? "" },
    { label: "ItemNumber", get: (r) => r.bo?.ItemNumber || "" },
    { label: "Site", get: (r) => r.bo?.Site || "" },
    { label: "Warehouse", get: (r) => r.bo?.Warehouse || "" },
    { label: "ScheduledDate", get: (r) => r.bo?.ScheduledDate || "" },
    { key: "notes", label: "Notes" },
    { key: "createdAt", label: "CreatedAt" },
    { key: "updatedAt", label: "UpdatedAt" },
  ];

  const handleExportCsv = () => {
    download(
      `bo_work_info_${new Date().toISOString().slice(0, 10)}.csv`,
      toCSV(enriched, exportHeaders)
    );
  };

  const handleExportXlsx = () => {
    // Sheet 1: detaljerte registreringer
    const rows = enriched.map((r) => {
      const obj = {};
      for (const h of exportHeaders) {
        obj[h.label] = typeof h.get === "function" ? h.get(r) : r[h.key];
      }
      return obj;
    });
    const wsData = XLSX.utils.json_to_sheet(rows, { header: exportHeaders.map((h) => h.label) });
    wsData["!cols"] = [
      { wch: 22 }, { wch: 16 }, { wch: 28 }, { wch: 6 }, { wch: 6 },
      { wch: 18 }, { wch: 18 }, { wch: 8 }, { wch: 10 },
      { wch: 14 }, { wch: 8 }, { wch: 12 }, { wch: 20 },
      { wch: 30 }, { wch: 22 }, { wch: 22 },
    ];

    // Sheet 2: sammendrag pr linje (KPI-grunnlag)
    const byLine = {};
    for (const r of enriched.filter((x) => x.end)) {
      const k = r.line;
      if (!byLine[k]) byLine[k] = { Line: k, Registrations: 0, TotalMinutes: 0, PersonHours: 0, UniqueBatchOrders: new Set() };
      byLine[k].Registrations++;
      byLine[k].TotalMinutes += r.durationMin || 0;
      byLine[k].PersonHours += r.personHours || 0;
      byLine[k].UniqueBatchOrders.add(r.batchOrderNumber);
    }
    const summary = Object.values(byLine)
      .sort((a, b) => a.Line - b.Line)
      .map((s) => ({
        Line: s.Line,
        Registrations: s.Registrations,
        UniqueBatchOrders: s.UniqueBatchOrders.size,
        TotalMinutes: s.TotalMinutes,
        PersonHours: +s.PersonHours.toFixed(2),
        AvgPersonHoursPerRegistration: +(s.PersonHours / s.Registrations).toFixed(2),
      }));
    const wsSummary = XLSX.utils.json_to_sheet(summary);
    wsSummary["!cols"] = [{ wch: 6 }, { wch: 14 }, { wch: 18 }, { wch: 14 }, { wch: 12 }, { wch: 26 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsData, "Registreringer");
    XLSX.utils.book_append_sheet(wb, wsSummary, "Sammendrag pr linje");
    XLSX.writeFile(wb, `bo_work_info_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <div className="fixed inset-0 z-40 bg-stone-950">
      <div className="h-full flex flex-col">
        <div className="px-8 py-5 border-b border-stone-800 flex items-center justify-between">
          <div>
            <div className="text-[10px] font-mono tracking-[0.3em] text-amber-400">BO_WORK_INFO</div>
            <h2 className="text-2xl text-stone-100 font-mono mt-1">Registreringer</h2>
          </div>
          <button onClick={onClose} className="text-stone-400 hover:text-stone-100 p-2">
            <X className="w-6 h-6" />
          </button>
        </div>

        <div className="px-8 py-5 border-b border-stone-800 grid grid-cols-2 md:grid-cols-4 gap-4">
          <div className="border border-stone-800 px-4 py-3 bg-stone-900/30">
            <div className="text-[10px] font-mono text-stone-500 tracking-wider">TOTALT</div>
            <div className="text-2xl font-mono text-stone-100 mt-1">{registrations.length}</div>
          </div>
          <div className="border border-stone-800 px-4 py-3 bg-stone-900/30">
            <div className="text-[10px] font-mono text-stone-500 tracking-wider">PÅGÅR</div>
            <div className="text-2xl font-mono text-amber-300 mt-1">{stats.open}</div>
          </div>
          <div className="border border-stone-800 px-4 py-3 bg-stone-900/30">
            <div className="text-[10px] font-mono text-stone-500 tracking-wider">FULLFØRT</div>
            <div className="text-2xl font-mono text-emerald-300 mt-1">{stats.closed}</div>
          </div>
          <div className="border border-stone-800 px-4 py-3 bg-stone-900/30">
            <div className="text-[10px] font-mono text-stone-500 tracking-wider">SUM PERSONTIMER</div>
            <div className="text-2xl font-mono text-stone-100 mt-1">{stats.totalMh}h</div>
          </div>
        </div>

        {Object.keys(stats.byLine).length > 0 && (
          <div className="px-8 py-4 border-b border-stone-800 flex items-center gap-3 flex-wrap">
            <div className="text-[10px] font-mono tracking-wider text-stone-500">FORDELING PR LINJE</div>
            {[1,2,3,4,5].map((n) => stats.byLine[n] && (
              <div key={n} className="flex items-center gap-2 text-xs font-mono">
                <LineChip line={n} />
                <span className="text-stone-300">{stats.byLine[n].toFixed(1)}h</span>
              </div>
            ))}
          </div>
        )}

        <div className="px-8 py-3 border-b border-stone-800 flex items-center justify-end gap-2">
          <button
            onClick={handleExportCsv}
            disabled={registrations.length === 0}
            className="inline-flex items-center gap-2 px-3 py-1.5 border border-stone-700 hover:border-amber-400/50 hover:bg-amber-400/5 text-stone-300 hover:text-amber-200 text-xs font-mono tracking-wider disabled:opacity-30 disabled:hover:border-stone-700 disabled:hover:bg-transparent"
          >
            <Download className="w-3.5 h-3.5" />
            EKSPORT CSV
          </button>
          <button
            onClick={handleExportXlsx}
            disabled={registrations.length === 0}
            className="inline-flex items-center gap-2 px-3 py-1.5 border border-emerald-500/40 hover:border-emerald-400 hover:bg-emerald-500/10 text-emerald-300 hover:text-emerald-200 text-xs font-mono tracking-wider disabled:opacity-30 disabled:hover:border-emerald-500/40 disabled:hover:bg-transparent"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            EKSPORT EXCEL
          </button>
        </div>

        <div className="flex-1 overflow-auto">
          {registrations.length === 0 ? (
            <div className="flex flex-col items-center justify-center h-full text-stone-500">
              <Activity className="w-12 h-12 mb-3" strokeWidth={1} />
              <div className="font-mono text-sm">Ingen registreringer ennå</div>
            </div>
          ) : (
            <table className="w-full text-sm">
              <thead className="bg-stone-900 sticky top-0">
                <tr className="text-left text-[10px] font-mono tracking-wider text-stone-500 border-b border-stone-800">
                  <th className="px-8 py-3">BATCH ORDER</th>
                  <th className="py-3">LINJE</th>
                  <th className="py-3">PERS</th>
                  <th className="py-3">START</th>
                  <th className="py-3">SLUTT</th>
                  <th className="py-3">VARIGHET</th>
                  <th className="py-3">MH</th>
                  <th className="py-3">STATUS</th>
                  <th className="py-3 pr-8"></th>
                </tr>
              </thead>
              <tbody>
                {enriched.sort((a, b) => (b.start || "").localeCompare(a.start || "")).map((r) => (
                  <tr key={r.id} className="border-b border-stone-900 hover:bg-stone-900/40">
                    <td className="px-8 py-3 font-mono text-stone-200">{r.batchOrderNumber}</td>
                    <td className="py-3"><LineChip line={r.line} /></td>
                    <td className="py-3 font-mono text-stone-300">
                      <span className="inline-flex items-center gap-1"><Users className="w-3 h-3 text-stone-500" />{r.people}</span>
                    </td>
                    <td className="py-3 font-mono text-stone-300 text-xs">{fmtDateTime(r.start)}</td>
                    <td className="py-3 font-mono text-stone-300 text-xs">{r.end ? fmtDateTime(r.end) : <span className="text-amber-400">— pågår</span>}</td>
                    <td className="py-3 font-mono text-stone-300">{fmtMinutes(r.durationMin)}</td>
                    <td className="py-3 font-mono text-amber-300">{r.personHours != null ? `${r.personHours}h` : "—"}</td>
                    <td className="py-3">
                      {r.end
                        ? <span className="inline-flex items-center gap-1 text-emerald-400 text-xs font-mono"><CheckCircle2 className="w-3 h-3" />OK</span>
                        : <span className="inline-flex items-center gap-1 text-amber-400 text-xs font-mono"><Clock className="w-3 h-3" />ÅPEN</span>}
                    </td>
                    <td className="py-3 pr-8 text-right">
                      <button onClick={() => onEdit(r)} className="text-stone-400 hover:text-amber-300 p-1">
                        <Edit2 className="w-4 h-4" />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </div>
  );
}

// ============ DEMO DATA ============
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
    Dimension: "1008",
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
    ScheduledQuantity: 1000,
    Dimension: "1008",
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
    Dimension: "1008",
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
    Dimension: "1008",
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
    Dimension: "1008",
  },
];

// ============ MAIN APP ============
export default function App() {
  const [orders, setOrders] = useState([]);
  const [sourceName, setSourceName] = useState("");
  const [registrations, setRegistrations] = useState([]);
  const [activeSessions, setActiveSessions] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("active");
  const [warehouseFilter, setWarehouseFilter] = useState("all");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [showOnlyWithReg, setShowOnlyWithReg] = useState(false);
  const [selectedOrder, setSelectedOrder] = useState(null);
  const [editingReg, setEditingReg] = useState(null);
  const [detailOrder, setDetailOrder] = useState(null);
  const [showPanel, setShowPanel] = useState(false);
  const [sortBy, setSortBy] = useState("date");
  const [sortDir, setSortDir] = useState("desc");

  useEffect(() => {
    Promise.all([loadRegistrations(), loadActiveSessions()]).then(([r, a]) => {
      setRegistrations(r);
      setActiveSessions(a);
      setLoading(false);
    });
  }, []);

  const persistRegistrations = async (next) => {
    setRegistrations(next);
    await saveRegistrations(next);
  };

  const persistActiveSessions = async (next) => {
    setActiveSessions(next);
    await saveActiveSessions(next);
  };

  // Pågående sesjon for en spesifikk batch order
  const activeSessionFor = (boNumber) =>
    activeSessions.find((s) => s.batchOrderNumber === boNumber) || null;

  const handleStartSession = async (session) => {
    // Erstatt evt. tidligere åpen sesjon på samme BO
    const others = activeSessions.filter((s) => s.batchOrderNumber !== session.batchOrderNumber);
    await persistActiveSessions([...others, session]);
  };

  const handleStopSession = async () => {
    // Sesjonen for valgt order ryddes — varigheten er nå overført til formularet i review-modus
    if (!selectedOrder) return;
    const next = activeSessions.filter((s) => s.batchOrderNumber !== selectedOrder.BatchOrderNumber);
    await persistActiveSessions(next);
  };

  const handleCancelActive = async () => {
    if (!selectedOrder) return;
    const next = activeSessions.filter((s) => s.batchOrderNumber !== selectedOrder.BatchOrderNumber);
    await persistActiveSessions(next);
    setSelectedOrder(null);
    setEditingReg(null);
  };

  const handleSaveRegistration = (reg) => {
    const idx = registrations.findIndex((r) => r.id === reg.id);
    const next = idx >= 0
      ? registrations.map((r) => (r.id === reg.id ? reg : r))
      : [...registrations, reg];
    persistRegistrations(next);
    // Rydd evt. pågående sesjon for denne BO
    const nextActive = activeSessions.filter((s) => s.batchOrderNumber !== reg.batchOrderNumber);
    if (nextActive.length !== activeSessions.length) {
      persistActiveSessions(nextActive);
    }
    setSelectedOrder(null);
    setEditingReg(null);
  };

  const handleDeleteRegistration = (id) => {
    persistRegistrations(registrations.filter((r) => r.id !== id));
    setSelectedOrder(null);
    setEditingReg(null);
  };

  const regCountByBO = useMemo(() => {
    const m = {};
    for (const r of registrations) m[r.batchOrderNumber] = (m[r.batchOrderNumber] || 0) + 1;
    return m;
  }, [registrations]);

  const warehouses = useMemo(() => {
    const set = new Set();
    for (const o of orders) if (o.Warehouse) set.add(o.Warehouse);
    return Array.from(set).sort();
  }, [orders]);

  const filtered = useMemo(() => {
    const q = search.toLowerCase().trim();
    const fromTs = dateFrom ? new Date(dateFrom + "T00:00:00").getTime() : null;
    const toTs = dateTo ? new Date(dateTo + "T23:59:59").getTime() : null;
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
      if (sortBy === "date") { av = a.ScheduledDate; bv = b.ScheduledDate; }
      else if (sortBy === "number") { av = a.BatchOrderNumber; bv = b.BatchOrderNumber; }
      else if (sortBy === "qty") { av = a.ScheduledQuantity; bv = b.ScheduledQuantity; }
      else { av = a.Status; bv = b.Status; }
      const cmp = av < bv ? -1 : av > bv ? 1 : 0;
      return sortDir === "asc" ? cmp : -cmp;
    });
    return list;
  }, [orders, search, statusFilter, warehouseFilter, dateFrom, dateTo, showOnlyWithReg, regCountByBO, sortBy, sortDir]);

  const hasActiveFilters = search || statusFilter !== "active" || warehouseFilter !== "all" || dateFrom || dateTo || showOnlyWithReg;
  const clearFilters = () => {
    setSearch(""); setStatusFilter("active"); setWarehouseFilter("all");
    setDateFrom(""); setDateTo(""); setShowOnlyWithReg(false);
  };

  const toggleSort = (col) => {
    if (sortBy === col) setSortDir(sortDir === "asc" ? "desc" : "asc");
    else { setSortBy(col); setSortDir("desc"); }
  };

  const SortHeader = ({ col, children, className = "" }) => (
    <th className={`py-3 cursor-pointer hover:text-stone-200 select-none ${className}`} onClick={() => toggleSort(col)}>
      <span className="inline-flex items-center gap-1">
        {children}
        {sortBy === col && (sortDir === "asc" ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />)}
      </span>
    </th>
  );

  if (loading) {
    return <div className="min-h-screen bg-stone-950 flex items-center justify-center text-stone-500 font-mono text-sm">Initialiserer ...</div>;
  }

  return (
    <div className="min-h-screen bg-stone-950 text-stone-100" style={{ fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&display=swap');
        body { background: #0c0a09; }
        input[type="datetime-local"]::-webkit-calendar-picker-indicator {
          filter: invert(0.6);
          cursor: pointer;
        }
        .font-mono { font-family: 'JetBrains Mono', monospace; }
      `}</style>

      {/* HEADER */}
      <header className="border-b border-stone-800 bg-stone-950 sticky top-0 z-30">
        <div className="px-6 md:px-8 py-4 flex items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <div className="w-8 h-8 bg-amber-400 flex items-center justify-center text-stone-950 font-bold font-mono">B</div>
            <div>
              <div className="text-[10px] font-mono tracking-[0.3em] text-amber-400">BATCH ORDER TRACKER</div>
              <div className="text-sm text-stone-400 font-mono">
                {orders.length > 0 ? (
                  <>
                    {orders.length} ORDERS // {registrations.length} REGISTRERINGER
                    {activeSessions.length > 0 && (
                      <span className="text-emerald-400 ml-1">// {activeSessions.length} PÅGÅR</span>
                    )}
                  </>
                ) : "INGEN DATA"}
                {sourceName && <span className="text-stone-600"> · {sourceName}</span>}
              </div>
            </div>
          </div>
          {orders.length > 0 && (
            <div className="flex items-center gap-2">
              <button
                onClick={() => setShowPanel(true)}
                className="inline-flex items-center gap-2 px-3 py-2 border border-stone-700 hover:border-amber-400/50 hover:bg-amber-400/5 text-stone-300 hover:text-amber-200 text-xs font-mono tracking-wider"
              >
                <Activity className="w-4 h-4" />
                BO_WORK_INFO
                {registrations.length > 0 && <span className="bg-amber-400 text-stone-950 px-1.5 font-bold">{registrations.length}</span>}
              </button>
              <button
                onClick={() => { setOrders([]); setSourceName(""); }}
                className="px-3 py-2 border border-stone-700 hover:border-stone-500 text-stone-400 hover:text-stone-200 text-xs font-mono tracking-wider"
              >
                BYTT KILDE
              </button>
            </div>
          )}
        </div>
      </header>

      {orders.length === 0 ? (
        <UploadPanel
          onLoad={(o, name) => { setOrders(o); setSourceName(name); }}
          onDemo={() => { setOrders(DEMO_DATA); setSourceName("DEMO"); }}
        />
      ) : (
        <>
          {/* FILTERS */}
          <div className="border-b border-stone-800 bg-stone-950">
            <div className="px-6 md:px-8 pt-4 pb-3 flex flex-wrap items-center gap-3">
              <div className="relative flex-1 min-w-[240px] max-w-md">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-stone-500" />
                <input
                  type="text"
                  placeholder="Søk batch order, vare, lager..."
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                  className="w-full bg-stone-900 border border-stone-800 pl-10 pr-3 py-2 text-sm text-stone-200 placeholder-stone-600 focus:border-amber-400 focus:outline-none font-mono"
                />
              </div>
              <div className="flex items-center gap-1 border border-stone-800">
                {[
                  { v: "all", l: "ALLE" },
                  { v: "active", l: "AKTIVE" },
                  { v: "Completed", l: "FULLFØRT" },
                ].map((s) => (
                  <button
                    key={s.v}
                    onClick={() => setStatusFilter(s.v)}
                    className={`px-3 py-1.5 text-[10px] font-mono tracking-wider ${
                      statusFilter === s.v ? "bg-amber-400 text-stone-950" : "text-stone-400 hover:text-stone-200"
                    }`}
                  >
                    {s.l}
                  </button>
                ))}
                <select
                  value={["all","active","Completed"].includes(statusFilter) ? "" : statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value || "all")}
                  className="bg-stone-900 border-l border-stone-800 px-2 py-1.5 text-[10px] text-stone-300 font-mono tracking-wider focus:outline-none"
                >
                  <option value="">SPESIFIKK ...</option>
                  <option value="Created">OPPRETTET</option>
                  <option value="StartedUp">STARTET</option>
                  <option value="ReportedFinished">FERDIGSTILT</option>
                  <option value="Completed">FULLFØRT</option>
                  <option value="CostEstimated">ESTIMERT</option>
                </select>
              </div>
              <label className="flex items-center gap-2 text-xs font-mono text-stone-400 cursor-pointer hover:text-stone-200">
                <input
                  type="checkbox"
                  checked={showOnlyWithReg}
                  onChange={(e) => setShowOnlyWithReg(e.target.checked)}
                  className="accent-amber-400"
                />
                MED REGISTRERING
              </label>
              <div className="ml-auto text-xs font-mono text-stone-500">
                {filtered.length} / {orders.length}
              </div>
            </div>
            <div className="px-6 md:px-8 pb-4 flex flex-wrap items-center gap-3">
              <div className="flex items-center gap-2">
                <span className="text-[10px] font-mono tracking-wider text-stone-500">LAGER</span>
                <select
                  value={warehouseFilter}
                  onChange={(e) => setWarehouseFilter(e.target.value)}
                  className="bg-stone-900 border border-stone-800 px-3 py-1.5 text-xs text-stone-200 font-mono focus:border-amber-400 focus:outline-none min-w-[140px]"
                >
                  <option value="all">Alle ({warehouses.length})</option>
                  {warehouses.map((w) => (
                    <option key={w} value={w}>{w}</option>
                  ))}
                </select>
              </div>
              <div className="flex items-center gap-2">
                <span className="text-[10px] font-mono tracking-wider text-stone-500">PLANLAGT</span>
                <input
                  type="date"
                  value={dateFrom}
                  onChange={(e) => setDateFrom(e.target.value)}
                  className="bg-stone-900 border border-stone-800 px-2 py-1.5 text-xs text-stone-200 font-mono focus:border-amber-400 focus:outline-none"
                />
                <span className="text-stone-600 text-xs font-mono">→</span>
                <input
                  type="date"
                  value={dateTo}
                  onChange={(e) => setDateTo(e.target.value)}
                  className="bg-stone-900 border border-stone-800 px-2 py-1.5 text-xs text-stone-200 font-mono focus:border-amber-400 focus:outline-none"
                />
                <div className="flex items-center gap-1 ml-1">
                  {(() => {
                    const toISO = (d) => {
                      const tz = d.getTimezoneOffset() * 60000;
                      return new Date(d - tz).toISOString().slice(0, 10);
                    };
                    const today = new Date();
                    today.setHours(0, 0, 0, 0);
                    const dayAt = (offset) => {
                      const d = new Date(today);
                      d.setDate(d.getDate() + offset);
                      return toISO(d);
                    };
                    const quickRanges = [
                      { label: "I GÅR", from: dayAt(-1), to: dayAt(-1) },
                      { label: "I DAG", from: dayAt(0), to: dayAt(0) },
                      { label: "I MORGEN", from: dayAt(1), to: dayAt(1) },
                      { label: "+2 DAGER", from: dayAt(0), to: dayAt(2) },
                    ];
                    return quickRanges.map((r) => {
                      const active = dateFrom === r.from && dateTo === r.to;
                      return (
                        <button
                          key={r.label}
                          onClick={() => { setDateFrom(r.from); setDateTo(r.to); }}
                          className={`px-2 py-1.5 text-[10px] font-mono tracking-wider border ${
                            active
                              ? "bg-amber-400 text-stone-950 border-amber-400"
                              : "border-stone-800 text-stone-400 hover:border-stone-600 hover:text-stone-200"
                          }`}
                        >
                          {r.label}
                        </button>
                      );
                    });
                  })()}
                </div>
              </div>
              {hasActiveFilters && (
                <button
                  onClick={clearFilters}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-[10px] font-mono tracking-wider text-rose-300 hover:text-rose-200 hover:bg-rose-500/10 border border-rose-500/30"
                >
                  <X className="w-3 h-3" />
                  NULLSTILL
                </button>
              )}
            </div>
          </div>

          {/* TABLE */}
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="bg-stone-900/60">
                <tr className="text-left text-[10px] font-mono tracking-wider text-stone-500 border-b border-stone-800">
                  <SortHeader col="number" className="pl-6 md:pl-8">BATCH ORDER</SortHeader>
                  <th className="py-3">VARE</th>
                  <th className="py-3">SITE</th>
                  <th className="py-3">LAGER</th>
                  <SortHeader col="qty">MENGDE</SortHeader>
                  <SortHeader col="date">PLANLAGT</SortHeader>
                  <SortHeader col="status">STATUS</SortHeader>
                  <th className="py-3">REG</th>
                  <th className="py-3 pr-6 md:pr-8 text-right"></th>
                </tr>
              </thead>
              <tbody>
                {filtered.slice(0, 500).map((o) => {
                  const regs = registrations.filter((r) => r.batchOrderNumber === o.BatchOrderNumber);
                  const active = activeSessionFor(o.BatchOrderNumber);
                  return (
                    <tr key={o.BatchOrderNumber} className={`border-b border-stone-900 hover:bg-stone-900/40 group ${active ? "bg-emerald-500/5" : ""}`}>
                      <td className="pl-6 md:pl-8 py-3 font-mono text-stone-100">
                        {active && (
                          <span className="inline-block w-1.5 h-1.5 bg-emerald-400 rounded-full mr-2 align-middle animate-pulse" title="Pågående registrering" />
                        )}
                        <button
                          onClick={() => setDetailOrder(o)}
                          className="hover:text-amber-300 hover:underline underline-offset-4 transition-colors"
                          title="Vis alle registreringer"
                        >
                          {o.BatchOrderNumber}
                        </button>
                      </td>
                      <td className="py-3 text-xs">
                        <div className="font-mono text-stone-200">{o.ItemNumber || "—"}</div>
                        {o.BatchOrderName && (
                          <div className="text-stone-400 mt-0.5 uppercase tracking-wide" style={{ fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif" }}>
                            {o.BatchOrderName}
                          </div>
                        )}
                      </td>
                      <td className="py-3 font-mono text-stone-300 text-xs">{o.Site || "—"}</td>
                      <td className="py-3 font-mono text-stone-300 text-xs">{o.Warehouse || "—"}</td>
                      <td className="py-3 font-mono text-stone-300 text-xs">{o.ScheduledQuantity || "—"}</td>
                      <td className="py-3 font-mono text-stone-300 text-xs">{fmtDate(o.ScheduledDate)}</td>
                      <td className="py-3"><StatusBadge status={o.Status} /></td>
                      <td className="py-3">
                        {active ? (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 bg-emerald-500/15 border border-emerald-400/40 text-emerald-300 text-[10px] font-mono font-bold tracking-wider">
                            <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full animate-pulse" />
                            PÅGÅR L{active.line}
                          </span>
                        ) : regs.length > 0 ? (
                          <button
                            onClick={() => setDetailOrder(o)}
                            className="flex items-center gap-1.5 hover:opacity-100 transition-opacity"
                            title="Vis alle registreringer"
                          >
                            <span className="inline-flex items-center justify-center min-w-[20px] h-5 px-1.5 bg-amber-400/15 border border-amber-400/40 text-amber-300 text-[10px] font-mono font-bold hover:bg-amber-400/25">
                              {regs.length}
                            </span>
                            <div className="flex gap-0.5">
                              {[...new Set(regs.map((r) => r.line))].sort().map((l) => (
                                <span key={l} className="w-1.5 h-4 bg-stone-700" style={{
                                  backgroundColor: ["#f43f5e","#f59e0b","#10b981","#0ea5e9","#8b5cf6"][l-1]
                                }} title={`Linje ${l}`} />
                              ))}
                            </div>
                          </button>
                        ) : <span className="text-stone-700">—</span>}
                      </td>
                      <td className="py-3 pr-6 md:pr-8 text-right">
                        <button
                          onClick={() => { setSelectedOrder(o); setEditingReg(null); }}
                          className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-mono font-bold transition-colors ${
                            active
                              ? "bg-emerald-500 hover:bg-emerald-400 text-stone-950"
                              : "bg-stone-800 hover:bg-amber-400 text-stone-300 hover:text-stone-950"
                          }`}
                        >
                          {active ? <><Square className="w-3.5 h-3.5 fill-current" /> ÅPNE</> : <><Plus className="w-3.5 h-3.5" /> REG</>}
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            {filtered.length > 500 && (
              <div className="px-6 md:px-8 py-4 text-xs font-mono text-stone-500 border-b border-stone-900">
                Viser de første 500 av {filtered.length} treff. Bruk søk og filtre for å snevre inn.
              </div>
            )}
            {filtered.length === 0 && (
              <div className="px-8 py-16 text-center text-stone-500 font-mono text-sm">
                Ingen batch orders matcher filtrene.
              </div>
            )}
          </div>
        </>
      )}

      {/* MODAL */}
      {selectedOrder && (
        <RegistrationForm
          batchOrder={selectedOrder}
          existing={editingReg}
          activeSession={!editingReg ? activeSessionFor(selectedOrder.BatchOrderNumber) : null}
          onSave={handleSaveRegistration}
          onCancel={() => { setSelectedOrder(null); setEditingReg(null); }}
          onDelete={handleDeleteRegistration}
          onStart={handleStartSession}
          onStop={handleStopSession}
          onCancelActive={handleCancelActive}
        />
      )}

      {/* PANEL */}
      {showPanel && (
        <RegistrationsPanel
          registrations={registrations}
          batchOrders={orders}
          onEdit={(reg) => {
            const bo = orders.find((o) => o.BatchOrderNumber === reg.batchOrderNumber);
            if (bo) {
              setShowPanel(false);
              setDetailOrder(bo);
            }
          }}
          onClose={() => setShowPanel(false)}
        />
      )}

      {/* DETAIL */}
      {detailOrder && !selectedOrder && (
        <BatchOrderDetailView
          batchOrder={detailOrder}
          registrations={registrations}
          activeSession={activeSessionFor(detailOrder.BatchOrderNumber)}
          onEdit={(reg) => {
            setSelectedOrder(detailOrder);
            setEditingReg(reg);
          }}
          onNew={() => {
            setSelectedOrder(detailOrder);
            setEditingReg(null);
          }}
          onResumeActive={() => {
            setSelectedOrder(detailOrder);
            setEditingReg(null);
          }}
          onClose={() => setDetailOrder(null)}
        />
      )}
    </div>
  );
}
