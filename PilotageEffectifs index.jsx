import React, { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, Legend, Area, AreaChart, ReferenceLine, Cell,
} from "recharts";
import {
  Upload, Users, Clock, TrendingUp, AlertTriangle, Download, Search,
  Calendar, ChevronLeft, ChevronRight, X, Check, Activity, Gauge,
  ArrowUpRight, ArrowDownRight, FileSpreadsheet, LayoutGrid, CalendarDays,
  Table2, Bell, Compass, RefreshCw, Building2, UserRound, ChevronDown,
  Sparkles,
} from "lucide-react";

/* ============================================================================
   PILOTAGE — centre de pilotage des effectifs & plannings
   Palette : encre #0B1E3B · acier #13294B · blanc #F6F8FB
             vert opérationnel #1FA97A · rouge alerte #E1443A · ambre #E8A33D
   ============================================================================ */

const C = {
  ink: "#081729",
  steel: "#0F2444",
  steel2: "#173259",
  steel3: "#1F3E6B",
  paper: "#F5F7FB",
  paperDim: "#EDF1F7",
  card: "#FFFFFF",
  good: "#1AA179",
  goodBg: "#E5F6F0",
  bad: "#E14640",
  badBg: "#FDEAEA",
  warn: "#E5A13B",
  warnBg: "#FCF2E1",
  signal: "#3E82F7",
  signalBg: "#E8F0FE",
  textDim: "#93A6C4",
  textMid: "#5B7093",
  border: "#233B62",
};

const DAY_LABELS = ["dim.", "lun.", "mar.", "mer.", "jeu.", "ven.", "sam."];
const MONTH_LABELS = ["janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.", "nov.", "déc."];

function pad2(n) { return String(n).padStart(2, "0"); }
function dateKey(d) { return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`; }
function fmtDateShort(d) { return `${DAY_LABELS[d.getDay()]} ${d.getDate()} ${MONTH_LABELS[d.getMonth()]}`; }
function fmtDateFull(d) { return `${DAY_LABELS[d.getDay()]} ${d.getDate()} ${MONTH_LABELS[d.getMonth()]} ${d.getFullYear()}`; }
function hoursLabel(h) {
  if (h == null || isNaN(h)) return "—";
  const hh = Math.floor(h);
  const mm = Math.round((h - hh) * 60);
  return mm === 0 ? `${hh}h` : `${hh}h${pad2(mm)}`;
}

function cellToMinutes(v) {
  if (v == null) return null;
  if (v instanceof Date) return v.getHours() * 60 + v.getMinutes() + v.getSeconds() / 60;
  if (typeof v === "number") { if (v >= 0 && v < 1) return Math.round(v * 24 * 60); return null; }
  if (typeof v === "string") {
    const m = v.trim().match(/^(\d{1,2})[:h](\d{2})/);
    if (m) return parseInt(m[1], 10) * 60 + parseInt(m[2], 10);
  }
  return null;
}
function cellToDurationHours(v) {
  if (v == null) return null;
  if (v instanceof Date) return v.getHours() + v.getMinutes() / 60;
  if (typeof v === "number") { if (v >= 0 && v < 2) return v * 24; return v; }
  if (typeof v === "string") {
    const m = v.trim().match(/^(\d{1,2})[:h](\d{2})/);
    if (m) return parseInt(m[1], 10) + parseInt(m[2], 10) / 60;
  }
  return null;
}
function normalizeAbsence(raw) {
  if (raw == null) return null;
  const s = String(raw).trim().toLowerCase();
  if (!s) return null;
  if (["repos", "rh", "off"].includes(s)) return "repos";
  if (["cp", "congé", "conge", "congés payés"].includes(s)) return "cp";
  if (["ca"].includes(s)) return "ca";
  if (["rc"].includes(s)) return "rc";
  if (["at", "accident"].includes(s)) return "at";
  if (["mal", "maladie", "arret", "arrêt"].includes(s)) return "maladie";
  return s;
}
const ABSENCE_LABEL = { repos: "Repos", cp: "Congé payé", ca: "Congé", rc: "Récup.", at: "Accident travail", maladie: "Maladie" };

/* ---------------------------------------------------------------------------
   PARSER — structure : ligne1 = dates par bloc de 4 col. à partir de "Début"
   ligne2 = en-têtes ; lignes suivantes = collaborateurs
--------------------------------------------------------------------------- */
function parseWorkbook(workbook) {
  const sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  if (!rows || rows.length < 3) throw new Error("Le fichier ne contient pas assez de lignes pour être un planning.");

  const dateRow = rows[0] || [];
  const headerRow = rows[1] || [];

  const blockStarts = [];
  headerRow.forEach((v, i) => { if (v && /^d[ée]but$/i.test(String(v).trim())) blockStarts.push(i); });
  if (blockStarts.length === 0) throw new Error('Colonnes "Début" introuvables en ligne 2 : structure du fichier non reconnue.');

  let fallbackBase = null;
  const blocks = blockStarts.map((c, idx) => {
    let d = dateRow[c];
    if (!(d instanceof Date) && typeof d === "number" && d > 20000) {
      const parsed = XLSX.SSF.parse_date_code(d);
      if (parsed) d = new Date(parsed.y, parsed.m - 1, parsed.d);
    }
    if (!(d instanceof Date)) {
      if (!fallbackBase) fallbackBase = new Date();
      d = new Date(fallbackBase); d.setDate(d.getDate() + idx);
    }
    return { col: c, date: new Date(d.getFullYear(), d.getMonth(), d.getDate()) };
  });

  const nameColGuess = headerRow.findIndex((v) => v && /nom/i.test(String(v)));
  const nameCol = nameColGuess >= 0 ? nameColGuess : 2;
  const matriculeCol = 0, equipeCol = 1;
  const lastBlockEnd = blocks[blocks.length - 1].col + 4;
  let reposCol = lastBlockEnd;
  headerRow.forEach((v, i) => { if (v && /repos/i.test(String(v))) reposCol = i; });

  const employees = [];
  const warnings = [];

  for (let r = 2; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const matricule = row[matriculeCol];
    const nom = row[nameCol];
    if (!matricule && !nom) continue;

    const equipe = row[equipeCol] != null ? String(row[equipeCol]).trim() : "Non affecté";
    const days = [];

    blocks.forEach((b) => {
      const debutRaw = row[b.col], finRaw = row[b.col + 1], pauseRaw = row[b.col + 2], dureeRaw = row[b.col + 3];
      const startMin = cellToMinutes(debutRaw);
      const endMinRaw = cellToMinutes(finRaw);
      const dureeH = cellToDurationHours(dureeRaw);

      let status = "vide", startMinutes = null, endMinutes = null, breakStart = null, breakLenMin = 0;

      if (startMin != null && endMinRaw != null) {
        status = "travail";
        startMinutes = startMin;
        endMinutes = endMinRaw <= startMin ? endMinRaw + 1440 : endMinRaw;
        const shiftLen = endMinutes - startMinutes;
        const workedMin = dureeH != null ? dureeH * 60 : shiftLen;
        breakLenMin = Math.max(0, Math.round(shiftLen - workedMin));
        const pauseMin = cellToMinutes(pauseRaw);
        if (pauseMin != null) {
          breakStart = pauseMin < startMinutes % 1440 ? pauseMin + 1440 : pauseMin;
          if (breakStart < startMinutes || breakStart > endMinutes) {
            warnings.push(`${nom || matricule} — pause incohérente le ${fmtDateShort(b.date)}`);
            breakStart = startMinutes + Math.max(0, (shiftLen - breakLenMin) / 2);
          }
        } else if (breakLenMin > 0) {
          breakStart = startMinutes + Math.max(0, (shiftLen - breakLenMin) / 2);
        }
      } else {
        const codeD = normalizeAbsence(debutRaw), codeF = normalizeAbsence(finRaw);
        status = codeD || codeF || "vide";
      }

      days.push({
        date: b.date, dateKey: dateKey(b.date), status,
        startMinutes, endMinutes, breakStart, breakLenMin,
        dureeH: status === "travail" ? (dureeH != null ? dureeH : (endMinutes - startMinutes - breakLenMin) / 60) : (dureeH || 0),
      });
    });

    const reposDeclares = row[reposCol];
    const totalH = days.reduce((s, d) => s + (d.dureeH || 0), 0);
    const joursTravailles = days.filter((d) => d.status === "travail").length;
    const joursRepos = days.filter((d) => d.status === "repos").length;

    employees.push({
      id: `${matricule || nom}-${r}`,
      matricule: matricule != null ? String(matricule).trim() : `—${r}`,
      equipe: equipe || "Non affecté",
      nom: nom != null ? String(nom).trim() : "(sans nom)",
      days, reposDeclares: typeof reposDeclares === "number" ? reposDeclares : joursRepos,
      totalH, joursTravailles, joursRepos,
    });
  }

  if (employees.length === 0) throw new Error("Aucune ligne collaborateur détectée dans le fichier.");
  const dates = blocks.map((b) => b.date);
  const teams = Array.from(new Set(employees.map((e) => e.equipe))).sort();
  return { employees, dates, teams, warnings, sheetName };
}

/* ---------------------------------------------------------------------------
   CALCULS — couverture horaire (gère les nuits, exclut les pauses)
--------------------------------------------------------------------------- */
function buildHourlyCoverage(employees) {
  // map: "YYYY-MM-DD|H" -> Set(matricule)
  const map = new Map();
  const addPresence = (baseDate, hourOffset, matricule) => {
    const d = new Date(baseDate);
    d.setDate(d.getDate() + Math.floor(hourOffset / 24));
    const h = ((hourOffset % 24) + 24) % 24;
    const key = `${dateKey(d)}|${h}`;
    if (!map.has(key)) map.set(key, new Set());
    map.get(key).add(matricule);
  };

  employees.forEach((emp) => {
    emp.days.forEach((day) => {
      if (day.status !== "travail" || day.startMinutes == null || day.endMinutes == null) return;
      const startH = day.startMinutes / 60, endH = day.endMinutes / 60;
      const bs = day.breakStart != null ? day.breakStart / 60 : null;
      const be = bs != null ? bs + day.breakLenMin / 60 : null;
      for (let h = Math.floor(startH); h < endH; h++) {
        const segStart = Math.max(h, startH), segEnd = Math.min(h + 1, endH);
        let covered = segEnd - segStart;
        if (bs != null && be != null) {
          const overlapStart = Math.max(segStart, bs), overlapEnd = Math.min(segEnd, be);
          if (overlapEnd > overlapStart) covered -= (overlapEnd - overlapStart);
        }
        if (covered > 0.001) addPresence(day.date, h, emp.matricule);
      }
    });
  });
  return map;
}

function coverageForDate(map, date) {
  const arr = [];
  for (let h = 0; h < 24; h++) {
    const key = `${dateKey(date)}|${h}`;
    const set = map.get(key);
    arr.push({ hour: h, effectif: set ? set.size : 0, agents: set ? Array.from(set) : [] });
  }
  return arr;
}

/* ---------------------------------------------------------------------------
   ALERTES
--------------------------------------------------------------------------- */
function computeAlerts(employees, coverageMap, dates, seuilBas) {
  const alerts = [];
  employees.forEach((emp) => {
    if (emp.totalH > 48) {
      alerts.push({ type: "surcharge", sev: "bad", emp: emp.nom, matricule: emp.matricule,
        msg: `${emp.nom} dépasse 48h sur la semaine (${hoursLabel(emp.totalH)})` });
    }
    emp.days.forEach((d) => {
      if (d.status === "travail" && d.dureeH > 12) {
        alerts.push({ type: "journee_longue", sev: "bad", emp: emp.nom, matricule: emp.matricule,
          msg: `${emp.nom} — journée de ${hoursLabel(d.dureeH)} le ${fmtDateShort(d.date)} (> 12h)` });
      }
      if (d.status === "travail" && d.dureeH > 6 && d.breakLenMin <= 0) {
        alerts.push({ type: "sans_pause", sev: "warn", emp: emp.nom, matricule: emp.matricule,
          msg: `${emp.nom} — aucune pause détectée le ${fmtDateShort(d.date)} (${hoursLabel(d.dureeH)})` });
      }
    });
  });
  dates.forEach((date) => {
    const cov = coverageForDate(coverageMap, date);
    cov.forEach((c) => {
      if (c.effectif < seuilBas) {
        alerts.push({ type: "sous_effectif", sev: "bad", emp: null, matricule: null,
          msg: `Sous-effectif le ${fmtDateShort(date)} à ${pad2(c.hour)}h — ${c.effectif} agent(s) présent(s)` });
      }
    });
  });
  return alerts;
}

/* ---------------------------------------------------------------------------
   EXPORTS
--------------------------------------------------------------------------- */
function downloadBlob(content, filename, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
function exportCSV(rows, headers, filename) {
  const esc = (v) => {
    const s = v == null ? "" : String(v);
    return /[;,"\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lines = [headers.map(esc).join(";"), ...rows.map((r) => headers.map((h) => esc(r[h])).join(";"))];
  downloadBlob("\uFEFF" + lines.join("\n"), filename, "text/csv;charset=utf-8");
}
function exportXLSX(sheets, filename) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, rows }) => {
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  const out = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  downloadBlob(out, filename, "application/octet-stream");
}

/* ============================================================================
   COMPOSANTS UI
   ============================================================================ */

function KpiCard({ icon: Icon, label, value, sub, tone = "signal", trend }) {
  const toneMap = {
    signal: { bg: C.signalBg, fg: C.signal },
    good: { bg: C.goodBg, fg: C.good },
    bad: { bg: C.badBg, fg: C.bad },
    warn: { bg: C.warnBg, fg: C.warn },
  };
  const t = toneMap[tone];
  return (
    <div style={{
      background: C.card, borderRadius: 14, padding: "18px 20px", flex: "1 1 200px",
      minWidth: 190, boxShadow: "0 1px 2px rgba(11,30,57,0.06), 0 8px 24px -16px rgba(11,30,57,0.18)",
      border: `1px solid ${C.paperDim || "#EDF1F7"}`, display: "flex", flexDirection: "column", gap: 10,
    }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <span style={{ fontSize: 12, fontWeight: 600, letterSpacing: 0.3, color: C.textMid, textTransform: "uppercase" }}>{label}</span>
        <div style={{ width: 30, height: 30, borderRadius: 9, background: t.bg, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <Icon size={15} color={t.fg} strokeWidth={2.4} />
        </div>
      </div>
      <div style={{ fontSize: 26, fontWeight: 800, color: C.ink, fontVariantNumeric: "tabular-nums", lineHeight: 1 }}>{value}</div>
      {sub && (
        <div style={{ fontSize: 12.5, color: trend === "up" ? C.good : trend === "down" ? C.bad : C.textMid, display: "flex", alignItems: "center", gap: 4 }}>
          {trend === "up" && <ArrowUpRight size={13} />}
          {trend === "down" && <ArrowDownRight size={13} />}
          {sub}
        </div>
      )}
    </div>
  );
}

function StatusPill({ status }) {
  const map = {
    travail: { bg: "#E5F6F0", fg: "#1AA179", label: "Travail" },
    repos: { bg: "#EEF1F6", fg: "#5B7093", label: "Repos" },
    vide: { bg: "#F5F5F5", fg: "#AAB4C4", label: "—" },
  };
  const s = map[status] || { bg: "#FCF2E1", fg: "#B5791E", label: ABSENCE_LABEL[status] || status };
  return (
    <span style={{ background: s.bg, color: s.fg, fontSize: 11.5, fontWeight: 700, padding: "3px 9px", borderRadius: 999, whiteSpace: "nowrap" }}>
      {s.label}
    </span>
  );
}

function Toggle({ options, value, onChange }) {
  return (
    <div style={{ display: "inline-flex", background: C.steel, borderRadius: 10, padding: 3, gap: 2 }}>
      {options.map((o) => (
        <button key={o.value} onClick={() => onChange(o.value)} style={{
          border: "none", cursor: "pointer", padding: "7px 13px", borderRadius: 8,
          fontSize: 12.5, fontWeight: 700, letterSpacing: 0.2,
          background: value === o.value ? C.card : "transparent",
          color: value === o.value ? C.ink : C.textDim,
          display: "flex", alignItems: "center", gap: 6, transition: "all .15s",
        }}>
          {o.icon && <o.icon size={13} />}
          {o.label}
        </button>
      ))}
    </div>
  );
}

function SectionTitle({ icon: Icon, title, desc, right }) {
  return (
    <div style={{ display: "flex", alignItems: "flex-end", justifyContent: "space-between", marginBottom: 14, gap: 12, flexWrap: "wrap" }}>
      <div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          {Icon && <Icon size={17} color={C.signal} strokeWidth={2.3} />}
          <h2 style={{ margin: 0, fontSize: 16.5, fontWeight: 800, color: C.ink, letterSpacing: -0.2 }}>{title}</h2>
        </div>
        {desc && <p style={{ margin: "4px 0 0", fontSize: 12.5, color: C.textMid }}>{desc}</p>}
      </div>
      {right}
    </div>
  );
}

function Card({ children, style, pad = 20 }) {
  return (
    <div style={{
      background: C.card, borderRadius: 16, padding: pad,
      boxShadow: "0 1px 2px rgba(11,30,57,0.05), 0 10px 28px -20px rgba(11,30,57,0.25)",
      ...style,
    }}>
      {children}
    </div>
  );
}

function coverageTone(effectif, seuilBas, seuilHaut) {
  if (effectif < seuilBas) return "bad";
  if (effectif > seuilHaut) return "warn";
  return "good";
}

/* ============================================================================
   IMPORT / PREVIEW
   ============================================================================ */

function ImportScreen({ onParsed, loadDemo }) {
  const [dragOver, setDragOver] = useState(false);
  const [error, setError] = useState(null);
  const [busy, setBusy] = useState(false);
  const inputRef = useRef(null);

  const handleFile = useCallback((file) => {
    if (!file) return;
    setError(null); setBusy(true);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const result = parseWorkbook(wb);
        onParsed(result, file.name);
      } catch (err) {
        setError(err.message || "Impossible de lire ce fichier.");
      } finally { setBusy(false); }
    };
    reader.onerror = () => { setError("Échec de la lecture du fichier."); setBusy(false); };
    reader.readAsArrayBuffer(file);
  }, [onParsed]);

  return (
    <div style={{ minHeight: "100%", display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
      <div style={{ maxWidth: 620, width: "100%" }}>
        <div style={{ textAlign: "center", marginBottom: 26 }}>
          <div style={{
            width: 52, height: 52, borderRadius: 15, background: "linear-gradient(135deg,#3E82F7,#1F3E6B)",
            display: "flex", alignItems: "center", justifyContent: "center", margin: "0 auto 16px",
            boxShadow: "0 10px 24px -10px rgba(62,130,247,0.6)",
          }}>
            <Gauge size={24} color="#fff" strokeWidth={2.3} />
          </div>
          <h1 style={{ margin: 0, fontSize: 24, fontWeight: 800, color: "#fff", letterSpacing: -0.4 }}>Centre de pilotage des effectifs</h1>
          <p style={{ margin: "8px 0 0", color: C.textDim, fontSize: 14, lineHeight: 1.5 }}>
            Importez un planning Excel pour générer automatiquement la couverture horaire,
            les indicateurs et les alertes d'exploitation.
          </p>
        </div>

        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files?.[0]); }}
          onClick={() => inputRef.current?.click()}
          style={{
            border: `2px dashed ${dragOver ? C.signal : "#2A4770"}`, borderRadius: 18, padding: "42px 24px",
            textAlign: "center", cursor: "pointer", background: dragOver ? "rgba(62,130,247,0.08)" : "rgba(255,255,255,0.02)",
            transition: "all .15s",
          }}>
          <input ref={inputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
            onChange={(e) => handleFile(e.target.files?.[0])} />
          <Upload size={30} color={C.signal} strokeWidth={1.8} style={{ marginBottom: 12 }} />
          <div style={{ color: "#fff", fontWeight: 700, fontSize: 15 }}>
            {busy ? "Analyse du fichier…" : "Glissez votre fichier .xlsx ici"}
          </div>
          <div style={{ color: C.textDim, fontSize: 13, marginTop: 4 }}>ou cliquez pour parcourir vos fichiers</div>
        </div>

        {error && (
          <div style={{ marginTop: 14, background: "rgba(225,70,64,0.12)", border: "1px solid rgba(225,70,64,0.4)",
            borderRadius: 10, padding: "12px 14px", color: "#FF9B96", fontSize: 13, display: "flex", gap: 9 }}>
            <AlertTriangle size={16} style={{ flexShrink: 0, marginTop: 1 }} />
            <span>{error}</span>
          </div>
        )}

        <div style={{ marginTop: 18, display: "flex", justifyContent: "center" }}>
          <button onClick={loadDemo} style={{
            background: "transparent", border: `1px solid ${C.border}`, color: C.textDim,
            borderRadius: 10, padding: "9px 16px", fontSize: 12.5, fontWeight: 600, cursor: "pointer",
            display: "flex", alignItems: "center", gap: 7,
          }}>
            <Sparkles size={14} /> Essayer avec un jeu de données de démonstration
          </button>
        </div>

        <div style={{ marginTop: 26, display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10 }}>
          {[
            ["Matricule, équipe, nom", "Une ligne par collaborateur"],
            ["Début / fin / pause / durée", "Un bloc de 4 colonnes par jour"],
            ["Repos & heures semaine", "Détectés automatiquement"],
          ].map(([t, d], i) => (
            <div key={i} style={{ background: "rgba(255,255,255,0.04)", borderRadius: 12, padding: "12px 13px" }}>
              <div style={{ color: "#fff", fontSize: 12, fontWeight: 700 }}>{t}</div>
              <div style={{ color: C.textDim, fontSize: 11.5, marginTop: 3 }}>{d}</div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function PreviewScreen({ result, fileName, onConfirm, onCancel }) {
  const { employees, dates, teams, warnings } = result;
  const sample = employees.slice(0, 6);
  return (
    <div style={{ minHeight: "100%", display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
      <div style={{ maxWidth: 900, width: "100%" }}>
        <Card style={{ background: C.card }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 4 }}>
            <FileSpreadsheet size={20} color={C.signal} />
            <h2 style={{ margin: 0, fontSize: 18, fontWeight: 800, color: C.ink }}>Prévisualisation de l'import</h2>
          </div>
          <p style={{ margin: "2px 0 18px", fontSize: 13, color: C.textMid }}>{fileName} · feuille « {result.sheetName} »</p>

          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 18 }}>
            {[
              [`${employees.length}`, "collaborateurs détectés"],
              [`${dates.length}`, "jours de planning"],
              [`${teams.length}`, "équipes"],
              [`${fmtDateShort(dates[0])} → ${fmtDateShort(dates[dates.length - 1])}`, "période"],
            ].map(([v, l], i) => (
              <div key={i} style={{ background: C.paper, borderRadius: 10, padding: "10px 14px", flex: "1 1 140px" }}>
                <div style={{ fontSize: 16, fontWeight: 800, color: C.ink }}>{v}</div>
                <div style={{ fontSize: 11.5, color: C.textMid }}>{l}</div>
              </div>
            ))}
          </div>

          {warnings.length > 0 && (
            <div style={{ background: C.warnBg, borderRadius: 10, padding: "10px 13px", marginBottom: 16, fontSize: 12.5, color: "#8A5A16" }}>
              <b>{warnings.length}</b> incohérence(s) mineure(s) détectée(s) et corrigée(s) automatiquement (pauses recalculées).
            </div>
          )}

          <div style={{ overflowX: "auto", border: `1px solid ${C.paperDim}`, borderRadius: 12 }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
              <thead>
                <tr style={{ background: C.paper }}>
                  {["Matricule", "Équipe", "Nom et prénom", "Jour 1", "Total sem."].map((h) => (
                    <th key={h} style={{ textAlign: "left", padding: "9px 12px", color: C.textMid, fontWeight: 700, fontSize: 11 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sample.map((e) => (
                  <tr key={e.id} style={{ borderTop: `1px solid ${C.paperDim}` }}>
                    <td style={{ padding: "9px 12px", fontWeight: 600, color: C.ink }}>{e.matricule}</td>
                    <td style={{ padding: "9px 12px", color: C.textMid }}>{e.equipe}</td>
                    <td style={{ padding: "9px 12px", color: C.ink }}>{e.nom}</td>
                    <td style={{ padding: "9px 12px" }}><StatusPill status={e.days[0]?.status} /></td>
                    <td style={{ padding: "9px 12px", fontWeight: 700, color: C.ink }}>{hoursLabel(e.totalH)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{ fontSize: 11.5, color: C.textMid, marginTop: 8 }}>… et {Math.max(0, employees.length - sample.length)} autres collaborateurs.</div>

          <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
            <button onClick={onCancel} style={{ background: "transparent", border: `1px solid ${C.paperDim}`, color: C.textMid,
              borderRadius: 10, padding: "10px 18px", fontWeight: 700, fontSize: 13, cursor: "pointer" }}>Annuler</button>
            <button onClick={onConfirm} style={{ background: C.ink, border: "none", color: "#fff",
              borderRadius: 10, padding: "10px 20px", fontWeight: 700, fontSize: 13, cursor: "pointer",
              display: "flex", alignItems: "center", gap: 7 }}>
              <Check size={15} /> Confirmer l'import
            </button>
          </div>
        </Card>
      </div>
    </div>
  );
}

/* ============================================================================
   NAVIGATION
   ============================================================================ */

const NAV = [
  { key: "dashboard", label: "Tableau de bord", icon: LayoutGrid },
  { key: "planning", label: "Planning", icon: CalendarDays },
  { key: "couverture", label: "Couverture horaire", icon: Table2 },
  { key: "graphiques", label: "Graphiques", icon: TrendingUp },
  { key: "alertes", label: "Alertes", icon: Bell },
  { key: "analyse", label: "Analyse", icon: Compass },
];

function TopBar({ view, setView, alertCount, fileName, onReimport, onExport }) {
  return (
    <div style={{ background: C.ink, borderBottom: `1px solid ${C.border}`, position: "sticky", top: 0, zIndex: 30 }}>
      <div style={{ padding: "14px 22px 0", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 16, flexWrap: "wrap" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 34, height: 34, borderRadius: 10, background: "linear-gradient(135deg,#3E82F7,#1F3E6B)",
            display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Gauge size={17} color="#fff" strokeWidth={2.4} />
          </div>
          <div>
            <div style={{ color: "#fff", fontWeight: 800, fontSize: 14.5, letterSpacing: -0.2, lineHeight: 1.1 }}>Pilotage effectifs</div>
            <div style={{ color: C.textDim, fontSize: 11, marginTop: 1 }}>{fileName}</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <button onClick={onExport} style={{ background: "rgba(255,255,255,0.06)", border: `1px solid ${C.border}`, color: "#fff",
            borderRadius: 9, padding: "8px 13px", fontSize: 12.5, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}>
            <Download size={14} /> Exporter
          </button>
          <button onClick={onReimport} style={{ background: "transparent", border: `1px solid ${C.border}`, color: C.textDim,
            borderRadius: 9, padding: "8px 13px", fontSize: 12.5, fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}>
            <RefreshCw size={13} /> Nouveau fichier
          </button>
        </div>
      </div>
      <div style={{ padding: "12px 22px 0", display: "flex", gap: 3, overflowX: "auto" }}>
        {NAV.map((n) => {
          const active = view === n.key;
          return (
            <button key={n.key} onClick={() => setView(n.key)} style={{
              background: "transparent", border: "none", cursor: "pointer",
              padding: "9px 14px", fontSize: 13, fontWeight: 700, whiteSpace: "nowrap",
              color: active ? "#fff" : C.textDim, borderBottom: `2.5px solid ${active ? C.signal : "transparent"}`,
              display: "flex", alignItems: "center", gap: 7, transition: "color .15s",
            }}>
              <n.icon size={15} />
              {n.label}
              {n.key === "alertes" && alertCount > 0 && (
                <span style={{ background: C.bad, color: "#fff", fontSize: 10.5, fontWeight: 800, borderRadius: 999,
                  minWidth: 17, height: 17, display: "flex", alignItems: "center", justifyContent: "center", padding: "0 4px" }}>
                  {alertCount}
                </span>
              )}
            </button>
          );
        })}
      </div>
    </div>
  );
}

function FilterBar({ teams, filterTeam, setFilterTeam, search, setSearch, extra }) {
  return (
    <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", marginBottom: 18 }}>
      <div style={{ position: "relative", flex: "1 1 240px", maxWidth: 320 }}>
        <Search size={15} color={C.textMid} style={{ position: "absolute", left: 12, top: 10 }} />
        <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Rechercher un collaborateur…"
          style={{ width: "100%", boxSizing: "border-box", border: `1px solid ${C.paperDim}`, borderRadius: 10,
            padding: "9px 12px 9px 34px", fontSize: 13, outline: "none", background: C.card, color: C.ink }} />
      </div>
      <select value={filterTeam} onChange={(e) => setFilterTeam(e.target.value)} style={{
        border: `1px solid ${C.paperDim}`, borderRadius: 10, padding: "9px 12px", fontSize: 13, fontWeight: 600,
        background: C.card, color: C.ink, cursor: "pointer" }}>
        <option value="">Toutes les équipes</option>
        {teams.map((t) => <option key={t} value={t}>{t}</option>)}
      </select>
      {extra}
    </div>
  );
}

function DateStepper({ date, dates, onChange }) {
  const idx = dates.findIndex((d) => dateKey(d) === dateKey(date));
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 6, background: C.card, border: `1px solid ${C.paperDim}`, borderRadius: 10, padding: 4 }}>
      <button disabled={idx <= 0} onClick={() => onChange(dates[idx - 1])} style={{
        border: "none", background: "transparent", cursor: idx <= 0 ? "default" : "pointer", padding: 6,
        opacity: idx <= 0 ? 0.3 : 1, display: "flex" }}>
        <ChevronLeft size={16} color={C.ink} />
      </button>
      <div style={{ fontSize: 13, fontWeight: 700, color: C.ink, minWidth: 120, textAlign: "center" }}>
        {fmtDateFull(date)}
      </div>
      <button disabled={idx >= dates.length - 1} onClick={() => onChange(dates[idx + 1])} style={{
        border: "none", background: "transparent", cursor: idx >= dates.length - 1 ? "default" : "pointer", padding: 6,
        opacity: idx >= dates.length - 1 ? 0.3 : 1, display: "flex" }}>
        <ChevronRight size={16} color={C.ink} />
      </button>
    </div>
  );
}

/* ============================================================================
   VUE — DASHBOARD
   ============================================================================ */

function DashboardView({ employees, dates, coverageMap, selectedDate, setSelectedDate, alerts }) {
  const today = new Date();
  const todayInRange = dates.some((d) => dateKey(d) === dateKey(today));
  const refDate = todayInRange ? today : dates[0];

  const cov = useMemo(() => coverageForDate(coverageMap, selectedDate), [coverageMap, selectedDate]);
  const agentsToday = useMemo(() => new Set(employees.flatMap((e) =>
    e.days.filter((d) => dateKey(d.date) === dateKey(selectedDate) && d.status === "travail").map(() => e.matricule)
  )).size, [employees, selectedDate]);

  const totalHeuresSemaine = useMemo(() => employees.reduce((s, e) => s + e.totalH, 0), [employees]);
  const effectifMoyen = useMemo(() => (cov.reduce((s, c) => s + c.effectif, 0) / 24).toFixed(1), [cov]);
  const peak = useMemo(() => cov.reduce((a, b) => (b.effectif > a.effectif ? b : a), cov[0]), [cov]);
  const low = useMemo(() => cov.reduce((a, b) => (b.effectif < a.effectif ? b : a), cov[0]), [cov]);
  const tauxOccupation = employees.length ? ((agentsToday / employees.length) * 100).toFixed(0) : 0;

  const chartData = cov.map((c) => ({ heure: `${pad2(c.hour)}h`, effectif: c.effectif }));

  return (
    <div>
      <SectionTitle icon={Gauge} title="Vue d'ensemble" desc={`Indicateurs pour le ${fmtDateFull(selectedDate)}`}
        right={<DateStepper date={selectedDate} dates={dates} onChange={setSelectedDate} />} />

      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 16 }}>
        <KpiCard icon={Users} label="Total agents" value={employees.length} sub="collaborateurs planifiés" tone="signal" />
        <KpiCard icon={UserRound} label="Agents planifiés" value={agentsToday} sub={`${tauxOccupation}% de l'effectif`} tone="good" />
        <KpiCard icon={Clock} label="Heures / semaine" value={hoursLabel(totalHeuresSemaine)} sub="cumul planning" tone="signal" />
        <KpiCard icon={Activity} label="Effectif moyen / h" value={effectifMoyen} sub="sur la journée" tone="signal" />
        <KpiCard icon={TrendingUp} label="Heure de pointe" value={`${pad2(peak.hour)}h`} sub={`${peak.effectif} agents présents`} tone="good" trend="up" />
        <KpiCard icon={AlertTriangle} label="Moins couverte" value={`${pad2(low.hour)}h`} sub={`${low.effectif} agent(s) présent(s)`} tone={low.effectif === 0 ? "bad" : "warn"} trend="down" />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr", gap: 16 }}>
        <Card>
          <SectionTitle icon={Activity} title="Courbe d'effectif horaire" desc="Nombre d'agents présents, pauses exclues" />
          <ResponsiveContainer width="100%" height={260}>
            <AreaChart data={chartData} margin={{ left: -18, right: 8 }}>
              <defs>
                <linearGradient id="fillCov" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="0%" stopColor={C.signal} stopOpacity={0.35} />
                  <stop offset="100%" stopColor={C.signal} stopOpacity={0.02} />
                </linearGradient>
              </defs>
              <CartesianGrid strokeDasharray="3 3" stroke={C.paperDim} vertical={false} />
              <XAxis dataKey="heure" tick={{ fontSize: 11, fill: C.textMid }} interval={1} axisLine={{ stroke: C.paperDim }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: C.textMid }} axisLine={false} tickLine={false} allowDecimals={false} />
              <Tooltip contentStyle={{ borderRadius: 10, border: `1px solid ${C.paperDim}`, fontSize: 12.5 }} />
              <Area type="monotone" dataKey="effectif" stroke={C.signal} strokeWidth={2.5} fill="url(#fillCov)" name="Agents présents" />
            </AreaChart>
          </ResponsiveContainer>
        </Card>

        <Card>
          <SectionTitle icon={Bell} title="Alertes actives" />
          <div style={{ display: "flex", flexDirection: "column", gap: 8, maxHeight: 260, overflowY: "auto" }}>
            {alerts.length === 0 && (
              <div style={{ fontSize: 12.5, color: C.textMid, padding: "20px 0", textAlign: "center" }}>
                Aucune alerte — planning conforme.
              </div>
            )}
            {alerts.slice(0, 8).map((a, i) => (
              <div key={i} style={{ background: a.sev === "bad" ? C.badBg : C.warnBg, borderRadius: 9, padding: "8px 11px", fontSize: 12 }}>
                <span style={{ color: a.sev === "bad" ? C.bad : "#8A5A16", fontWeight: 500 }}>{a.msg}</span>
              </div>
            ))}
            {alerts.length > 8 && <div style={{ fontSize: 11.5, color: C.textMid, textAlign: "center" }}>+ {alerts.length - 8} autres</div>}
          </div>
        </Card>
      </div>
    </div>
  );
}

/* ============================================================================
   VUE — PLANNING
   ============================================================================ */

function PlanningView({ employees, dates, teams }) {
  const [mode, setMode] = useState("semaine");
  const [filterTeam, setFilterTeam] = useState("");
  const [search, setSearch] = useState("");
  const [dayIdx, setDayIdx] = useState(0);
  const [weekStart, setWeekStart] = useState(0);

  const filtered = useMemo(() => employees.filter((e) =>
    (!filterTeam || e.equipe === filterTeam) &&
    (!search || e.nom.toLowerCase().includes(search.toLowerCase()) || e.matricule.toLowerCase().includes(search.toLowerCase()))
  ), [employees, filterTeam, search]);

  const weekDates = dates.slice(weekStart, weekStart + 7);

  return (
    <div>
      <SectionTitle icon={CalendarDays} title="Planning" desc="Vue jour, semaine ou mois de l'ensemble des collaborateurs"
        right={<Toggle value={mode} onChange={setMode} options={[
          { value: "jour", label: "Jour", icon: Calendar },
          { value: "semaine", label: "Semaine", icon: CalendarDays },
          { value: "mois", label: "Mois", icon: LayoutGrid },
        ]} />} />

      <FilterBar teams={teams} filterTeam={filterTeam} setFilterTeam={setFilterTeam} search={search} setSearch={setSearch}
        extra={
          mode === "jour" ? <DateStepper date={dates[dayIdx]} dates={dates} onChange={(d) => setDayIdx(dates.findIndex((x) => dateKey(x) === dateKey(d)))} /> :
          mode === "semaine" ? (
            <div style={{ display: "flex", alignItems: "center", gap: 6, background: C.card, border: `1px solid ${C.paperDim}`, borderRadius: 10, padding: 4 }}>
              <button disabled={weekStart <= 0} onClick={() => setWeekStart(Math.max(0, weekStart - 7))}
                style={{ border: "none", background: "transparent", cursor: "pointer", padding: 6, opacity: weekStart <= 0 ? 0.3 : 1 }}>
                <ChevronLeft size={16} color={C.ink} />
              </button>
              <div style={{ fontSize: 12.5, fontWeight: 700, color: C.ink, minWidth: 160, textAlign: "center" }}>
                {weekDates[0] && fmtDateShort(weekDates[0])} → {weekDates[weekDates.length - 1] && fmtDateShort(weekDates[weekDates.length - 1])}
              </div>
              <button disabled={weekStart + 7 >= dates.length} onClick={() => setWeekStart(Math.min(dates.length - 1, weekStart + 7))}
                style={{ border: "none", background: "transparent", cursor: "pointer", padding: 6, opacity: weekStart + 7 >= dates.length ? 0.3 : 1 }}>
                <ChevronRight size={16} color={C.ink} />
              </button>
            </div>
          ) : null
        } />

      {mode === "jour" && (
        <Card pad={0} style={{ overflow: "hidden" }}>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
              <thead>
                <tr style={{ background: C.paper }}>
                  {["Matricule", "Équipe", "Nom et prénom", "Début", "Fin", "Pause", "Durée"].map((h) => (
                    <th key={h} style={{ textAlign: "left", padding: "10px 14px", color: C.textMid, fontWeight: 700, fontSize: 11.5 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.map((e) => {
                  const d = e.days[dayIdx];
                  if (!d) return null;
                  return (
                    <tr key={e.id} style={{ borderTop: `1px solid ${C.paperDim}` }}>
                      <td style={{ padding: "9px 14px", fontWeight: 600, color: C.ink }}>{e.matricule}</td>
                      <td style={{ padding: "9px 14px", color: C.textMid }}>{e.equipe}</td>
                      <td style={{ padding: "9px 14px", color: C.ink }}>{e.nom}</td>
                      {d.status === "travail" ? (
                        <>
                          <td style={{ padding: "9px 14px", fontVariantNumeric: "tabular-nums" }}>{minToClock(d.startMinutes)}</td>
                          <td style={{ padding: "9px 14px", fontVariantNumeric: "tabular-nums" }}>{minToClock(d.endMinutes)}</td>
                          <td style={{ padding: "9px 14px", fontVariantNumeric: "tabular-nums", color: C.textMid }}>{d.breakLenMin > 0 ? `${Math.round(d.breakLenMin)} min` : "—"}</td>
                          <td style={{ padding: "9px 14px", fontWeight: 700 }}>{hoursLabel(d.dureeH)}</td>
                        </>
                      ) : (
                        <td colSpan={4} style={{ padding: "9px 14px" }}><StatusPill status={d.status} /></td>
                      )}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </Card>
      )}

      {mode === "semaine" && (
        <Card pad={0} style={{ overflow: "hidden" }}>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
              <thead>
                <tr style={{ background: C.paper }}>
                  <th style={{ textAlign: "left", padding: "10px 14px", color: C.textMid, fontWeight: 700, fontSize: 11.5, position: "sticky", left: 0, background: C.paper }}>Collaborateur</th>
                  {weekDates.map((d) => (
                    <th key={dateKey(d)} style={{ textAlign: "center", padding: "10px 10px", color: C.textMid, fontWeight: 700, fontSize: 11.5, minWidth: 92 }}>{fmtDateShort(d)}</th>
                  ))}
                  <th style={{ textAlign: "center", padding: "10px 14px", color: C.textMid, fontWeight: 700, fontSize: 11.5 }}>Total</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((e) => (
                  <tr key={e.id} style={{ borderTop: `1px solid ${C.paperDim}` }}>
                    <td style={{ padding: "9px 14px", fontWeight: 600, color: C.ink, position: "sticky", left: 0, background: C.card, whiteSpace: "nowrap" }}>
                      {e.nom}
                      <div style={{ fontSize: 10.5, color: C.textMid, fontWeight: 500 }}>{e.equipe}</div>
                    </td>
                    {weekDates.map((wd, wi) => {
                      const d = e.days[weekStart + wi];
                      return (
                        <td key={wi} style={{ padding: "7px 8px", textAlign: "center" }}>
                          {d?.status === "travail" ? (
                            <div style={{ fontSize: 11.5, fontWeight: 700, color: C.ink }}>
                              {minToClock(d.startMinutes)}–{minToClock(d.endMinutes)}
                            </div>
                          ) : <StatusPill status={d?.status} />}
                        </td>
                      );
                    })}
                    <td style={{ padding: "9px 14px", textAlign: "center", fontWeight: 800, color: C.ink }}>
                      {hoursLabel(weekDates.reduce((s, wd, wi) => s + (e.days[weekStart + wi]?.dureeH || 0), 0))}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      )}

      {mode === "mois" && (
        <Card pad={0} style={{ overflow: "hidden" }}>
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr style={{ background: C.paper }}>
                  <th style={{ textAlign: "left", padding: "10px 14px", color: C.textMid, fontWeight: 700, fontSize: 11.5, position: "sticky", left: 0, background: C.paper }}>Collaborateur</th>
                  {dates.map((d) => (
                    <th key={dateKey(d)} style={{ textAlign: "center", padding: "8px 4px", color: C.textMid, fontWeight: 700, fontSize: 10.5, minWidth: 34 }}>
                      {d.getDate()}<div style={{ fontSize: 9, fontWeight: 500 }}>{DAY_LABELS[d.getDay()]}</div>
                    </th>
                  ))}
                  <th style={{ textAlign: "center", padding: "8px 10px", color: C.textMid, fontWeight: 700, fontSize: 11 }}>Total</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((e) => (
                  <tr key={e.id} style={{ borderTop: `1px solid ${C.paperDim}` }}>
                    <td style={{ padding: "7px 14px", fontWeight: 600, color: C.ink, position: "sticky", left: 0, background: C.card, whiteSpace: "nowrap", fontSize: 12 }}>{e.nom}</td>
                    {e.days.map((d, i) => (
                      <td key={i} style={{ padding: "4px 2px", textAlign: "center" }}>
                        <div title={`${d.status}${d.status === "travail" ? " " + hoursLabel(d.dureeH) : ""}`} style={{
                          width: 22, height: 22, borderRadius: 6, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "center",
                          fontSize: 8.5, fontWeight: 800,
                          background: d.status === "travail" ? C.signalBg : d.status === "repos" ? "#EEF1F6" : C.warnBg,
                          color: d.status === "travail" ? C.signal : d.status === "repos" ? C.textMid : "#8A5A16",
                        }}>
                          {d.status === "travail" ? Math.round(d.dureeH) : d.status === "repos" ? "R" : (d.status[0] || "").toUpperCase()}
                        </div>
                      </td>
                    ))}
                    <td style={{ padding: "7px 10px", textAlign: "center", fontWeight: 800, color: C.ink, fontSize: 12 }}>{hoursLabel(e.totalH)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      )}
    </div>
  );
}
function minToClock(min) {
  if (min == null || isNaN(min)) return "—";
  const h = Math.floor(min / 60) % 24, m = Math.round(min % 60);
  return `${pad2(h)}h${pad2(m)}`;
}

/* ============================================================================
   VUE — COUVERTURE HORAIRE
   ============================================================================ */

function CouvertureView({ dates, coverageMap, selectedDate, setSelectedDate, seuilBas, setSeuilBas, seuilHaut, setSeuilHaut }) {
  const cov = useMemo(() => coverageForDate(coverageMap, selectedDate), [coverageMap, selectedDate]);
  const [openHour, setOpenHour] = useState(null);
  const max = Math.max(...cov.map((c) => c.effectif), 1);

  return (
    <div>
      <SectionTitle icon={Table2} title="Couverture horaire" desc="Effectif présent par tranche d'une heure, pauses exclues"
        right={<DateStepper date={selectedDate} dates={dates} onChange={setSelectedDate} />} />

      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", gap: 24, flexWrap: "wrap", alignItems: "center" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 12.5, fontWeight: 700, color: C.textMid }}>Seuil sous-effectif</span>
            <input type="number" value={seuilBas} min={0} onChange={(e) => setSeuilBas(Number(e.target.value))}
              style={{ width: 60, border: `1px solid ${C.paperDim}`, borderRadius: 8, padding: "6px 8px", fontSize: 13, fontWeight: 700 }} />
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 12.5, fontWeight: 700, color: C.textMid }}>Seuil sur-effectif</span>
            <input type="number" value={seuilHaut} min={0} onChange={(e) => setSeuilHaut(Number(e.target.value))}
              style={{ width: 60, border: `1px solid ${C.paperDim}`, borderRadius: 8, padding: "6px 8px", fontSize: 13, fontWeight: 700 }} />
          </div>
          <div style={{ display: "flex", gap: 14, marginLeft: "auto" }}>
            {[["Sous-effectif", C.bad], ["Effectif optimal", C.good], ["Sur-effectif", C.warn]].map(([l, c]) => (
              <div key={l} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: C.textMid, fontWeight: 600 }}>
                <span style={{ width: 10, height: 10, borderRadius: 3, background: c, display: "inline-block" }} /> {l}
              </div>
            ))}
          </div>
        </div>
      </Card>

      <Card pad={0} style={{ overflow: "hidden" }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ background: C.paper }}>
                <th style={{ textAlign: "left", padding: "10px 16px", color: C.textMid, fontWeight: 700, fontSize: 11.5 }}>Heure</th>
                <th style={{ textAlign: "left", padding: "10px 16px", color: C.textMid, fontWeight: 700, fontSize: 11.5 }}>Effectif présent</th>
                <th style={{ textAlign: "left", padding: "10px 16px", color: C.textMid, fontWeight: 700, fontSize: 11.5, width: "40%" }}>Charge relative</th>
                <th style={{ textAlign: "left", padding: "10px 16px", color: C.textMid, fontWeight: 700, fontSize: 11.5 }}>Statut</th>
              </tr>
            </thead>
            <tbody>
              {cov.map((c) => {
                const tone = coverageTone(c.effectif, seuilBas, seuilHaut);
                const toneColor = tone === "bad" ? C.bad : tone === "warn" ? C.warn : C.good;
                const toneBg = tone === "bad" ? C.badBg : tone === "warn" ? C.warnBg : C.goodBg;
                const toneLabel = tone === "bad" ? "Sous-effectif" : tone === "warn" ? "Sur-effectif" : "Optimal";
                const open = openHour === c.hour;
                return (
                  <React.Fragment key={c.hour}>
                    <tr onClick={() => setOpenHour(open ? null : c.hour)} style={{ borderTop: `1px solid ${C.paperDim}`, cursor: "pointer" }}>
                      <td style={{ padding: "9px 16px", fontWeight: 700, color: C.ink, fontVariantNumeric: "tabular-nums" }}>{pad2(c.hour)}h00</td>
                      <td style={{ padding: "9px 16px", fontWeight: 800, color: C.ink }}>{c.effectif}</td>
                      <td style={{ padding: "9px 16px" }}>
                        <div style={{ background: C.paperDim, borderRadius: 999, height: 8, width: "100%", overflow: "hidden" }}>
                          <div style={{ width: `${(c.effectif / max) * 100}%`, height: "100%", background: toneColor, borderRadius: 999 }} />
                        </div>
                      </td>
                      <td style={{ padding: "9px 16px" }}>
                        <span style={{ background: toneBg, color: toneColor, fontSize: 11.5, fontWeight: 700, padding: "3px 9px", borderRadius: 999 }}>{toneLabel}</span>
                        <ChevronDown size={13} style={{ marginLeft: 8, transform: open ? "rotate(180deg)" : "none", transition: "transform .15s" }} />
                      </td>
                    </tr>
                    {open && (
                      <tr>
                        <td colSpan={4} style={{ padding: "10px 16px 16px", background: C.paper }}>
                          {c.agents.length === 0 ? (
                            <span style={{ fontSize: 12.5, color: C.textMid }}>Aucun agent présent sur cette tranche.</span>
                          ) : (
                            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
                              {c.agents.map((m) => (
                                <span key={m} style={{ background: C.card, border: `1px solid ${C.paperDim}`, borderRadius: 999, padding: "4px 10px", fontSize: 11.5, fontWeight: 600, color: C.ink }}>{m}</span>
                              ))}
                            </div>
                          )}
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

/* ============================================================================
   VUE — GRAPHIQUES
   ============================================================================ */

function heatColor(v, max) {
  if (max <= 0) return C.paperDim;
  const ratio = Math.min(1, v / max);
  // interpolate paperDim -> signal
  const c1 = [237, 241, 247], c2 = [62, 130, 247];
  const r = Math.round(c1[0] + (c2[0] - c1[0]) * ratio);
  const g = Math.round(c1[1] + (c2[1] - c1[1]) * ratio);
  const b = Math.round(c1[2] + (c2[2] - c1[2]) * ratio);
  return `rgb(${r},${g},${b})`;
}

function GraphiquesView({ employees, dates, coverageMap, teams }) {
  const [heatDates, setHeatDates] = useState(0);
  const visibleDates = dates.slice(heatDates, heatDates + 10);

  const lineData = useMemo(() => {
    // average per hour across the whole period
    const sums = new Array(24).fill(0);
    dates.forEach((d) => {
      const cov = coverageForDate(coverageMap, d);
      cov.forEach((c) => { sums[c.hour] += c.effectif; });
    });
    return sums.map((s, h) => ({ heure: `${pad2(h)}h`, effectif: Math.round((s / dates.length) * 10) / 10 }));
  }, [dates, coverageMap]);

  const teamData = useMemo(() => {
    const map = new Map();
    employees.forEach((e) => map.set(e.equipe, (map.get(e.equipe) || 0) + 1));
    return Array.from(map, ([equipe, effectif]) => ({ equipe, effectif })).sort((a, b) => b.effectif - a.effectif);
  }, [employees]);

  const hoursData = useMemo(() =>
    employees.map((e) => ({ nom: e.nom.split(" ")[0] + " " + (e.nom.split(" ")[1]?.[0] || ""), heures: Math.round(e.totalH * 10) / 10 }))
      .sort((a, b) => b.heures - a.heures).slice(0, 18)
  , [employees]);

  const teamColors = [C.signal, C.good, C.warn, "#8B6FE0", "#3EC2C2", C.bad];

  const maxCov = useMemo(() => {
    let m = 0;
    dates.forEach((d) => coverageForDate(coverageMap, d).forEach((c) => { if (c.effectif > m) m = c.effectif; }));
    return m;
  }, [dates, coverageMap]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <Card>
        <SectionTitle icon={Activity} title="Courbe d'effectif horaire moyen" desc="Moyenne du nombre d'agents présents, toutes journées confondues" />
        <ResponsiveContainer width="100%" height={260}>
          <LineChart data={lineData} margin={{ left: -18, right: 8 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.paperDim} vertical={false} />
            <XAxis dataKey="heure" tick={{ fontSize: 11, fill: C.textMid }} interval={1} axisLine={{ stroke: C.paperDim }} tickLine={false} />
            <YAxis tick={{ fontSize: 11, fill: C.textMid }} axisLine={false} tickLine={false} />
            <Tooltip contentStyle={{ borderRadius: 10, border: `1px solid ${C.paperDim}`, fontSize: 12.5 }} />
            <Line type="monotone" dataKey="effectif" stroke={C.signal} strokeWidth={2.5} dot={false} name="Agents présents (moy.)" />
          </LineChart>
        </ResponsiveContainer>
      </Card>

      <Card>
        <SectionTitle icon={LayoutGrid} title="Heatmap de couverture" desc="Lignes = jours · colonnes = heures · intensité = effectif présent"
          right={
            <div style={{ display: "flex", gap: 6 }}>
              <button disabled={heatDates <= 0} onClick={() => setHeatDates(Math.max(0, heatDates - 10))}
                style={{ border: `1px solid ${C.paperDim}`, background: C.card, borderRadius: 8, padding: 6, cursor: "pointer", opacity: heatDates <= 0 ? 0.3 : 1 }}>
                <ChevronLeft size={14} />
              </button>
              <button disabled={heatDates + 10 >= dates.length} onClick={() => setHeatDates(Math.min(dates.length - 1, heatDates + 10))}
                style={{ border: `1px solid ${C.paperDim}`, background: C.card, borderRadius: 8, padding: 6, cursor: "pointer", opacity: heatDates + 10 >= dates.length ? 0.3 : 1 }}>
                <ChevronRight size={14} />
              </button>
            </div>
          } />
        <div style={{ overflowX: "auto" }}>
          <table style={{ borderCollapse: "collapse", fontSize: 10.5 }}>
            <thead>
              <tr>
                <th style={{ padding: "3px 8px", textAlign: "left", color: C.textMid, fontSize: 10.5, position: "sticky", left: 0, background: C.card }}></th>
                {Array.from({ length: 24 }, (_, h) => (
                  <th key={h} style={{ padding: "3px 2px", color: C.textMid, fontWeight: 600, fontSize: 9.5, width: 26 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {visibleDates.map((d) => {
                const cov = coverageForDate(coverageMap, d);
                return (
                  <tr key={dateKey(d)}>
                    <td style={{ padding: "2px 8px", fontWeight: 700, color: C.ink, whiteSpace: "nowrap", fontSize: 11, position: "sticky", left: 0, background: C.card }}>{fmtDateShort(d)}</td>
                    {cov.map((c) => (
                      <td key={c.hour} title={`${c.effectif} agents à ${pad2(c.hour)}h`} style={{
                        width: 26, height: 22, background: heatColor(c.effectif, maxCov), textAlign: "center",
                        color: c.effectif / maxCov > 0.55 ? "#fff" : C.textMid, fontWeight: 700, fontSize: 9.5, border: "2px solid #fff",
                      }}>{c.effectif || ""}</td>
                    ))}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <Card>
          <SectionTitle icon={Building2} title="Répartition par équipe" desc="Nombre de collaborateurs" />
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={teamData} margin={{ left: -18, right: 8 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.paperDim} vertical={false} />
              <XAxis dataKey="equipe" tick={{ fontSize: 11.5, fill: C.textMid, fontWeight: 700 }} axisLine={{ stroke: C.paperDim }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: C.textMid }} axisLine={false} tickLine={false} allowDecimals={false} />
              <Tooltip contentStyle={{ borderRadius: 10, border: `1px solid ${C.paperDim}`, fontSize: 12.5 }} />
              <Bar dataKey="effectif" radius={[8, 8, 0, 0]}>
                {teamData.map((_, i) => <Cell key={i} fill={teamColors[i % teamColors.length]} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>

        <Card>
          <SectionTitle icon={Clock} title="Heures travaillées par collaborateur" desc="Top 18 · cumul sur la période" />
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={hoursData} layout="vertical" margin={{ left: 10, right: 20 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.paperDim} horizontal={false} />
              <XAxis type="number" tick={{ fontSize: 10.5, fill: C.textMid }} axisLine={false} tickLine={false} />
              <YAxis type="category" dataKey="nom" tick={{ fontSize: 10, fill: C.textMid }} axisLine={false} tickLine={false} width={90} />
              <Tooltip contentStyle={{ borderRadius: 10, border: `1px solid ${C.paperDim}`, fontSize: 12.5 }} />
              <Bar dataKey="heures" radius={[0, 6, 6, 0]} fill={C.signal}>
                {hoursData.map((d, i) => <Cell key={i} fill={d.heures > 48 ? C.bad : C.signal} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </Card>
      </div>
    </div>
  );
}

/* ============================================================================
   VUE — ALERTES
   ============================================================================ */

const ALERT_META = {
  surcharge: { label: "Dépassement 48h/semaine", icon: Clock },
  journee_longue: { label: "Journée > 12h", icon: AlertTriangle },
  sans_pause: { label: "Absence de pause", icon: AlertTriangle },
  sous_effectif: { label: "Effectif insuffisant", icon: Users },
};

function AlertesView({ alerts }) {
  const [filterType, setFilterType] = useState("");
  const grouped = useMemo(() => {
    const m = new Map();
    alerts.forEach((a) => m.set(a.type, (m.get(a.type) || 0) + 1));
    return m;
  }, [alerts]);
  const filtered = filterType ? alerts.filter((a) => a.type === filterType) : alerts;

  return (
    <div>
      <SectionTitle icon={Bell} title="Alertes d'exploitation" desc={`${alerts.length} alerte(s) détectée(s) sur la période`} />

      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 18 }}>
        <button onClick={() => setFilterType("")} style={{
          border: `1px solid ${!filterType ? C.ink : C.paperDim}`, background: !filterType ? C.ink : C.card,
          color: !filterType ? "#fff" : C.ink, borderRadius: 10, padding: "8px 14px", fontSize: 12.5, fontWeight: 700, cursor: "pointer" }}>
          Toutes ({alerts.length})
        </button>
        {Object.entries(ALERT_META).map(([key, meta]) => (
          <button key={key} onClick={() => setFilterType(key)} style={{
            border: `1px solid ${filterType === key ? C.ink : C.paperDim}`, background: filterType === key ? C.ink : C.card,
            color: filterType === key ? "#fff" : C.ink, borderRadius: 10, padding: "8px 14px", fontSize: 12.5, fontWeight: 700, cursor: "pointer",
            display: "flex", alignItems: "center", gap: 6 }}>
            <meta.icon size={13} /> {meta.label} ({grouped.get(key) || 0})
          </button>
        ))}
      </div>

      <Card pad={0} style={{ overflow: "hidden" }}>
        {filtered.length === 0 ? (
          <div style={{ padding: "40px 0", textAlign: "center", color: C.textMid, fontSize: 13 }}>
            <Check size={22} color={C.good} style={{ marginBottom: 8 }} /><br />Aucune alerte sur cette catégorie.
          </div>
        ) : (
          <div>
            {filtered.map((a, i) => {
              const meta = ALERT_META[a.type];
              const toneColor = a.sev === "bad" ? C.bad : C.warn;
              const toneBg = a.sev === "bad" ? C.badBg : C.warnBg;
              return (
                <div key={i} style={{ display: "flex", alignItems: "center", gap: 12, padding: "13px 18px", borderTop: i ? `1px solid ${C.paperDim}` : "none" }}>
                  <div style={{ width: 32, height: 32, borderRadius: 9, background: toneBg, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                    {meta && <meta.icon size={15} color={toneColor} />}
                  </div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 13, color: C.ink, fontWeight: 600 }}>{a.msg}</div>
                    <div style={{ fontSize: 11, color: C.textMid, marginTop: 1 }}>{meta?.label}</div>
                  </div>
                  <span style={{ background: toneBg, color: toneColor, fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 999 }}>
                    {a.sev === "bad" ? "Critique" : "Attention"}
                  </span>
                </div>
              );
            })}
          </div>
        )}
      </Card>
    </div>
  );
}

/* ============================================================================
   VUE — ANALYSE DE COUVERTURE
   ============================================================================ */

function AnalyseView({ dates, coverageMap, selectedDate, setSelectedDate }) {
  const [queryHour, setQueryHour] = useState(10);
  const [target, setTarget] = useState(15);

  const cov = useMemo(() => coverageForDate(coverageMap, selectedDate), [coverageMap, selectedDate]);
  const atHour = cov.find((c) => c.hour === queryHour) || { effectif: 0, agents: [] };
  const peak = cov.reduce((a, b) => (b.effectif > a.effectif ? b : a), cov[0]);
  const low = cov.reduce((a, b) => (b.effectif < a.effectif ? b : a), cov[0]);
  const gap = Math.max(0, target - atHour.effectif);
  const chartData = cov.map((c) => ({ heure: `${pad2(c.hour)}h`, effectif: c.effectif }));

  const gapsAllDay = cov.map((c) => ({ heure: `${pad2(c.hour)}h`, manque: Math.max(0, target - c.effectif) }));
  const totalGapHours = gapsAllDay.reduce((s, g) => s + g.manque, 0);

  return (
    <div>
      <SectionTitle icon={Compass} title="Analyse de couverture" desc="Interrogez la couverture horaire du jour sélectionné"
        right={<DateStepper date={selectedDate} dates={dates} onChange={setSelectedDate} />} />

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
        <Card>
          <div style={{ fontSize: 12.5, fontWeight: 700, color: C.textMid, marginBottom: 10 }}>Combien de personnes sont présentes à une heure donnée ?</div>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
            <span style={{ fontSize: 13, color: C.textMid }}>Heure :</span>
            <input type="range" min={0} max={23} value={queryHour} onChange={(e) => setQueryHour(Number(e.target.value))} style={{ flex: 1 }} />
            <span style={{ fontSize: 15, fontWeight: 800, color: C.ink, minWidth: 38 }}>{pad2(queryHour)}h</span>
          </div>
          <div style={{ display: "flex", alignItems: "baseline", gap: 8 }}>
            <span style={{ fontSize: 36, fontWeight: 800, color: C.signal }}>{atHour.effectif}</span>
            <span style={{ fontSize: 13, color: C.textMid }}>agent(s) présent(s) à {pad2(queryHour)}h</span>
          </div>
          {atHour.agents.length > 0 && (
            <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginTop: 10 }}>
              {atHour.agents.map((m) => (
                <span key={m} style={{ background: C.paper, borderRadius: 999, padding: "3px 9px", fontSize: 11, fontWeight: 600, color: C.textMid }}>{m}</span>
              ))}
            </div>
          )}
        </Card>

        <Card>
          <div style={{ fontSize: 12.5, fontWeight: 700, color: C.textMid, marginBottom: 10 }}>Heure la plus chargée / la moins couverte</div>
          <div style={{ display: "flex", gap: 14 }}>
            <div style={{ flex: 1, background: C.goodBg, borderRadius: 12, padding: 14 }}>
              <div style={{ fontSize: 11.5, fontWeight: 700, color: C.good }}>PLUS CHARGÉE</div>
              <div style={{ fontSize: 26, fontWeight: 800, color: C.ink, marginTop: 4 }}>{pad2(peak.hour)}h</div>
              <div style={{ fontSize: 12, color: C.textMid }}>{peak.effectif} agents présents</div>
            </div>
            <div style={{ flex: 1, background: C.badBg, borderRadius: 12, padding: 14 }}>
              <div style={{ fontSize: 11.5, fontWeight: 700, color: C.bad }}>MOINS COUVERTE</div>
              <div style={{ fontSize: 26, fontWeight: 800, color: C.ink, marginTop: 4 }}>{pad2(low.hour)}h</div>
              <div style={{ fontSize: 12, color: C.textMid }}>{low.effectif} agent(s) présent(s)</div>
            </div>
          </div>
        </Card>
      </div>

      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10, marginBottom: 10 }}>
          <div style={{ fontSize: 12.5, fontWeight: 700, color: C.textMid }}>Quel est le besoin supplémentaire pour atteindre un effectif cible ?</div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 12.5, color: C.textMid }}>Effectif cible / heure :</span>
            <input type="number" value={target} min={0} onChange={(e) => setTarget(Number(e.target.value))}
              style={{ width: 60, border: `1px solid ${C.paperDim}`, borderRadius: 8, padding: "6px 8px", fontSize: 13, fontWeight: 700 }} />
          </div>
        </div>
        <div style={{ display: "flex", gap: 20, marginBottom: 12, flexWrap: "wrap" }}>
          <div>
            <span style={{ fontSize: 24, fontWeight: 800, color: gap > 0 ? C.bad : C.good }}>{gap > 0 ? `+${gap}` : "0"}</span>
            <span style={{ fontSize: 12, color: C.textMid, marginLeft: 6 }}>agent(s) à ajouter à {pad2(queryHour)}h pour atteindre {target}</span>
          </div>
          <div>
            <span style={{ fontSize: 24, fontWeight: 800, color: C.ink }}>{totalGapHours}</span>
            <span style={{ fontSize: 12, color: C.textMid, marginLeft: 6 }}>agent-heures manquants sur la journée entière</span>
          </div>
        </div>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={gapsAllDay} margin={{ left: -18, right: 8 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.paperDim} vertical={false} />
            <XAxis dataKey="heure" tick={{ fontSize: 10.5, fill: C.textMid }} interval={1} axisLine={{ stroke: C.paperDim }} tickLine={false} />
            <YAxis tick={{ fontSize: 11, fill: C.textMid }} axisLine={false} tickLine={false} allowDecimals={false} />
            <Tooltip contentStyle={{ borderRadius: 10, border: `1px solid ${C.paperDim}`, fontSize: 12.5 }} />
            <Bar dataKey="manque" radius={[6, 6, 0, 0]} fill={C.bad} name="Manque d'effectif" />
          </BarChart>
        </ResponsiveContainer>
      </Card>

      <Card>
        <div style={{ fontSize: 12.5, fontWeight: 700, color: C.textMid, marginBottom: 10 }}>Effectif vs objectif — vue complète de la journée</div>
        <ResponsiveContainer width="100%" height={240}>
          <AreaChart data={chartData} margin={{ left: -18, right: 8 }}>
            <defs>
              <linearGradient id="fillAnalyse" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor={C.signal} stopOpacity={0.35} />
                <stop offset="100%" stopColor={C.signal} stopOpacity={0.02} />
              </linearGradient>
            </defs>
            <CartesianGrid strokeDasharray="3 3" stroke={C.paperDim} vertical={false} />
            <XAxis dataKey="heure" tick={{ fontSize: 10.5, fill: C.textMid }} interval={1} axisLine={{ stroke: C.paperDim }} tickLine={false} />
            <YAxis tick={{ fontSize: 11, fill: C.textMid }} axisLine={false} tickLine={false} allowDecimals={false} />
            <Tooltip contentStyle={{ borderRadius: 10, border: `1px solid ${C.paperDim}`, fontSize: 12.5 }} />
            <ReferenceLine y={target} stroke={C.bad} strokeDasharray="5 4" label={{ value: "Cible", fontSize: 11, fill: C.bad }} />
            <Area type="monotone" dataKey="effectif" stroke={C.signal} strokeWidth={2.5} fill="url(#fillAnalyse)" name="Effectif présent" />
          </AreaChart>
        </ResponsiveContainer>
      </Card>
    </div>
  );
}

/* ============================================================================
   APP ROOT
   ============================================================================ */

function generateDemoData() {
  // Small synthetic dataset shaped like the real file, for a quick preview.
  const teams = ["EER", "TEL BAQ", "CHAT BAQ"];
  const names = [
    "RAKOTO Andry", "RASOA Voahangy", "RABE Solofo", "RAZAFY Nirina", "RANDRIA Fetra",
    "ANDRY Hery", "NIRINA Tojo", "MIORA Lala", "TAHIANA Fy", "HASINA Njara",
    "FANOMEZANTSOA Lova", "MAMPIONONA Rija", "VOLATIANA Sitraka", "ANDRIAMBOLA Hanta", "RIVO Tsanta",
  ];
  const shifts = [
    [6, 15], [8, 17], [9, 18], [12, 21], [19, 6], [14, 22], [7, 16],
  ];
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const dates = Array.from({ length: 7 }, (_, i) => { const d = new Date(today); d.setDate(d.getDate() - 3 + i); return d; });

  const employees = names.map((nom, i) => {
    const equipe = teams[i % teams.length];
    const days = dates.map((date, di) => {
      const isRepos = (i + di) % 6 === 5;
      if (isRepos) return { date, dateKey: dateKey(date), status: "repos", startMinutes: null, endMinutes: null, breakStart: null, breakLenMin: 0, dureeH: 0 };
      const [sh, eh] = shifts[(i + di) % shifts.length];
      const startMinutes = sh * 60;
      let endMinutes = eh * 60; if (endMinutes <= startMinutes) endMinutes += 1440;
      const shiftLen = endMinutes - startMinutes;
      const breakLenMin = shiftLen > 6 * 60 ? 60 : 0;
      const breakStart = startMinutes + shiftLen / 2 - breakLenMin / 2;
      const dureeH = (shiftLen - breakLenMin) / 60;
      return { date, dateKey: dateKey(date), status: "travail", startMinutes, endMinutes, breakStart, breakLenMin, dureeH };
    });
    const totalH = days.reduce((s, d) => s + d.dureeH, 0);
    return {
      id: `DEMO${i}`, matricule: `CN0${1000 + i}`, equipe, nom, days,
      reposDeclares: days.filter((d) => d.status === "repos").length, totalH,
      joursTravailles: days.filter((d) => d.status === "travail").length,
      joursRepos: days.filter((d) => d.status === "repos").length,
    };
  });

  return { employees, dates, teams: Array.from(new Set(employees.map((e) => e.equipe))).sort(), warnings: [], sheetName: "Démonstration" };
}

export default function App() {
  const [stage, setStage] = useState("import"); // import | preview | app
  const [parsed, setParsed] = useState(null);
  const [fileName, setFileName] = useState("");
  const [view, setView] = useState("dashboard");
  const [selectedDate, setSelectedDate] = useState(null);
  const [seuilBas, setSeuilBas] = useState(8);
  const [seuilHaut, setSeuilHaut] = useState(35);

  const handleParsed = useCallback((result, name) => { setParsed(result); setFileName(name); setStage("preview"); }, []);
  const confirmImport = useCallback(() => {
    setSelectedDate(parsed.dates[0]);
    setStage("app");
  }, [parsed]);
  const loadDemo = useCallback(() => {
    const demo = generateDemoData();
    setParsed(demo); setFileName("Jeu de démonstration"); setStage("preview");
  }, []);

  const coverageMap = useMemo(() => (parsed ? buildHourlyCoverage(parsed.employees) : new Map()), [parsed]);
  const alerts = useMemo(() => (parsed ? computeAlerts(parsed.employees, coverageMap, parsed.dates, seuilBas) : []), [parsed, coverageMap, seuilBas]);

  const handleExport = useCallback(() => {
    if (!parsed) return;
    const choice = window.confirm("Exporter en Excel (OK) ou en CSV (Annuler) ? Pour un export PDF, utilisez l'impression du navigateur après export.");
    const perAgentRows = parsed.employees.map((e) => ({
      Matricule: e.matricule, Équipe: e.equipe, "Nom et prénom": e.nom,
      "Jours travaillés": e.joursTravailles, "Jours de repos": e.joursRepos,
      "Total heures semaine": Math.round(e.totalH * 100) / 100,
    }));
    const coverageRows = [];
    parsed.dates.forEach((d) => {
      coverageForDate(coverageMap, d).forEach((c) => {
        coverageRows.push({ Date: fmtDateShort(d), Heure: `${pad2(c.hour)}h00`, "Effectif présent": c.effectif });
      });
    });
    if (choice) {
      exportXLSX([
        { name: "Effectifs", rows: perAgentRows },
        { name: "Couverture horaire", rows: coverageRows },
      ], "pilotage-effectifs.xlsx");
    } else {
      exportCSV(perAgentRows, ["Matricule", "Équipe", "Nom et prénom", "Jours travaillés", "Jours de repos", "Total heures semaine"], "effectifs.csv");
    }
  }, [parsed, coverageMap]);

  if (stage === "import") {
    return (
      <div style={{ minHeight: "100vh", background: `radial-gradient(1200px 600px at 20% -10%, #14295080, transparent), ${C.ink}`, fontFamily: "'Inter','Segoe UI',system-ui,sans-serif" }}>
        <ImportScreen onParsed={handleParsed} loadDemo={loadDemo} />
      </div>
    );
  }
  if (stage === "preview") {
    return (
      <div style={{ minHeight: "100vh", background: `radial-gradient(1200px 600px at 20% -10%, #14295080, transparent), ${C.ink}`, fontFamily: "'Inter','Segoe UI',system-ui,sans-serif" }}>
        <PreviewScreen result={parsed} fileName={fileName} onConfirm={confirmImport} onCancel={() => setStage("import")} />
      </div>
    );
  }

  const filteredForCoverage = parsed.employees;

  return (
    <div style={{ minHeight: "100vh", background: C.paper, fontFamily: "'Inter','Segoe UI',system-ui,sans-serif" }}>
      <TopBar view={view} setView={setView} alertCount={alerts.length} fileName={fileName}
        onReimport={() => { setStage("import"); setParsed(null); }} onExport={handleExport} />
      <div style={{ maxWidth: 1360, margin: "0 auto", padding: "22px 22px 60px" }}>
        {view === "dashboard" && (
          <DashboardView employees={parsed.employees} dates={parsed.dates} coverageMap={coverageMap}
            selectedDate={selectedDate} setSelectedDate={setSelectedDate} alerts={alerts} />
        )}
        {view === "planning" && (
          <PlanningView employees={parsed.employees} dates={parsed.dates} teams={parsed.teams} />
        )}
        {view === "couverture" && (
          <CouvertureView dates={parsed.dates} coverageMap={coverageMap} selectedDate={selectedDate} setSelectedDate={setSelectedDate}
            seuilBas={seuilBas} setSeuilBas={setSeuilBas} seuilHaut={seuilHaut} setSeuilHaut={setSeuilHaut} />
        )}
        {view === "graphiques" && (
          <GraphiquesView employees={parsed.employees} dates={parsed.dates} coverageMap={coverageMap} teams={parsed.teams} />
        )}
        {view === "alertes" && <AlertesView alerts={alerts} />}
        {view === "analyse" && (
          <AnalyseView dates={parsed.dates} coverageMap={coverageMap} selectedDate={selectedDate} setSelectedDate={setSelectedDate} />
        )}
      </div>
    </div>
  );
}
