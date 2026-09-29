import React, { useState, useMemo, useCallback, useEffect } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  PieChart, Pie, Cell, ResponsiveContainer, Tooltip,
  BarChart, Bar, XAxis, YAxis, CartesianGrid,
  ComposedChart, Line, Legend, ReferenceLine,
} from "recharts";

/* ------------------------------------------------------------------ */
/*  נתוני דמו — מהדוח (מתעדכנים אוטומטית בעת חיבור הגיליון)             */
/* ------------------------------------------------------------------ */

const DEMO = {
  meta: { budget: 677000, spent: 263033, pace: 41.9, utilPct: 38.9, month: "חודש יולי" },
  totalsRow: { budget: 637000, spent: 198993 },
  harvard: {
    platforms: [
      { name: "Linkedin", budget: 40000, spent: 23543.5, gross: 195, cpl: 120.7, quality: 9, cpql: 2615.9, qPct: 4.6, inProc: 70, inProcPct: 35.9 },
    ],
    total: { name: "Total", budget: 40000, spent: 23543.5, gross: 195, cpl: 120.7, quality: 9, cpql: 2615.9, qPct: 4.6, inProc: 70, inProcPct: 35.9 },
  },
  summary: {
    actual: { gross: 2114, quality: 136, qualityPct: 6, cpl: 94, cpql: 1597, inProcess: 699 },
    target: { gross: 3931, quality: 476, qualityPct: 12, cpl: 172, cpql: 1422 },
  },
  platforms: [
    { name: "PMAX",               budget: 190000, spent: 66483, gross: 345, cpl: 193, quality: 20, cpql: 3324,  qPct: 5.8,  inProc: 144, inProcPct: 42 },
    { name: "PMAX Degree",        budget: 170000, spent: 45426, gross: 192, cpl: 237, quality: 19, cpql: 2391,  qPct: 9.9,  inProc: 85,  inProcPct: 44 },
    { name: "חיפוש",              budget: 70000,  spent: 25562, gross: 795, cpl: 32,  quality: 67, cpql: 382,   qPct: 8.4,  inProc: 237, inProcPct: 30, drill: true },
    { name: "Demandgen",          budget: 60000,  spent: 15522, gross: 23,  cpl: 675, quality: 1,  cpql: 15522, qPct: 4.3,  inProc: 11,  inProcPct: 48 },
    { name: "פייסבוק תואר ראשון", budget: 55000,  spent: 20455, gross: 194, cpl: 105, quality: 7,  cpql: 2922,  qPct: 3.6,  inProc: 75,  inProcPct: 39 },
    { name: "פייסבוק תואר שני",   budget: 45000,  spent: 19808, gross: 115, cpl: 172, quality: 4,  cpql: 4952,  qPct: 3.5,  inProc: 51,  inProcPct: 44 },
    { name: "פייסבוק משאבי אנוש", budget: 5000,   spent: 1200,  gross: 16,  cpl: 75,  quality: 1,  cpql: 1200,  qPct: 6.3,  inProc: 1,   inProcPct: 6 },
    { name: "UGC - Facebook",     budget: 40000,  spent: 5276,  gross: 19,  cpl: 278, quality: 2,  cpql: 2638,  qPct: 10.5, inProc: 8,   inProcPct: 42 },
    { name: "Linkedin - HR2",     budget: 7000,   spent: 772,   gross: 2,   cpl: 386, quality: 0,  cpql: null,  qPct: 0,    inProc: 0,   inProcPct: 0 },
    { name: "Linkedin - Law2",    budget: 30000,  spent: 8392,  gross: 238, cpl: 35,  quality: 4,  cpql: 2098,  qPct: 1.7,  inProc: null, inProcPct: null },
    { name: "Tiktok - LP+LG",     budget: 25000,  spent: 0,     gross: null, cpl: null, quality: null, cpql: null, qPct: null, inProc: null, inProcPct: null },
    { name: "Tiktok - LP - NEW",  budget: 0,      spent: 5619,  gross: 8,   cpl: 702, quality: 1,  cpql: 5619,  qPct: 12.5, inProc: 5,   inProcPct: 63 },
  ],
  campaigns: [
    { name: "מותג",                 spent: 3359,  share: 13, gross: 201, cpl: 17,  quality: 28, cpql: 120,  qPct: 13.9, inProc: 69, inProcPct: 34 },
    { name: "מותג+תואר",            spent: 10006, share: 39, gross: 65,  cpl: 154, quality: 13, cpql: 770,  qPct: 20.0, inProc: 24, inProcPct: 37 },
    { name: "גנרי ראשון",           spent: 10537, share: 41, gross: 69,  cpl: 153, quality: 4,  cpql: 2634, qPct: 5.8,  inProc: 24, inProcPct: 35 },
    { name: "משאבי אנוש תואר שני",  spent: 1660,  share: 6,  gross: 3,   cpl: 553, quality: 0,  cpql: null, qPct: 0,    inProc: 3,  inProcPct: 100 },
    { name: "גנרי - RLSA",          spent: 10584, share: 41, gross: 25,  cpl: 423, quality: 3,  cpql: 3528, qPct: 12.0, inProc: 8,  inProcPct: 32 },
    { name: "Direct (unknown leads)", spent: null, share: null, gross: 12, cpl: null, quality: 1, cpql: null, qPct: 8.3, inProc: 6, inProcPct: null },
    { name: "GMB",                  spent: null,  share: null, gross: 378, cpl: null, quality: 12, cpql: null, qPct: 3.2, inProc: 59, inProcPct: null },
    { name: "callbox",              spent: null,  share: null, gross: 36,  cpl: null, quality: 6,  cpql: null, qPct: 16.7, inProc: 7, inProcPct: null },
  ],
};

/* נתוני דוגמה לביצועי יממה — מוצגים רק עד שנשמרות שתי תמונות מצב אמיתיות */
const DEMO_DAY_HARVARD = [
  { name: "Linkedin", type: "כלי", spent: 1850, gross: 12, quality: 1 },
];
/* נתוני דוגמה לטאב "מבט על" — מוצגים רק עד שנצברים שני ימים בארכיון */
/* נתוני דוגמה ל-YoY — מיוצרים כדי להראות איך הטאב נראה לפני שנטען גיליון אמיתי.
   מבנה זהה לפלט של parseYoY: { year: { "MM-DD": {digital, site, total} } } */
const DEMO_YOY = (() => {
  const now = new Date(); const prevY = now.getFullYear() - 1; const mo = now.getMonth() + 1;
  const mm = String(mo).padStart(2, "0");
  const daysInM = new Date(prevY, mo, 0).getDate();
  const leads = {}, budget = {};
  const seed = [22, 28, 25, 31, 14, 9, 12, 34, 27, 24, 20, 30, 13, 8, 23, 38, 30, 26, 18, 12, 9, 33, 41, 29, 22, 19, 15, 27, 36, 30, 24];
  for (let d = 1; d <= daysInM; d++) {
    const dt = new Date(prevY, mo - 1, d).getDay();
    let grossBase = seed[(d - 1) % seed.length];
    if (dt === 5 || dt === 6) grossBase = Math.max(2, Math.round(grossBase * 0.4));
    const gDig = Math.round(grossBase * 0.7), gSite = grossBase - Math.round(grossBase * 0.7);
    const qRate = 0.12 + (d % 5) * 0.01; /* ~12-16% איכות */
    const qDig = Math.round(gDig * qRate), qSite = Math.round(gSite * (qRate + 0.03));
    const key = `${mm}-${String(d).padStart(2, "0")}`;
    leads[key] = { gross: { digital: gDig, site: gSite, total: gDig + gSite }, quality: { digital: qDig, site: qSite, total: qDig + qSite } };
    /* תקציב ~ ₪150-220 לליד ברוטו דיגיטלי */
    budget[key] = { digital: Math.round(gDig * (170 + (d % 7) * 8)), site: 0, total: Math.round(gDig * (170 + (d % 7) * 8)) };
  }
  return { leads: { [prevY]: leads }, budget: { [prevY]: budget } };
})();

const DEMO_TREND = [
  { label: "05.07", "תקציב יומי": 11200, "איכותיים תארים": 5, "תקציב הארווארד": 1150, "איכותיים הארווארד": 0, "לידים ברוטו אתר": 31, "איכותיים אתר": 4 },
  { label: "06.07", "תקציב יומי": 12800, "איכותיים תארים": 7, "תקציב הארווארד": 1240, "איכותיים הארווארד": 1, "לידים ברוטו אתר": 38, "איכותיים אתר": 6 },
  { label: "07.07", "תקציב יומי": 9800,  "איכותיים תארים": 4, "תקציב הארווארד": 1090, "איכותיים הארווארד": 0, "לידים ברוטו אתר": 26, "איכותיים אתר": 3 },
  { label: "08.07", "תקציב יומי": 13600, "איכותיים תארים": 9, "תקציב הארווארד": 1310, "איכותיים הארווארד": 1, "לידים ברוטו אתר": 41, "איכותיים אתר": 7 },
  { label: "09.07", "תקציב יומי": 12100, "איכותיים תארים": 6, "תקציב הארווארד": 1180, "איכותיים הארווארד": 0, "לידים ברוטו אתר": 35, "איכותיים אתר": 5 },
  { label: "10.07", "תקציב יומי": 8400,  "איכותיים תארים": 3, "תקציב הארווארד": 960,  "איכותיים הארווארד": 0, "לידים ברוטו אתר": 22, "איכותיים אתר": 2 },
  { label: "11.07", "תקציב יומי": 7900,  "איכותיים תארים": 3, "תקציב הארווארד": 890,  "איכותיים הארווארד": 1, "לידים ברוטו אתר": 19, "איכותיים אתר": 2 },
  { label: "12.07", "תקציב יומי": 14200, "איכותיים תארים": 8, "תקציב הארווארד": 1380, "איכותיים הארווארד": 1, "לידים ברוטו אתר": 44, "איכותיים אתר": 6 },
  { label: "13.07", "תקציב יומי": 13100, "איכותיים תארים": 7, "תקציב הארווארד": 1290, "איכותיים הארווארד": 0, "לידים ברוטו אתר": 39, "איכותיים אתר": 5 },
  { label: "14.07", "תקציב יומי": 12500, "איכותיים תארים": 6, "תקציב הארווארד": 1220, "איכותיים הארווארד": 1, "לידים ברוטו אתר": 36, "איכותיים אתר": 4 },
  { label: "15.07", "תקציב יומי": 11800, "איכותיים תארים": 5, "תקציב הארווארד": 1160, "איכותיים הארווארד": 0, "לידים ברוטו אתר": 33, "איכותיים אתר": 5 },
  { label: "16.07", "תקציב יומי": 13900, "איכותיים תארים": 8, "תקציב הארווארד": 1340, "איכותיים הארווארד": 1, "לידים ברוטו אתר": 42, "איכותיים אתר": 7 },
  { label: "17.07", "תקציב יומי": 10600, "איכותיים תארים": 4, "תקציב הארווארד": 1050, "איכותיים הארווארד": 0, "לידים ברוטו אתר": 28, "איכותיים אתר": 3 },
  { label: "18.07", "תקציב יומי": 12300, "איכותיים תארים": 7, "תקציב הארווארד": 1210, "איכותיים הארווארד": 1, "לידים ברוטו אתר": 37, "איכותיים אתר": 6 },
];

/* נתוני דוגמה לבורר הקמפיינים בגרף האיכות (מבט על) — מוצגים עד שנצבר ארכיון אמיתי */
const DEMO_CAMP_QUALITY = {
  "PMAX": { q: [2,3,1,3,2,1,1,3,2,2,2,3,1,2], cpl: [2900,2600,3400,2400,2800,3600,3500,2300,2500,2700,2800,2300,3300,2600] },
  "PMAX Degree": { q: [1,2,1,2,1,0,1,2,2,1,1,2,1,2], cpl: [3100,2500,3300,2600,3200,null,3000,2500,2600,3100,3200,2600,3200,2600] },
  "חיפוש": { q: [2,2,1,3,2,1,1,2,2,2,1,2,1,2], cpl: [420,410,520,340,430,600,590,360,380,420,540,360,560,410] },
  "פייסבוק תואר ראשון": { q: [0,1,0,1,1,0,0,1,1,1,0,1,0,1], cpl: [null,2900,null,2700,2900,null,null,2700,2900,2800,null,2700,null,2900] },
  "פייסבוק תואר שני": { q: [0,0,1,0,0,1,0,1,0,0,1,0,1,0], cpl: [null,null,4900,null,null,5100,null,4800,null,null,5000,null,4900,null] },
  "חיפוש · מותג": { q: [3,4,2,4,3,2,1,4,3,3,2,4,2,3], cpl: [110,100,140,95,115,150,175,95,120,110,140,95,145,115] },
  "חיפוש · גנרי ראשון": { q: [1,1,0,2,1,0,0,2,1,1,0,2,0,1], cpl: [2600,2500,null,2300,2500,null,null,2300,2600,2500,null,2300,null,2600] },
  "חיפוש · גנרי - RLSA": { q: [0,1,1,1,0,0,1,1,1,0,1,1,0,1], cpl: [null,3500,3400,3300,null,null,3600,3300,3400,null,3500,3300,null,3500] },
};

const DEMO_DAY = [
  { name: "PMAX",               type: "כלי",   spent: 2450, gross: 14, quality: 1 },
  { name: "חיפוש",              type: "כלי",   spent: 980,  gross: 31, quality: 3 },
  { name: "PMAX Degree",        type: "כלי",   spent: 1720, gross: 7,  quality: 1 },
  { name: "פייסבוק תואר ראשון", type: "כלי",   spent: 810,  gross: 9,  quality: 0 },
  { name: "Demandgen",          type: "כלי",   spent: 640,  gross: 1,  quality: 0 },
  { name: "מותג",               type: "חיפוש", spent: 130,  gross: 8,  quality: 1 },
  { name: "גנרי ראשון",         type: "חיפוש", spent: 420,  gross: 3,  quality: 0 },
];

/* ------------------------------------------------------------------ */
/*  עזרי פורמט, פרסינג ואחסון                                           */
/* ------------------------------------------------------------------ */

const toNum = (v) => {
  if (v === null || v === undefined) return null;
  const s = String(v).replace(/[₪%\s"]/g, "").replace(/,/g, "").trim();
  if (s === "" || s === "-" || s.includes("#DIV") || s.includes("#REF") || s.toUpperCase().includes("N/A")) return null;
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : null;
};

/* אחוזים: בגיליון חלק מהאחוזים שמורים כשברים (0.069). אם המקור לא הכיל % והערך <= 1.5 — זה שבר */
const toPct = (raw) => {
  const n = toNum(raw);
  if (n === null) return null;
  if (String(raw).includes("%")) return n;
  return Math.abs(n) <= 1.5 ? n * 100 : n;
};

const nis = (v, dec = 0) =>
  v === null || v === undefined ? "—" : "₪" + Number(v).toLocaleString("he-IL", { maximumFractionDigits: dec });
const num = (v) => (v === null || v === undefined ? "—" : Number(v).toLocaleString("he-IL"));
const pct = (v, dec = 0) => (v === null || v === undefined ? "—" : Number(v).toFixed(dec) + "%");
const heDate = (d) => new Date(d).toLocaleDateString("he-IL", { day: "numeric", month: "long" }) + " · " +
  new Date(d).toLocaleTimeString("he-IL", { hour: "2-digit", minute: "2-digit" });

/* אחסון תלת-שכבתי: window.storage (קלוד) + localStorage (דפדפן/ורסל) + זיכרון (גיבוי לסשן).
   כתיבה נשמרת לכל השכבות במקביל — כך הארכיון עובד בכל סביבה. */
const memStore = {};
const LS_PREFIX = "peres:";
/* ================= אחסון ענן (Supabase) — מקור אמת משותף לכל המכשירים =================
   כל שמירה נכתבת גם לענן וגם מקומית; כל קריאה מנסה קודם את הענן.
   כך כל דפדפן/מכשיר רואה את אותה היסטוריה. אם אין רשת — עובדים מקומית וממשיכים כרגיל. */
const CLOUD = {
  url: "https://ydeahusxyptavxbyemtp.supabase.co",
  key: "sb_publishable_zSyDYgBupJcKu6cuzB7b_g_0Cc3V-lz",
  table: "archive",
};
const cloudHeaders = () => ({
  apikey: CLOUD.key,
  Authorization: "Bearer " + CLOUD.key,
  "Content-Type": "application/json",
});
const cloudFetch = (path, opts = {}) => {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), 12000);
  return fetch(`${CLOUD.url}/rest/v1/${CLOUD.table}${path}`, { ...opts, headers: { ...cloudHeaders(), ...(opts.headers || {}) }, signal: ctrl.signal })
    .finally(() => clearTimeout(t));
};
const cloud = {
  enabled: true, /* מכובה אוטומטית אם הענן לא נגיש, כדי לא להאט את האפליקציה */
  async get(k) {
    if (!this.enabled) return undefined;
    try {
      const res = await cloudFetch(`?key=eq.${encodeURIComponent(k)}&select=value`);
      if (!res.ok) throw new Error(res.status);
      const rows = await res.json();
      return rows.length ? rows[0].value : null;
    } catch { return undefined; } /* undefined = הענן לא ענה; null = המפתח לא קיים בענן */
  },
  async set(k, v) {
    if (!this.enabled) return false;
    try {
      const res = await cloudFetch(`?on_conflict=key`, {
        method: "POST",
        headers: { Prefer: "resolution=merge-duplicates" },
        body: JSON.stringify([{ key: k, value: v, updated_at: new Date().toISOString() }]),
      });
      return res.ok;
    } catch { return false; }
  },
  async list(prefix) {
    if (!this.enabled) return undefined;
    try {
      const res = await cloudFetch(`?key=like.${encodeURIComponent(prefix + "*")}&select=key`);
      if (!res.ok) throw new Error(res.status);
      return (await res.json()).map((r) => r.key);
    } catch { return undefined; }
  },
  async ping() {
    try {
      const res = await cloudFetch(`?select=key&limit=1`);
      this.enabled = res.ok;
    } catch { this.enabled = false; }
    return this.enabled;
  },
};

const store = {
  async get(k) {
    /* קודם הענן — מקור האמת המשותף */
    const c = await cloud.get(k);
    if (c !== undefined && c !== null) {
      try { memStore[k] = JSON.stringify(c); localStorage.setItem(LS_PREFIX + k, JSON.stringify(c)); } catch {}
      return c;
    }
    if (c === null) return null; /* הענן ענה שאין מפתח כזה */
    /* הענן לא זמין — נופלים חזרה למקומי */
    try {
      if (window.storage) {
        const r = await window.storage.get(k);
        if (r && r.value !== undefined && r.value !== null) return JSON.parse(r.value);
      }
    } catch {}
    try {
      const v = localStorage.getItem(LS_PREFIX + k);
      if (v !== null) return JSON.parse(v);
    } catch {}
    return memStore[k] !== undefined ? JSON.parse(memStore[k]) : null;
  },
  async set(k, v) {
    const s = JSON.stringify(v);
    memStore[k] = s;
    try { localStorage.setItem(LS_PREFIX + k, s); } catch {}
    try { if (window.storage) await window.storage.set(k, s); } catch {}
    await cloud.set(k, v); /* דחיפה לענן — כך כל המכשירים רואים */
  },
  async list(prefix) {
    const keys = new Set();
    const c = await cloud.list(prefix);
    if (c !== undefined) c.forEach((k) => keys.add(k));
    Object.keys(memStore).filter((k) => k.startsWith(prefix)).forEach((k) => keys.add(k));
    try {
      if (window.storage) {
        const r = await window.storage.list(prefix);
        ((r && r.keys) || []).forEach((k) => keys.add(k));
      }
    } catch {}
    try {
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k && k.startsWith(LS_PREFIX + prefix)) keys.add(k.slice(LS_PREFIX.length));
      }
    } catch {}
    return [...keys];
  },
  /* דחיפת כל הארכיון המקומי לענן — פעולה חד-פעמית להעברת היסטוריה קיימת */
  async pushAllToCloud() {
    const localKeys = new Set();
    Object.keys(memStore).forEach((k) => localKeys.add(k));
    try {
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k && k.startsWith(LS_PREFIX)) localKeys.add(k.slice(LS_PREFIX.length));
      }
    } catch {}
    let n = 0;
    for (const k of localKeys) {
      let v = null;
      try { const raw = memStore[k] ?? localStorage.getItem(LS_PREFIX + k); if (raw != null) v = JSON.parse(raw); } catch {}
      if (v !== null) { if (await cloud.set(k, v)) n++; }
    }
    return n;
  },
};

/* מפתחות ארכיון יומי: peres:day:YYYY-MM-DD */
const DAY_PREFIX = "peres:day:";
const isoDay = (d) =>
  `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
const dayKey = (iso) => DAY_PREFIX + iso;
const heDayLabel = (iso) => {
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, m - 1, d).toLocaleDateString("he-IL", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
};

function mapHeader(rows, anchor) {
  for (let i = 0; i < rows.length; i++) {
    const idx = rows[i].findIndex((c) => String(c || "").trim() === anchor);
    if (idx !== -1) return { rowIdx: i, cols: rows[i].map((c) => String(c || "").trim()) };
  }
  return null;
}
const findCol = (cols, ...keys) =>
  cols.findIndex((c) => keys.some((k) => String(c).replace(/\s/g, "").includes(k.replace(/\s/g, ""))));

function parseReport(csvText) {
  const rows = Papa.parse(csvText).data;
  const out = { meta: {}, summary: null, platforms: [], campaigns: [], harvard: null, totalsRow: null };

  /* --- בלוק מטא עליון: תקציב, קצב קלנדרי, מימוש, חודש --- */
  for (let i = 0; i < Math.min(rows.length, 16); i++) {
    const label = String(rows[i][0] || "").trim();
    const val = rows[i][1];
    if (label.startsWith("תקציב מדיה") && out.meta.budget === undefined) out.meta.budget = toNum(val);
    else if (label.startsWith("תקציב שמומש") && out.meta.spent === undefined) out.meta.spent = toNum(val);
    else if (label.startsWith("קצב התקדמות קלנדרי")) out.meta.pace = toPct(val);
    else if (label.startsWith("מימוש תקציב")) out.meta.utilPct = toPct(val);
    else if (label === "קמפיין") out.meta.month = String(val || "").trim();
    else if (label.startsWith("מועד דיווח")) out.meta.reportDate = String(val || "").trim();
  }

  /* --- ריכוז עליון: בפועל / יעד --- */
  const sumHdr = rows.findIndex(
    (r) => r.some((c) => String(c).includes("לידים ברוטו")) && r.some((c) => String(c).includes("לידים איכותיים"))
  );
  if (sumHdr !== -1) {
    const cols = rows[sumHdr].map((c) => String(c || "").trim());
    const ci = {
      gross: findCol(cols, "לידים ברוטו"),
      quality: findCol(cols, "לידים איכותיים"),
      qualityPct: findCol(cols, "% לידים איכותיים", "%לידים", "אחוז איכות"),
      cpl: findCol(cols, "עלות לליד ברוטו", "עלות ליד ברוטו"),
      cpql: findCol(cols, "עלות לליד איכותי", "עלות ליד איכותי"),
      inProcess: findCol(cols, "לידים בתהליך"),
      invalid: findCol(cols, "לא תקינים"),
    };
    const pick = (label) => rows.slice(sumHdr + 1, sumHdr + 6).find((r) => r.some((c) => String(c).trim() === label));
    const a = pick("בפועל"), t = pick("יעד");
    const read = (r) =>
      r && {
        gross: toNum(r[ci.gross]), quality: toNum(r[ci.quality]), qualityPct: toPct(r[ci.qualityPct]),
        cpl: ci.cpl !== -1 ? toNum(r[ci.cpl]) : null, cpql: ci.cpql !== -1 ? toNum(r[ci.cpql]) : null,
        inProcess: toNum(r[ci.inProcess]),
        invalid: ci.invalid !== -1 ? toNum(r[ci.invalid]) : null,
      };
    if (a && t) out.summary = { actual: read(a), target: read(t) };
  }

  /* שורה "חיה" = יש בה תקציב, מימוש או לידים. מסנן שורות ריקות / #N/A */
  const alive = (row) => (row.spent || 0) > 0 || (row.budget || 0) > 0 || (row.gross || 0) > 0;

  /* --- טבלת פלטפורמות (BOF) --- */
  const ph = mapHeader(rows, "פלטפורמה");
  if (ph) {
    const c = ph.cols;
    const ci = {
      name: c.indexOf("פלטפורמה"),
      budget: findCol(c, "תקציב"),
      spent: findCol(c, "תקציב שמומש"),
      gross: findCol(c, "לידים ברוטו"),
      cpl: findCol(c, "עלות ליד ברוטו", "עלות לליד ברוטו"),
      quality: findCol(c, "לידים איכותיים"),
      cpql: findCol(c, "עלות ליד איכותי", "עלות לליד איכותי"),
      qPct: findCol(c, "אחוז איכות"),
      inProc: findCol(c, "לידים בתהליך"),
      inProcPct: findCol(c, "אחוז הלידים בתהליך"),
    };
    for (let i = ph.rowIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      const name = String(r[ci.name] || "").trim();
      if (!name) continue;
      if (/^total$|סה"כ|סהכ/i.test(name)) {
        /* שורת ה-Total של הטבלה — מקור האמת ליעד התקציב והמימוש של חוברת התארים */
        out.totalsRow = {
          budget: toNum(r[ci.budget]), spent: toNum(r[ci.spent]),
          gross: toNum(r[ci.gross]), quality: toNum(r[ci.quality]),
          inProc: toNum(r[ci.inProc]),
        };
        break;
      }
      const row = {
        name, budget: toNum(r[ci.budget]) ?? 0, spent: toNum(r[ci.spent]) ?? 0,
        gross: toNum(r[ci.gross]), cpl: toNum(r[ci.cpl]),
        quality: toNum(r[ci.quality]), cpql: toNum(r[ci.cpql]),
        qPct: toPct(r[ci.qPct]), inProc: toNum(r[ci.inProc]), inProcPct: toPct(r[ci.inProcPct]),
        drill: name.includes("חיפוש"),
      };
      if (alive(row)) out.platforms.push(row);
    }
  }

  /* --- דריל-דאון קמפייני חיפוש --- */
  /* שורת הכותרת חייבת להכיל גם "קמפיין" וגם "תקציב שמומש" — כדי לא להתבלבל עם בלוק המטא למעלה */
  let ch = null;
  for (let i = 0; i < rows.length; i++) {
    const cells = rows[i].map((x) => String(x || "").trim());
    if (cells.includes("קמפיין") && cells.some((x) => x.includes("תקציב שמומש"))) {
      ch = { rowIdx: i, cols: cells };
      break;
    }
  }
  if (ch) {
    const c = ch.cols;
    const ci = {
      name: c.indexOf("קמפיין"),
      spent: findCol(c, "תקציב שמומש"),
      share: findCol(c, "נתח מעלות"),
      gross: findCol(c, "לידים ברוטו"),
      cpl: findCol(c, "עלות ליד ברוטו", "עלות לליד ברוטו"),
      quality: findCol(c, "לידים איכותיים"),
      cpql: findCol(c, "עלות ליד איכותי", "עלות לליד איכותי"),
      qPct: findCol(c, "אחוז איכות"),
      inProc: findCol(c, "לידים בתהליך"),
      inProcPct: findCol(c, "אחוז הלידים בתהליך"),
    };
    for (let i = ch.rowIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      const name = String(r[ci.name] || "").trim();
      if (!name) continue;
      if (/^total$|סה"כ|סהכ/i.test(name)) break;
      const row = {
        name, spent: toNum(r[ci.spent]), share: toPct(r[ci.share]),
        gross: toNum(r[ci.gross]), cpl: toNum(r[ci.cpl]),
        quality: toNum(r[ci.quality]), cpql: toNum(r[ci.cpql]),
        qPct: toPct(r[ci.qPct]), inProc: toNum(r[ci.inProc]), inProcPct: toPct(r[ci.inProcPct]),
      };
      /* מקורות אורגניים (GMB/callbox/Direct) חיים גם בלי תקציב */
      if (alive(row) || (row.gross || 0) > 0) out.campaigns.push(row);
    }
  }

  /* --- הארווארד: סקשן נפרד אחרי כותרת "הארווארד - ריכוז נתונים" --- */
  const hIdx = rows.findIndex((r) => r.some((c) => String(c).includes("הארווארד") && String(c).includes("ריכוז")));
  if (hIdx !== -1) {
    let hh = -1;
    for (let i = hIdx + 1; i < Math.min(rows.length, hIdx + 12); i++) {
      if (rows[i].some((x) => String(x || "").trim() === "פלטפורמה")) { hh = i; break; }
    }
    if (hh !== -1) {
      const c = rows[hh].map((x) => String(x || "").trim());
      const ci = {
        name: c.indexOf("פלטפורמה"),
        budget: findCol(c, "תקציב"),
        spent: findCol(c, "תקציב שמומש"),
        gross: findCol(c, "לידים ברוטו"),
        cpl: findCol(c, "עלות ליד ברוטו", "עלות לליד ברוטו"),
        quality: findCol(c, "לידים איכותיים"),
        cpql: findCol(c, "עלות ליד איכותי", "עלות לליד איכותי"),
        qPct: findCol(c, "אחוז איכות"),
        inProc: findCol(c, "לידים בתהליך"),
        inProcPct: findCol(c, "אחוז הלידים בתהליך"),
      };
      const H = { platforms: [], total: null };
      for (let i = hh + 1; i < rows.length; i++) {
        const r = rows[i];
        const name = String(r[ci.name] || "").trim();
        if (!name) { if (H.platforms.length || H.total) break; else continue; }
        const row = {
          name, budget: toNum(r[ci.budget]) ?? 0, spent: toNum(r[ci.spent]) ?? 0,
          gross: toNum(r[ci.gross]), cpl: toNum(r[ci.cpl]),
          quality: toNum(r[ci.quality]), cpql: toNum(r[ci.cpql]),
          qPct: ci.qPct !== -1 ? toPct(r[ci.qPct]) : null,
          inProc: toNum(r[ci.inProc]), inProcPct: toPct(r[ci.inProcPct]),
        };
        if (row.qPct === null && (row.gross || 0) > 0 && row.quality !== null) row.qPct = (row.quality / row.gross) * 100;
        if (/^total$|סה"כ/i.test(name)) { H.total = row; break; }
        if (alive(row)) H.platforms.push(row);
      }
      if (H.platforms.length) out.harvard = H;
    }
  }

  const ok = out.summary && out.platforms.length > 0;
  return ok ? { meta: out.meta, summary: out.summary, platforms: out.platforms, campaigns: out.campaigns, harvard: out.harvard, totalsRow: out.totalsRow || null } : null;
}

export { parseReport };

/* פרסר לגיליון ה-YoY: שורה per ליד עם תאריך / מקור / איכות.
   מחזיר מפה: { "MM-DD": {digital, site} } של תקינים יומיים, מקובצת לפי יום-בשנה. */
/* מזהה אם עמודת תאריכים היא dd/mm (ישראלי) או mm/dd (אמריקאי, למשל מ-Supermetrics).
   אם באיזשהו ערך החלק הראשון > 12 → הוא חייב להיות יום → פורמט dd/mm.
   אם השני > 12 → הוא חייב להיות יום → פורמט mm/dd. ברירת מחדל: dd/mm. */
const _detectDateOrder = (samples) => {
  let firstGt12 = false, secondGt12 = false;
  for (const s of samples) {
    const m = String(s).trim().match(/(\d{1,2})[\/.\-](\d{1,2})[\/.\-]\d{2,4}/);
    if (!m) continue;
    const a = +m[1], b = +m[2];
    if (a > 12) firstGt12 = true;
    if (b > 12) secondGt12 = true;
  }
  if (secondGt12 && !firstGt12) return "mdy"; /* mm/dd/yyyy */
  return "dmy"; /* ברירת מחדל dd/mm/yyyy */
};
const _yoyParseDate = (s, order = "dmy") => {
  const m = String(s).trim().match(/(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/);
  if (!m) return null;
  let a = +m[1], b = +m[2], y = m[3]; y = y.length === 2 ? "20" + y : y;
  const d = order === "mdy" ? b : a;
  const mo = order === "mdy" ? a : b;
  return { y: +y, mo, d, mmdd: `${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}` };
};
const _isSiteSrc = (s) => { const v = String(s || "").trim(); return v.includes("אתר") || /site|website|magazine|מגזין/i.test(v); };

/* גיליון לידים (ברוטו): שורה per ליד עם תאריך/מקור/איכות.
   מחזיר { year: { "MM-DD": {gross:{digital,site,total}, quality:{digital,site,total}} } } */
function parseYoY(csvText) {
  const rows = Papa.parse(csvText).data;
  if (!rows.length) return null;
  let hdrIdx = -1, ci = {};
  for (let i = 0; i < Math.min(rows.length, 8); i++) {
    const cells = rows[i].map((c) => String(c || "").trim());
    const dCol = cells.findIndex((c) => c.includes("תאריך"));
    const qCol = cells.findIndex((c) => c === "איכות" || c.includes("איכות"));
    if (dCol !== -1 && qCol !== -1) { hdrIdx = i; ci = { date: dCol, source: cells.findIndex((c) => c.includes("מקור")), quality: qCol }; break; }
  }
  if (hdrIdx === -1) return null;
  const order = _detectDateOrder(rows.slice(hdrIdx + 1, hdrIdx + 400).map((r) => r[ci.date]));
  const byYear = {};
  for (let i = hdrIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const dt = _yoyParseDate(r[ci.date], order);
    if (!dt) continue;
    const q = String(r[ci.quality] || "").trim();
    const isSite = _isSiteSrc(ci.source !== -1 ? r[ci.source] : "");
    byYear[dt.y] = byYear[dt.y] || {};
    const b = byYear[dt.y][dt.mmdd] || { gross: { digital: 0, site: 0, total: 0 }, quality: { digital: 0, site: 0, total: 0 } };
    const k = isSite ? "site" : "digital";
    b.gross[k]++; b.gross.total++;
    if (q === "תקין") { b.quality[k]++; b.quality.total++; }
    byYear[dt.y][dt.mmdd] = b;
  }
  return byYear;
}
export { parseYoY };

/* גיליון תקציב: Data source / Date / Cost (ILS). מחזיר { year: { "MM-DD": {digital, site, total} } } בש"ח */
function parseYoYBudget(csvText) {
  const rows = Papa.parse(csvText).data;
  if (!rows.length) return null;
  let hdrIdx = -1, ci = {};
  for (let i = 0; i < Math.min(rows.length, 8); i++) {
    const cells = rows[i].map((c) => String(c || "").trim());
    const dCol = cells.findIndex((c) => /date|תאריך/i.test(c));
    /* עדיפות לעמודת עלות בשקלים (ILS/₪/שקל); רק אם אין — עמודת עלות כללית (למשל USD) */
    let cCol = cells.findIndex((c) => /cost|עלות|תקציב|spend/i.test(c) && /ils|₪|שקל|nis/i.test(c));
    if (cCol === -1) cCol = cells.findIndex((c) => /cost|עלות|תקציב|spend/i.test(c) && !/usd|\$|dollar/i.test(c));
    if (cCol === -1) cCol = cells.findIndex((c) => /cost|עלות|תקציב|spend/i.test(c));
    if (dCol !== -1 && cCol !== -1) { hdrIdx = i; ci = { date: dCol, cost: cCol, source: cells.findIndex((c) => /source|מקור|פלטפורמה/i.test(c)) }; break; }
  }
  if (hdrIdx === -1) return null;
  const order = _detectDateOrder(rows.slice(hdrIdx + 1, hdrIdx + 400).map((r) => r[ci.date]));
  const byYear = {};
  for (let i = hdrIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const dt = _yoyParseDate(r[ci.date], order);
    if (!dt) continue;
    const cost = toNum(r[ci.cost]);
    if (cost === null) continue;
    const isSite = _isSiteSrc(ci.source !== -1 ? r[ci.source] : "");
    byYear[dt.y] = byYear[dt.y] || {};
    const b = byYear[dt.y][dt.mmdd] || { digital: 0, site: 0, total: 0 };
    if (isSite) b.site += cost; else b.digital += cost;
    b.total += cost;
    byYear[dt.y][dt.mmdd] = b;
  }
  return byYear;
}
export { parseYoYBudget };

/* חגים עבריים לחודש גרגוריאני נתון — מזוהים דרך לוח השנה העברי של הדפדפן.
   מחזיר [{day, label}] לימים שהם חג/ערב חג באותו חודש. */
function hebrewHolidaysForMonth(gy, gm) {
  const out = [];
  try {
    const fmt = new Intl.DateTimeFormat("en-u-ca-hebrew", { day: "numeric", month: "long" });
    const daysInM = new Date(gy, gm, 0).getDate();
    /* מיפוי אירועים עבריים מרכזיים לפי (חודש עברי, יום עברי). ערב חג = היום שלפני */
    const HEB = {
      "Tishri-1": "ראש השנה", "Tishri-10": "יום כיפור", "Tishri-15": "סוכות",
      "Tishri-22": "שמחת תורה", "Nisan-15": "פסח", "Nisan-21": "שביעי של פסח",
      "Sivan-6": "שבועות", "Shevat-15": 'ט"ו בשבט', "Adar-14": "פורים", "Kislev-25": "חנוכה",
      "Iyar-5": "יום העצמאות", "Iyar-4": "יום הזיכרון", "Av-9": "תשעה באב",
    };
    let prevLabel = null;
    for (let d = 1; d <= daysInM; d++) {
      const parts = fmt.formatToParts(new Date(gy, gm - 1, d));
      const hMonth = parts.find((p) => p.type === "month")?.value.replace(/\s/g, "");
      const hDay = parts.find((p) => p.type === "day")?.value;
      const key = `${hMonth}-${hDay}`;
      if (HEB[key]) out.push({ day: d, label: HEB[key] });
    }
    /* הוספת "ערב" ליום שלפני חג מרכזי אם הוא באותו חודש */
    const majors = new Set(["ראש השנה", "יום כיפור", "סוכות", "פסח", "שבועות"]);
    const withEves = [];
    out.forEach((h) => { if (majors.has(h.label) && h.day > 1) withEves.push({ day: h.day - 1, label: "ערב " + h.label }); });
    return [...out, ...withEves].sort((a, b) => a.day - b.day);
  } catch { return out; }
}
export { hebrewHolidaysForMonth };

/* מהקישור נגזרים כמה מסלולי גישה — מנסים אותם לפי הסדר עד שאחד מצליח */
function toCsvUrls(link) {
  const m = String(link).match(/\/d\/(?:e\/)?([\w-]+)/);
  if (!m) return [];
  const id = m[1];
  const gid = (String(link).match(/[#&?]gid=(\d+)/) || [])[1] || "0";
  if (link.includes("/d/e/")) return [`https://docs.google.com/spreadsheets/d/e/${id}/pub?output=csv&gid=${gid}`];
  return [
    `https://docs.google.com/spreadsheets/d/${id}/export?format=csv&gid=${gid}`,
    `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&gid=${gid}`,
  ];
}

/* fetch עם timeout — כדי שבקשה תקועה לא תקפיא את הכפתורים לנצח */
const fetchWithTimeout = (url, ms = 15000) => {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), ms);
  return fetch(url, { signal: ctrl.signal, redirect: "follow" }).finally(() => clearTimeout(t));
};

/* קידוד/פענוח מצב הדשבורד לתוך קישור (hash) — הקישור נושא את הנתונים עצמם */
const encodeState = (obj) =>
  btoa(unescape(encodeURIComponent(JSON.stringify(obj)))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
const decodeState = (s) =>
  JSON.parse(decodeURIComponent(escape(atob(s.replace(/-/g, "+").replace(/_/g, "/")))));

function currentBaseUrl() {
  try { return window.top.location.href.split("#")[0]; } catch {}
  return window.location.href.split("#")[0];
}

/* ------------------------------------------------------------------ */
/*  דלתא בין דוחות                                                      */
/* ------------------------------------------------------------------ */

function computeDelta(curr, prev) {
  if (!prev || !prev.data) return null;
  const p = prev.data;
  const d = { date: prev.savedAt, summary: {}, platforms: {} };
  const S = curr.summary.actual, PS = p.summary.actual;
  ["gross", "quality", "cpl", "cpql", "qualityPct", "inProcess"].forEach((k) => {
    if (S[k] !== null && PS[k] !== null && S[k] !== undefined && PS[k] !== undefined) d.summary[k] = S[k] - PS[k];
  });
  const spentNow = curr.platforms.reduce((s, x) => s + (x.spent || 0), 0);
  const spentPrev = p.platforms.reduce((s, x) => s + (x.spent || 0), 0);
  d.summary.spent = spentNow - spentPrev;
  curr.platforms.forEach((pl) => {
    const old = p.platforms.find((x) => x.name === pl.name);
    if (old) d.platforms[pl.name] = {
      spent: (pl.spent || 0) - (old.spent || 0),
      gross: pl.gross !== null && old.gross !== null ? pl.gross - old.gross : null,
      quality: pl.quality !== null && old.quality !== null ? pl.quality - old.quality : null,
    };
  });
  d.campaigns = {};
  (curr.campaigns || []).forEach((cm) => {
    const old = (p.campaigns || []).find((x) => x.name === cm.name);
    if (old) d.campaigns[cm.name] = {
      spent: (cm.spent || 0) - (old.spent || 0),
      gross: cm.gross !== null && old.gross !== null ? cm.gross - old.gross : null,
      quality: cm.quality !== null && old.quality !== null ? cm.quality - old.quality : null,
    };
  });
  /* הארווארד — דלתות לפלטפורמות + סיכום */
  d.harvard = {};
  d.harvardSummary = null;
  if (curr.harvard && p.harvard) {
    (curr.harvard.platforms || []).forEach((pl) => {
      const old = (p.harvard.platforms || []).find((x) => x.name === pl.name);
      if (old) d.harvard[pl.name] = {
        spent: (pl.spent || 0) - (old.spent || 0),
        gross: pl.gross !== null && old.gross !== null ? pl.gross - old.gross : null,
        quality: pl.quality !== null && old.quality !== null ? pl.quality - old.quality : null,
      };
    });
    const sum = (h) => (h.platforms || []).reduce(
      (a, b) => ({ spent: a.spent + (b.spent || 0), gross: a.gross + (b.gross || 0), quality: a.quality + (b.quality || 0) }),
      { spent: 0, gross: 0, quality: 0 }
    );
    const cs = sum(curr.harvard), ps = sum(p.harvard);
    d.harvardSummary = { spent: cs.spent - ps.spent, gross: cs.gross - ps.gross, quality: cs.quality - ps.quality };
  }
  return d;
}

const Delta = ({ v, inverse, suffix = "", money }) => {
  if (v === null || v === undefined || v === 0 || Number.isNaN(v)) return null;
  const good = inverse ? v < 0 : v > 0;
  const fmt = money ? nis(Math.abs(v)) : Math.abs(v) % 1 ? Math.abs(v).toFixed(1) : num(Math.abs(v));
  return (
    <span style={{ color: good ? "#37C9A9" : "#F0685F", fontSize: 12, fontWeight: 700, marginRight: 6, direction: "ltr", display: "inline-block" }}>
      {v > 0 ? "▲" : "▼"} {fmt}{suffix}
    </span>
  );
};

/* ------------------------------------------------------------------ */
/*  תובנות אוטומטיות                                                    */
/* ------------------------------------------------------------------ */

/* מיצוי מדדי-על מדוח מפורסר בודד — לפילוח תארים / הארווארד / אתר אינו רלוונטי כאן */
function reportTotals(rep) {
  if (!rep) return null;
  const budgetTarget = (rep.totalsRow && rep.totalsRow.budget) || (rep.platforms || []).reduce((a, b) => a + (b.budget || 0), 0);
  const spent = (rep.totalsRow && rep.totalsRow.spent) || (rep.platforms || []).reduce((a, b) => a + (b.spent || 0), 0);
  /* איכותיים · תארים בלבד — משורת ה-Total של ה-BOF (ללא הארווארד). אם אין, נגזור מהבלוק המשולב פחות הארווארד */
  const harvardQuality = rep.harvard
    ? ((rep.harvard.total && rep.harvard.total.quality) ?? (rep.harvard.platforms || []).reduce((a, b) => a + (b.quality || 0), 0))
    : null;
  const harvardSpent = rep.harvard
    ? ((rep.harvard.total && rep.harvard.total.spent) ?? (rep.harvard.platforms || []).reduce((a, b) => a + (b.spent || 0), 0))
    : null;
  let degQuality;
  if (rep.totalsRow && rep.totalsRow.quality != null) degQuality = rep.totalsRow.quality;
  else {
    const combined = rep.summary && rep.summary.actual ? (rep.summary.actual.quality || 0) : 0;
    degQuality = combined - (harvardQuality || 0);
  }
  const cpqlTotal = rep.summary && rep.summary.actual ? rep.summary.actual.cpql : null;
  return { budgetTarget, spent, degQuality, harvardQuality, harvardSpent, cpqlTotal };
}

function buildInsights(data, delta, isActive = () => true) {
  const out = [];
  const P = data.platforms.filter((p) => (p.spent || 0) > 1000 && isActive(p.name));
  const withQ = P.filter((p) => (p.quality || 0) >= 2 && p.cpql);
  const anySpend = data.platforms.some((p) => (p.spent || 0) > 0);
  /* מקורות ללא תקציב (חוברת האתר) — תובנות סביב לידים ואיכות בלבד */
  if (!anySpend && data.platforms.length >= 1) {
    const top = [...data.platforms].filter((p) => p.gross).sort((a, b) => b.gross - a.gross)[0];
    const totGross = data.platforms.reduce((a, b) => a + (b.gross || 0), 0);
    if (top && totGross)
      out.push({ icon: "🌐", tone: "info", text: `${top.name} הוא מקור הלידים המרכזי — ${num(top.gross)} לידים (${Math.round((top.gross / totGross) * 100)}% מהסך).` });
    const s0 = data.summary;
    if (s0.actual.invalid !== null && s0.actual.invalid !== undefined && s0.actual.gross)
      out.push({
        icon: "🧹", tone: s0.actual.invalid / s0.actual.gross > 0.4 ? "bad" : "warn",
        text: `${num(s0.actual.invalid)} לידים לא תקינים מתוך ${num(s0.actual.gross)} ברוטו (${Math.round((s0.actual.invalid / s0.actual.gross) * 100)}%) — ${s0.actual.invalid / s0.actual.gross > 0.4 ? "שיעור גבוה; שווה לבדוק ולידציה בטפסים וסינון ספאם" : "שווה מעקב על איכות הטפסים"}.`,
      });
    const bestQ = [...data.platforms].filter((p) => (p.quality || 0) >= 2).sort((a, b) => (b.qPct || 0) - (a.qPct || 0))[0];
    if (bestQ) out.push({ icon: "⭐", tone: "good", text: `${bestQ.name} מוביל באחוז איכות — ${pct(bestQ.qPct, 1)} (${num(bestQ.quality)} איכותיים).` });
    if (delta && delta.summary && delta.summary.gross)
      out.push({ icon: "📈", tone: "info", text: `ביממה האחרונה נוספו ${num(delta.summary.gross)} לידים מהאתר${delta.summary.quality ? `, מתוכם ${num(delta.summary.quality)} איכותיים` : ""}.` });
  } else if (data.platforms.length === 1) {
    const p = data.platforms[0];
    const util = p.budget ? ((p.spent || 0) / p.budget) * 100 : null;
    const pace = data.meta && data.meta.pace != null ? data.meta.pace : null;
    if (util !== null && pace !== null) {
      const gap = Math.round(util - pace);
      out.push({
        icon: gap >= 0 ? "⏱️" : "🐢", tone: gap >= -5 ? "good" : "warn",
        text: gap >= 0
          ? `${p.name} מנצל ${Math.round(util)}% מהתקציב, ${gap}+ נק' מעל הקצב הקלנדרי (${Math.round(pace)}%) — קצב הוצאה בריא.`
          : `${p.name} מנצל ${Math.round(util)}% מהתקציב מול קצב קלנדרי של ${Math.round(pace)}% — פיגור של ${Math.abs(gap)} נק'. אם היעד הוא מיצוי החודש, שקלו הגברת בידים/קהלים.`,
      });
    }
    if (delta && delta.platforms[p.name]) {
      const d = delta.platforms[p.name];
      const dayCpql = (d.quality || 0) > 0 ? d.spent / d.quality : null;
      if (dayCpql !== null && p.cpql)
        out.push({
          icon: dayCpql <= p.cpql ? "📈" : "📉", tone: dayCpql <= p.cpql ? "good" : "warn",
          text: `עלות לליד איכותי ביממה האחרונה: ${nis(dayCpql)} — ${dayCpql <= p.cpql ? "טובה מהממוצע המצטבר" : "יקרה מהממוצע המצטבר"} (${nis(p.cpql)}).`,
        });
      else if ((d.spent || 0) > 500 && (d.quality || 0) === 0)
        out.push({ icon: "🧐", tone: "warn", text: `ביממה האחרונה מומשו ${nis(d.spent)} בלי ליד איכותי חדש — ${num(d.gross ?? 0)} לידים ברוטו נכנסו וממתינים לטיוב.` });
    }
    if (p.inProc)
      out.push({ icon: "⏳", tone: "info", text: `${num(p.inProc)} לידים בתהליך${p.inProcPct ? ` (${pct(p.inProcPct)} מהברוטו)` : ""} — שם טמון הפוטנציאל המיידי לשיפור אחוז האיכות.` });
  } else if (withQ.length) {
    const best = [...withQ].sort((a, b) => a.cpql - b.cpql)[0];
    out.push({ icon: "🏆", tone: "good", text: `${best.name} הוא הכלי היעיל ביותר לליד איכותי — ${nis(best.cpql)} בלבד (${pct(best.qPct, 1)} איכות). שווה לבחון הסטת תקציב לכיוונו.` });
    const worst = [...withQ].sort((a, b) => b.cpql - a.cpql)[0];
    if (worst.name !== best.name && worst.cpql > best.cpql * 4)
      out.push({ icon: "⚠️", tone: "bad", text: `${worst.name} יקר משמעותית — ${nis(worst.cpql)} לליד איכותי, פי ${Math.round(worst.cpql / best.cpql)} מ־${best.name}. מומלץ לבחון קריאייטיב/קהלים או צמצום.` });
  }
  const burned = P.filter((p) => (p.spent || 0) > 3000 && (p.quality || 0) === 0);
  burned.forEach((p) => out.push({ icon: "🔻", tone: "bad", text: `${p.name} שרף ${nis(p.spent)} בלי ליד איכותי אחד — לעצור ולבדוק לפני המשך השקעה.` }));
  /* כלים כבויים — או להתעלם, או להמליץ להפעיל אם היו יעילים */
  const inactiveGood = data.platforms.filter((p) => !isActive(p.name) && (p.quality || 0) >= 2 && p.cpql);
  inactiveGood.sort((a, b) => a.cpql - b.cpql).slice(0, 2).forEach((p) =>
    out.push({ icon: "💤", tone: "warn", text: `${p.name} מסומן ככבוי, אבל היסטורית הביא ${num(p.quality)} לידים איכותיים ב־${nis(p.cpql)} לליד — שווה לשקול הפעלה מחדש.` })
  );
  const s = data.summary;
  if (s.actual.qualityPct !== null && s.target.qualityPct)
    out.push({
      icon: s.actual.qualityPct >= s.target.qualityPct ? "✅" : "🎯", tone: s.actual.qualityPct >= s.target.qualityPct ? "good" : "warn",
      text: s.actual.qualityPct >= s.target.qualityPct
        ? `אחוז האיכות (${pct(s.actual.qualityPct)}) עומד ביעד (${pct(s.target.qualityPct)}).`
        : `אחוז האיכות עומד על ${pct(s.actual.qualityPct)} מול יעד של ${pct(s.target.qualityPct)} — הפער הוא בעיקר בטיוב הלידים, לא בכמות. ${num(s.actual.inProcess)} לידים עדיין בתהליך ויכולים לסגור את הפער.`,
    });
  const camps = data.campaigns.filter((c) => c.spent && isActive(c.name));
  if (camps.length) {
    const bestC = [...camps].filter((c) => (c.quality || 0) >= 2).sort((a, b) => (b.qPct || 0) - (a.qPct || 0))[0];
    if (bestC) out.push({ icon: "🔍", tone: "good", text: `בחיפוש: קמפיין "${bestC.name}" מוביל באיכות עם ${pct(bestC.qPct, 1)} ו־${nis(bestC.cpql)} לליד איכותי.` });
    const zero = camps.find((c) => (c.spent || 0) > 1000 && (c.quality || 0) === 0);
    if (zero) out.push({ icon: "🧯", tone: "warn", text: `קמפיין "${zero.name}" מימש ${nis(zero.spent)} ללא לידים איכותיים — כדאי לבחון מילות מפתח שליליות והתאמת ה־LP.` });
  }
  if (delta) {
    const movers = Object.entries(delta.platforms)
      .filter(([, d]) => d.quality)
      .sort((a, b) => Math.abs(b[1].quality) - Math.abs(a[1].quality))[0];
    if (movers && movers[1].quality)
      out.push({
        icon: movers[1].quality > 0 ? "📈" : "📉", tone: movers[1].quality > 0 ? "good" : "bad",
        text: `מאז הדוח הקודם (${heDate(delta.date)}): ${movers[0]} ${movers[1].quality > 0 ? "הוסיף" : "איבד"} ${Math.abs(movers[1].quality)} לידים איכותיים.`,
      });
    if (delta.summary.spent)
      out.push({ icon: "💸", tone: "info", text: `מומשו ${nis(Math.abs(delta.summary.spent))} מאז הדוח הקודם, שהניבו ${num(delta.summary.gross ?? 0)} לידים חדשים (${num(delta.summary.quality ?? 0)} איכותיים).` });
  }
  return out.slice(0, 7);
}

/* ------------------------------------------------------------------ */
/*  צבעים                                                               */
/* ------------------------------------------------------------------ */

const C = {
  bg: "#0D1626", panel: "#152238", panelSoft: "#1B2B45", line: "#26395A",
  text: "#E9EFFA", dim: "#8FA3C2",
  amber: "#F5B841", teal: "#37C9A9", coral: "#F0685F", blue: "#5E93F5", violet: "#9A7BF7",
};
const PIE_COLORS = [C.amber, C.blue, C.teal, C.coral, C.violet, "#E58BC0", "#7FD1F0", "#C9D66B", "#F58F5E", "#6BC5D6", "#B48EAD", "#8AC98A"];

const qualityColor = (p) => {
  if (p === null || p === undefined) return C.dim;
  if (p >= 10) return C.teal;
  if (p >= 6) return C.amber;
  return C.coral;
};

/* ------------------------------------------------------------------ */
/*  מצב סטורי — "המצגת של הבוקר"                                        */
/* ------------------------------------------------------------------ */

function StoryMode({ slides, onClose, onShare }) {
  const [idx, setIdx] = useState(0);
  const [paused, setPaused] = useState(false);
  const next = () => (idx < slides.length - 1 ? setIdx(idx + 1) : onClose());
  useEffect(() => {
    const onKey = (e) => {
      if (e.key === "Escape") onClose();
      if (e.key === " ") { e.preventDefault(); setPaused((p) => !p); }
      if (e.key === "ArrowLeft") setIdx((i) => Math.min(i + 1, slides.length - 1));
      if (e.key === "ArrowRight") setIdx((i) => Math.max(i - 1, 0));
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [slides.length, onClose]);
  const s = slides[idx];
  return (
    <div className="story-overlay" dir="rtl">
      <div className="story-card" style={{ background: s.bg }}>
        <div className="story-progress">
          {slides.map((_, i) => (
            <div key={i} className="story-seg">
              <div className="story-seg-fill" key={`f${i}-${idx}`}
                   style={{ width: i < idx ? "100%" : i === idx ? undefined : "0%", animationPlayState: paused ? "paused" : "running" }}
                   data-active={i === idx ? "1" : "0"}
                   onAnimationEnd={i === idx ? next : undefined} />
            </div>
          ))}
        </div>
        <button className="story-close" onClick={onClose} aria-label="סגירה">✕</button>
        <div className="story-actions">
          <button className="story-share" onClick={onShare}>📤 שיתוף</button>
          <button className="story-share" onClick={() => setPaused((p) => !p)} aria-label={paused ? "המשך" : "השהיה"}>
            {paused ? "▶ המשך" : "⏸ השהיה"}
          </button>
        </div>
        {paused && <div className="story-paused-tag">מושהה · הזמן עצר בשבילך ⏸</div>}
        <div className="story-body">
          <div className="story-kicker">{s.kicker}</div>
          <div className="story-big" style={{ color: s.color }}>{s.big}</div>
          {s.text && <div className="story-text">{s.text}</div>}
          {s.rows && (
            <div className="story-rows">
              {s.rows.map((r, i) => (
                <div className="story-row" key={i}>
                  <span className="story-row-label">{r.label}</span>
                  <span className="story-row-value" style={{ color: r.color || C.text }}>{r.value}</span>
                </div>
              ))}
            </div>
          )}
          {s.foot && <div className="story-foot">{s.foot}</div>}
        </div>
        <div className="story-nav prev" onClick={() => setIdx((i) => Math.max(i - 1, 0))} />
        <div className="story-nav next" onClick={next} />
        <div className="story-count">{idx + 1} / {slides.length}</div>
      </div>
    </div>
  );
}

function buildSlides(data, totals, delta, insights, isActive = () => true, label = "תארים") {
  const s = data.summary;
  const now = new Date();
  const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  /* קצב קלנדרי — עדיפות לערך מהדוח עצמו (בלוק המטא), אחרת לפי תאריך היום */
  const pace = data.meta && data.meta.pace != null ? Math.round(data.meta.pace) : Math.round((now.getDate() / daysInMonth) * 100);
  const util = Math.round(totals.util);
  const gap = util - pace;
  const g1 = "linear-gradient(160deg,#0D1626 0%,#1B2B45 100%)";
  const dayCPL = (m) => (m.gross > 0 ? m.spent / m.gross : null);
  const dayCPQL = (m) => (m.quality > 0 ? m.spent / m.quality : null);

  /* תנועות היממה האחרונה (דלתא מול הדוח הקודם) — כלים + קמפייני חיפוש */
  const plMovers = delta ? Object.entries(delta.platforms).map(([name, d]) => ({ name, raw: name, ...d })) : [];
  const cmMovers = delta && delta.campaigns
    ? Object.entries(delta.campaigns).map(([name, d]) => ({ name: `חיפוש · ${name}`, raw: name, ...d })) : [];
  const all = [...plMovers, ...cmMovers];
  const topGross = all.filter((m) => (m.gross || 0) > 0).sort((a, b) => b.gross - a.gross)[0];
  const topQuality = all.filter((m) => (m.quality || 0) > 0).sort((a, b) => b.quality - a.quality)[0];
  const weakGross = plMovers
    .filter((m) => isActive(m.raw))
    .filter((m) => m.spent > 300 && ((m.gross || 0) <= 0 || (dayCPL(m) && dayCPL(m) > (s.actual.cpl || 100) * 2)))
    .sort((a, b) => b.spent - a.spent).slice(0, 3);
  const weakQuality = plMovers
    .filter((m) => isActive(m.raw))
    .filter((m) => m.spent > 300 && (m.quality || 0) <= 0)
    .sort((a, b) => b.spent - a.spent).slice(0, 3);

  /* חלופות כשאין עדיין דוח קודם */
  const noDeltaNote = "אין עדיין דוח קודם להשוואה — מוצגים נתונים מצטברים";
  const bestGrossAll = [...data.platforms].filter((p) => p.gross).sort((a, b) => b.gross - a.gross)[0];
  const bestQualityAll = [...data.platforms].filter((p) => p.quality).sort((a, b) => b.quality - a.quality)[0];
  const weakGrossAll = data.platforms.filter((p) => isActive(p.name) && (p.spent || 0) > 3000 && p.cpl && p.cpl > (s.actual.cpl || 100) * 3).slice(0, 3);
  const weakQualityAll = data.platforms.filter((p) => isActive(p.name) && (p.spent || 0) > 3000 && (p.quality || 0) === 0).slice(0, 3);

  const weakRow = (m, quality) => ({
    label: m.name,
    value: `${nis(m.spent)} · ${num(quality ? m.quality ?? 0 : m.gross ?? 0)} לידים · ${nis(quality ? dayCPQL(m) : dayCPL(m)) || "—"} לליד`,
    color: C.coral,
  });
  const weakRowAll = (p, quality) => ({
    label: p.name,
    value: `${nis(p.spent)} · ${num(quality ? p.quality : p.gross)} לידים · ${nis(quality ? p.cpql : p.cpl)} לליד`,
    color: C.coral,
  });

  const hasBudget = (totals.budget || 0) > 0;
  const slides = [
    /* 1 — נתונים כלליים: תקציב/לידים, קצב קלנדרי */
    hasBudget ? {
      kicker: now.toLocaleDateString("he-IL", { weekday: "long", day: "numeric", month: "long" }) + ` · ${label} · תמונת מצב`,
      big: `${util}% ניצול`,
      color: util >= pace ? C.teal : C.amber,
      rows: [
        { label: "תקציב שנוצל", value: `${nis(totals.spent)} מתוך ${nis(totals.budget)}`, color: C.amber },
        { label: "קצב קלנדרי של החודש", value: `${pace}% (יום ${now.getDate()} מתוך ${daysInMonth})` },
        {
          label: "ניצול מול קצב",
          value: `אנחנו ב־${util}% ניצול מול ${pace}% קצב קלנדרי (${gap >= 0 ? "+" : "−"}${Math.abs(gap)}%)`,
          color: gap >= 0 ? C.teal : C.coral,
        },
        ...(delta && delta.summary.spent ? [{ label: "מומש ביממה האחרונה", value: nis(delta.summary.spent), color: C.blue }] : []),
      ],
      foot: delta ? `השוואה מול הדוח מ־${heDate(delta.date)}` : noDeltaNote,
      bg: g1,
    } : {
      kicker: now.toLocaleDateString("he-IL", { weekday: "long", day: "numeric", month: "long" }) + ` · ${label} · תמונת מצב`,
      big: `${num(s.actual.gross)} לידים`,
      color: C.blue,
      rows: [
        ...(s.target.gross ? [{ label: "עמידה ביעד ברוטו", value: `${Math.round((s.actual.gross / s.target.gross) * 100)}% מתוך ${num(s.target.gross)}` }] : []),
        ...(s.target.quality ? [{ label: "עמידה ביעד איכותיים", value: `${Math.round((s.actual.quality / s.target.quality) * 100)}% מתוך ${num(s.target.quality)}`, color: C.teal }] : []),
        { label: "קצב קלנדרי של החודש", value: `${pace}% (יום ${now.getDate()} מתוך ${daysInMonth})` },
        ...(delta && delta.summary.gross ? [{ label: "לידים ביממה האחרונה", value: `+${num(delta.summary.gross)}`, color: C.blue }] : []),
      ],
      foot: delta ? `השוואה מול הדוח מ־${heDate(delta.date)}` : noDeltaNote,
      bg: g1,
    },
    /* 2 — ברוטו */
    {
      kicker: "לידים ברוטו",
      big: num(s.actual.gross),
      color: C.blue,
      rows: [
        ...(s.target.gross ? [{ label: "עמידה ביעד", value: `${Math.round((s.actual.gross / s.target.gross) * 100)}% מתוך ${num(s.target.gross)}` }] : []),
        ...(s.actual.cpl !== null && s.actual.cpl !== undefined ? [{ label: "עלות לליד ברוטו", value: s.target.cpl ? `${nis(s.actual.cpl)} (יעד: ${nis(s.target.cpl)})` : nis(s.actual.cpl), color: s.target.cpl ? (s.actual.cpl <= s.target.cpl ? C.teal : C.coral) : C.text }] : []),
        ...(delta && delta.summary.gross ? [{ label: "לידים ביממה האחרונה", value: `+${num(delta.summary.gross)}`, color: C.teal }] : []),
        topGross
          ? { label: "מוביל היממה", value: `${topGross.name} · ${num(topGross.gross)} לידים (${nis(dayCPL(topGross))} לליד)`, color: C.amber }
          : bestGrossAll
            ? { label: "המוביל המצטבר", value: `${bestGrossAll.name} · ${num(bestGrossAll.gross)} לידים`, color: C.amber }
            : null,
      ].filter(Boolean),
      foot: delta ? null : noDeltaNote,
      bg: g1,
    },
    /* 3 — איכות */
    {
      kicker: "לידים איכותיים",
      big: num(s.actual.quality),
      color: C.teal,
      rows: [
        ...(s.target.quality ? [{ label: "עמידה ביעד", value: `${Math.round((s.actual.quality / s.target.quality) * 100)}% מתוך ${num(s.target.quality)}` }] : []),
        { label: "אחוז איכות", value: s.target.qualityPct ? `${pct(s.actual.qualityPct)} (יעד: ${pct(s.target.qualityPct)})` : pct(s.actual.qualityPct, 1), color: s.target.qualityPct ? (s.actual.qualityPct >= s.target.qualityPct ? C.teal : C.coral) : C.text },
        ...(s.actual.cpql !== null && s.actual.cpql !== undefined ? [{ label: "עלות לליד איכותי", value: s.target.cpql ? `${nis(s.actual.cpql)} (יעד: ${nis(s.target.cpql)})` : nis(s.actual.cpql), color: s.target.cpql ? (s.actual.cpql <= s.target.cpql ? C.teal : C.coral) : C.text }] : []),
        ...(delta && delta.summary.quality ? [{ label: "איכותיים ביממה האחרונה", value: `+${num(delta.summary.quality)}`, color: C.teal }] : []),
        topQuality
          ? { label: "מוביל היממה", value: `${topQuality.name} · ${num(topQuality.quality)} איכותיים (${nis(dayCPQL(topQuality))} לליד)`, color: C.amber }
          : bestQualityAll
            ? { label: "המוביל המצטבר", value: `${bestQualityAll.name} · ${num(bestQualityAll.quality)} איכותיים`, color: C.amber }
            : null,
        { label: "בתהליך", value: `${num(s.actual.inProcess)} לידים שעוד יכולים להפוך לאיכותיים`, color: C.violet },
      ].filter(Boolean),
      foot: delta ? null : noDeltaNote,
      bg: g1,
    },
  ];
  const weakSlides = hasBudget ? [
    /* 4 — לשים לב · ברוטו */
    {
      kicker: "🚨 לשים לב · ברוטו" + (delta ? " · היממה האחרונה" : ""),
      big: (delta ? weakGross : weakGrossAll).length ? `${(delta ? weakGross : weakGrossAll).length} כלים` : "הכל תקין ✅",
      color: (delta ? weakGross : weakGrossAll).length ? C.coral : C.teal,
      rows: delta
        ? (weakGross.length ? weakGross.map((m) => weakRow(m, false)) : [{ label: "אין חריגים", value: "כל הכלים שהוציאו תקציב הביאו לידים בעלות סבירה", color: C.teal }])
        : (weakGrossAll.length ? weakGrossAll.map((p) => weakRowAll(p, false)) : [{ label: "אין חריגים", value: "אין כלים עם עלות לליד חריגה", color: C.teal }]),
      foot: "חריג = תקציב שמומש בלי לידים, או עלות לליד כפולה ומעלה מהממוצע",
      bg: g1,
    },
    /* 5 — לשים לב · איכות */
    {
      kicker: "🚨 לשים לב · איכות" + (delta ? " · היממה האחרונה" : ""),
      big: (delta ? weakQuality : weakQualityAll).length ? `${(delta ? weakQuality : weakQualityAll).length} כלים` : "הכל תקין ✅",
      color: (delta ? weakQuality : weakQualityAll).length ? C.coral : C.teal,
      rows: delta
        ? (weakQuality.length ? weakQuality.map((m) => weakRow(m, true)) : [{ label: "אין חריגים", value: "כל הכלים שהוציאו תקציב הביאו גם לידים איכותיים", color: C.teal }])
        : (weakQualityAll.length ? weakQualityAll.map((p) => weakRowAll(p, true)) : [{ label: "אין חריגים", value: "אין כלים ששרפו תקציב בלי איכותיים", color: C.teal }]),
      foot: "חריג = תקציב שמומש בלי ליד איכותי אחד",
      bg: g1,
    },
  ] : [];
  slides.push(...weakSlides);
  if (insights.length) slides.push({
    kicker: "💡 תובנת היום",
    big: insights[0].icon,
    color: C.amber,
    text: insights[0].text,
    foot: "כל התובנות המלאות — בדשבורד",
    bg: g1,
  });
  return slides;
}

/* ------------------------------------------------------------------ */
/*  רכיבים                                                              */
/* ------------------------------------------------------------------ */

function KpiCard({ title, actual, target, format, inverse, sub, delta, deltaMoney, deltaSuffix }) {
  const hasTarget = target !== null && target !== undefined && target !== 0;
  const prog = hasTarget ? (inverse ? (target / actual) * 100 : (actual / target) * 100) : null;
  const clamped = prog === null ? null : Math.max(0, Math.min(prog, 130));
  const good = prog !== null && prog >= (inverse ? 100 : 90);
  return (
    <div className="kpi">
      <div className="kpi-title">{title}</div>
      <div className="kpi-value">
        {format(actual)}
        <Delta v={delta} inverse={inverse} money={deltaMoney} suffix={deltaSuffix} />
      </div>
      <div className="kpi-target">{hasTarget ? `יעד: ${format(target)}` : ""}{hasTarget && sub ? " · " : ""}{sub || ""}</div>
      {prog !== null && (
        <div className="kpi-bar-wrap">
          <div className="kpi-bar" style={{ width: `${(clamped / 130) * 100}%`, background: good ? C.teal : prog >= 50 ? C.amber : C.coral }} />
          <div className="kpi-goal-line" style={{ right: `${(100 / 130) * 100}%` }} />
        </div>
      )}
      {prog !== null && (
        <div className="kpi-pct" style={{ color: good ? C.teal : prog >= 50 ? C.amber : C.coral }}>
          {Math.round(prog)}% עמידה ביעד
        </div>
      )}
    </div>
  );
}

function Th({ children, k, sort, setSort }) {
  const active = sort.key === k;
  return (
    <th onClick={() => setSort({ key: k, dir: active && sort.dir === "desc" ? "asc" : "desc" })}
        style={{ cursor: "pointer", color: active ? C.amber : undefined }}>
      {children}{active ? (sort.dir === "desc" ? " ▾" : " ▴") : ""}
    </th>
  );
}

/* ------------------------------------------------------------------ */
/*  אפליקציה                                                            */
/* ------------------------------------------------------------------ */

export default function App() {
  const [data, setData] = useState(DEMO);
  const [prevSnap, setPrevSnap] = useState(null);
  const [link, setLink] = useState("");
  const [status, setStatus] = useState({ kind: "demo", msg: "מציג נתוני דוגמה מהדוח — חברו את הגיליון לעדכון יומי" });
  const [showConnect, setShowConnect] = useState(false);
  const [pasteMode, setPasteMode] = useState(false);
  const [csvPaste, setCsvPaste] = useState("");
  const [sort, setSort] = useState({ key: "spent", dir: "desc" });
  const [drillOpen, setDrillOpen] = useState(true);
  const [updatedAt, setUpdatedAt] = useState(null);
  const [story, setStory] = useState(false);
  const [daySort, setDaySort] = useState({ key: "spent", dir: "desc" });
  const [board, setBoard] = useState("main"); /* main = תארים · harvard = הארווארד · site = אתר */
  const [siteData, setSiteData] = useState(null);
  const [sitePrev, setSitePrev] = useState(null);
  const [siteLink, setSiteLink] = useState("https://docs.google.com/spreadsheets/d/1HJ8LFliiXjbKJiSi8NAcg6tNiRYKGsjH8yN4QgVJgLo/edit?usp=sharing");
  const [siteUpdatedAt, setSiteUpdatedAt] = useState(null);
  const [autoSync, setAutoSync] = useState(false);
  const [loading, setLoading] = useState(false);
  const [pendingAutoFetch, setPendingAutoFetch] = useState(null);
  /* ארכיון ולוח שנה */
  const [calOpen, setCalOpen] = useState(false);
  const [calYear, setCalYear] = useState(new Date().getFullYear());
  const [dayIndex, setDayIndex] = useState([]);          /* תאריכים שיש להם דוח שמור */
  const [trendData, setTrendData] = useState([]);        /* סדרת מגמה יומית לטאב "מבט על" */
  const [qualSel, setQualSel] = useState("__all__");      /* בחירת קמפיין בגרף האיכות של מבט-על */
  const [cloudUp, setCloudUp] = useState(false);          /* האם אחסון הענן זמין */
  const [cloudBusy, setCloudBusy] = useState(false);
  const [yoyData, setYoyData] = useState(null);           /* לידים: { year: { "MM-DD": {gross, quality} } } */
  const [yoyBudget, setYoyBudget] = useState(null);       /* תקציב: { year: { "MM-DD": {digital,site,total} } } */
  const [yoyLink, setYoyLink] = useState("");             /* קישור חוברת ה-YoY (כוללת את שתי הלשוניות) */
  const [yoyUpdatedAt, setYoyUpdatedAt] = useState(null);
  const [yoyMetric, setYoyMetric] = useState("quality");  /* gross | quality | qpct | cpl | cpql */
  const [overviewTab, setOverviewTab] = useState("trends"); /* trends | yoy — תת-ניווט בתוך "מבט על" */
  const [curDaily, setCurDaily] = useState(null);         /* {day: total} — לידים איכותיים יומיים של השנה, מהארכיון */
  const [userEvents, setUserEvents] = useState([]);       /* אירועים ידניים לציר ה-YoY */
  const [evName, setEvName] = useState("");
  const [evDay, setEvDay] = useState("");
  const [viewingDay, setViewingDay] = useState(null);    /* צפייה בדוח היסטורי */
  const [calPaste, setCalPaste] = useState(null);        /* {date} — הזנת דוח ליום ספציפי */
  const [calPasteText, setCalPasteText] = useState("");

  const refreshDayIndex = useCallback(async () => {
    const keys = await store.list(DAY_PREFIX);
    setDayIndex(keys.map((k) => k.slice(DAY_PREFIX.length)).sort());
  }, []);

  const refreshTrend = useCallback(async () => {
    const mainKeys = (await store.list(DAY_PREFIX)).map((k) => k.slice(DAY_PREFIX.length));
    const siteKeys = (await store.list("peres:siteday:")).map((k) => k.slice("peres:siteday:".length));
    const allDays = [...new Set([...mainKeys, ...siteKeys])].sort().slice(-32); /* עד חודש אחורה */
    const rows = [];
    for (const iso of allDays) {
      const mainRec = mainKeys.includes(iso) ? await store.get(dayKey(iso)) : null;
      const siteRec = siteKeys.includes(iso) ? await store.get("peres:siteday:" + iso) : null;
      const mt = mainRec ? reportTotals(mainRec.data) : null;
      const sSum = siteRec && siteRec.data && siteRec.data.summary ? siteRec.data.summary.actual : null;
      /* מפת קמפיינים מצטברת ליום — לבחירת קמפיין ספציפי בגרף האיכות */
      const plat = mainRec ? Object.fromEntries([
        ...(mainRec.data.platforms || []).map((x) => [x.name, { spent: x.spent || 0, quality: x.quality || 0 }]),
        ...(mainRec.data.campaigns || []).filter((c) => (c.spent || 0) > 0 || (c.quality || 0) > 0)
          .map((c) => [`חיפוש · ${c.name}`, { spent: c.spent || 0, quality: c.quality || 0 }]),
      ]) : null;
      rows.push({
        iso, label: iso.split("-").reverse().slice(0, 2).join("."),
        plat,
        budgetTarget: mt ? mt.budgetTarget : null,
        spent: mt ? Math.round(mt.spent) : null,
        degQuality: mt ? mt.degQuality : null,
        harvardSpent: mt && mt.harvardSpent !== null ? Math.round(mt.harvardSpent) : null,
        harvardQuality: mt ? mt.harvardQuality : null,
        siteQuality: sSum ? (sSum.quality || 0) : null,
        siteGross: sSum ? (sSum.gross || 0) : null,
        cpqlTotal: mt && mt.cpqlTotal ? Math.round(mt.cpqlTotal) : null,
      });
    }
    setTrendData(rows);
  }, []);

  /* הדוח הקודם = הדוח השמור בארכיון ליום הקרוב ביותר שלפני תאריך הארכוב הנוכחי.
     כך הדלתא היומית תמיד נכונה — לא משנה כמה פעמים רועננו הנתונים באותו יום */
  const findPrevSnapFor = useCallback(async (iso) => {
    const keys = await store.list(DAY_PREFIX);
    const days = keys.map((k) => k.slice(DAY_PREFIX.length)).filter((d) => d < iso).sort();
    if (!days.length) return null;
    return await store.get(dayKey(days[days.length - 1]));
  }, []);

  const archiveDateOf = (rec) => {
    if (rec && rec.archivedFor) return rec.archivedFor;
    if (rec && rec.savedAt) { const d = new Date(rec.savedAt); d.setDate(d.getDate() - 1); return isoDay(d); }
    return yesterdayIso();
  };

  /* תאריך הארכוב — ברירת מחדל: יום לפני תאריך ההעלאה. ניתן לשינוי ידני למקרים חריגים */
  const yesterdayIso = () => { const y = new Date(); y.setDate(y.getDate() - 1); return isoDay(y); };
  const [archiveDate, setArchiveDate] = useState(yesterdayIso());
  useEffect(() => { if (showConnect) setArchiveDate(yesterdayIso()); }, [showConnect]);
  const [activeMap, setActiveMap] = useState({});
  const [shareModal, setShareModal] = useState(null);

  const isActive = useCallback((name) => activeMap[name] !== false, [activeMap]);
  const toggleActive = useCallback(async (name) => {
    setActiveMap((prev) => {
      const next = { ...prev, [name]: prev[name] === false };
      store.set("peres:active", next);
      return next;
    });
  }, []);

  /* טעינת תמונת מצב — קודם מקישור משותף (hash), אחר כך מהאחסון */
  useEffect(() => {
    (async () => {
      /* בדיקת זמינות הענן — קובעת אם קוראים/כותבים אליו */
      const cloudOk = await cloud.ping();
      setCloudUp(cloudOk);
      /* קישור משותף? */
      try {
        const hash = (window.location.hash || "").replace(/^#/, "");
        const m = hash.match(/r=([\w\-_]+)/);
        if (m) {
          const st = decodeState(m[1]);
          if (st && st.data) {
            setData(st.data);
            if (st.prev) setPrevSnap(st.prev);
            if (st.activeMap) setActiveMap(st.activeMap);
            setUpdatedAt(st.savedAt ? new Date(st.savedAt) : null);
            setStatus({ kind: "ok", msg: `נטענה תמונת מצב מקישור משותף${st.savedAt ? ` (${heDate(st.savedAt)})` : ""}` });
            return;
          }
        }
      } catch {}
      const savedLink = await store.get("peres:link");
      if (savedLink) setLink(savedLink);
      const savedSiteLink = await store.get("peres:site:link");
      if (savedSiteLink) setSiteLink(savedSiteLink);
      const siteLatest = await store.get("peres:site:latest");
      if (siteLatest && siteLatest.data) { setSiteData(siteLatest.data); setSiteUpdatedAt(new Date(siteLatest.savedAt)); }
      const siteP = await store.get("peres:site:previous");
      if (siteP) setSitePrev(siteP);
      const yoyLinkSaved = await store.get("peres:yoy:link");
      if (yoyLinkSaved) setYoyLink(yoyLinkSaved);
      const yoySaved = await store.get("peres:yoy:data");
      if (yoySaved && yoySaved.byYear) { setYoyData(yoySaved.byYear); setYoyBudget(yoySaved.budget || null); setYoyUpdatedAt(new Date(yoySaved.savedAt)); }
      const evSaved = await store.get("peres:events");
      if (Array.isArray(evSaved)) setUserEvents(evSaved);
      const savedAuto = await store.get("peres:autosync");
      if (savedAuto) setAutoSync(true);
      if (savedAuto && savedLink) setPendingAutoFetch(savedLink);
      const latest = await store.get("peres:latest");
      const prev = await store.get("peres:previous");
      const savedActive = await store.get("peres:active");
      if (savedActive) setActiveMap(savedActive);
      if (latest && latest.data) {
        setData(latest.data);
        setUpdatedAt(new Date(latest.savedAt));
        setStatus({ kind: "ok", msg: `נטענה תמונת המצב האחרונה שנשמרה (${heDate(latest.savedAt)})` });
        const archPrev = await findPrevSnapFor(archiveDateOf(latest));
        setPrevSnap(archPrev || prev || null);
      } else if (prev) setPrevSnap(prev);
    })();
    refreshDayIndex();
    refreshTrend();
  }, [refreshDayIndex, refreshTrend]);

  const loadCsv = useCallback(async (text, targetDay = null, archTo = null) => {
    const parsed = parseReport(text);
    if (parsed) {
      const now = new Date().toISOString();
      if (targetDay) {
        /* הזנה ידנית ליום ספציפי מלוח השנה — נשמר לארכיון בלבד */
        await store.set(dayKey(targetDay), { savedAt: now, reportDay: targetDay, data: parsed });
        await refreshDayIndex();
        setCalPaste(null); setCalPasteText("");
        setStatus({ kind: "ok", msg: `הדוח נשמר בארכיון לתאריך ${heDayLabel(targetDay)}` });
        return;
      }
      /* archTo — סנכרון אוטומטי מעביר תאריך טרי, כדי שטאב שנשאר פתוח ימים לא יארכב לתאריך ישן */
      const yIso = archTo || (/^\d{4}-\d{2}-\d{2}$/.test(archiveDate) ? archiveDate : yesterdayIso());
      const latest = await store.get("peres:latest");
      await store.set("peres:latest", { savedAt: now, archivedFor: yIso, data: parsed });
      /* ארכוב אוטומטי: דוח שנטען היום משקף את אתמול — נשמר לתאריך יום קודם (או לתאריך שנבחר ידנית) */
      await store.set(dayKey(yIso), { savedAt: now, reportDay: yIso, data: parsed });
      await refreshDayIndex();
      refreshTrend();
      /* הדלתא נגזרת מהארכיון: מול הדוח של היום השמור הקרוב שלפני */
      const prev = await findPrevSnapFor(yIso);
      setPrevSnap(prev);
      if (prev) await store.set("peres:previous", prev);
      setData(parsed);
      setViewingDay(null);
      setUpdatedAt(new Date(now));
        setStatus({ kind: "ok", msg: (prev ? `הנתונים נטענו — הדלתות מול הדוח של ${(prev.reportDay || "").split("-").reverse().join(".") || "היום הקודם"}` : "הנתונים נטענו ונשמרו. הדלתות יופיעו כשיהיה יום קודם בארכיון") + ` · אורכב לתאריך ${yIso.split("-").reverse().join(".")}` });
      setShowConnect(false);
    } else {
      setStatus({ kind: "err", msg: "לא זוהה מבנה הדוח בקובץ. ודאו שהלשונית הנכונה מקושרת (עם טבלאות פלטפורמה/קמפיין)." });
      if (targetDay) setCalPaste({ date: targetDay, error: true });
    }
  }, [refreshDayIndex, archiveDate, findPrevSnapFor, refreshTrend]);

  /* פתיחת דוח היסטורי מהארכיון — הדלתא מחושבת מול היום השמור הקרוב שלפניו */
  const openDay = useCallback(async (iso) => {
    const snap = await store.get(dayKey(iso));
    if (!snap || !snap.data) return;
    const before = dayIndex.filter((d) => d < iso);
    const prevIso = before.length ? before[before.length - 1] : null;
    const prevDaySnap = prevIso ? await store.get(dayKey(prevIso)) : null;
    setData(snap.data);
    setPrevSnap(prevDaySnap && prevDaySnap.data ? prevDaySnap : null);
    setViewingDay(iso);
    setUpdatedAt(new Date(snap.savedAt));
    setCalOpen(false);
    setStatus({ kind: "ok", msg: `צפייה בדוח מהארכיון · ${heDayLabel(iso)}${prevIso ? ` · דלתות מול ${prevIso.split("-").reverse().slice(0, 2).join(".")}` : ""}` });
  }, [dayIndex]);

  const backToToday = useCallback(async () => {
    const latest = await store.get("peres:latest");
    const prev = latest ? await findPrevSnapFor(archiveDateOf(latest)) : await store.get("peres:previous");
    setData(latest && latest.data ? latest.data : DEMO);
    setPrevSnap(prev || null);
    setViewingDay(null);
    setUpdatedAt(latest ? new Date(latest.savedAt) : null);
    setStatus({ kind: "ok", msg: "חזרת לדוח הנוכחי" });
  }, [findPrevSnapFor]);

  /* העלאת קובץ CSV / XLSX — עובר דרך אותו loadCsv, כולל ארכוב אוטומטי ליום קודם */
  const handleFile = useCallback((e, targetDay = null) => {
    const file = e.target.files && e.target.files[0];
    e.target.value = ""; /* מאפשר לבחור שוב את אותו קובץ */
    if (!file) return;
    const isExcel = /\.(xlsx|xlsm|xls)$/i.test(file.name);
    const reader = new FileReader();
    reader.onerror = () => setStatus({ kind: "err", msg: "קריאת הקובץ נכשלה — נסו שוב" });
    if (isExcel) {
      reader.onload = () => {
        try {
          const wb = XLSX.read(reader.result, { type: "array" });
          /* עדיפות ראשונה: הלשונית "נתונים מתחילת החודש". אחרת — הלשונית הראשונה עם מבנה הדוח */
          let chosen = null;
          const exact = wb.SheetNames.find((n) => n.trim() === "נתונים מתחילת החודש");
          if (exact) chosen = XLSX.utils.sheet_to_csv(wb.Sheets[exact]);
          if (!chosen) {
            for (const name of wb.SheetNames) {
              const csv = XLSX.utils.sheet_to_csv(wb.Sheets[name]);
              if (csv.includes("פלטפורמה") && csv.includes("לידים ברוטו")) { chosen = csv; break; }
            }
          }
          if (!chosen) chosen = XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]);
          loadCsv(chosen, targetDay);
        } catch {
          setStatus({ kind: "err", msg: "לא הצלחתי לקרוא את קובץ האקסל — ודאו שהוא אינו מוגן בסיסמה" });
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      reader.onload = () => loadCsv(String(reader.result || ""), targetDay);
      reader.readAsText(file);
    }
  }, [loadCsv]);

  /* גיבוי/ייבוא ארכיון — מאפשר לסנכרן את כל הדוחות בין דפדפנים ומחשבים */
  const exportArchive = useCallback(async () => {
    const keys = await store.list("peres:");
    const dump = {};
    for (const k of keys) { const v = await store.get(k); if (v !== null) dump[k] = v; }
    const blob = new Blob(
      [JSON.stringify({ app: "peres-dashboard", exportedAt: new Date().toISOString(), data: dump })],
      { type: "application/json" }
    );
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `peres-archive-${isoDay(new Date())}.json`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 3000);
    setStatus({ kind: "ok", msg: `גיבוי הארכיון ירד כקובץ (${Object.keys(dump).length} רשומות). פתחו את הדשבורד בדפדפן אחר ולחצו "ייבוא" כדי לסנכרן.` });
  }, []);

  const importArchive = useCallback((e) => {
    const f = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!f) return;
    const r = new FileReader();
    r.onload = async () => {
      try {
        const j = JSON.parse(String(r.result));
        const entries = Object.entries(j.data || {});
        if (!entries.length) throw new Error();
        for (const [k, v] of entries) if (k.startsWith("peres:") && v !== null) await store.set(k, v);
        await refreshDayIndex();
        await backToToday();
        setStatus({ kind: "ok", msg: `הארכיון יובא בהצלחה (${entries.length} רשומות) — כל הדוחות זמינים עכשיו גם כאן` });
      } catch {
        setStatus({ kind: "err", msg: "קובץ הגיבוי לא תקין — ודאו שזה קובץ שיוצא מכפתור הגיבוי" });
      }
    };
    r.readAsText(f);
  }, [refreshDayIndex, backToToday]);

  /* ---------- חוברת האתר: מקור נתונים נפרד עם אותה לוגיקת יממה ---------- */
  const loadSiteCsv = useCallback(async (text) => {
    const parsed = parseReport(text);
    if (!parsed) {
      setStatus({ kind: "err", msg: "לא זוהה מבנה דוח בחוברת האתר — ודאו שהלשונית המקושרת בנויה באותו מבנה (פלטפורמה / לידים ברוטו)" });
      return false;
    }
    const now = new Date().toISOString();
    const yIso = yesterdayIso();
    const latest = await store.get("peres:site:latest");
    if (latest && latest.archivedFor !== yIso) {
      await store.set("peres:site:previous", latest);
      setSitePrev(latest);
    } else if (!latest) {
      setSitePrev(null);
    }
    await store.set("peres:site:latest", { savedAt: now, archivedFor: yIso, data: parsed });
    await store.set("peres:siteday:" + yIso, { savedAt: now, reportDay: yIso, data: parsed });
    setSiteData(parsed);
    setSiteUpdatedAt(new Date(now));
    return true;
  }, []);

  const fetchSite = useCallback(async (theLink, silent = false) => {
    const urls = toCsvUrls(theLink);
    if (!urls.length) { if (!silent) setStatus({ kind: "err", msg: "קישור חוברת האתר לא זוהה כקישור Google Sheets תקין" }); return; }
    if (!silent) setStatus({ kind: "load", msg: "טוען את חוברת האתר…" });

    /* שלב א': משיכת CSV לפי ה-gid שבקישור (אם יש) */
    for (const url of urls) {
      try {
        const res = await fetchWithTimeout(url, 15000);
        if (!res.ok) throw new Error("HTTP " + res.status);
        const text = await res.text();
        if (/^\s*</.test(text)) throw new Error("not-csv");
        const ok = await loadSiteCsv(text);
        if (ok) {
          await store.set("peres:site:link", theLink);
          refreshTrend();
          if (!silent) setStatus({ kind: "ok", msg: "חוברת האתר נטענה ונשמרה" });
          return;
        }
        break; /* ה-CSV נמשך אך אינו במבנה דוח — כנראה לשונית לא נכונה; עוברים לאיתור אוטומטי */
      } catch { /* מסלול הבא */ }
    }

    /* שלב ב': מושכים את כל החוברת כ-XLSX ומאתרים אוטומטית את לשונית הדוח */
    const m = String(theLink).match(/\/d\/([\w-]+)/);
    if (m && !theLink.includes("/d/e/")) {
      try {
        const res = await fetchWithTimeout(`https://docs.google.com/spreadsheets/d/${m[1]}/export?format=xlsx`, 25000);
        if (res.ok) {
          const buf = await res.arrayBuffer();
          const wb = XLSX.read(buf, { type: "array" });
          let chosen = null;
          for (const name of wb.SheetNames) {
            const csv = XLSX.utils.sheet_to_csv(wb.Sheets[name]);
            if (csv.includes("פלטפורמה") && csv.includes("לידים ברוטו")) {
              chosen = csv;
              if (name.includes("אתר") || csv.includes("- אתר")) break; /* עדיפות ללשונית האתר */
            }
          }
          if (chosen) {
            const ok = await loadSiteCsv(chosen);
            if (ok) {
              await store.set("peres:site:link", theLink);
              refreshTrend();
              if (!silent) setStatus({ kind: "ok", msg: "חוברת האתר נטענה — הלשונית אותרה אוטומטית בתוך החוברת" });
              return;
            }
          }
        }
      } catch { /* ממשיכים לשגיאה */ }
    }

    if (!silent) setStatus({ kind: "err", msg: "לא הצלחתי למשוך את חוברת האתר. שימו לב: זו חוברת נפרדת עם הרשאות משלה — השיתוף של חוברת התארים לא חל עליה. בחוברת האתר: שיתוף ← גישה כללית ← כל מי שיש לו את הקישור ← צפייה, ואז נסו שוב. (לחלופין: קובץ ← שיתוף ← פרסום באינטרנט ← בחרו את לשונית האתר + CSV, והדביקו את הקישור שנוצר בשדה חוברת האתר.)" });
  }, [loadSiteCsv, refreshTrend]);

  /* ---------- YoY: משיכת גיליון ההשוואה ההיסטורי (2025) ---------- */
  const saveYoY = useCallback(async (leads, budget) => {
    const now = new Date().toISOString();
    await store.set("peres:yoy:data", { savedAt: now, byYear: leads, budget: budget || null });
    setYoyData(leads);
    setYoyBudget(budget || null);
    setYoyUpdatedAt(new Date(now));
  }, []);

  const fetchYoY = useCallback(async (theLink, silent = false) => {
    const m = String(theLink).match(/\/d\/([\w-]+)/);
    if (!m) { if (!silent) setStatus({ kind: "err", msg: "קישור חוברת ה-YoY לא זוהה כקישור Google Sheets תקין" }); return; }
    if (!silent) setStatus({ kind: "load", msg: "טוען את חוברת ה-YoY…" });
    /* מושכים את כל החוברת כ-XLSX ומזהים אוטומטית את לשונית הלידים ואת לשונית התקציב */
    if (!theLink.includes("/d/e/")) {
      try {
        const res = await fetchWithTimeout(`https://docs.google.com/spreadsheets/d/${m[1]}/export?format=xlsx`, 25000);
        if (res.ok) {
          const wb = XLSX.read(await res.arrayBuffer(), { type: "array" });
          /* ממפים כל לשונית ל-CSV ומסווגים לפי שם + תוכן */
          const sheets = wb.SheetNames.map((name) => ({ name, csv: XLSX.utils.sheet_to_csv(wb.Sheets[name]) }));
          const isLeads = (s) => s.csv.includes("תאריך") && s.csv.includes("איכות");
          const isBudget = (s) => /cost|עלות|תקציב|spend/i.test((s.csv.split("\n")[0] || "")) || /תקציב|cost|spend/i.test(s.name);
          /* לידים: עדיפות מוחלטת ללשונית "ברוטו" (כל הלידים). רק אם אין — כל לשונית לידים אחרת */
          const bruttoSheet = sheets.find((s) => /ברוטו/.test(s.name) && isLeads(s));
          const anyLeadsSheet = sheets.find((s) => isLeads(s) && !/תקינ/.test(s.name)) || sheets.find(isLeads);
          const leadsCsv = bruttoSheet ? bruttoSheet.csv : (anyLeadsSheet ? anyLeadsSheet.csv : null);
          /* תקציב: לשונית עם "תקציב"/Cost בשם או בכותרת (ולא לשונית הלידים) */
          const budgetSheet = sheets.find((s) => isBudget(s) && !isLeads(s)) || sheets.find((s) => isBudget(s));
          const budgetCsv = budgetSheet && budgetSheet.csv !== leadsCsv ? budgetSheet.csv : null;
          const leads = leadsCsv ? parseYoY(leadsCsv) : null;
          const budget = budgetCsv ? parseYoYBudget(budgetCsv) : null;
          if (leads && Object.keys(leads).length) {
            await saveYoY(leads, budget);
            await store.set("peres:yoy:link", theLink);
            if (!silent) {
              const budgetOk = budget && Object.keys(budget).length > 0;
              setStatus({ kind: "ok", msg: budgetOk ? "חוברת ה-YoY נטענה (לידים + תקציב)" : "חוברת ה-YoY נטענה (לידים בלבד — לא נמצאה לשונית תקציב תקינה)" });
            }
            return;
          }
        }
      } catch {}
    }
    /* fallback: CSV בודד (רק לידים) */
    for (const url of toCsvUrls(theLink)) {
      try {
        const res = await fetchWithTimeout(url, 15000);
        if (!res.ok) throw new Error();
        const txt = await res.text();
        if (/^\s*</.test(txt)) throw new Error();
        const leads = parseYoY(txt);
        if (leads && Object.keys(leads).length) {
          await saveYoY(leads, null);
          await store.set("peres:yoy:link", theLink);
          if (!silent) setStatus({ kind: "ok", msg: "חוברת ה-YoY נטענה (לידים בלבד)" });
          return;
        }
      } catch {}
    }
    if (!silent) setStatus({ kind: "err", msg: "לא הצלחתי למשוך את חוברת ה-YoY — ודאו שהיא משותפת לצפייה, ושיש בה לשונית לידים (תאריך/מקור/איכות) ולשונית תקציב (Date/Cost)" });
  }, [saveYoY]);

  const fetchYoYRef = React.useRef(null);
  fetchYoYRef.current = fetchYoY;

  const fetchSiteRef = React.useRef(null);
  fetchSiteRef.current = fetchSite;

  /* סנכרון יומי מתוזמן: כל עוד הדשבורד פתוח, מדי דקה נבדק אם עברה השעה 10:00 והיום טרם סונכרן */
  useEffect(() => {
    if (!autoSync || !link) return;
    const tick = async () => {
      const now = new Date();
      if (now.getHours() < 10) return;
      const today = isoDay(now);
      const last = await store.get("peres:lastAutoSyncDay");
      if (last === today) return;
      await store.set("peres:lastAutoSyncDay", today);
      if (fetchFromLinkRef.current) fetchFromLinkRef.current(link, true);
      if (fetchSiteRef.current && siteLink) fetchSiteRef.current(siteLink, true);
      if (fetchYoYRef.current && yoyLink) fetchYoYRef.current(yoyLink, true);
    };
    tick();
    const id = setInterval(tick, 60000);
    return () => clearInterval(id);
  }, [autoSync, link, siteLink, yoyLink]);

  /* סנכרון אוטומטי בפתיחה — רץ פעם אחת כשיש קישור שמור והאפשרות מופעלת */
  useEffect(() => {
    if (pendingAutoFetch && fetchFromLinkRef.current) {
      const l = pendingAutoFetch;
      setPendingAutoFetch(null);
      fetchFromLinkRef.current(l, true);
      if (fetchSiteRef.current && siteLink) fetchSiteRef.current(siteLink, true);
      if (fetchYoYRef.current && yoyLink) fetchYoYRef.current(yoyLink, true);
    }
  }, [pendingAutoFetch, siteLink, yoyLink]);

  const fetchFromLink = useCallback(async (theLink, silent = false) => {
    const urls = toCsvUrls(theLink);
    if (!urls.length) { if (!silent) setStatus({ kind: "err", msg: "הקישור לא זוהה כקישור Google Sheets תקין" }); return; }
    setLoading(true);
    setStatus({ kind: "load", msg: silent ? "מסנכרן אוטומטית מהגיליון…" : "טוען נתונים מהגיליון…" });
    let ok = false;
    for (const url of urls) {
      try {
        const res = await fetchWithTimeout(url, 15000);
        if (!res.ok) throw new Error("HTTP " + res.status);
        const text = await res.text();
        /* אם חזר HTML (עמוד התחברות של גוגל) — הגיליון לא נגיש, מנסים מסלול הבא */
        if (/^\s*</.test(text)) throw new Error("not-csv");
        await loadCsv(text, null, silent ? yesterdayIso() : null);
        await store.set("peres:link", theLink);
        ok = true;
        break;
      } catch { /* מנסים את המסלול הבא */ }
    }
    setLoading(false);
    if (!ok) {
      setStatus({
        kind: "err",
        msg: silent
          ? "הסנכרון האוטומטי נכשל — ודאו שהגיליון משותף (כל מי שיש לו קישור → צפייה) או מפורסם כ-CSV (קובץ ← שיתוף ← פרסום באינטרנט)"
          : "לא הצלחתי למשוך את הגיליון (הבקשה נחסמה או שאין הרשאה). שתי דרכים לפתור: (1) בגיליון: שיתוף ← כל מי שיש לו קישור ← צפייה, ואז נסו שוב. (2) הדרך האמינה ביותר: קובץ ← שיתוף ← פרסום באינטרנט ← בחרו את הלשונית + CSV, והדביקו כאן את הקישור שנוצר. לחלופין — העלו קובץ או הדביקו ידנית.",
      });
      if (!silent) setPasteMode(true);
    }
  }, [loadCsv]);

  const fetchFromLinkRef = React.useRef(null);
  fetchFromLinkRef.current = fetchFromLink;

  const fetchSheet = useCallback(() => fetchFromLink(link), [link, fetchFromLink]);

  const fullDelta = useMemo(() => computeDelta(data, prevSnap), [data, prevSnap]);

  const siteDelta = useMemo(() => (siteData ? computeDelta(siteData, sitePrev) : null), [siteData, sitePrev]);

  /* סדרת "לידים איכותיים יומיים של השנה" מהארכיון: תקינים(יום) = מצטבר(יום) − מצטבר(יום קודם).
     תקינים מצטברים ליום = תארים(שורת Total ב-BOF) + הארווארד + אתר. */
  useEffect(() => {
    if (!(board === "summary" && overviewTab === "yoy")) return;
    (async () => {
      const rd = new Date(); rd.setDate(rd.getDate() - 1);
      const y = rd.getFullYear(), mo = rd.getMonth() + 1, mm = String(mo).padStart(2, "0");
      const daysInM = new Date(y, mo, 0).getDate();
      const cumOfDay = async (iso) => {
        const main = await store.get(dayKey(iso));
        const site = await store.get("peres:siteday:" + iso);
        if (!main && !site) return null;
        let q = 0, b = 0;
        if (main && main.data) {
          const rt = reportTotals(main.data);
          if (rt) { q += (rt.degQuality || 0) + (rt.harvardQuality || 0); b += (rt.spent || 0) + (rt.harvardSpent || 0); }
        }
        if (site && site.data && site.data.summary) q += (site.data.summary.actual.quality || 0); /* לאתר אין תקציב */
        return { q, b };
      };
      const cumByDay = {};
      for (let d = 1; d <= daysInM; d++) {
        const iso = `${y}-${mm}-${String(d).padStart(2, "0")}`;
        cumByDay[d] = await cumOfDay(iso);
      }
      /* מצטבר רציף של 2026 + יומי + עלות לליד איכותי מצטברת */
      const daily = {}, cumCur = {}, cpqlCur = {}; let lastQ = null;
      for (let d = 1; d <= daysInM; d++) {
        const rec = cumByDay[d];
        if (rec == null) { daily[d] = null; cumCur[d] = lastQ; cpqlCur[d] = null; continue; }
        cumCur[d] = rec.q;
        cpqlCur[d] = rec.q > 0 ? Math.round(rec.b / rec.q) : null;
        if (lastQ == null) { daily[d] = null; lastQ = rec.q; continue; }
        daily[d] = Math.max(rec.q - lastQ, 0);
        lastQ = rec.q;
      }
      setCurDaily({ daily, cum: cumCur, cpql: cpqlCur });
    })();
  }, [board, overviewTab, updatedAt, siteUpdatedAt, dayIndex]);

  /* ה"חוברת" הפעילה: תארים / הארווארד / אתר — לכל אחת נתונים, דלתות, תובנות וסטורי משלה */
  const view = useMemo(() => {
    const hasBudget = (d, tr) => (tr && tr.budget) || (d.platforms || []).some((x) => (x.budget || 0) > 0 || (x.spent || 0) > 0);
    if (board === "site" && siteData) {
      return { data: siteData, delta: siteDelta, label: "אתר", isHarvard: false, totalsRow: siteData.totalsRow || null, noBudget: !hasBudget(siteData, siteData.totalsRow) };
    }
    if (board === "harvard" && data.harvard) {
      const H = data.harvard;
      const t = H.total || H.platforms.reduce(
        (a, b) => ({ spent: (a.spent || 0) + (b.spent || 0), gross: (a.gross || 0) + (b.gross || 0), quality: (a.quality || 0) + (b.quality || 0), inProc: (a.inProc || 0) + (b.inProc || 0) }),
        { spent: 0, gross: 0, quality: 0, inProc: 0 }
      );
      const vData = {
        meta: data.meta,
        summary: {
          actual: {
            gross: t.gross, quality: t.quality,
            qualityPct: (t.gross || 0) > 0 && t.quality !== null ? (t.quality / t.gross) * 100 : null,
            cpl: t.cpl ?? ((t.gross || 0) > 0 ? t.spent / t.gross : null),
            cpql: t.cpql ?? ((t.quality || 0) > 0 ? t.spent / t.quality : null),
            inProcess: t.inProc,
          },
          target: { gross: null, quality: null, qualityPct: null, cpl: null, cpql: null },
        },
        platforms: H.platforms,
        campaigns: [],
      };
      const vDelta = fullDelta
        ? { date: fullDelta.date, summary: fullDelta.harvardSummary || {}, platforms: fullDelta.harvard || {}, campaigns: {} }
        : null;
      return { data: vData, delta: vDelta, label: "הארווארד", isHarvard: true, totalsRow: H.total, noBudget: false };
    }
    /* חוברת התארים: הבלוק העליון בגיליון מסכם תארים+הארווארד יחד.
       כאן בונים סיכום תארים-בלבד משורת ה-Total של טבלת ה-BOF (שכבר מנקה את הארווארד),
       ושומרים את היעדים מהבלוק העליון. */
    const degSummary = (() => {
      const base = data.summary;
      const tr = data.totalsRow;
      if (!base || !tr || tr.gross == null || tr.quality == null) return base;
      const gross = tr.gross, quality = tr.quality, spent = tr.spent;
      /* לידים בתהליך של התארים = סכום עמודת "בתהליך" בטבלת ה-BOF (ללא הארווארד) */
      const inProc = (data.platforms || []).reduce((a, b) => a + (b.inProc || 0), 0);
      return {
        actual: {
          gross, quality,
          qualityPct: gross > 0 ? (quality / gross) * 100 : null,
          cpl: spent != null && gross > 0 ? spent / gross : (base.actual.cpl ?? null),
          cpql: spent != null && quality > 0 ? spent / quality : (base.actual.cpql ?? null),
          inProcess: inProc || base.actual.inProcess,
        },
        target: base.target,
      };
    })();
    const degData = { ...data, summary: degSummary };
    /* דלתא תארים-בלבד: משורת ה-Total של ה-BOF מול אותה שורה בדוח הקודם */
    let degDelta = fullDelta;
    if (fullDelta && prevSnap && prevSnap.data) {
      const pd = prevSnap.data;
      const ptr = pd.totalsRow;
      if (ptr && ptr.gross != null && data.totalsRow) {
        const tr = data.totalsRow;
        const pInProc = (pd.platforms || []).reduce((a, b) => a + (b.inProc || 0), 0);
        degDelta = {
          ...fullDelta,
          summary: {
            ...fullDelta.summary,
            gross: (tr.gross ?? 0) - (ptr.gross ?? 0),
            quality: (tr.quality ?? 0) - (ptr.quality ?? 0),
            spent: (tr.spent ?? 0) - (ptr.spent ?? 0),
            inProcess: (degSummary.actual.inProcess ?? 0) - pInProc,
            qualityPct: (degSummary.actual.qualityPct ?? 0) -
              ((ptr.gross > 0 ? (ptr.quality / ptr.gross) * 100 : 0)),
            cpl: (degSummary.actual.cpl ?? 0) - (ptr.spent != null && ptr.gross > 0 ? ptr.spent / ptr.gross : 0),
            cpql: (degSummary.actual.cpql ?? 0) - (ptr.spent != null && ptr.quality > 0 ? ptr.spent / ptr.quality : 0),
          },
        };
      }
    }
    return { data: degData, delta: degDelta, label: "תארים", isHarvard: false, totalsRow: data.totalsRow || null, noBudget: false };
  }, [data, fullDelta, prevSnap, board, siteData, siteDelta]);

  const delta = view.delta;
  const s = view.data.summary;

  const totals = useMemo(() => {
    const p = view.data.platforms;
    const sumBudget = p.reduce((s, x) => s + (x.budget || 0), 0);
    const sumSpent = p.reduce((s, x) => s + (x.spent || 0), 0);
    /* יעד התקציב נמשך משורת ה-Total של הטבלה בגיליון (מקור האמת); סכימת השורות היא גיבוי בלבד */
    const tr = view.totalsRow;
    const budget = tr && tr.budget ? tr.budget : sumBudget;
    const spent = tr && tr.spent ? tr.spent : sumSpent;
    return { budget, spent, util: budget ? (spent / budget) * 100 : 0 };
  }, [view]);

  const insights = useMemo(() => buildInsights(view.data, view.delta, isActive), [view, isActive]);
  /* שורות ביצועי היממה — כלים + קמפייני חיפוש מתוך הדלתא */
  const dayRows = useMemo(() => {
    if (!view.delta) return null;
    const rows = [];
    Object.entries(view.delta.platforms).forEach(([name, d]) => rows.push({ name, type: "כלי", ...d }));
    Object.entries(view.delta.campaigns || {}).forEach(([name, d]) => rows.push({ name, type: "חיפוש", ...d }));
    return rows.filter((r) => (r.spent || 0) !== 0 || (r.gross || 0) !== 0 || (r.quality || 0) !== 0);
  }, [view]);

  const sortedPlatforms = useMemo(() => {
    const arr = [...view.data.platforms];
    arr.sort((a, b) => {
      const va = a[sort.key], vb = b[sort.key];
      if (va === null || va === undefined) return 1;
      if (vb === null || vb === undefined) return -1;
      return sort.dir === "desc" ? vb - va : va - vb;
    });
    return arr;
  }, [view, sort]);

  const pieData = useMemo(() => {
    if (view.noBudget) return view.data.platforms.filter((p) => (p.gross || 0) > 0).map((p) => ({ name: p.name, value: p.gross }));
    return view.data.platforms.filter((p) => p.spent > 0).map((p) => ({ name: p.name, value: p.spent }));
  }, [view]);
  const pieTotal = useMemo(() => pieData.reduce((s, x) => s + x.value, 0), [pieData]);
  const barData = useMemo(
    () => view.data.platforms.filter((p) => p.budget > 0 || p.spent > 0).map((p) => ({ name: p.name, "תקציב": p.budget, "מומש": p.spent })),
    [view]
  );



  /* ---------- שיתוף ---------- */
  const buildShareLink = useCallback(() => {
    const state = {
      v: 1,
      savedAt: (updatedAt || new Date()).toISOString(),
      data,
      prev: prevSnap,
      activeMap,
    };
    return currentBaseUrl() + "#r=" + encodeState(state);
  }, [data, prevSnap, activeMap, updatedAt]);

  const shareOut = useCallback(async (title, text, withLink = true) => {
    const url = withLink ? buildShareLink() : null;
    try {
      if (navigator.share) { await navigator.share(url ? { title, text, url } : { title, text }); return; }
    } catch (e) {
      if (e && e.name === "AbortError") return; /* המשתמש ביטל — לא שגיאה */
    }
    let copied = false;
    const toCopy = url || text;
    try { await navigator.clipboard.writeText(toCopy); copied = true; } catch {}
    if (!copied) {
      try {
        const ta = document.createElement("textarea");
        ta.value = toCopy;
        ta.style.position = "fixed"; ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.focus(); ta.select();
        copied = document.execCommand("copy");
        document.body.removeChild(ta);
      } catch {}
    }
    setShareModal({ title, text, url, copied });
  }, [buildShareLink]);

  const shareDashboard = useCallback(() => {
    const d = new Date().toLocaleDateString("he-IL", { day: "numeric", month: "long" });
    const vsT = (a, t) => (t ? ` / יעד ${num(t)} (${Math.round((a / t) * 100)}%)` : "");
    const lines = [
      `📊 המרכז האקדמי פרס · ${view.label} · דשבורד יומי · ${d}`,
      ``,
      ...(totals.budget > 0 ? [`💰 תקציב: ${nis(totals.spent)} מתוך ${nis(totals.budget)} (${Math.round(totals.util)}%)`] : []),
      `📥 לידים ברוטו: ${num(s.actual.gross)}${vsT(s.actual.gross, s.target.gross)} · ${nis(s.actual.cpl)} לליד`,
      `⭐ איכותיים: ${num(s.actual.quality)}${vsT(s.actual.quality, s.target.quality)} · ${nis(s.actual.cpql)} לליד · ${pct(s.actual.qualityPct, 1)} איכות`,
      `⏳ בתהליך: ${num(s.actual.inProcess)} לידים`,
    ];
    if (delta) {
      lines.push(``, `🔄 מאז הדוח הקודם (${heDate(delta.date)}):`);
      if (delta.summary.spent) lines.push(`   מומש: ${nis(delta.summary.spent)}+`);
      if (delta.summary.gross) lines.push(`   לידים חדשים: ${num(delta.summary.gross)}+`);
      if (delta.summary.quality) lines.push(`   איכותיים חדשים: ${num(delta.summary.quality)}+`);
    }
    const top3 = [...view.data.platforms].sort((a, b) => (b.spent || 0) - (a.spent || 0)).slice(0, 3);
    lines.push(``, `🏗️ הכלים המובילים במימוש:`);
    top3.forEach((p) => lines.push(totals.budget > 0
      ? `   ${p.name}: ${nis(p.spent)} · ${num(p.gross)} לידים · ${num(p.quality)} איכותיים`
      : `   ${p.name}: ${num(p.gross)} לידים · ${num(p.quality)} איכותיים`));
    if (insights.length) lines.push(``, `💡 ${insights[0].text}`);
    shareOut(`דשבורד יומי · ${view.label} — המרכז האקדמי פרס`, lines.join("\n"));
  }, [view, totals, delta, insights, s, shareOut]);

  const shareStory = useCallback(() => {
    const slides = buildSlides(view.data, totals, view.delta, insights, isActive, view.label);
    const lines = [`⚡ הסטורי היומי · ${view.label} · המרכז האקדמי פרס · ${new Date().toLocaleDateString("he-IL")}`, ``];
    slides.forEach((sl) => {
      lines.push(`◾ ${sl.kicker}: ${sl.big}`);
      if (sl.text) lines.push(`   ${sl.text}`);
      (sl.rows || []).forEach((r) => lines.push(`   • ${r.label}: ${r.value}`));
      lines.push(``);
    });
    shareOut(`הסטורי היומי · ${view.label} — המרכז האקדמי פרס`, lines.join("\n"));
  }, [view, totals, insights, isActive, shareOut]);

  return (
    <div dir="rtl" className="root">
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;700;900&display=swap');
        * { box-sizing: border-box; margin: 0; scrollbar-width: thin; scrollbar-color: ${C.panelSoft} transparent; }
        ::-webkit-scrollbar { height: 8px; width: 8px; }
        ::-webkit-scrollbar-track { background: transparent; }
        ::-webkit-scrollbar-thumb { background: ${C.panelSoft}; border-radius: 8px; border: 1px solid ${C.line}; }
        ::-webkit-scrollbar-thumb:hover { background: ${C.line}; }
        ::-webkit-scrollbar-corner { background: transparent; }
        .root { min-height: 100vh; background: ${C.bg}; color: ${C.text}; font-family: 'Heebo', system-ui, sans-serif; padding: 24px clamp(12px, 3vw, 40px) 60px; font-feature-settings: 'tnum'; }
        .head { display:flex; flex-wrap:wrap; gap:16px; align-items:flex-end; justify-content:space-between; margin-bottom: 6px; }
        h1 { font-size: clamp(20px, 2.6vw, 30px); font-weight: 900; letter-spacing: -0.5px; }
        h1 span { color: ${C.amber}; }
        .sub { color: ${C.dim}; font-size: 13px; margin-top: 4px; }
        .btn { background: ${C.amber}; color:#1a1206; border:none; border-radius:10px; padding:9px 18px; font-family:inherit; font-weight:700; font-size:14px; cursor:pointer; }
        .btn:hover { filter: brightness(1.08); }
        .btn.ghost { background:transparent; color:${C.dim}; border:1px solid ${C.line}; }
        .file-btn { display:inline-flex; align-items:center; gap:6px; }
        .file-btn input[type="file"] { display:none; }
        .btn.story-btn { background: linear-gradient(90deg, ${C.violet}, ${C.blue}); color:#fff; }
        .btn:disabled { opacity:.5; cursor:default; }
        .board-tabs { display:flex; align-items:center; gap:8px; margin-top:14px; flex-wrap:wrap; }
        .tab { background:${C.panel}; border:1px solid ${C.line}; color:${C.dim}; border-radius:12px 12px 0 0; padding:9px 22px; font-family:inherit; font-size:14.5px; font-weight:700; cursor:pointer; border-bottom:none; }
        .tab.on { background:${C.panelSoft}; color:${C.amber}; border-color:${C.amber}; }
        .tab:hover:not(.on) { color:${C.text}; }
        .tab-meta { color:${C.dim}; font-size:12px; margin-right:8px; }
        .subtab { background:transparent; border:1px solid ${C.line}; color:${C.dim}; border-radius:20px; padding:6px 18px; font-family:inherit; font-size:13px; font-weight:700; cursor:pointer; }
        .subtab.on { background:rgba(245,184,65,.12); color:${C.amber}; border-color:${C.amber}; }
        .subtab:hover:not(.on) { color:${C.text}; }
        .status { font-size:13px; padding:8px 14px; border-radius:10px; margin:14px 0 20px; border:1px solid ${C.line}; background:${C.panel}; color:${C.dim}; }
        .status.err { border-color:${C.coral}; color:${C.coral}; }
        .status.ok { border-color:${C.teal}; color:${C.teal}; }
        .connect { background:${C.panel}; border:1px solid ${C.line}; border-radius:14px; padding:18px; margin-bottom:22px; }
        .connect input, .connect textarea { width:100%; background:${C.bg}; border:1px solid ${C.line}; color:${C.text}; border-radius:10px; padding:10px 12px; font-family:inherit; font-size:14px; margin:8px 0; direction:ltr; }
        .connect textarea { min-height:110px; font-size:12px; }
        .archive-date-row { display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin:2px 0 10px; font-size:13px; }
        .autosync { display:flex; align-items:center; gap:8px; font-size:13px; color:${C.text}; margin:2px 0 10px; cursor:pointer; }
        .autosync input { accent-color:${C.amber}; width:16px; height:16px; cursor:pointer; }
        .archive-date-row input[type="date"] { width:auto; background:${C.bg}; border:1px solid ${C.line}; color:${C.text}; border-radius:8px; padding:6px 10px; font-family:inherit; font-size:13px; margin:0; color-scheme:dark; }
        .grid-kpi { display:grid; grid-template-columns:repeat(auto-fit, minmax(190px, 1fr)); gap:14px; margin-bottom:26px; }
        .kpi { background:${C.panel}; border:1px solid ${C.line}; border-radius:14px; padding:16px 18px 14px; }
        .kpi-title { color:${C.dim}; font-size:13px; font-weight:500; }
        .kpi-value { font-size:28px; font-weight:900; margin:4px 0 2px; letter-spacing:-0.5px; display:flex; align-items:baseline; flex-wrap:wrap; }
        .kpi-target { color:${C.dim}; font-size:12px; }
        .kpi-bar-wrap { position:relative; height:6px; background:${C.panelSoft}; border-radius:4px; margin-top:12px; overflow:hidden; }
        .kpi-bar { height:100%; border-radius:4px; transition:width .6s ease; }
        .kpi-goal-line { position:absolute; top:-2px; bottom:-2px; width:2px; background:${C.text}; opacity:.55; }
        .kpi-pct { font-size:12px; font-weight:700; margin-top:6px; }
        .two-col { display:grid; grid-template-columns: 1.4fr 1fr; gap:16px; margin-bottom:26px; }
        @media (max-width: 900px){ .two-col { grid-template-columns:1fr; } }
        .panel { background:${C.panel}; border:1px solid ${C.line}; border-radius:14px; padding:18px; }
        .panel h2 { font-size:16px; font-weight:700; margin-bottom:4px; }
        .panel .hint { color:${C.dim}; font-size:12px; margin-bottom:12px; }
        table { width:100%; border-collapse:collapse; font-size:13.5px; }
        th { text-align:right; color:${C.dim}; font-weight:500; font-size:12px; padding:8px 10px; border-bottom:1px solid ${C.line}; white-space:nowrap; user-select:none; }
        td { padding:9px 10px; border-bottom:1px solid ${C.panelSoft}; white-space:nowrap; }
        tr.click { cursor:pointer; }
        tr.click:hover td { background:${C.panelSoft}; }
        .qdot { display:inline-block; width:8px; height:8px; border-radius:50%; margin-left:6px; vertical-align:middle; }
        .util-cell { display:flex; align-items:center; gap:8px; }
        .util-track { flex:1; min-width:70px; height:5px; background:${C.panelSoft}; border-radius:3px; overflow:hidden; }
        .util-fill { height:100%; border-radius:3px; }
        .group-row td { background:${C.panelSoft}; font-size:13px; font-weight:900; border-bottom:1px solid ${C.line}; padding-top:11px; padding-bottom:11px; }
        .qual-select { background:${C.bg}; border:1px solid ${C.line}; color:${C.text}; border-radius:9px; padding:7px 12px; font-family:inherit; font-size:13px; font-weight:700; cursor:pointer; max-width:60%; }
        .qual-select:hover { border-color:${C.amber}; }
        .qual-select option { background:${C.panel}; color:${C.text}; }
        .tag { font-size:11px; background:${C.panelSoft}; color:${C.amber}; border-radius:6px; padding:2px 8px; margin-right:8px; font-weight:700; }
        .drill { margin-top:-10px; margin-bottom:26px; border:1px solid ${C.line}; border-top:none; border-radius:0 0 14px 14px; background:${C.panelSoft}; padding:14px 18px 18px; }
        .table-wrap { overflow-x:auto; }
        .foot { color:${C.dim}; font-size:12px; text-align:center; margin-top:30px; }
        .legend { display:flex; flex-direction:column; gap:7px; font-size:12.5px; max-height:300px; overflow:auto; }
        .legend-row { display:flex; align-items:center; gap:8px; }
        .legend-dot { width:10px; height:10px; border-radius:3px; flex-shrink:0; }
        .legend-name { flex:1; color:${C.text}; }
        .legend-val { color:${C.dim}; }
        .pie-flex { display:flex; gap:14px; align-items:center; }
        @media (max-width:600px){ .pie-flex { flex-direction:column; } }
        .insight-list { display:flex; flex-direction:column; gap:10px; }
        .insight { display:flex; gap:12px; align-items:flex-start; background:${C.panelSoft}; border-radius:12px; padding:12px 14px; font-size:14px; line-height:1.55; border-right:3px solid ${C.line}; }
        .insight.good { border-right-color:${C.teal}; }
        .insight.bad { border-right-color:${C.coral}; }
        .insight.warn { border-right-color:${C.amber}; }
        .insight .ico { font-size:18px; line-height:1.3; }
        .story-overlay { position:fixed; inset:0; background:rgba(5,10,20,.85); backdrop-filter:blur(6px); display:flex; align-items:center; justify-content:center; z-index:100; }
        .story-card { position:relative; width:min(430px, 94vw); height:min(760px, 92vh); border-radius:22px; border:1px solid ${C.line}; overflow:hidden; display:flex; flex-direction:column; box-shadow:0 30px 80px rgba(0,0,0,.6); }
        .story-progress { display:flex; gap:5px; padding:14px 16px 0; }
        .story-seg { flex:1; height:3px; background:rgba(255,255,255,.18); border-radius:2px; overflow:hidden; }
        .story-seg-fill { height:100%; background:#fff; }
        @keyframes fill { from { width:0%; } to { width:100%; } }
        .story-seg-fill[data-active="1"] { animation: fill 8s linear forwards; }
        .story-close { position:absolute; top:24px; left:16px; background:rgba(255,255,255,.12); border:none; color:#fff; width:32px; height:32px; border-radius:50%; cursor:pointer; font-size:14px; z-index:3; }
        .story-actions { position:absolute; top:24px; right:16px; display:flex; gap:8px; z-index:3; }
        .story-share { background:rgba(255,255,255,.12); border:none; color:#fff; height:32px; padding:0 14px; border-radius:16px; cursor:pointer; font-size:13px; font-family:inherit; font-weight:700; }
        .story-share:hover { background:rgba(255,255,255,.22); }
        .story-paused-tag { position:absolute; top:64px; right:16px; z-index:3; font-size:11.5px; color:${C.amber}; background:rgba(245,184,65,.12); border:1px solid rgba(245,184,65,.4); border-radius:8px; padding:3px 10px; }
        .switch { position:relative; width:36px; height:19px; border-radius:12px; padding:0; cursor:pointer; margin-left:10px; vertical-align:middle; transition:background .25s, border-color .25s, box-shadow .25s; }
        .switch::after { content:""; position:absolute; top:2px; width:13px; height:13px; border-radius:50%; transition:right .25s, background .25s, box-shadow .25s; }
        .switch.on { background:rgba(55,201,169,.28); border:1px solid ${C.teal}; box-shadow:0 0 10px rgba(55,201,169,.25); }
        .switch.on::after { right:3px; background:${C.teal}; box-shadow:0 0 6px rgba(55,201,169,.8); }
        .switch.off { background:${C.panelSoft}; border:1px solid ${C.line}; }
        .switch.off::after { right:19px; background:${C.dim}; }
        .switch:hover { border-color:${C.amber}; }
        .share-overlay { position:fixed; inset:0; background:rgba(5,10,20,.8); backdrop-filter:blur(4px); display:flex; align-items:center; justify-content:center; z-index:200; padding:16px; }
        .share-modal { background:${C.panel}; border:1px solid ${C.line}; border-radius:16px; padding:18px; width:min(560px, 96vw); display:flex; flex-direction:column; gap:12px; box-shadow:0 30px 80px rgba(0,0,0,.6); }
        .share-head { display:flex; justify-content:space-between; align-items:center; }
        .share-copied { font-size:13px; font-weight:700; }
        .share-text { width:100%; min-height:180px; background:${C.bg}; border:1px solid ${C.line}; color:${C.text}; border-radius:10px; padding:12px; font-family:inherit; font-size:13px; line-height:1.6; resize:vertical; direction:rtl; }
        .share-link-row { display:flex; gap:8px; }
        .share-link { flex:1; background:${C.bg}; border:1px solid ${C.line}; color:${C.blue}; border-radius:10px; padding:10px 12px; font-size:12px; font-family:monospace; min-width:0; }
        .hist-banner { display:flex; align-items:center; justify-content:space-between; gap:12px; flex-wrap:wrap; background:rgba(154,123,247,.1); border:1px solid ${C.violet}; border-radius:12px; padding:10px 16px; margin:-8px 0 20px; font-size:14px; }
        .cal-cta { display:flex; align-items:center; justify-content:space-between; gap:14px; cursor:pointer; transition:border-color .15s; }
        .cal-cta:hover { border-color:${C.amber}; }
        .cal-cta-arrow { color:${C.amber}; font-size:20px; }
        .cloud-pill { font-size:11.5px; font-weight:700; border-radius:8px; padding:5px 10px; border:1px solid ${C.line}; }
        .cloud-pill.up { color:${C.teal}; border-color:rgba(55,201,169,.45); background:rgba(55,201,169,.1); }
        .cloud-pill.down { color:${C.amber}; border-color:rgba(245,184,65,.4); background:rgba(245,184,65,.1); }
        .cal-overlay { position:fixed; inset:0; background:rgba(5,10,20,.85); backdrop-filter:blur(6px); display:flex; align-items:center; justify-content:center; z-index:150; padding:14px; }
        .cal-sheet { background:${C.panel}; border:1px solid ${C.line}; border-radius:18px; padding:18px; width:min(1060px, 96vw); max-height:92vh; overflow-y:auto; box-shadow:0 30px 80px rgba(0,0,0,.6); }
        .cal-head { display:flex; align-items:center; justify-content:space-between; margin-bottom:8px; }
        .cal-year-nav { display:flex; align-items:center; gap:12px; }
        .cal-year { font-size:24px; font-weight:900; color:${C.amber}; letter-spacing:1px; }
        .cal-legend { display:flex; gap:20px; flex-wrap:wrap; color:${C.dim}; font-size:12px; margin-bottom:14px; }
        .cal-legend > span { display:flex; align-items:center; gap:6px; }
        .cal-dot { width:9px; height:9px; border-radius:50%; background:${C.panelSoft}; border:1px solid ${C.line}; display:inline-block; }
        .cal-dot.has { background:${C.teal}; border-color:${C.teal}; }
        .cal-months { display:grid; grid-template-columns:repeat(auto-fill, minmax(230px, 1fr)); gap:14px; }
        .cal-month { background:${C.bg}; border:1px solid ${C.line}; border-radius:12px; padding:12px; }
        .cal-month-name { font-weight:900; font-size:14px; margin-bottom:8px; color:${C.text}; }
        .cal-grid { display:grid; grid-template-columns:repeat(7, 1fr); gap:3px; }
        .cal-dow { text-align:center; color:${C.dim}; font-size:10.5px; padding-bottom:3px; }
        .cal-day { aspect-ratio:1; border:none; background:transparent; color:${C.dim}; font-family:inherit; font-size:11.5px; border-radius:7px; cursor:pointer; display:flex; align-items:center; justify-content:center; }
        .cal-day:hover { background:${C.panelSoft}; color:${C.text}; }
        .cal-day.has { background:rgba(55,201,169,.18); color:${C.teal}; font-weight:900; border:1px solid rgba(55,201,169,.5); }
        .cal-day.has:hover { background:rgba(55,201,169,.32); }
        .cal-day.today { outline:2px solid ${C.amber}; outline-offset:-2px; }
        .story-rows { display:flex; flex-direction:column; gap:10px; margin-top:4px; }
        .story-row { display:flex; flex-direction:column; gap:2px; background:rgba(255,255,255,.05); border:1px solid ${C.line}; border-radius:12px; padding:10px 14px; }
        .story-row-label { color:${C.dim}; font-size:12.5px; font-weight:700; }
        .story-row-value { font-size:15.5px; font-weight:700; line-height:1.4; }
        .story-body { flex:1; display:flex; flex-direction:column; justify-content:center; padding:14px 30px 30px; gap:14px; text-align:right; overflow-y:auto; }
        .story-kicker { color:${C.dim}; font-size:15px; font-weight:700; letter-spacing:.5px; }
        .story-big { font-size:clamp(34px, 8vw, 52px); font-weight:900; letter-spacing:-1px; line-height:1.05; }
        .story-text { font-size:19px; line-height:1.55; color:${C.text}; }
        .story-foot { font-size:13px; color:${C.dim}; border-top:1px solid ${C.line}; padding-top:12px; }
        .story-nav { position:absolute; top:0; bottom:0; width:38%; z-index:2; cursor:pointer; }
        .story-nav.prev { right:0; } .story-nav.next { left:0; }
        .story-count { position:absolute; bottom:14px; left:0; right:0; text-align:center; color:${C.dim}; font-size:12px; }
        @media (prefers-reduced-motion: reduce){ .kpi-bar{ transition:none; } .story-seg-fill[data-active="1"]{ animation:none; width:30%; } }
      `}</style>

      {/* ---------- כותרת ---------- */}
      <div className="head">
        <div>
          <h1>המרכז האקדמי פרס · <span>{view.isHarvard ? "הארווארד" : board === "site" ? "אתר" : board === "summary" ? (overviewTab === "yoy" ? "השוואת שנים" : "מבט על") : "דשבורד יומי"}</span></h1>
          <div className="sub">
            {board === "site"
              ? (siteUpdatedAt ? `חוברת האתר עודכנה: ${heDate(siteUpdatedAt)}` : "חוברת האתר — טרם נטענה")
              : (updatedAt ? `עודכן לאחרונה: ${heDate(updatedAt)}` : "מקור: דוח ריכוז נתונים — Google Sheets")}
            {delta && ` · דלתות מול הדוח מ־${heDate(delta.date)}`}
          </div>
        </div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button className="btn story-btn" onClick={() => setStory(true)}>▶ הסטורי היומי</button>
          <button className="btn ghost" onClick={shareDashboard}>📤 שיתוף הדשבורד</button>
          {link && <button className="btn ghost" disabled={loading} onClick={fetchSheet}>{loading ? "⏳ מסנכרן…" : "🔄 סנכרון עכשיו"}</button>}
          <button className="btn" onClick={() => setShowConnect((v) => !v)}>
            {showConnect ? "סגירה" : "חיבור הגיליון"}
          </button>
        </div>
      </div>

      <div className="board-tabs">
        <button className={`tab ${board === "summary" ? "on" : ""}`} onClick={() => { refreshTrend(); setBoard("summary"); }}>📈 מבט על</button>
        <button className={`tab ${board === "main" ? "on" : ""}`} onClick={() => setBoard("main")}>🎓 תארים</button>
        <button className={`tab ${board === "harvard" ? "on" : ""}`} onClick={() => setBoard("harvard")}>
          🏛️ הארווארד{!data.harvard ? " (אין נתונים)" : ""}
        </button>
        <button className={`tab ${board === "site" ? "on" : ""}`} onClick={() => setBoard("site")}>
          🌐 אתר{!siteData ? " (טרם נטען)" : ""}
        </button>
        {view.data.meta && view.data.meta.month && (
          <span className="tab-meta">{view.data.meta.month}{view.data.meta.reportDate ? ` · מועד דיווח: ${view.data.meta.reportDate}` : ""}</span>
        )}
      </div>

      <div className={`status ${status.kind === "err" ? "err" : status.kind === "ok" ? "ok" : ""}`}>{status.msg}</div>

      {viewingDay && (
        <div className="hist-banner">
          <span>🕰️ אתה צופה בדוח מהארכיון — <b>{heDayLabel(viewingDay)}</b></span>
          <button className="btn" onClick={backToToday}>חזרה לדוח הנוכחי</button>
        </div>
      )}

      {board === "site" && !siteData && (
        <div className="panel" style={{ marginBottom: 26 }}>
          <h2>🌐 חוברת האתר</h2>
          <div className="hint" style={{ marginBottom: 10 }}>
            טרם נטענו נתוני האתר. שימו לב: זו חוברת Google Sheets נפרדת — צריך לפתוח אותה לצפייה (שיתוף ← כל מי שיש לו קישור ← צפייה) או לפרסם כ-CSV, בנפרד מחוברת התארים. אחרי זה המשיכה תתבצע אוטומטית בכל פתיחה ומדי יום ב־10:00. חשוב: העתיקו את הקישור מתוך חוברת האתר כשהלשונית הרלוונטית פתוחה, כדי שה-gid יצביע על הלשונית הנכונה.
          </div>
          <button className="btn" onClick={() => fetchSite(siteLink)}>משיכת נתוני האתר עכשיו</button>
        </div>
      )}

      {board === "harvard" && !data.harvard && (
        <div className="panel" style={{ marginBottom: 26 }}>
          <h2>🏛️ הארווארד</h2>
          <div className="hint" style={{ marginBottom: 0 }}>
            בדוח הנוכחי לא נמצא סקשן "הארווארד - ריכוז נתונים". טענו דוח שמכיל אותו — והחוברת תתמלא אוטומטית עם KPI, טבלה, תובנות וסטורי משלה.
          </div>
        </div>
      )}

      {/* ---------- חיבור גיליון ---------- */}
      {showConnect && (
        <div className="connect">
          <strong>הדביקו את הקישור לדוח</strong>
          <div style={{ color: C.dim, fontSize: 12, marginTop: 4 }}>
            שלוש דרכים לטעון: קישור לגיליון (קובץ ← שיתוף ← פרסום באינטרנט ← CSV), העלאת קובץ CSV, או הדבקה ידנית. בכל שלוש הדרכים הדוח נשמר אוטומטית בארכיון לתאריך של יום קודם, וטעינה נוספת באותו יום מחליפה את הגרסה הקודמת של אותו תאריך — תמיד נשמר העדכון האחרון.
          </div>
          <input placeholder="https://docs.google.com/spreadsheets/d/..." value={link} onChange={(e) => setLink(e.target.value)} />
          <div className="archive-date-row">
            <span>📆 יאורכב לתאריך:</span>
            <input type="date" value={archiveDate} onChange={(e) => setArchiveDate(e.target.value)} />
            <span style={{ color: C.dim, fontSize: 12 }}>
              ברירת מחדל: יום לפני ההעלאה (הדוח משקף את אתמול). שנו רק אם מעלים דוח באיחור.
            </span>
          </div>
          <div style={{ marginTop: 10 }}>
            <strong style={{ fontSize: 13 }}>🌐 קישור חוברת האתר (אופציונלי)</strong>
            <input placeholder="https://docs.google.com/spreadsheets/d/... (חוברת האתר)" value={siteLink} onChange={(e) => setSiteLink(e.target.value)} />
            <button className="btn ghost" onClick={() => fetchSite(siteLink)}>טעינת חוברת האתר</button>
          </div>
          <div style={{ marginTop: 10 }}>
            <strong style={{ fontSize: 13 }}>📅 קישור גיליון YoY · השוואת שנים (אופציונלי)</strong>
            <input placeholder="https://docs.google.com/spreadsheets/d/... (גיליון YoY — תאריך/מקור/איכות)" value={yoyLink} onChange={(e) => setYoyLink(e.target.value)} />
            <button className="btn ghost" onClick={() => fetchYoY(yoyLink)}>טעינת גיליון YoY</button>
          </div>
          <label className="autosync">
            <input type="checkbox" checked={autoSync} onChange={async (e) => {
              setAutoSync(e.target.checked);
              await store.set("peres:autosync", e.target.checked);
              if (e.target.checked && link) await store.set("peres:link", link);
            }} />
            🔄 סנכרון אוטומטי — כל החוברות (תארים + אתר + YoY) יימשכו מהקישורים השמורים בכל פתיחה של הדשבורד, וגם מדי יום ב־10:00 (כשהדשבורד פתוח בדפדפן)
          </label>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <button className="btn" disabled={loading} onClick={fetchSheet}>{loading ? "⏳ טוען…" : "טעינה מהקישור"}</button>
            <label className="btn file-btn">
              📁 העלאת קובץ CSV / XLSX
              <input type="file" accept=".csv,.xlsx,.xlsm,.xls,text/csv,text/plain" onChange={(e) => handleFile(e)} />
            </label>
            <button className="btn ghost" onClick={() => setPasteMode((v) => !v)}>הדבקת נתונים ידנית</button>
          </div>
          {pasteMode && (
            <>
              <div style={{ color: C.dim, fontSize: 12, marginTop: 12 }}>
                לחלופין: בגיליון בחרו קובץ ← הורדה ← CSV, פתחו את הקובץ והדביקו את התוכן כאן.
              </div>
              <textarea placeholder="הדביקו כאן את תוכן ה-CSV..." value={csvPaste} onChange={(e) => setCsvPaste(e.target.value)} />
              <button className="btn" onClick={() => loadCsv(csvPaste)}>טעינת הנתונים</button>
            </>
          )}
        </div>
      )}

      {board !== "summary" && (board !== "site" || siteData) && (<>
      {/* ---------- KPI ---------- */}
      <div className="grid-kpi">
        {!view.noBudget && (
          <KpiCard title="מימוש תקציב כולל" actual={totals.spent} target={totals.budget} format={(v) => nis(v)}
                   sub={`${Math.round(totals.util)}% מומש`} delta={delta?.summary.spent} deltaMoney />
        )}
        <KpiCard title="לידים ברוטו" actual={s.actual.gross} target={s.target.gross} format={num} delta={delta?.summary.gross} />
        <KpiCard title="לידים איכותיים" actual={s.actual.quality} target={s.target.quality} format={num} delta={delta?.summary.quality} />
        <KpiCard title="% לידים איכותיים" actual={s.actual.qualityPct} target={s.target.qualityPct} format={(v) => pct(v, 1)}
                 delta={delta?.summary.qualityPct} deltaSuffix=" נק'" />
        {!view.noBudget && (
          <KpiCard title="עלות לליד ברוטו" actual={s.actual.cpl} target={s.target.cpl} format={(v) => nis(v)} inverse delta={delta?.summary.cpl} deltaMoney />
        )}
        {!view.noBudget && (
          <KpiCard title="עלות לליד איכותי" actual={s.actual.cpql} target={s.target.cpql} format={(v) => nis(v)} inverse delta={delta?.summary.cpql} deltaMoney />
        )}
        {view.noBudget && s.actual.inProcess !== null && s.actual.inProcess !== undefined && (
          <KpiCard title="לידים בתהליך" actual={s.actual.inProcess} target={null} format={num} />
        )}
        {view.noBudget && s.actual.invalid !== null && s.actual.invalid !== undefined && (
          <KpiCard title="לידים לא תקינים" actual={s.actual.invalid} target={null} format={num}
                   sub={s.actual.gross ? `${Math.round((s.actual.invalid / s.actual.gross) * 100)}% מהברוטו` : ""} />
        )}
      </div>

      {/* ---------- תובנות ---------- */}
      <div className="panel" style={{ marginBottom: 26 }}>
        <h2>💡 תובנות</h2>
        <div className="hint">מחושבות אוטומטית מהנתונים{delta ? " כולל דלתות מול הדוח הקודם" : ""}</div>
        <div className="insight-list">
          {insights.map((ins, i) => (
            <div key={i} className={`insight ${ins.tone}`}>
              <span className="ico">{ins.icon}</span><span>{ins.text}</span>
            </div>
          ))}
        </div>
      </div>

      {/* ---------- גרפים ---------- */}
      <div className="two-col" style={view.noBudget ? { gridTemplateColumns: "1fr" } : undefined}>
        {!view.noBudget && (
        <div className="panel">
          <h2>תקציב מול מימוש לפי כלי</h2>
          <div className="hint">₪ — עמודה כהה: תקציב · עמודה צהובה: מומש בפועל · ריחוף מציג % מימוש</div>
          <div style={{ width: "100%", height: 360, direction: "ltr" }}>
            <ResponsiveContainer>
              <BarChart data={barData} layout="vertical" margin={{ top: 0, right: 6, left: 6, bottom: 0 }}>
                <CartesianGrid stroke={C.panelSoft} horizontal={false} />
                <XAxis type="number" tick={{ fill: C.dim, fontSize: 11 }} tickFormatter={(v) => "₪" + (v / 1000) + "K"} reversed />
                <YAxis type="category" dataKey="name" orientation="right" width={155}
                       tick={{ fill: C.text, fontSize: 11.5 }} tickMargin={8} interval={0} />
                <Tooltip
                  contentStyle={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 10, direction: "rtl", color: C.text }}
                  labelStyle={{ color: C.text, fontWeight: 700 }} itemStyle={{ color: C.text }}
                  formatter={(v, n, props) => {
                    const budget = props && props.payload ? props.payload["תקציב"] : 0;
                    if (n === "מומש" && budget > 0) return [`${nis(v)} (${Math.round((v / budget) * 100)}% מימוש)`, n];
                    return [nis(v), n];
                  }}
                />
                <Bar dataKey="תקציב" fill={C.panelSoft} stroke={C.line} radius={[4, 0, 0, 4]} barSize={10} />
                <Bar dataKey="מומש" fill={C.amber} radius={[4, 0, 0, 4]} barSize={10} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
        )}

        <div className="panel">
          {pieData.length > 1 ? (
            <>
              <h2>{view.noBudget ? "התפלגות הלידים בין המקורות" : "התפלגות ההוצאה בין הכלים"}</h2>
              <div className="hint">{view.noBudget
                ? `נתח של כל מקור מתוך ${num(pieTotal)} לידים ברוטו`
                : `נתח של כל כלי מתוך ${nis(totals.spent)} שמומשו עד כה (לא אחוז ניצול תקציב)`}</div>
              <div className="pie-flex">
                <div style={{ width: 210, height: 260, flexShrink: 0, direction: "ltr" }}>
                  <ResponsiveContainer>
                    <PieChart>
                      <Pie data={pieData} dataKey="value" nameKey="name" innerRadius={55} outerRadius={95} paddingAngle={2} stroke={C.bg}>
                        {pieData.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                      </Pie>
                      <Tooltip
                        contentStyle={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 10, direction: "rtl", color: C.text }}
                        labelStyle={{ color: C.text, fontWeight: 700 }} itemStyle={{ color: C.text, fontWeight: 700 }}
                        formatter={(v, n) => [view.noBudget
                          ? `${num(v)} לידים · ${Math.round((v / pieTotal) * 100)}%`
                          : `${nis(v)} · ${Math.round((v / pieTotal) * 100)}% מההוצאה`, n]}
                      />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
                <div className="legend">
                  {pieData.map((p, i) => (
                    <div className="legend-row" key={p.name}>
                      <span className="legend-dot" style={{ background: PIE_COLORS[i % PIE_COLORS.length] }} />
                      <span className="legend-name">{p.name}</span>
                      <span className="legend-val">{view.noBudget ? num(p.value) : nis(p.value)} · {Math.round((p.value / pieTotal) * 100)}%</span>
                    </div>
                  ))}
                </div>
              </div>
            </>
          ) : (
            /* כלי יחיד — התפלגות בין כלים חסרת משמעות, במקומה: דונאט מימוש תקציב */
            <>
              <h2>מימוש התקציב{pieData.length === 1 ? ` · ${pieData[0].name}` : ""}</h2>
              <div className="hint">{nis(totals.spent)} מומשו מתוך {nis(totals.budget)}</div>
              <div className="pie-flex">
                <div style={{ width: 210, height: 260, flexShrink: 0, direction: "ltr", position: "relative" }}>
                  <ResponsiveContainer>
                    <PieChart>
                      <Pie
                        data={[
                          { name: "מומש", value: totals.spent },
                          { name: "יתרת תקציב", value: Math.max(totals.budget - totals.spent, 0) },
                        ]}
                        dataKey="value" nameKey="name" innerRadius={62} outerRadius={95} paddingAngle={2} stroke={C.bg}
                        startAngle={90} endAngle={-270}
                      >
                        <Cell fill={C.amber} />
                        <Cell fill={C.panelSoft} />
                      </Pie>
                      <Tooltip
                        contentStyle={{ background: C.panel, border: `1px solid ${C.line}`, borderRadius: 10, direction: "rtl", color: C.text }}
                        labelStyle={{ color: C.text, fontWeight: 700 }} itemStyle={{ color: C.text, fontWeight: 700 }}
                        formatter={(v, n) => [nis(v), n]}
                      />
                    </PieChart>
                  </ResponsiveContainer>
                  <div style={{ position: "absolute", inset: 0, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", pointerEvents: "none" }}>
                    <span style={{ fontSize: 30, fontWeight: 900, color: C.amber }}>{Math.round(totals.util)}%</span>
                    <span style={{ fontSize: 11.5, color: C.dim }}>מימוש</span>
                  </div>
                </div>
                <div className="legend">
                  <div className="legend-row">
                    <span className="legend-dot" style={{ background: C.amber }} />
                    <span className="legend-name">מומש</span>
                    <span className="legend-val">{nis(totals.spent)} · {Math.round(totals.util)}%</span>
                  </div>
                  <div className="legend-row">
                    <span className="legend-dot" style={{ background: C.panelSoft, border: `1px solid ${C.line}` }} />
                    <span className="legend-name">יתרת תקציב</span>
                    <span className="legend-val">{nis(Math.max(totals.budget - totals.spent, 0))} · {Math.max(100 - Math.round(totals.util), 0)}%</span>
                  </div>
                </div>
              </div>
            </>
          )}
        </div>
      </div>

      {/* ---------- טבלת כלים ---------- */}
      <div className="panel" style={{ borderRadius: drillOpen && !view.isHarvard && view.data.campaigns.length > 0 ? "14px 14px 0 0" : 14 }}>
        <h2>{view.isHarvard ? "הארווארד · ביצועים לפי כלי" : view.noBudget ? "ביצועים לפי מקור (Organic)" : "ביצועים לפי כלי (BOF)"}</h2>
        <div className="hint">{view.isHarvard
          ? "לחיצה על כותרת ממיינת · חצים ירוקים/אדומים = שינוי מאז הדוח הקודם"
          : 'לחיצה על כותרת ממיינת · לחיצה על שורת "חיפוש" פותחת דריל־דאון · חצים ירוקים/אדומים = שינוי מאז הדוח הקודם'}</div>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>
                <th>{view.noBudget ? "מקור" : "כלי"}</th>
                {!view.noBudget && <Th k="budget" sort={sort} setSort={setSort}>תקציב</Th>}
                {!view.noBudget && <Th k="spent" sort={sort} setSort={setSort}>מומש</Th>}
                {!view.noBudget && <th>% מימוש</th>}
                <Th k="gross" sort={sort} setSort={setSort}>לידים ברוטו</Th>
                {!view.noBudget && <Th k="cpl" sort={sort} setSort={setSort}>עלות לליד</Th>}
                <Th k="quality" sort={sort} setSort={setSort}>לידים איכותיים</Th>
                {!view.noBudget && <Th k="cpql" sort={sort} setSort={setSort}>עלות לליד איכותי</Th>}
                <Th k="qPct" sort={sort} setSort={setSort}>% איכות</Th>
                <Th k="inProc" sort={sort} setSort={setSort}>בתהליך</Th>
              </tr>
            </thead>
            <tbody>
              {sortedPlatforms.map((p) => {
                const util = p.budget ? (p.spent / p.budget) * 100 : null;
                const pd = delta?.platforms[p.name];
                const on = isActive(p.name);
                return (
                  <tr key={p.name} className={p.drill ? "click" : ""} style={{ opacity: on ? 1 : 0.45 }}
                      onClick={() => p.drill && setDrillOpen((v) => !v)}>
                    <td style={{ fontWeight: 700 }}>
                      <button className={`switch ${on ? "on" : "off"}`} aria-label={on ? "פעיל" : "כבוי"}
                              title={on ? "פעיל — לחיצה מסמנת ככבוי (יוחרג מתובנות והתראות)" : "כבוי — לחיצה מחזירה לפעיל"}
                              onClick={(e) => { e.stopPropagation(); toggleActive(p.name); }} />
                      {p.name}{p.drill && <span className="tag">{drillOpen ? "דריל־דאון ▴" : "דריל־דאון ▾"}</span>}
                    </td>
                    {!view.noBudget && <td>{p.budget ? nis(p.budget) : "—"}</td>}
                    {!view.noBudget && <td style={{ fontWeight: 700 }}>{nis(p.spent)}{pd && <Delta v={pd.spent} money inverse />}</td>}
                    {!view.noBudget && (
                      <td>
                        <div className="util-cell">
                          <div className="util-track">
                            <div className="util-fill" style={{ width: `${Math.min(util ?? 0, 100)}%`, background: util > 90 ? C.coral : C.blue }} />
                          </div>
                          <span style={{ color: C.dim, fontSize: 12, minWidth: 34 }}>{util === null ? "—" : Math.round(util) + "%"}</span>
                        </div>
                      </td>
                    )}
                    <td>{num(p.gross)}{pd && <Delta v={pd.gross} />}</td>
                    {!view.noBudget && <td>{nis(p.cpl)}</td>}
                    <td style={{ fontWeight: 700, color: C.teal }}>{num(p.quality)}{pd && <Delta v={pd.quality} />}</td>
                    {!view.noBudget && <td>{nis(p.cpql)}</td>}
                    <td><span className="qdot" style={{ background: qualityColor(p.qPct) }} />{pct(p.qPct, 1)}</td>
                    <td>{num(p.inProc)}{p.inProcPct !== null && p.inProcPct !== undefined ? ` (${pct(p.inProcPct)})` : ""}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* ---------- דריל-דאון חיפוש ---------- */}
      {drillOpen && !view.isHarvard && view.data.campaigns.length > 0 && (
        <div className="drill">
          <h2 style={{ fontSize: 15, fontWeight: 700, margin: "4px 0 2px" }}>דריל־דאון · קמפייני גוגל חיפוש</h2>
          <div style={{ color: C.dim, fontSize: 12, marginBottom: 10 }}>פירוט הקמפיינים המרכיבים את כלי החיפוש, כולל מקורות אורגניים (GMB / callbox / Direct)</div>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>קמפיין</th><th>מומש</th><th>נתח מעלות החיפוש</th><th>לידים ברוטו</th><th>עלות לליד</th>
                  <th>לידים איכותיים</th><th>עלות לליד איכותי</th><th>% איכות</th><th>בתהליך</th>
                </tr>
              </thead>
              <tbody>
                {data.campaigns.map((cmp) => {
                  const on = isActive(cmp.name);
                  return (
                    <tr key={cmp.name} style={{ opacity: on ? 1 : 0.45 }}>
                      <td style={{ fontWeight: 700 }}>
                        <button className={`switch ${on ? "on" : "off"}`} aria-label={on ? "פעיל" : "כבוי"}
                                title={on ? "פעיל — לחיצה מסמנת ככבוי (יוחרג מתובנות והתראות)" : "כבוי — לחיצה מחזירה לפעיל"}
                                onClick={() => toggleActive(cmp.name)} />
                        {cmp.name}
                      </td>
                      <td>{nis(cmp.spent)}</td>
                      <td>{pct(cmp.share)}</td>
                      <td>{num(cmp.gross)}</td>
                      <td>{nis(cmp.cpl)}</td>
                      <td style={{ fontWeight: 700, color: C.teal }}>{num(cmp.quality)}</td>
                      <td>{nis(cmp.cpql)}</td>
                      <td><span className="qdot" style={{ background: qualityColor(cmp.qPct) }} />{pct(cmp.qPct, 1)}</td>
                      <td>{num(cmp.inProc)}{cmp.inProcPct !== null && cmp.inProcPct !== undefined ? ` (${pct(cmp.inProcPct)})` : ""}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* ---------- ביצועי היממה האחרונה (דלתא) ---------- */}
      {board !== "site" && (
      <div className="panel" style={{ marginBottom: 26, marginTop: drillOpen && !view.isHarvard && view.data.campaigns.length > 0 ? 0 : 26 }}>
        <h2>📅 ביצועי קמפיינים · היממה האחרונה{view.isHarvard ? " · הארווארד" : ""}</h2>
        <div className="hint">
          {dayRows
            ? `השינוי בין הדוח הנוכחי לדוח מ־${heDate(delta.date)} — ${view.isHarvard ? "כלי הארווארד בלבד" : "מה כל כלי וקמפיין עשה בפועל ביממה"} · לחיצה על כותרת ממיינת`
            : "נתוני דוגמה להמחשה — הטבלה תתמלא אוטומטית בנתונים אמיתיים אחרי הטעינה השנייה של הדוח (צריך שתי תמונות מצב כדי לחשב יממה)"}
        </div>
        {(() => {
          const base = (dayRows || (view.isHarvard ? DEMO_DAY_HARVARD : DEMO_DAY)).map((r) => ({
            ...r,
            cplDay: (r.gross || 0) > 0 ? (r.spent || 0) / r.gross : null,
            cpqlDay: (r.quality || 0) > 0 ? (r.spent || 0) / r.quality : null,
          }));
          const sortRows = (arr) => [...arr].sort((a, b) => {
            const va = a[daySort.key], vb = b[daySort.key];
            if (va === null || va === undefined) return 1;
            if (vb === null || vb === undefined) return -1;
            return daySort.dir === "desc" ? vb - va : va - vb;
          });
          const trend = (r) => {
            if (!isActive(r.name)) return { txt: "כבוי", color: C.dim };
            if ((r.quality || 0) > 0) return { txt: "🟢 מייצר איכות", color: C.teal };
            if (!view.noBudget && (r.spent || 0) > 300 && (r.gross || 0) <= 0) return { txt: "🔴 שורף בלי לידים", color: C.coral };
            if ((r.gross || 0) > 0) return { txt: "🟡 לידים בלי איכות", color: C.amber };
            return { txt: "—", color: C.dim };
          };
          const rowsFor = (rows, maxSpent) => sortRows(rows).map((r) => {
            const t = trend(r);
            const on = isActive(r.name);
            return (
              <tr key={`${r.type}-${r.name}`} style={{ opacity: on ? 1 : 0.45 }}>
                <td style={{ fontWeight: 700 }}>{r.name}</td>
                {!view.noBudget && (
                  <td>
                    <div className="util-cell">
                      <div className="util-track"><div className="util-fill" style={{ width: `${((r.spent || 0) / maxSpent) * 100}%`, background: C.amber }} /></div>
                      <span style={{ minWidth: 60, fontWeight: 700 }}>{nis(r.spent)}</span>
                    </div>
                  </td>
                )}
                <td>{num(r.gross)}</td>
                {!view.noBudget && <td>{nis(r.cplDay)}</td>}
                <td style={{ fontWeight: 700, color: (r.quality || 0) > 0 ? C.teal : undefined }}>{num(r.quality)}</td>
                {!view.noBudget && <td>{nis(r.cpqlDay)}</td>}
                <td style={{ color: t.color, fontSize: 12.5, fontWeight: 700 }}>{t.txt}</td>
              </tr>
            );
          });
          const tools = base.filter((r) => r.type === "כלי");
          const search = base.filter((r) => r.type === "חיפוש");
          const maxTools = Math.max(...tools.map((r) => r.spent || 0), 1);
          const maxSearch = Math.max(...search.map((r) => r.spent || 0), 1);
          return (
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>{view.noBudget ? "מקור" : "כלי / קמפיין"}</th>
                    {!view.noBudget && <Th k="spent" sort={daySort} setSort={setDaySort}>מומש ביממה</Th>}
                    <Th k="gross" sort={daySort} setSort={setDaySort}>לידים ביממה</Th>
                    {!view.noBudget && <Th k="cplDay" sort={daySort} setSort={setDaySort}>עלות לליד יומית</Th>}
                    <Th k="quality" sort={daySort} setSort={setDaySort}>איכותיים ביממה</Th>
                    {!view.noBudget && <Th k="cpqlDay" sort={daySort} setSort={setDaySort}>עלות לאיכותי יומית</Th>}
                    <th>מגמה</th>
                  </tr>
                </thead>
                <tbody>
                  <tr className="group-row"><td colSpan={7} style={{ color: C.blue }}>🏗️ {view.isHarvard ? "כלי הארווארד" : view.noBudget ? "מקורות האתר" : "כלים · רמת מאקרו"}</td></tr>
                  {tools.length ? rowsFor(tools, maxTools)
                    : <tr><td colSpan={7} style={{ color: C.dim }}>אין תנועות ברמת הכלים ביממה האחרונה</td></tr>}
                  {!view.isHarvard && !view.noBudget && (
                    <>
                      <tr className="group-row"><td colSpan={7} style={{ color: C.violet }}>🔍 דריל־דאון · קמפייני חיפוש</td></tr>
                      {search.length ? rowsFor(search, maxSearch)
                        : <tr><td colSpan={7} style={{ color: C.dim }}>אין תנועות בקמפייני החיפוש ביממה האחרונה</td></tr>}
                    </>
                  )}
                </tbody>
              </table>
            </div>
          );
        })()}
        {dayRows && (
          <div style={{ color: C.dim, fontSize: 12, marginTop: 10 }}>
            "יממה" = הפרש בין הדוח הנוכחי לדוח של היום הקודם בארכיון. שורות כבויות מוצגות אך לא נכללות בהתראות.
          </div>
        )}
      </div>
      )}

      </>)}

      {/* ---------- סיכום נתונים · מגמות ---------- */}
      {board === "summary" && (
        <div className="board-tabs" style={{ marginBottom: 16, marginTop: -6 }}>
          <button className={`subtab ${overviewTab === "trends" ? "on" : ""}`} onClick={() => { refreshTrend(); setOverviewTab("trends"); }}>מגמות</button>
          <button className={`subtab ${overviewTab === "yoy" ? "on" : ""}`} onClick={() => setOverviewTab("yoy")}>השוואת שנים · YoY</button>
        </div>
      )}

      {board === "summary" && overviewTab === "trends" && (() => {
        const t = trendData;
        const isDemo = t.length < 2;
        const latest = isDemo
          ? { spent: 263033, budgetTarget: 637000, degQuality: 164, harvardQuality: 9, harvardSpent: 23543, cpqlTotal: 1604 }
          : ([...t].reverse().find((r) => r.spent !== null) || {});
        const latestSite = isDemo ? { siteQuality: 98 } : ([...t].reverse().find((r) => r.siteQuality !== null) || {});
        /* "מה הגיע בכל יום" = הפרש בין כל יום שמור ליום השמור שלפניו (שלילי מתאפס — למשל במעבר חודש) */
        const dd = (a, b) => (a !== null && a !== undefined && b !== null && b !== undefined ? Math.max(b - a, 0) : null);
        let daily = [];
        for (let i = 1; i < t.length; i++) {
          const a = t[i - 1], b = t[i];
          daily.push({
            label: b.label,
            "תקציב יומי": dd(a.spent, b.spent),
            "איכותיים תארים": dd(a.degQuality, b.degQuality),
            "תקציב הארווארד": dd(a.harvardSpent, b.harvardSpent),
            "איכותיים הארווארד": dd(a.harvardQuality, b.harvardQuality),
            "לידים ברוטו אתר": dd(a.siteGross, b.siteGross),
            "איכותיים אתר": dd(a.siteQuality, b.siteQuality),
          });
        }
        if (isDemo) daily = DEMO_TREND;
        const has = (key) => daily.some((r) => r[key] !== null && r[key] !== undefined);

        /* --- בורר קמפיין לגרף האיכות: "כללי" או קמפיין ספציפי, איכותיים + עלות לליד יומיים --- */
        const campNames = (() => {
          if (isDemo) return Object.keys(DEMO_CAMP_QUALITY);
          const set = new Set();
          t.forEach((r) => r.plat && Object.keys(r.plat).forEach((n) => set.add(n)));
          return [...set].sort();
        })();
        const effSel = (qualSel === "__all__" || campNames.includes(qualSel)) ? qualSel : "__all__";
        const qualDaily = (() => {
          if (isDemo) {
            if (effSel === "__all__")
              return DEMO_TREND.map((r) => ({ label: r.label, q: r["איכותיים תארים"], cpl: Math.round((r["תקציב יומי"] || 0) / Math.max(r["איכותיים תארים"] || 1, 1)) }));
            const c = DEMO_CAMP_QUALITY[effSel];
            return DEMO_TREND.map((r, i) => ({ label: r.label, q: c.q[i], cpl: c.cpl[i] }));
          }
          const out = [];
          for (let i = 1; i < t.length; i++) {
            const a = t[i - 1], b = t[i];
            if (effSel === "__all__") {
              const q = dd(a.degQuality, b.degQuality);
              const sp = dd(a.spent, b.spent);
              out.push({ label: b.label, q, cpl: q ? Math.round((sp || 0) / q) : null });
            } else {
              const pa = a.plat && a.plat[effSel], pb = b.plat && b.plat[effSel];
              if (pa && pb) {
                const q = Math.max((pb.quality || 0) - (pa.quality || 0), 0);
                const sp = Math.max((pb.spent || 0) - (pa.spent || 0), 0);
                out.push({ label: b.label, q, cpl: q ? Math.round(sp / q) : null });
              } else out.push({ label: b.label, q: null, cpl: null });
            }
          }
          return out;
        })();
        /* התקציב במבט-על = תארים + הארווארד ביחד (שניהם בתשלום) */
        const paidSpent = (latest.spent || 0) + (latest.harvardSpent || 0);
        const paidBudget = (latest.budgetTarget || 0) + (isDemo ? 40000 : 0); /* הארווארד יעד ~40K בדמו */
        const totalQ = (latest.degQuality || 0) + (latest.harvardQuality || 0) + (latestSite.siteQuality || 0);
        const range = isDemo ? `${DEMO_TREND[0].label} – ${DEMO_TREND[DEMO_TREND.length - 1].label}` : (t.length ? `${t[0].label} – ${t[t.length - 1].label}` : "");
        const overviewCards = [
          { title: "תקציב כולל (תארים + הארווארד)", val: nis(paidSpent), sub: paidBudget ? `מתוך ${nis(paidBudget)} · ${Math.round((paidSpent / paidBudget) * 100)}%` : "", color: C.amber },
          { title: "לידים איכותיים · טוטאל", val: num(totalQ), sub: "תארים + הארווארד + אתר", color: C.teal },
          { title: "עלות לליד איכותי · טוטאל", val: totalQ > 0 && paidSpent > 0 ? nis(paidSpent / totalQ) : (latest.cpqlTotal ? nis(latest.cpqlTotal) : "—"),
            sub: totalQ > 0 && paidSpent > 0 ? `${nis(paidSpent)} ÷ ${num(totalQ)} איכותיים (כולל אתר)` : "", color: C.coral },
          { title: "לידים איכותיים · תארים", val: num(latest.degQuality ?? 0), color: C.blue },
          { title: "לידים איכותיים · הארווארד", val: latest.harvardQuality !== null && latest.harvardQuality !== undefined ? num(latest.harvardQuality) : "—", color: C.violet },
          { title: "לידים איכותיים · אתר", val: latestSite.siteQuality !== null && latestSite.siteQuality !== undefined ? num(latestSite.siteQuality) : "—", color: C.teal },
        ];
        const tip = (fmt) => ({
          contentStyle: { background: C.panel, border: `1px solid ${C.line}`, borderRadius: 10, direction: "rtl", color: C.text },
          labelStyle: { color: C.text, fontWeight: 700 }, itemStyle: { color: C.text },
          formatter: fmt,
        });
        const axis = { tick: { fill: C.dim, fontSize: 11 }, stroke: C.line };
        const chartCard = (title, hint, children, empty) => (
          <div className="panel" style={{ marginBottom: 18 }}>
            <h2>{title}</h2>
            <div className="hint">{hint}</div>
            {empty ? (
              <div className="hint" style={{ marginBottom: 0 }}>{empty}</div>
            ) : (
              <div style={{ width: "100%", height: 280, direction: "ltr" }}>
                <ResponsiveContainer>{children}</ResponsiveContainer>
              </div>
            )}
          </div>
        );
        return (
          <>
            {isDemo && (
              <div className="status" style={{ borderColor: C.amber, color: C.amber }}>
                🧪 נתוני דוגמה להמחשה — כך ייראה "מבט על" כשיצטברו ימים בארכיון. הגרפים יתמלאו אוטומטית בנתונים אמיתיים ברגע שיהיו לפחות שני ימים שמורים (טעינה יומית או השלמה דרך לוח השנה).
              </div>
            )}
            <div className="grid-kpi" style={{ marginBottom: 20 }}>
              {overviewCards.map((c, i) => (
                <div className="kpi" key={i}>
                  <div className="kpi-title">{c.title}</div>
                  <div className="kpi-value" style={{ color: c.color }}>{c.val}</div>
                  <div className="kpi-target">{c.sub || (range && `טווח: ${range}`)}</div>
                </div>
              ))}
            </div>
            <div className="panel" style={{ marginBottom: 18 }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" }}>
                <h2>⭐ איכות לפי קמפיין · לפי יום</h2>
                <select className="qual-select" value={effSel} onChange={(e) => setQualSel(e.target.value)}>
                  <option value="__all__">כל הקמפיינים (כללי · תארים)</option>
                  {campNames.map((n) => <option key={n} value={n}>{n}</option>)}
                </select>
              </div>
              <div className="hint">
                {effSel === "__all__" ? "סך הלידים האיכותיים והעלות לליד איכותי בכל יום — כלל התארים" : `קמפיין "${effSel}" — לידים איכותיים ועלות לליד איכותי בכל יום`}
                {isDemo ? " · נתוני דוגמה" : ""}
              </div>
              {qualDaily.length === 0 || qualDaily.every((r) => r.q === null) ? (
                <div className="hint" style={{ marginBottom: 0 }}>אין עדיין נתונים לקמפיין הזה בארכיון — הם ייצברו ככל שייטענו דוחות</div>
              ) : (
                <div style={{ width: "100%", height: 280, direction: "ltr" }}>
                  <ResponsiveContainer>
                    <ComposedChart data={qualDaily} margin={{ top: 6, right: 8, left: 8, bottom: 0 }}>
                      <CartesianGrid stroke={C.panelSoft} vertical={false} />
                      <XAxis dataKey="label" {...axis} />
                      <YAxis yAxisId="q" {...axis} allowDecimals={false} />
                      <YAxis yAxisId="cpl" orientation="right" {...axis} tickFormatter={(v) => "₪" + (v >= 1000 ? Math.round(v / 1000) + "K" : Math.round(v))} />
                      <Tooltip {...tip((v, n) => [n.includes("עלות") ? nis(v) : num(v), n])} />
                      <Legend wrapperStyle={{ direction: "rtl", fontSize: 12 }} />
                      <Bar yAxisId="q" dataKey="q" name="לידים איכותיים" fill={C.teal} radius={[4, 4, 0, 0]} maxBarSize={34} />
                      <Line yAxisId="cpl" type="monotone" dataKey="cpl" name="עלות לליד איכותי" stroke={C.coral} strokeWidth={2.5} dot={{ r: 3 }} connectNulls />
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              )}
            </div>

            {chartCard("🎓 תארים · תקציב ולידים איכותיים לפי יום",
              "עמודות: כמה תקציב מומש בכל יום · קו: כמה לידים איכותיים הגיעו באותו יום",
              <ComposedChart data={daily} margin={{ top: 6, right: 8, left: 8, bottom: 0 }}>
                <CartesianGrid stroke={C.panelSoft} vertical={false} />
                <XAxis dataKey="label" {...axis} />
                <YAxis yAxisId="money" {...axis} tickFormatter={(v) => "₪" + Math.round(v / 1000) + "K"} />
                <YAxis yAxisId="q" orientation="right" {...axis} allowDecimals={false} />
                <Tooltip {...tip((v, n) => [n.includes("תקציב") ? nis(v) : num(v), n])} />
                <Legend wrapperStyle={{ direction: "rtl", fontSize: 12 }} />
                <Bar yAxisId="money" dataKey="תקציב יומי" fill={C.amber} radius={[4, 4, 0, 0]} maxBarSize={34} />
                <Line yAxisId="q" type="monotone" dataKey="איכותיים תארים" name="לידים איכותיים" stroke={C.blue} strokeWidth={2.5} dot={{ r: 3 }} connectNulls />
              </ComposedChart>,
              has("תקציב יומי") || has("איכותיים תארים") ? null : "אין עדיין נתוני תארים בארכיון"
            )}

            {chartCard("🏛️ הארווארד · תקציב ולידים איכותיים לפי יום",
              "עמודות: מימוש יומי של הארווארד · קו: לידים איכותיים שהגיעו באותו יום",
              <ComposedChart data={daily} margin={{ top: 6, right: 8, left: 8, bottom: 0 }}>
                <CartesianGrid stroke={C.panelSoft} vertical={false} />
                <XAxis dataKey="label" {...axis} />
                <YAxis yAxisId="money" {...axis} tickFormatter={(v) => "₪" + (v >= 1000 ? Math.round(v / 1000) + "K" : Math.round(v))} />
                <YAxis yAxisId="q" orientation="right" {...axis} allowDecimals={false} />
                <Tooltip {...tip((v, n) => [n.includes("תקציב") ? nis(v) : num(v), n])} />
                <Legend wrapperStyle={{ direction: "rtl", fontSize: 12 }} />
                <Bar yAxisId="money" dataKey="תקציב הארווארד" fill={C.violet} radius={[4, 4, 0, 0]} maxBarSize={34} />
                <Line yAxisId="q" type="monotone" dataKey="איכותיים הארווארד" name="לידים איכותיים" stroke={C.teal} strokeWidth={2.5} dot={{ r: 3 }} connectNulls />
              </ComposedChart>,
              has("תקציב הארווארד") || has("איכותיים הארווארד") ? null : "אין עדיין נתוני הארווארד בארכיון"
            )}

            {chartCard("🌐 אתר · לידים לפי יום",
              "עמודות: לידים ברוטו שהגיעו מהאתר בכל יום · קו: כמה מהם איכותיים (ללא תקציב — מקור אורגני)",
              <ComposedChart data={daily} margin={{ top: 6, right: 8, left: 8, bottom: 0 }}>
                <CartesianGrid stroke={C.panelSoft} vertical={false} />
                <XAxis dataKey="label" {...axis} />
                <YAxis {...axis} allowDecimals={false} />
                <Tooltip {...tip((v, n) => [num(v), n])} />
                <Legend wrapperStyle={{ direction: "rtl", fontSize: 12 }} />
                <Bar dataKey="לידים ברוטו אתר" name="לידים ברוטו" fill={C.blue} radius={[4, 4, 0, 0]} maxBarSize={34} />
                <Line type="monotone" dataKey="איכותיים אתר" name="לידים איכותיים" stroke={C.teal} strokeWidth={2.5} dot={{ r: 3 }} connectNulls />
              </ComposedChart>,
              has("לידים ברוטו אתר") || has("איכותיים אתר") ? null : "היסטוריית האתר מתחילה להיצבר מעכשיו — אחרי יומיים של משיכות יופיע כאן הגרף"
            )}

            <div className="foot" style={{ marginTop: 0, marginBottom: 26 }}>
              "יום" = ההפרש בין דוח לדוח שלפניו בארכיון · עד 31 הימים האחרונים · ימים חסרים בארכיון אפשר להשלים דרך לוח השנה
            </div>
          </>
        );
      })()}

      {/* ---------- YoY · השוואת שנים (תת-ניווט בתוך "מבט על") ---------- */}
      {board === "summary" && overviewTab === "yoy" && (() => {
        const now = new Date();
        const reportDate = new Date(now); reportDate.setDate(reportDate.getDate() - 1);
        const y = reportDate.getFullYear(), prevY = y - 1;
        const mo = reportDate.getMonth() + 1, dayOfMonth = reportDate.getDate();
        const mm = String(mo).padStart(2, "0");
        const heMonth = reportDate.toLocaleDateString("he-IL", { month: "long" });
        const fullDate = `${dayOfMonth}.${mo}`;

        const isDemoYoY = !(yoyData && yoyData[prevY]);
        const leadsYoY = isDemoYoY ? DEMO_YOY.leads : yoyData;
        const budgetYoY = isDemoYoY ? DEMO_YOY.budget : yoyBudget;

        const METRICS = {
          gross:   { label: "לידים ברוטו",       fmt: num,              kind: "count", better: "up" },
          quality: { label: "לידים איכותיים",    fmt: num,              kind: "count", better: "up" },
          qpct:    { label: "% לידים איכותיים",   fmt: (v) => pct(v, 1), kind: "ratio", better: "up" },
          cpl:     { label: "עלות לליד ברוטו",    fmt: (v) => nis(v),    kind: "cost",  better: "down" },
          cpql:    { label: "עלות לליד איכותי",   fmt: (v) => nis(v),    kind: "cost",  better: "down" },
        };
        const metric = "quality"; /* תצוגה כללית — הטבלה המסכמת מכסה את כל המדדים */
        const M = METRICS[metric];
        const metricVal = (g, q, b) => {
          if (metric === "gross") return g;
          if (metric === "quality") return q;
          if (metric === "qpct") return g > 0 ? (q / g) * 100 : null;
          if (metric === "cpl") return g > 0 ? b / g : null;
          if (metric === "cpql") return q > 0 ? b / q : null;
          return null;
        };

        const daysInMonth = new Date(y, mo, 0).getDate();
        const histSeries = (() => {
          const lm = (leadsYoY && leadsYoY[prevY]) || {}, bm = (budgetYoY && budgetYoY[prevY]) || {};
          const out = []; let cg = 0, cq = 0, cb = 0;
          for (let d = 1; d <= daysInMonth; d++) {
            const key = `${mm}-${String(d).padStart(2, "0")}`;
            const L = lm[key], B = bm[key];
            if (L) { cg += L.gross.total; cq += L.quality.total; }
            if (B) cb += B.total;
            const dg = L ? L.gross.total : 0, dq = L ? L.quality.total : 0, db = B ? B.total : 0;
            out.push({ day: d, val: metricVal(cg, cq, cb), dayVal: metricVal(dg, dq, db), _g: cg, _q: cq, _b: cb,
              cpqlCum: cq > 0 ? Math.round(cb / cq) : null });
          }
          return out;
        })();
        const histAtDay = histSeries[Math.min(dayOfMonth, histSeries.length) - 1] || { val: null, _g: 0, _q: 0, _b: 0 };

        const curBase = (() => {
          if (isDemoYoY) {
            const h = histAtDay;
            return { gross: Math.round(h._g * 1.1), quality: Math.round(h._q * 1.15), budget: Math.round(h._b * 1.08) };
          }
          const t = reportTotals(data);
          const degG = data.totalsRow && data.totalsRow.gross != null ? data.totalsRow.gross : (data.summary ? (data.summary.actual.gross || 0) : 0);
          const harvG = data.harvard && data.harvard.total ? (data.harvard.total.gross || 0) : 0;
          const siteG = siteData && siteData.summary ? (siteData.summary.actual.gross || 0) : 0;
          const degQ = t ? t.degQuality : 0;
          const harvQ = t && t.harvardQuality != null ? t.harvardQuality : 0;
          const siteQ = siteData && siteData.summary ? (siteData.summary.actual.quality || 0) : 0;
          const spent = (t ? t.spent : 0) + (t && t.harvardSpent ? t.harvardSpent : 0);
          return { gross: degG + harvG + siteG, quality: degQ + harvQ + siteQ, budget: spent };
        })();
        const curVal = metricVal(curBase.gross, curBase.quality, curBase.budget);
        /* עלות לליד איכותי מצטברת של השנה (הוצאה מצטברת ÷ איכותיים מצטברים) — יציב, בר-השוואה */
        const curCPQL = curBase.quality > 0 ? Math.round(curBase.budget / curBase.quality) : null;
        const histCPQL = histAtDay._q > 0 ? Math.round(histAtDay._b / histAtDay._q) : null;
        /* מדדי הסיכום לטבלה (עד יום הדיווח, מצטבר) */
        const curQpct = curBase.gross > 0 ? (curBase.quality / curBase.gross) * 100 : null;
        const histQpct = histAtDay._g > 0 ? (histAtDay._q / histAtDay._g) * 100 : null;
        const pctDiff = (c, p) => (p != null && p !== 0 && c != null) ? Math.round(((c - p) / Math.abs(p)) * 100) : null;

        const holidays = [
          ...hebrewHolidaysForMonth(y, mo).map((h) => ({ ...h, kind: "cur" })),
          ...hebrewHolidaysForMonth(prevY, mo).map((h) => ({ ...h, kind: "prev" })),
          ...userEvents.filter((e) => { const dt = e.date && e.date.match(/^\d{4}-(\d{2})-\d{2}$/); return dt && +dt[1] === mo; })
            .map((e) => ({ day: +e.date.split("-")[2], label: e.label, kind: "user" })),
        ];

        const cdDaily = (curDaily && curDaily.daily) || {};
        const hasCurDaily = Object.values(cdDaily).some((v) => v != null);
        const chartData = histSeries.map((r) => ({
          day: `${String(r.day).padStart(2, "0")}.${mo}`,
          [`${prevY}`]: r.val,
          [`${prevY}_daily`]: r.dayVal,
          [`${y}`]: r.day === dayOfMonth ? curVal : null,
          [`${y}_daily`]: cdDaily[r.day] != null ? cdDaily[r.day] : null,
          [`cpql_${prevY}`]: r.cpqlCum,
          [`cpql_${y}`]: (curDaily && curDaily.cpql && curDaily.cpql[r.day] != null) ? curDaily.cpql[r.day] : (r.day === dayOfMonth ? curCPQL : null),
        }));

        const yoyDiff = (histAtDay.val != null && histAtDay.val !== 0 && curVal != null)
          ? Math.round(((curVal - histAtDay.val) / Math.abs(histAtDay.val)) * 100) : null;
        const yoyGood = yoyDiff == null ? null : (M.better === "up" ? yoyDiff >= 0 : yoyDiff <= 0);
        const tip = (fmt) => ({
          contentStyle: { background: C.panel, border: `1px solid ${C.line}`, borderRadius: 10, direction: "rtl", color: C.text },
          labelStyle: { color: C.text, fontWeight: 700 }, itemStyle: { color: C.text }, formatter: fmt,
        });
        const axis = { tick: { fill: C.dim, fontSize: 11 }, stroke: C.line };
        const yTickFmt = (v) => M.kind === "cost" ? "₪" + (v >= 1000 ? Math.round(v / 1000) + "K" : Math.round(v)) : M.kind === "ratio" ? v + "%" : num(v);

        return (
          <>
            {isDemoYoY && (
              <div className="status" style={{ borderColor: C.amber, color: C.amber }}>
                🧪 נתוני דוגמה להמחשה — כך ייראה טאב ה-YoY. לחיבור נתונים אמיתיים: הזינו את קישור חוברת ה-YoY בחלון "חיבור הגיליון" (לשונית לידים: תאריך/מקור/איכות · לשונית תקציב: Date/Cost).
              </div>
            )}
            <div className="grid-kpi" style={{ marginBottom: 20 }}>
              <div className="kpi">
                <div className="kpi-title">{M.label} · {heMonth} {y} (עד {fullDate})</div>
                <div className="kpi-value" style={{ color: C.teal }}>{curVal == null ? "—" : M.fmt(curVal)}</div>
                <div className="kpi-target">ברוטו {num(curBase.gross)} · איכותי {num(curBase.quality)}{curBase.budget ? ` · תקציב ${nis(curBase.budget)}` : ""}</div>
              </div>
              <div className="kpi">
                <div className="kpi-title">אשתקד · {heMonth} {prevY} (עד {fullDate})</div>
                <div className="kpi-value" style={{ color: C.dim }}>{histAtDay.val == null ? "—" : M.fmt(histAtDay.val)}</div>
                <div className="kpi-target">ברוטו {num(histAtDay._g)} · איכותי {num(histAtDay._q)}{histAtDay._b ? ` · תקציב ${nis(histAtDay._b)}` : ""}</div>
              </div>
              <div className="kpi">
                <div className="kpi-title">שינוי שנתי (YoY)</div>
                <div className="kpi-value" style={{ color: yoyDiff == null ? C.text : yoyGood ? C.teal : C.coral }}>
                  {yoyDiff == null ? "—" : (yoyDiff >= 0 ? "+" : "") + yoyDiff + "%"}
                </div>
                <div className="kpi-target">{yoyDiff == null ? "" : (M.better === "down" ? (yoyGood ? "זול יותר מאשתקד ✓" : "יקר יותר מאשתקד") : (yoyGood ? "גבוה מאשתקד ✓" : "נמוך מאשתקד"))}</div>
              </div>
            </div>

            <div className="panel" style={{ marginBottom: 26 }}>
              <h2>📋 סיכום השוואת שנים · עד {fullDate}</h2>
              <div className="hint">כל המדדים המצטברים עד יום הדיווח, {y} מול {prevY} — התמונה המלאה במבט אחד</div>
              <div className="table-wrap">
                <table>
                  <thead>
                    <tr>
                      <th>מדד</th><th>{y} (השנה)</th><th>{prevY} (אשתקד)</th><th>שינוי</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(() => {
                      const rowsDef = [
                        { label: "לידים ברוטו", c: curBase.gross, p: histAtDay._g, fmt: num, better: "up" },
                        { label: "לידים איכותיים", c: curBase.quality, p: histAtDay._q, fmt: num, better: "up" },
                        { label: "% לידים איכותיים", c: curQpct, p: histQpct, fmt: (v) => pct(v, 1), pts: true, better: "up" },
                        { label: "תקציב שנוצל", c: curBase.budget, p: histAtDay._b, fmt: (v) => nis(v), better: "down", skipIfZero: true },
                        { label: "עלות לליד ברוטו", c: curBase.gross > 0 ? curBase.budget / curBase.gross : null, p: histAtDay._g > 0 ? histAtDay._b / histAtDay._g : null, fmt: (v) => nis(v), better: "down", skipIfZero: true },
                        { label: "עלות לליד איכותי", c: curCPQL, p: histCPQL, fmt: (v) => nis(v), better: "down", skipIfZero: true },
                      ];
                      return rowsDef.filter((r) => !(r.skipIfZero && !r.c && !r.p)).map((r, i) => {
                        const diff = r.pts
                          ? (r.c != null && r.p != null ? +(r.c - r.p).toFixed(1) : null)
                          : pctDiff(r.c, r.p);
                        const good = diff == null ? null : (r.better === "up" ? diff >= 0 : diff <= 0);
                        const emphasize = r.label === "% לידים איכותיים";
                        return (
                          <tr key={i} style={emphasize ? { background: "rgba(55,201,169,.06)" } : undefined}>
                            <td style={{ fontWeight: emphasize ? 900 : 700 }}>{r.label}</td>
                            <td style={{ fontWeight: 700, color: emphasize ? C.teal : undefined }}>{r.c == null ? "—" : r.fmt(r.c)}</td>
                            <td style={{ color: C.dim }}>{r.p == null ? "—" : r.fmt(r.p)}</td>
                            <td style={{ fontWeight: 700, color: diff == null ? C.dim : good ? C.teal : C.coral }}>
                              {diff == null ? "—" : r.pts ? `${diff >= 0 ? "+" : ""}${diff} נק'` : `${diff >= 0 ? "+" : ""}${diff}%`}
                            </td>
                          </tr>
                        );
                      });
                    })()}
                  </tbody>
                </table>
              </div>
              <div style={{ color: C.dim, fontSize: 12, marginTop: 10 }}>
                ירוק = טוב יותר מאשתקד (בעלויות: נמוך יותר) · % איכות מוצג בהפרש נקודות אחוז
              </div>
            </div>

            <div className="panel" style={{ marginBottom: 18 }}>
              <h2>📅 {M.label} מצטבר · {heMonth}: {y} מול {prevY}</h2>
              <div className="hint">
                קו {prevY} = מצטבר אשתקד · הנקודה של {y} = המצב עד יום הדיווח ({fullDate})
                {holidays.length ? ` · ירוק = חג ${y} · כחול-עמום = חג ${prevY} · כתום = אירוע שלכם` : ""}
              </div>
              <div style={{ width: "100%", height: 320, direction: "ltr" }}>
                <ResponsiveContainer>
                  <ComposedChart data={chartData} margin={{ top: 6, right: 8, left: 8, bottom: 0 }}>
                    <CartesianGrid stroke={C.panelSoft} vertical={false} />
                    <XAxis dataKey="day" {...axis} />
                    <YAxis {...axis} allowDecimals={M.kind !== "count"} tickFormatter={yTickFmt} />
                    <Tooltip {...tip((v, n, p) => {
                      if (v == null) return ["—", n];
                      if (n.indexOf("אשתקד") !== -1 && p && p.payload && M.kind === "count") return [`${M.fmt(v)} מצטבר · ${num(p.payload[`${prevY}_daily`] || 0)} ביום`, n];
                      return [M.fmt(v), n];
                    })} />
                    <Legend wrapperStyle={{ direction: "rtl", fontSize: 12 }} payload={[
                      { value: `${prevY} (אשתקד)`, type: "line", color: C.dim },
                      { value: `${y} (השנה)`, type: "line", color: C.teal },
                    ]} />
                    {holidays.map((h, i) => {
                      const col = h.kind === "user" ? C.amber : h.kind === "prev" ? "#9BB0D4" : C.teal;
                      const lbl = h.kind === "prev" ? `${h.label} '${String(prevY).slice(2)}` : h.kind === "cur" ? `${h.label} '${String(y).slice(2)}` : h.label;
                      return (
                        <ReferenceLine key={"h" + i} x={`${String(h.day).padStart(2, "0")}.${mo}`}
                          stroke={col} strokeDasharray={h.kind === "prev" ? "2 4" : "5 3"} strokeOpacity={h.kind === "prev" ? 0.75 : 1}
                          label={{ value: lbl, fill: col, fontSize: 11, fontWeight: 700, angle: -90, position: "insideTopRight", offset: 8 }} />
                      );
                    })}
                    <Line type="monotone" dataKey={`${prevY}`} name={`${prevY} (אשתקד)`} stroke={C.dim} strokeWidth={2} dot={false} connectNulls />
                    <Line type="monotone" dataKey={`${y}`} name={`${y} (השנה)`} stroke={C.teal} strokeWidth={3} dot={{ r: 4 }} connectNulls />
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
              <div style={{ color: C.dim, fontSize: 12, marginTop: 10 }}>
                ההשוואה "תפוח לתפוח" — אותו חלון ימים (1–{fullDate}). הדוח מדבר על יום קודם, לכן משווים את יום הדיווח מול אותו יום אשתקד. {yoyUpdatedAt ? `נתוני ${prevY} עודכנו: ${heDate(yoyUpdatedAt)}` : ""}
              </div>
              <div style={{ marginTop: 12, display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                <input className="qual-select" style={{ maxWidth: 160 }} placeholder="שם אירוע" value={evName} onChange={(e) => setEvName(e.target.value)} />
                <input className="qual-select" type="number" min="1" max="31" style={{ maxWidth: 90 }} placeholder="יום" value={evDay} onChange={(e) => setEvDay(e.target.value)} />
                <button className="btn ghost" onClick={async () => {
                  if (!evName || !evDay) return;
                  const iso = `${y}-${mm}-${String(+evDay).padStart(2, "0")}`;
                  const next = [...userEvents, { date: iso, label: evName }];
                  setUserEvents(next); await store.set("peres:events", next);
                  setEvName(""); setEvDay("");
                }}>➕ הוספת אירוע לציר</button>
                {userEvents.filter((e) => e.date && +e.date.split("-")[1] === mo).map((e, i) => (
                  <span key={"e" + i} className="tag" style={{ cursor: "pointer" }} title="לחצו להסרה"
                    onClick={async () => { const next = userEvents.filter((x) => x !== e); setUserEvents(next); await store.set("peres:events", next); }}>
                    {e.label} ({e.date.split("-")[2]}.{mo}) ✕
                  </span>
                ))}
              </div>
            </div>

            {metric === "quality" && (
              <div className="panel" style={{ marginBottom: 18 }}>
                <h2>💰 עלות לליד איכותי מצטברת · {heMonth}: {y} מול {prevY}</h2>
                <div className="hint">הוצאה מצטברת ÷ לידים איכותיים מצטברים — מדד יציב ובר-השוואה (נמוך = יעיל יותר)</div>
                <div style={{ width: "100%", height: 260, direction: "ltr" }}>
                  <ResponsiveContainer>
                    <ComposedChart data={chartData} margin={{ top: 6, right: 8, left: 8, bottom: 0 }}>
                      <CartesianGrid stroke={C.panelSoft} vertical={false} />
                      <XAxis dataKey="day" {...axis} />
                      <YAxis {...axis} tickFormatter={(v) => "₪" + (v >= 1000 ? Math.round(v / 1000) + "K" : Math.round(v))} />
                      <Tooltip {...tip((v, n) => [v == null ? "—" : nis(v), n])} />
                      <Legend wrapperStyle={{ direction: "rtl", fontSize: 12 }} />
                      {holidays.map((h, i) => {
                        const col = h.kind === "user" ? C.amber : h.kind === "prev" ? "#9BB0D4" : C.teal;
                        return (
                          <ReferenceLine key={"ch" + i} x={`${String(h.day).padStart(2, "0")}.${mo}`}
                            stroke={col} strokeDasharray={h.kind === "prev" ? "2 4" : "5 3"} strokeOpacity={h.kind === "prev" ? 0.6 : 0.9} />
                        );
                      })}
                      <Line type="monotone" dataKey={`cpql_${prevY}`} name={`${prevY} (אשתקד)`} stroke={C.dim} strokeWidth={2} dot={false} connectNulls />
                      <Line type="monotone" dataKey={`cpql_${y}`} name={`${y} (השנה)`} stroke={C.coral} strokeWidth={3} dot={false} connectNulls />
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              </div>
            )}

            {(metric === "gross" || metric === "quality") && (
              <div className="panel" style={{ marginBottom: 18 }}>
                <h2>📊 {M.label} לפי יום · {heMonth}: {y} מול {prevY}</h2>
                <div className="hint">
                  כמה {M.label} נכנסו בכל יום בנפרד (לא מצטבר) — הפיקים היומיים.
                  {hasCurDaily ? ` נתוני ${y} מחושבים מהפרש הדוחות היומיים בארכיון.` : ` נתוני ${y} ייכנסו ככל שייצבר ארכיון יומי.`}
                </div>
                <div style={{ width: "100%", height: 260, direction: "ltr" }}>
                  <ResponsiveContainer>
                    <ComposedChart data={chartData} margin={{ top: 6, right: 8, left: 8, bottom: 0 }}>
                      <CartesianGrid stroke={C.panelSoft} vertical={false} />
                      <XAxis dataKey="day" {...axis} />
                      <YAxis {...axis} allowDecimals={false} />
                      <Tooltip {...tip((v, n) => [v == null ? "—" : num(v), n])} />
                      <Legend wrapperStyle={{ direction: "rtl", fontSize: 12 }} />
                      {holidays.map((h, i) => {
                        const col = h.kind === "user" ? C.amber : h.kind === "prev" ? "#9BB0D4" : C.teal;
                        return (
                          <ReferenceLine key={"dh" + i} x={`${String(h.day).padStart(2, "0")}.${mo}`}
                            stroke={col} strokeDasharray={h.kind === "prev" ? "2 4" : "5 3"} strokeOpacity={h.kind === "prev" ? 0.6 : 0.9} />
                        );
                      })}
                      {hasCurDaily && <Bar dataKey={`${y}_daily`} name={`${y} · יומי`} fill={C.teal} radius={[3, 3, 0, 0]} maxBarSize={14} />}
                      <Bar dataKey={`${prevY}_daily`} name={`${prevY} · יומי`} fill="#7FB2E8" radius={[3, 3, 0, 0]} maxBarSize={14} />
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              </div>
            )}
          </>
        );
      })()}

      {/* ---------- ארכיון · לוח שנה ---------- */}
      <div className="panel cal-cta" style={{ marginBottom: 26 }} onClick={() => { refreshDayIndex(); setCalOpen(true); }}>
        <div>
          <h2>📆 ארכיון דוחות · לוח שנה</h2>
          <div className="hint" style={{ marginBottom: 0 }}>
            {dayIndex.length
              ? `${dayIndex.length} דוחות שמורים · האחרון: ${dayIndex[dayIndex.length - 1].split("-").reverse().join(".")} · לחצו לפתיחת לוח שנה שנתי`
              : "כל דוח שנטען נשמר אוטומטית לתאריך של יום קודם · לחצו לפתיחת לוח השנה, צפייה בדוחות קודמים והזנת דוחות לימים ספציפיים"}
          </div>
        </div>
        <div className="cal-cta-arrow">◀</div>
      </div>

      <div className="foot">
        ירוק = איכות ≥10% · צהוב = 6–10% · אדום = מתחת ל־6% · הקו הלבן בכרטיסי היעד מסמן 100% עמידה ביעד · ▲▼ = שינוי מאז הדוח הקודם
      </div>

      {story && <StoryMode slides={buildSlides(view.data, totals, view.delta, insights, isActive, view.label)} onClose={() => setStory(false)} onShare={shareStory} />}

      {/* ---------- לוח שנה שנתי ---------- */}
      {calOpen && (
        <div className="cal-overlay" dir="rtl" onClick={() => setCalOpen(false)}>
          <div className="cal-sheet" onClick={(e) => e.stopPropagation()}>
            <div className="cal-head">
              <div className="cal-year-nav">
                <button className="btn ghost" onClick={() => setCalYear((y) => y - 1)}>▶ {calYear - 1}</button>
                <span className="cal-year">{calYear}</span>
                <button className="btn ghost" onClick={() => setCalYear((y) => y + 1)}>{calYear + 1} ◀</button>
              </div>
              <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                <span className={`cloud-pill ${cloudUp ? "up" : "down"}`} title={cloudUp ? "מחובר לאחסון ענן — כל המכשירים חולקים את אותה היסטוריה" : "אחסון ענן לא זמין — עובד מקומית"}>
                  {cloudUp ? "☁️ ענן מחובר" : "⚠️ ענן מנותק"}
                </span>
                <button className="btn ghost" disabled={!cloudUp || cloudBusy} onClick={async () => {
                  setCloudBusy(true);
                  const n = await store.pushAllToCloud();
                  await refreshDayIndex(); await refreshTrend();
                  setCloudBusy(false);
                  setStatus({ kind: "ok", msg: `הועלו ${n} רשומות לענן — ההיסטוריה הקיימת זמינה עכשיו לכל המכשירים` });
                }} title="דחיפת כל הארכיון המקומי לענן — פעולה חד-פעמית להעברת היסטוריה קיימת">
                  {cloudBusy ? "⏳ מעלה…" : "☁️⬆️ העלאה לענן"}
                </button>
                <button className="btn ghost" onClick={exportArchive} title="הורדת כל הארכיון כקובץ — גיבוי מקומי">⬇️ גיבוי</button>
                <label className="btn ghost file-btn" title="ייבוא קובץ גיבוי שהורד בדפדפן אחר">
                  ⬆️ ייבוא
                  <input type="file" accept=".json,application/json" onChange={importArchive} />
                </label>
                <button className="btn ghost" onClick={async () => {
                  const y = new Date(); y.setDate(y.getDate() - 1);
                  const iso = viewingDay || isoDay(y);
                  await store.set(dayKey(iso), { savedAt: new Date().toISOString(), reportDay: iso, data });
                  await refreshDayIndex();
                  setStatus({ kind: "ok", msg: `הבורד הנוכחי נשמר לארכיון לתאריך ${iso.split("-").reverse().join(".")}` });
                }}>💾 שמירת הבורד הנוכחי</button>
                <button className="story-close" style={{ position: "static" }} onClick={() => setCalOpen(false)}>✕</button>
              </div>
            </div>
            <div className="cal-legend">
              <span><span className="cal-dot has" /> יש דוח שמור — לחיצה פותחת אותו</span>
              <span><span className="cal-dot" /> אין דוח — לחיצה מאפשרת להזין דוח לאותו יום</span>
              <span>{cloudUp ? "☁️ הארכיון מסונכרן בענן — כל המכשירים רואים את אותה היסטוריה אוטומטית" : "💾 הענן לא זמין כרגע — עובד מקומית; לסנכרון ידני: גיבוי כאן ← ייבוא שם"}</span>
            </div>
            <div className="cal-months">
              {Array.from({ length: 12 }, (_, mi) => {
                const monthName = new Date(calYear, mi, 1).toLocaleDateString("he-IL", { month: "long" });
                const firstDow = new Date(calYear, mi, 1).getDay(); /* 0=ראשון */
                const daysInM = new Date(calYear, mi + 1, 0).getDate();
                const todayIso = isoDay(new Date());
                return (
                  <div className="cal-month" key={mi}>
                    <div className="cal-month-name">{monthName}</div>
                    <div className="cal-grid">
                      {["א", "ב", "ג", "ד", "ה", "ו", "ש"].map((d) => (
                        <div key={d} className="cal-dow">{d}</div>
                      ))}
                      {Array.from({ length: firstDow }, (_, i) => <div key={"e" + i} />)}
                      {Array.from({ length: daysInM }, (_, di) => {
                        const iso = `${calYear}-${String(mi + 1).padStart(2, "0")}-${String(di + 1).padStart(2, "0")}`;
                        const has = dayIndex.includes(iso);
                        const isToday = iso === todayIso;
                        return (
                          <button
                            key={iso}
                            className={`cal-day ${has ? "has" : ""} ${isToday ? "today" : ""}`}
                            title={has ? "פתיחת הדוח של " + iso : "הזנת דוח ליום " + iso}
                            onClick={() => (has ? openDay(iso) : (setCalPaste({ date: iso }), setCalPasteText("")))}
                          >
                            {di + 1}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {/* ---------- הזנת דוח ליום ספציפי ---------- */}
      {calPaste && (
        <div className="share-overlay" onClick={() => setCalPaste(null)}>
          <div className="share-modal" dir="rtl" onClick={(e) => e.stopPropagation()}>
            <div className="share-head">
              <strong>הזנת דוח לתאריך {heDayLabel(calPaste.date)}</strong>
              <button className="story-close" style={{ position: "static" }} onClick={() => setCalPaste(null)}>✕</button>
            </div>
            <div style={{ color: calPaste.error ? C.coral : C.dim, fontSize: 13 }}>
              {calPaste.error
                ? "לא זוהה מבנה הדוח — ודאו שהעתקתם את כל הגיליון כולל הכותרות"
                : "הדביקו את תוכן ה-CSV של הדוח כפי שהיה באותו יום (קובץ ← הורדה ← CSV בגיליון). הדוח יישמר בארכיון לתאריך הזה."}
            </div>
            <textarea className="share-text" placeholder="הדביקו כאן את תוכן ה-CSV..."
                      value={calPasteText} onChange={(e) => setCalPasteText(e.target.value)} />
            <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
              <button className="btn" onClick={() => loadCsv(calPasteText, calPaste.date)}>שמירה לארכיון</button>
              <label className="btn file-btn">
                📁 או העלאת קובץ CSV / XLSX
                <input type="file" accept=".csv,.xlsx,.xlsm,.xls,text/csv,text/plain" onChange={(e) => handleFile(e, calPaste.date)} />
              </label>
              <button className="btn ghost" onClick={() => setCalPaste(null)}>ביטול</button>
            </div>
          </div>
        </div>
      )}

      {/* ---------- חלון שיתוף ---------- */}
      {shareModal && (
        <div className="share-overlay" onClick={() => setShareModal(null)}>
          <div className="share-modal" dir="rtl" onClick={(e) => e.stopPropagation()}>
            <div className="share-head">
              <strong>{shareModal.title}</strong>
              <button className="story-close" style={{ position: "static" }} onClick={() => setShareModal(null)}>✕</button>
            </div>
            {shareModal.url && (
              <>
                <div className="share-copied" style={{ color: shareModal.copied ? C.teal : C.dim }}>
                  {shareModal.copied ? "✓ הקישור הועתק ללוח — מי שיפתח אותו יראה בדיוק את תמונת המצב הזו" : "🔗 קישור לדשבורד עם הנתונים הנוכחיים:"}
                </div>
                <div className="share-link-row">
                  <input className="share-link" readOnly value={shareModal.url} onFocus={(e) => e.target.select()} dir="ltr" />
                  <button className="btn" onClick={async () => {
                    let ok = false;
                    try { await navigator.clipboard.writeText(shareModal.url); ok = true; } catch {}
                    if (!ok) {
                      const inp = document.querySelector(".share-link");
                      if (inp) { inp.focus(); inp.select(); try { ok = document.execCommand("copy"); } catch {} }
                    }
                    setShareModal({ ...shareModal, copied: ok });
                  }}>🔗 העתקת קישור</button>
                </div>
                <div style={{ color: C.dim, fontSize: 12 }}>
                  הקישור מכיל את הנתונים עצמם — הוא יעבוד לכל מי שיש לו גישה לכתובת שבה הדשבורד מפורסם, גם בלי גישה לגיליון.
                </div>
              </>
            )}
            <div className="share-copied" style={{ color: C.dim, marginTop: shareModal.url ? 6 : 0 }}>או סיכום טקסט להדבקה:</div>
            <textarea className="share-text" readOnly value={shareModal.text} onFocus={(e) => e.target.select()} />
            <div style={{ display: "flex", gap: 10 }}>
              <button className="btn ghost" onClick={async () => {
                let ok = false;
                try { await navigator.clipboard.writeText(shareModal.text); ok = true; } catch {}
                if (!ok) {
                  const ta = document.querySelector(".share-text");
                  if (ta) { ta.focus(); ta.select(); try { ok = document.execCommand("copy"); } catch {} }
                }
                setShareModal({ ...shareModal, copied: ok });
              }}>📋 העתקת הסיכום</button>
              <button className="btn ghost" onClick={() => setShareModal(null)}>סגירה</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
