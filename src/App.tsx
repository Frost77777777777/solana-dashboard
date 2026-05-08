import React, { useRef, useState, useMemo, useEffect, useCallback, memo, Component } from "react";
import { createPortal } from "react-dom";
import { Upload, X, Search, Sun, Moon, TrendingUp, TrendingDown, RefreshCw, Store, CalendarDays, Building2, ChevronDown, HardDrive, Menu } from "lucide-react";
import * as XLSX from "xlsx";
import { AreaChart, Area, BarChart, Bar, ComposedChart, LineChart, Line, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend, ReferenceLine } from "recharts";

/* ─── types ──────────────────────────────────────────────────── */
interface Row { [key: string]: unknown }

interface Cols {
  revenue:     string | null;
  delivery:    string | null;
  commission:  string | null;
  debt:        string | null;
  brand:       string | null; // магазин column (single-sheet mode)
  date:        string | null;
  status:      string | null;
  refusalDate: string | null; // дата_відмови
  reason:      string | null; // причина відмови
  phone:       string | null; // окремий телефон (якщо є)
  city:        string | null; // місто / регіон
  product:     string | null;
  customer:    string | null;
  quantity:    string | null;
  orderId:     string | null; // TTN / order number — for cross-sheet deduplication
}

interface FileData {
  fileName:     string;
  columns:      string[];  // all unique columns across sheets
  rows:         Row[];
  cols:         Cols;
  isMultiSheet: boolean;
  sheetNames:   string[];
  cacheVersion: string;    // must equal STORAGE_KEY — prevents stale HMR data from polluting cache
}

const STORAGE_KEY   = "orbit_analytics_v24";
const HUBBER_KEY    = "orbit_hubber_v1";
const FILTERS_KEY   = "orbit_filters_v1";
const PRESET_BRANDS = ["Каста", "Розетка", "Хаббер", "Шафа"];

/* ─── Base Blue design-system palette ───────────────────────── */
// Index 0 = leader (Deep Navy), index 1 = Electric Blue, rest lighter
const BASE_BLUE = ["#0052FF","#3376FF","#66A0FF","#99C5FF","#CCE2FF","#6B7280"] as const;
const NO_DATE_KEY   = "~~~~"; // sorts after all month keys (e.g. "2026-12")

/* ─── financial parsers ──────────────────────────────────────── */
// parseFinancial — matches reference "MASTER BLUE" parse():
//   strips whitespace only, replaces comma→dot, no large-number guard.
//   Use for income / fee / ship (financial amounts can legitimately be large).
function parseFinancial(val: unknown): number {
  if (val == null || val === "") return 0;
  if (typeof val === "number") return isFinite(val) ? val : 0;
  const n = parseFloat(
    String(val)
      .replace(/\s/g, "")          // strip whitespace (spaces, nbsp)
      .replace(",", ".")            // UA decimal comma → dot
  );
  return isFinite(n) ? n : 0;
}

// toNum — general safe parser with anti-ID guard:
//   strips ALL non-numeric chars, returns 0 for TTNs / phone numbers > 50 000.
//   Use for debt, quantity, and any column that may hold numeric IDs.
function toNum(val: unknown): number {
  if (val == null || val === "" || val === false) return 0;
  const n = parseFloat(String(val).replace(/[^\d.-]/g, "").replace(",", ".")) || 0;
  if (!isFinite(n) || isNaN(n)) return 0;
  if (Math.abs(n) > 50_000) return 0;
  return n;
}
// Alias for any legacy call sites
const parseNum = toNum;

/* ─── column detection ───────────────────────────────────────── */
// Keywords are checked in PRIORITY ORDER — the first keyword that has a matching
// column wins, regardless of which column appears first in the array.
function findCol(cols: string[], ...kw: string[]): string | null {
  const lc = cols.map(c => c.toLowerCase().trim());
  // Phase 1: exact match, keyword priority
  for (const k of kw) {
    const idx = lc.findIndex(c => c === k.toLowerCase());
    if (idx >= 0) return cols[idx];
  }
  // Phase 2: substring match, keyword priority
  for (const k of kw) {
    const idx = lc.findIndex(c => c.includes(k.toLowerCase()));
    if (idx >= 0) return cols[idx];
  }
  return null;
}

function detectCols(columns: string[], rows: Row[]): Cols {
  const status =
    findCol(columns, "статус", "status") ??
    columns.find(c =>
      rows.slice(0, 300).some(r => {
        const v = String(r[c] ?? "").toLowerCase();
        return v.includes("відмова") || v.includes("повернення") || v.includes("успіш");
      })
    ) ?? null;

  const refusalDate = findCol(columns, "ттн повернення", "дата_відмови", "дата відмови", "refusal_date", "дата повернення", "причина повернення");
  const brand       = findCol(columns, "магазин", "бренд", "brand", "shop", "store");
  const date        = columns.find(c => {
    const cl = c.toLowerCase();
    return (cl.includes("дата") || cl.includes("місяць") || cl.includes("date") || cl.includes("month"))
      && !cl.includes("відмов") && !cl.includes("повернен");
  }) ?? null;

  // For delivery cost: when Excel has two "доставка" columns, XLSX renames the second as "доставка_1".
  // The FIRST is the address text, the LAST is the cost (negative UAH).
  // Match: "доставка", "доставка_1", "доставка_2", etc. — always pick the last one.
  const deliveryAll = columns.filter(c => /^доставка(_\d+)?$/i.test(c.trim()));
  let delivery: string | null = null;
  let cityFromDelivery: string | null = null;

  if (deliveryAll.length >= 2) {
    // Two доставка cols: first = address, last = numeric cost
    cityFromDelivery = deliveryAll[0];
    delivery         = deliveryAll[deliveryAll.length - 1];
  } else if (deliveryAll.length === 1) {
    const col    = deliveryAll[0];
    // Sample up to 60 non-empty values to decide: numeric → cost; text → address
    const sample = rows.slice(0, 100)
      .map(r => String(r[col] ?? "").trim())
      .filter(v => v.length > 0)
      .slice(0, 60);
    const numericCount = sample.filter(v => !isNaN(parseFloat(v.replace(/[\s,]/g, "")))).length;
    const isNumeric    = sample.length > 0 && numericCount / sample.length >= 0.5;
    if (isNumeric) {
      delivery = col;           // numeric → delivery cost column
    } else {
      cityFromDelivery = col;   // text addresses → city column
      delivery = null;
    }
  } else {
    delivery = findCol(columns, "delivery", "shipping");
  }

  return {
    revenue:     findCol(columns, "дохід", "дохід від продажу", "income", "сума замовлення", "сума зам", "revenue", "amount"),
    delivery,
    commission:  findCol(columns, "комісія банку", "комісія", "commission", "fee"),
    debt:        findCol(columns, "борг", "debt"),
    product:     findCol(columns, "назва товару", "товар", "артикул", "sku", "продукт", "product", "item", "найменування", "назва"),
    customer:    findCol(columns, "піб", "прізвище", "ім'я", "клієнт", "покупець", "customer", "отримувач", "одержувач", "телефон", "phone"),
    quantity:    findCol(columns, "кількість", "к-сть", "qty", "quantity"),
    orderId:     findCol(columns, "ттн", "ttn", "номер накладної", "номер відправлення", "номер замовлення", "order id", "track"),
    reason:      findCol(columns, "причина відмови", "причина повернення", "причина", "коментар", "comment", "reason", "rejection reason"),
    phone:       findCol(columns, "телефон", "phone", "тел", "контакт", "contact", "мобільний"),
    city:        cityFromDelivery
                 ?? findCol(columns, "місто", "місто отримувача", "місто одержувача", "населений пункт", "city", "регіон", "область", "region", "district"),
    brand, date, status, refusalDate,
  };
}

/* ─── city name extractor ────────────────────────────────────── */
// Handles addresses like "Нова пошта №106, Одеса, вул. Хрещатик 1"
// → finds the first comma-segment that is NOT a carrier/branch name → "Одеса"
// Also normalises "м. Київ" → "Київ", strips postal codes.
const CARRIER_RE = /нова[\s_]?пошта|нп\s*[№#\d]|укрпошта|nova[\s_]?poshta|meest|justin|відділення|поштомат|сільпо|silpo|rozetka|розетка|[№#]\s*\d|\bвідд\b/i;
function extractCity(raw: string): string {
  const parts = raw.split(",").map(p => p.trim()).filter(p => p.length > 0);
  // Find the first segment that doesn't look like a carrier, branch number, or street
  let city = "";
  for (const part of parts) {
    if (CARRIER_RE.test(part))       continue; // skip carrier/branch
    if (/^[№#\d]/.test(part))       continue; // skip pure branch numbers
    if (/^вул\.|^пров\.|^пр\.|^бул\.|^просп/i.test(part)) break; // stop at street
    city = part;
    break;
  }
  if (!city && parts.length > 0) city = parts[0]; // fallback: first segment
  // Strip city-type abbreviations: "м.", "смт.", "с." etc.
  city = city.replace(/^(м\.\s*|м\s+|смт\.\s*|смт\s+|с\.\s+|с-ще\s+)/i, "").trim();
  // Strip leading postal codes
  city = city.replace(/^\d{5}\s*/, "").trim();
  // Strip trailing store/branch noise like "Сільпо", "Billa" etc. after the city name
  city = city.replace(/\s+(сільпо|silpo|billa|novus|atb|атб|metro|метро|ашан)\s*$/i, "").trim();
  if (!city) return "";
  return city.charAt(0).toUpperCase() + city.slice(1);
}

/* ─── multi-sheet product / customer fallback lookup ─────────── */
// Each sheet may use a different column name for the same concept.
// Try the cols-detected primary first, then the known candidate list.
const PRODUCT_CANDIDATES = ["назва товару","товар","артикул","sku","продукт","product","item","найменування","назва"];
const CUSTOMER_CANDIDATES = ["піб","прізвище","клієнт","покупець","customer","отримувач","одержувач","телефон","phone"];

/* ─── product key normalizer — strips colors/sizes, keeps Name + Model ── */
// Input: "Бюстгальтер 264 Рожева перлина 80C"
// Output: "Бюстгальтер, 264"
const COLORS_UA = new Set([
  "білий","біла","біле","білі","чорний","чорна","чорне","чорні",
  "рожевий","рожева","рожеве","рожеві","перлина",
  "сірий","сіра","сіре","сірі","бежевий","бежева","бежеве","бежеві",
  "синій","синя","синє","сині","червоний","червона","червоне","червоні",
  "зелений","зелена","зелене","зелені","жовтий","жовта","жовте","жовті",
  "фіолетовий","фіолетова","фіолетове","фіолетові",
  "пудровий","пудрова","пудрове","пудрові",
  "шоколадний","шоколадна","молочний","молочна",
  "карамель","вишневий","вишнева","тілесний","тілесна",
  "бузковий","бузкова","персиковий","персикова","кораловий","коралова",
  "nude","ivory","black","white","cream","pink","gray","grey","beige",
]);

/* Neon palette for Income-by-Year chart — high contrast on #000000 */
const HUBBER_NEON_COLORS = [
  "#14F195", // Solana Green
  "#9945FF", // Cyber Purple
  "#00C2FF", // Electric Blue
  "#FFD700", // Gold
  "#FF0040", // Ruby
  "#FF6B00", // Neon Orange
  "#00FFD1", // Mint
  "#FF69B4", // Hot Pink
  "#7B68EE", // Slate Blue
  "#ADFF2F", // Lime
];
function hubberYearColor(year: string, years: string[]): string {
  const idx = years.indexOf(year);
  return HUBBER_NEON_COLORS[idx >= 0 ? idx % HUBBER_NEON_COLORS.length : 0];
}

function normalizeProductKey(raw: string): string {
  const tokens = raw.split(/[\s,;/\\]+/).filter(t => t.length > 0);
  let productName = "";
  let modelNum    = "";
  for (const token of tokens) {
    const lower = token.toLowerCase();
    if (COLORS_UA.has(lower)) continue;                              // skip color words
    if (/^\d{2}[A-Ha-h]$/.test(token)) continue;                   // skip bra sizes: 65B, 80C…
    if (/^(XS|S|M|L|XL|XXL|2XL|3XL|XXXL)$/i.test(token)) continue; // skip clothing sizes
    if (/^[3-5]\d$/.test(token)) continue;                          // skip numeric sizes 30-59
    if (/^\d{2,5}([/-]\d+)?$/.test(token) && !modelNum) { modelNum = token; continue; } // model number
    if (!productName && /[а-яА-ЯїЇіІєЄ]/.test(token)) {
      productName = token.charAt(0).toUpperCase() + token.slice(1).toLowerCase();
    }
  }
  if (productName && modelNum) return `${productName}, ${modelNum}`;
  if (productName) return productName;
  if (modelNum)    return modelNum;
  return tokens[0] || raw;
}

function getRowProduct(row: Row, primaryCol: string | null): string {
  const candidates = primaryCol ? [primaryCol, ...PRODUCT_CANDIDATES] : PRODUCT_CANDIDATES;
  for (const c of candidates) {
    const v = String(row[c] ?? "").trim();
    if (v) return v;
  }
  return "";
}

function getRowCustomer(row: Row, primaryCol: string | null): string {
  const candidates = primaryCol ? [primaryCol, ...CUSTOMER_CANDIDATES] : CUSTOMER_CANDIDATES;
  for (const c of candidates) {
    const v = String(row[c] ?? "").trim();
    if (v) return v;
  }
  return "";
}

function getRowPhone(row: Row, phoneCol: string | null): string {
  if (phoneCol) {
    const v = String(row[phoneCol] ?? "").trim();
    if (v) return v;
  }
  // fallback: look for any value that looks like a phone number
  for (const c of CUSTOMER_CANDIDATES) {
    const v = String(row[c] ?? "").trim();
    if (/^\+?[\d\s\-()]{7,}$/.test(v)) return v;
  }
  return "";
}

/* ─── phone masking — partial hide for privacy ───────────────── */
// +380671231234 → +38067***234   |  0671231234 → 067***234
function maskPhone(raw: string): string {
  const s = raw.trim();
  const digits = s.replace(/\D/g, "");
  if (digits.length < 7) return s; // not a phone — return as-is
  // Keep first 5 chars + last 4 digits, mask the middle
  const suffix = s.slice(-4);
  const prefix = s.slice(0, Math.max(2, s.length - suffix.length - 3));
  return `${prefix}***${suffix}`;
}

/* ─── looks-like-phone heuristic ─────────────────────────────── */
function isPhoneString(s: string): boolean {
  return /^\+?[\d\s\-().]{7,}$/.test(s.trim()) && s.replace(/\D/g,"").length >= 7;
}

/* ─── refusal detection ──────────────────────────────────────── */
// A row is a refusal when:
//   1. ттн повернення column has a value (TTN return number, length > 5)
//   2. дохід < 0 (marketplace chargeback — the negative value IS the loss)
//   3. status contains відмова/повернення
// NOTE: do NOT zero the income for refusals. Negative дохід already represents the penalty.
function isRefusal(row: Row, c: Cols): boolean {
  if (c.revenue) {
    const inc = parseNum(row[c.revenue]);
    if (inc < 0) return true;
  }
  if (c.refusalDate) {
    const v = String(row[c.refusalDate] ?? "").trim();
    if (v.length > 5) return true;
  }
  if (c.status) {
    const v = String(row[c.status] ?? "").toLowerCase().trim();
    if (v.includes("відмова") || v.includes("повернення") || v.includes("refus")) return true;
  }
  return false;
}

/* ─── row net income (fallback — prefer r._net from stampRows) ── */
function rowNet(row: Row, c: Cols): number {
  if (row._net !== undefined) return row._net as number;
  const inc  = parseFinancial(c.revenue    ? row[c.revenue]    : null);
  const fee  = Math.abs(parseFinancial(c.commission ? row[c.commission] : null));
  const ship = Math.abs(parseFinancial(c.delivery   ? row[c.delivery]   : null));
  return inc - (fee + ship);
}

/* ─── BRUTE FORCE row processor ────────────────────────────────
   Exact formula as specified. No column detection indirection.
   Hardcoded Ukrainian column names. Called at upload + restore. ── */
function stampRows(rows: Row[], cols: Cols): void {
  // Resolve column names once — use detected names (case-preserving) with hardcoded fallbacks.
  // This is the root fix: if the file has "Дохід" (capital) and we searched "дохід" (lower),
  // the detected cols.revenue will have the EXACT key; hardcoded lookup returns 0.
  const revenueCol  = cols.revenue  ?? "дохід";
  const commCol     = cols.commission ?? "комісія банку";
  const deliveryCol = cols.delivery;            // already detected, never hardcode
  const brandCol    = cols.brand    ?? "магазин";
  const dateCol     = cols.date     ?? "Дата";

  // Debug: show which column names are actually being used (printed once, helps diagnose)
  console.log("[stampRows] cols →", { revenueCol, commCol, deliveryCol, brandCol, dateCol });
  // Sample first row's raw values for those columns (shows if lookup is working)
  if (rows.length > 0) {
    const s = rows[0];
    console.log("[stampRows] row[0] sample →", {
      [revenueCol]:  s[revenueCol],
      [commCol]:     s[commCol],
      [deliveryCol ?? "delivery(null)"]: deliveryCol ? s[deliveryCol] : "N/A",
      [brandCol]:    s[brandCol],
      [dateCol]:     s[dateCol],
    });
  }

  for (const r of rows) {

    // ── getVal: strip currency/space formatting, return numeric value ──
    const getVal = (key: string | null): number => {
      if (!key) return 0;
      const v = r[key];
      if (v == null || v === "") return 0;
      if (typeof v === "number") return isFinite(v) ? v : 0;
      // String: strip spaces/currency symbols but keep minus sign and decimal
      const n = parseFloat(
        String(v)
          .replace(/\s/g, "")          // remove all whitespace (incl. nbsp)
          .replace(",", ".")           // UA decimal comma
          .replace(/[^\d.+-]/g, "")   // keep digits, dot, plus, minus
      );
      return isNaN(n) ? 0 : n;
    };

    const income = getVal(revenueCol);
    const fee    = Math.abs(getVal(commCol));
    const ship   = Math.abs(getVal(deliveryCol));
    const net    = income - fee - ship;

    r._gross = income;
    r._fee   = fee;
    r._ship  = ship;
    // CLEANING RULE (user spec): single-row net outside ±50,000 = parser bug → zero
    r._net   = Math.abs(net) > 50_000 ? 0 : net;

    // brand/marketplace — use detected column name, not hardcoded "магазин"
    const rawMkt = String(r[brandCol] ?? "").trim();
    r._mkt = rawMkt
      ? rawMkt.charAt(0).toUpperCase() + rawMkt.slice(1).toLowerCase()
      : "";

    // date → YYYY-MM key
    // Strategy: for DD.MM.YYYY strings (Ukrainian standard), extract directly — no Date object,
    // no timezone conversion. Falls back to UTC methods for other formats/numbers.
    const rawDate = r[dateCol];
    r._monthKey = extractMonthKey(rawDate);
  }

  // First 5 marketplace rows — shows exact income/fee/ship/net to verify polarity
  const sampleRows = rows.filter(r => r._mkt !== "").slice(0, 5);
  if (sampleRows.length) {
    console.log("[stampRows] First 5 marketplace rows:");
    console.table(sampleRows.map(r => ({
      "_mkt":     r._mkt,
      "_monthKey":r._monthKey,
      "income":   Math.round((r._gross as number) * 100) / 100,
      "fee":      Math.round((r._fee   as number) * 100) / 100,
      "ship":     Math.round((r._ship  as number) * 100) / 100,
      "_net":     Math.round((r._net   as number) * 100) / 100,
    })));
  }

  // MANDATORY CONSOLE CHECK — filtered to marketplace rows only (matches KPI on screen)
  const mktRows = rows.filter(r => r._mkt !== "");
  const total   = mktRows.reduce((s, r) => s + (r._net as number), 0);
  console.log("CRITICAL CHECK: Total Net Sum is:", Math.round(total * 100) / 100);
  console.log("Rows processed:", mktRows.length);
  // Breakdown by marketplace
  const byMkt: Record<string, number> = {};
  for (const r of mktRows) {
    const m = r._mkt as string;
    byMkt[m] = (byMkt[m] || 0) + (r._net as number);
  }
  console.table(byMkt);
  // Per-month breakdown (for filter verification)
  const byMonth: Record<string, { "Net ₴": number; Рядки: number }> = {};
  for (const r of mktRows) {
    const mk = String(r._monthKey);
    const label = mk === "No Date" ? "Без дати" : mk;
    if (!byMonth[label]) byMonth[label] = { "Net ₴": 0, Рядки: 0 };
    byMonth[label]["Net ₴"] += r._net as number;
    byMonth[label].Рядки++;
  }
  // Round for readability
  for (const k of Object.keys(byMonth)) byMonth[k]["Net ₴"] = Math.round(byMonth[k]["Net ₴"]);
  console.log("By Month:");
  console.table(byMonth);

  // DIAGNOSIS: show raw дата values for rows in unexpected months
  // Jan-Apr 2026 are expected; Dec 2025 and May-Dec 2026 need inspection
  const unexpectedMonths = Object.keys(byMonth).filter(k => k !== "Без дати" && k !== "2026-01" && k !== "2026-02" && k !== "2026-03" && k !== "2026-04");
  if (unexpectedMonths.length > 0) {
    console.log("⚠ Rows outside Jan-Apr 2026 — raw дата values:");
    const diagRows = mktRows
      .filter(r => unexpectedMonths.includes(String(r._monthKey)))
      .slice(0, 10)
      .map(r => ({
        _mkt:     r._mkt,
        _monthKey:r._monthKey,
        дата_raw: r[dateCol],
        _net:     Math.round((r._net as number) * 100) / 100,
      }));
    console.table(diagRows);
  }
}

/* ─── extractMonthKey — timezone-safe YYYY-MM from any date value ── */
// Handles Ukrainian DD.MM.YYYY strings WITHOUT creating a Date object, so
// there is zero timezone shift. Falls back to UTC-based parsing for other
// formats and Excel serial numbers.
function extractMonthKey(val: unknown): string {
  if (val == null || val === "") return "No Date";

  // ── JS Date object (from cellDates:true) ─────────────────────────────────
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "No Date";
    return `${val.getUTCFullYear()}-${String(val.getUTCMonth() + 1).padStart(2, "0")}`;
  }

  // ── Numeric Excel serial ──────────────────────────────────────────────────
  if (typeof val === "number") return excelSerialToMonthKey(val);

  // ── String value ──────────────────────────────────────────────────────────
  if (typeof val === "string") {
    const s = val.trim();
    if (!s) return "No Date";

    // ── YYYY-MM-DD (ISO 8601) — read month directly, no Date() ──
    const iso = s.match(/^(\d{4})-(\d{2})-\d{2}/);
    if (iso) return `${iso[1]}-${iso[2]}`;

    // ── DD.MM.YYYY / DD/MM/YYYY / DD.MM.YY / DD/MM/YY — Ukrainian convention ──
    // XLSX outputs dates from Ukrainian Excel files as DD.MM.YYYY (dots, 4-digit year)
    // OR DD/MM/YY (slashes, 2-digit year) — BOTH are day-first (Ukrainian/European).
    // CRITICAL: do NOT treat slash-separated as US MM/DD — this file uses DD/MM.
    // Diagnostic evidence: "08/01/26" = January 8, "05/03/26" = March 5 (not May/Aug).
    const dmyMatch = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{2,4})$/);
    if (dmyMatch) {
      const day = parseInt(dmyMatch[1], 10);
      const mon = parseInt(dmyMatch[2], 10);
      let   yr  = parseInt(dmyMatch[3], 10);
      // 2-digit year: 00-69 → 2000-2069, 70-99 → 1970-1999
      if (dmyMatch[3].length <= 2) yr += yr < 70 ? 2000 : 1900;
      if (day >= 1 && day <= 31 && mon >= 1 && mon <= 12 && yr >= 2000)
        return `${yr}-${String(mon).padStart(2, "0")}`;
    }

    // ── Numeric serial string — XLSX "General" format, e.g. "46023" ──
    const asNum = parseFloat(s);
    if (!isNaN(asNum) && asNum >= 25569 && asNum < 60000) {
      // Make sure the string really is an integer serial, not a decimal
      const rounded = Math.round(asNum);
      if (String(rounded) === s || `${rounded}.0` === s || s.match(/^\d+\.0+$/))
        return excelSerialToMonthKey(asNum);
    }

    // ── Last resort: native Date() — uses UTC month to avoid timezone shift ──
    const d = new Date(s);
    if (!isNaN(d.getTime())) return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
  }

  return "No Date";
}

function excelSerialToMonthKey(serial: number): string {
  if (serial < 1) return "No Date";
  const d = new Date((serial - 25569) * 86400 * 1000);
  if (isNaN(d.getTime())) return "No Date";
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}`;
}

/* ─── date helpers ───────────────────────────────────────────── */
function parseDate(val: unknown): Date | null {
  if (val == null || val === "") return null;
  // JS Date object (from cellDates:true when raw=true)
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  // Numeric Excel serial (e.g. 45037 = 2023-04-01)
  if (typeof val === "number") {
    if (val < 1) return null; // 0 = no date sentinel in some files
    const d = new Date((val - 25569) * 86400 * 1000);
    return isNaN(d.getTime()) ? null : d;
  }
  if (typeof val === "string") {
    const s = val.trim();
    if (!s) return null;
    // Numeric string serial — XLSX emits these when cell format is "General"
    // Valid Excel date range: ~25569 (1970-01-01) to ~60000 (2064-03-07)
    const asNum = parseFloat(s);
    if (!isNaN(asNum) && asNum >= 25569 && asNum < 60000 && String(asNum) === s.replace(/\.0+$/, "")) {
      const d = new Date((asNum - 25569) * 86400 * 1000);
      if (!isNaN(d.getTime())) return d;
    }
    // ISO 8601 / US M/D/YYYY — handled natively by Date constructor
    const d1 = new Date(s);
    if (!isNaN(d1.getTime())) return d1;
    // DD.MM.YYYY  or  DD/MM/YYYY  (Ukrainian convention)
    const m = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{2,4})$/);
    if (m) {
      const yr = m[3].length === 2 ? "20" + m[3] : m[3];
      const d2 = new Date(`${yr}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`);
      if (!isNaN(d2.getTime())) return d2;
    }
  }
  return null;
}
function toMonthKey(d: Date) { return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`; }
const MUK = ["Січень","Лютий","Березень","Квітень","Травень","Червень","Липень","Серпень","Вересень","Жовтень","Листопад","Грудень"];
function toMonthLabel(k: string) { const [y,m]=k.split("-"); return `${MUK[parseInt(m,10)-1]} ${y}`; }
function toMonthShort(k: string) { const [y,m]=k.split("-"); return `${MUK[parseInt(m,10)-1].slice(0,3)} '${y.slice(2)}`; }

/* ─── currency formatting ────────────────────────────────────── */
function fmt(n: number): string {
  const sign = n < 0 ? "−" : "";
  const [int, dec] = Math.abs(n).toFixed(2).split(".");
  return sign + int.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + "." + dec + " ₴";
}
function fmtK(n: number): string {
  if (!isFinite(n)) return "—";
  if (Math.abs(n) >= 1_000_000) return (n/1_000_000).toFixed(1)+"M ₴";
  if (Math.abs(n) >= 1_000) return (n/1_000).toFixed(0)+"k ₴";
  return fmt(n);
}

/* ─── Hubber quick data (sidebar mini-module) ─────────────────── */
interface HubberQuick {
  fileName: string;
  years: string[];
  months: string[];
  values: Record<string, Record<string, number>>;
  yearTotals: Record<string, number>;
}
const HUB_YEARS = ["2017","2018","2019","2020","2021","2022","2023","2024","2025","2026"];
function parseHubberQuick(file: File): Promise<HubberQuick> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target?.result, { type:"array" });
        const ws = wb.Sheets[wb.SheetNames.find(n=>n.trim()==="Дохід")??wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json<(string|number|null)[]>(ws, { header:1, defval:null });
        if (rows.length < 2) throw new Error("Empty sheet");
        const header = rows[0] as (string|number|null)[];
        const colMap: {col:number; label:string}[] = [];
        for (let c=1; c<header.length; c++) {
          const lbl = String(header[c]??"").trim();
          if (!lbl || lbl.toLowerCase().startsWith("unnamed")) continue;
          colMap.push({ col:c, label:lbl });
        }
        const yearCols = colMap.filter(x=>HUB_YEARS.includes(x.label));
        const months: string[] = [];
        const values: Record<string,Record<string,number>> = {};
        const yearTotals: Record<string,number> = {};
        for (let r=1; r<rows.length; r++) {
          const row = rows[r] as (string|number|null)[];
          const month = String(row[0]??"").trim();
          if (!month) continue;
          const isTotal = /^(всього|разом|итого)/i.test(month);
          if (isTotal) {
            for (const {col,label} of yearCols) {
              const n = Number(row[col]);
              if (isFinite(n)) yearTotals[label] = n;
            }
            continue;
          }
          months.push(month);
          for (const {col,label} of yearCols) {
            if (!values[label]) values[label] = {};
            const n = Number(row[col]);
            values[label][month] = isFinite(n) ? n : 0;
          }
        }
        resolve({ fileName:file.name, years:yearCols.map(x=>x.label), months, values, yearTotals });
      } catch(err) { reject(err instanceof Error?err:new Error(String(err))); }
    };
    reader.onerror = () => reject(new Error("Read error"));
    reader.readAsArrayBuffer(file);
  });
}

/* ─── theme ──────────────────────────────────────────────────── */
interface T { bg:string; card:string; nav:string; border:string; text:string; sub:string; dim:string; in:string; blue:string; em:string; red:string; amb:string; dark:boolean }
const DK: T = { bg:"#000000", card:"rgba(10,10,10,0.95)", nav:"rgba(5,5,5,0.97)", border:"rgba(255,255,255,0.10)", text:"#FFFFFF", sub:"rgba(200,200,200,0.7)", dim:"rgba(160,160,160,0.45)", in:"rgba(255,255,255,0.04)", blue:"#0EA5E9", em:"#14F195", red:"#E29578", amb:"#8A9A5B", dark:true };
const LT: T = {
  bg:     "#F8FAFC",
  card:   "#FFFFFF",
  nav:    "#FFFFFF",
  border: "#E5E7EB",
  text:   "#000000",
  sub:    "#111111",
  dim:    "#374151",
  in:     "#F1F5F9",
  blue:   "#0052FF",
  em:     "#16A34A",
  red:    "#FF4D4D",
  amb:    "#374151",
  dark:   false,
};

function glass(t: T, glow?: string): React.CSSProperties {
  if (t.dark) {
    return {
      background: t.card,
      border: `1px solid ${glow ? glow+"44" : t.border}`,
      borderRadius: 16,
      boxShadow: glow ? `0 0 24px ${glow}18` : "0 2px 12px rgba(0,0,0,0.35)",
    };
  }
  return {
    background: "#FFFFFF",
    border: `1px solid ${glow ? glow+"30" : "#E5E7EB"}`,
    borderRadius: 12,
    boxShadow: glow
      ? `0 0 16px ${glow}14, 0 2px 8px rgba(0,0,0,0.04)`
      : "0 1px 2px rgba(0,0,0,0.04), 0 4px 16px rgba(0,0,0,0.04)",
  };
}

/* ─── sub-components ─────────────────────────────────────────── */

/* Marketplace brand logo — pure SVG, no external assets */
/* ─── prev-month key helper ──────────────────────────────────── */
function prevMonthKey(mk: string): string {
  const [y, m] = mk.split("-").map(Number);
  if (m === 1) return `${y - 1}-12`;
  return `${y}-${String(m - 1).padStart(2, "0")}`;
}

/* ─── LFL growth badge ───────────────────────────────────────── */
function LflBadge({ current, previous, fmt: fmtFn, t }: { current: number; previous: number | null; fmt: (v: number) => string; t?: T }) {
  if (previous === null || previous === 0) return null;
  const delta = current - previous;
  const pct   = (delta / Math.abs(previous)) * 100;
  const up    = delta >= 0;
  const sign  = up ? "+" : "";
  const subColor = t?.dark ? "rgba(255,255,255,0.35)" : "#9CA3AF";
  // Pastel green chip for positive, soft red chip for negative
  const chipBg    = up ? "#DCFCE7" : "rgba(239,68,68,0.08)";
  const chipColor = up ? "#15803D" : "#DC2626";
  return (
    <div style={{ display:"flex", alignItems:"center", gap:5, marginTop:4 }}>
      <span style={{ display:"inline-flex", alignItems:"center", padding:"2px 7px", borderRadius:4, background:chipBg, fontSize:11, fontWeight:700, color:chipColor, letterSpacing:"-0.01em" }}>
        {up ? "↑" : "↓"} {sign}{pct.toFixed(1)}%
      </span>
      <span style={{ fontSize:10, color:subColor }}>vs {fmtFn(previous)}</span>
    </div>
  );
}

/* ─── Chart Error Boundary ────────────────────────────────────── */
class ChartErrorBoundary extends Component<
  { children: React.ReactNode; label?: string; t: T },
  { hasError: boolean; message: string }
> {
  constructor(props: { children: React.ReactNode; label?: string; t: T }) {
    super(props);
    this.state = { hasError: false, message: "" };
  }
  static getDerivedStateFromError(err: Error) {
    return { hasError: true, message: err?.message ?? "Unknown error" };
  }
  render() {
    const { t, label = "Chart" } = this.props;
    if (this.state.hasError) {
      return (
        <div style={{
          display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center",
          minHeight:180, gap:10, color:t.dim, fontSize:12,
        }}>
          <span style={{ fontSize:22, opacity:0.5 }}>⚠</span>
          <span style={{ fontWeight:600, color:t.sub }}>{label} — розрахунок завершився помилкою</span>
          <span style={{ fontSize:11, color:t.dim, maxWidth:280, textAlign:"center", opacity:0.7 }}>{this.state.message}</span>
          <button
            onClick={() => this.setState({ hasError:false, message:"" })}
            style={{ marginTop:4, padding:"4px 12px", borderRadius:6, border:`1px solid ${t.border}`, background:"transparent", color:t.sub, fontSize:11, cursor:"pointer" }}
          >Спробувати знову</button>
        </div>
      );
    }
    return this.props.children;
  }
}

/* ─── count-up animation component ───────────────────────────── */
function AnimNum({ value, fmt: fmtFn }: { value: number; fmt: (v: number) => string }) {
  const [disp, setDisp] = useState(value);
  const prevRef = useRef(value);
  const rafRef  = useRef<number>(0);
  useEffect(() => {
    const from = prevRef.current;
    const to   = value;
    // Skip animation when value hasn't changed meaningfully (avoids wasted RAF loops)
    if (Math.abs(to - from) < 0.001 * (Math.abs(to) || 1)) {
      setDisp(to);
      prevRef.current = to;
      return;
    }
    cancelAnimationFrame(rafRef.current);
    const dur   = 280;
    const start = performance.now();
    function tick(now: number) {
      const p = Math.min((now - start) / dur, 1);
      const e = 1 - Math.pow(1 - p, 3); // ease-out cubic
      setDisp(from + (to - from) * e);
      if (p < 1) {
        rafRef.current = requestAnimationFrame(tick);
      } else {
        prevRef.current = to;
      }
    }
    rafRef.current = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(rafRef.current);
  }, [value]);
  return <>{fmtFn(disp)}</>;
}

/* ─── memoized KPI row ──────────────────────────────────────── */
interface KpiRowProps {
  kpi: { net:number; returnRate:number; refs:number; orders:number; successOrders:number; debt:number; logistics:number; del:number; com:number; grossIncome:number; };
  prevKpi: { net:number; orders:number; logistics:number; } | null;
  hubberLfl: { curr:number; prev:number; pct:number; year:string; prevYear:string; monthName:string } | null;
  filteredCount: number;
  syncError: boolean;
  debtCol: string | null;
  t: T;
  fmt: (v:number)=>string;
}
const KPI_CARD_BASE: React.CSSProperties = {
  borderRadius:12,
  boxShadow:"0 1px 4px rgba(0,0,0,0.06)",
  padding:"22px 22px 18px",
  display:"flex", flexDirection:"column", justifyContent:"space-between",
  minHeight:136,
  contain:"layout",
};
const KPI_LABEL: React.CSSProperties = {
  fontSize:10, fontWeight:700, letterSpacing:"0.09em",
  textTransform:"uppercase" as const,
  color:"#6B7280",
};
const KPI_NUM: React.CSSProperties = {
  fontSize:26, fontWeight:900, letterSpacing:"-0.03em", lineHeight:1,
  margin:"8px 0 4px",
};

const KpiRow = memo(function KpiRow({ kpi, prevKpi, hubberLfl, filteredCount, syncError, debtCol, t, fmt }: KpiRowProps) {
  const cardBg: React.CSSProperties = {
    background: t.dark ? "rgba(10,14,26,1)" : "#ffffff",
  };
  return (
    <div className="kpi-cards-grid" style={{ display:"grid", gridTemplateColumns:"repeat(5,1fr)", gap:10, alignItems:"stretch" }}>

      {/* 1 — Net Income */}
      <div className="kpi-card" style={{ ...KPI_CARD_BASE, ...cardBg, border:`1px solid ${kpi.net<0 ? t.red+"44" : t.border}`, borderLeft: kpi.net<0 ? `3px solid ${t.red}` : kpi.net>0 ? `3px solid ${t.em}` : `1px solid ${t.border}` }}>
        <div>
          <div style={{ display:"flex", alignItems:"center", gap:5, flexWrap:"wrap" }}>
            <span style={{ ...KPI_LABEL, color:t.dim }}>Чистий Дохід</span>
            {syncError && <span style={{ fontSize:9, fontWeight:800, padding:"1px 5px", borderRadius:4, background:"#ff3b3b22", border:"1px solid #ff3b3b88", color:"#ff3b3b" }}>⚠ Sync</span>}
          </div>
          <div style={{ display:"flex", alignItems:"flex-end", gap:5, ...KPI_NUM, color:kpi.net>=0?t.em:t.red }}>
            <AnimNum value={kpi.net} fmt={fmt}/>
            {kpi.net>=0
              ? <TrendingUp size={14} style={{ color:t.amb, marginBottom:2, flexShrink:0 }}/>
              : <TrendingDown size={14} style={{ color:t.red, marginBottom:2, flexShrink:0 }}/>}
          </div>
          <LflBadge current={kpi.net} previous={prevKpi?.net??null} fmt={fmt} t={t}/>
          {hubberLfl && (
            <div style={{ display:"flex", alignItems:"center", gap:5, marginTop:3 }}>
              <span style={{ fontSize:8, fontWeight:700, letterSpacing:"0.06em", color:"#6B7280", textTransform:"uppercase" as const }}>LFL vs {hubberLfl.prevYear}</span>
              <span style={{ display:"inline-flex", alignItems:"center", padding:"2px 7px", borderRadius:4, background:hubberLfl.pct>=0?"#DCFCE7":"rgba(239,68,68,0.08)", fontSize:11, fontWeight:700, color:hubberLfl.pct>=0?"#15803D":"#DC2626", letterSpacing:"-0.01em" }}>
                {hubberLfl.pct>=0?"↑":"↓"} {hubberLfl.pct>=0?"+":""}{hubberLfl.pct.toFixed(1)}%
              </span>
              <span style={{ fontSize:9, color:"#9CA3AF" }}>{hubberLfl.monthName}</span>
            </div>
          )}
        </div>
        <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:4, paddingTop:10, borderTop:`1px solid ${t.border}`, marginTop:10 }}>
          <div style={{ display:"flex", flexDirection:"column", gap:1 }}>
            <span style={{ fontSize:9, color:"#9CA3AF", letterSpacing:"0.05em", textTransform:"uppercase" as const }}>Лог</span>
            <strong style={{ fontSize:11, fontWeight:700, color:t.amb }}>{fmt(kpi.logistics)}</strong>
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:1 }}>
            <span style={{ fontSize:9, color:"#9CA3AF", letterSpacing:"0.05em", textTransform:"uppercase" as const }}>Дост</span>
            <strong style={{ fontSize:11, fontWeight:700, color:t.red }}>{fmt(kpi.del)}</strong>
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:1 }}>
            <span style={{ fontSize:9, color:"#9CA3AF", letterSpacing:"0.05em", textTransform:"uppercase" as const }}>Кому</span>
            <strong style={{ fontSize:11, fontWeight:700, color:t.amb }}>{fmt(kpi.com)}</strong>
          </div>
        </div>
      </div>

      {/* 2 — Return rate */}
      <div className="kpi-card" style={{ ...KPI_CARD_BASE, ...cardBg, border:`1px solid ${t.border}` }}>
        <div>
          <span style={{ ...KPI_LABEL }}>Відмови %</span>
          <div style={{ ...KPI_NUM, color:kpi.returnRate>0?t.red:"#374151" }}><AnimNum value={kpi.returnRate} fmt={v=>v.toFixed(1)+"%"}/></div>
          {kpi.orders>0 && kpi.returnRate>0 && (
            <div style={{ position:"relative", height:4, borderRadius:99, background:t.dark?"rgba(255,255,255,0.08)":"rgba(255,77,77,0.10)", overflow:"hidden", marginTop:8 }}>
              <div style={{ width:`${Math.min(kpi.returnRate,100)}%`, height:"100%", borderRadius:99, background:t.red, transition:"width 0.5s ease" }}/>
            </div>
          )}
        </div>
        <div style={{ paddingTop:10, borderTop:`1px solid ${t.border}`, marginTop:10, display:"flex", alignItems:"center", gap:6 }}>
          <span style={{ fontSize:10, fontWeight:700, color:kpi.refs>0?t.red:"#9CA3AF" }}>{kpi.refs}</span>
          <span style={{ fontSize:10, color:"#9CA3AF" }}>замовлень відмовлено</span>
        </div>
      </div>

      {/* 3 — Orders */}
      <div className="kpi-card" style={{ ...KPI_CARD_BASE, ...cardBg, border:`1px solid ${t.border}` }}>
        <div>
          <span style={{ ...KPI_LABEL }}>Замовлення</span>
          <div style={{ ...KPI_NUM, color:t.blue }}><AnimNum value={kpi.orders} fmt={v=>Math.round(v).toLocaleString()}/></div>
          <LflBadge current={kpi.orders} previous={prevKpi?.orders??null} fmt={v=>Math.round(v).toLocaleString()} t={t}/>
        </div>
        <div style={{ paddingTop:10, borderTop:`1px solid ${t.border}`, marginTop:10 }}>
          <span style={{ fontSize:10, color:"#9CA3AF" }}>Успішних: <strong style={{ color:t.em }}>{kpi.successOrders.toLocaleString()}</strong></span>
        </div>
      </div>

      {/* 4 — Debt / Receivables */}
      {(()=>{
        const hasDebt = kpi.debt > 0;
        return (
          <div className="kpi-card" style={{ ...KPI_CARD_BASE, ...cardBg, border:`1px solid ${hasDebt ? t.red+"44" : t.border}`, borderLeft: hasDebt ? `3px solid ${t.red}` : `1px solid ${t.border}` }}>
            <div>
              <span style={{ ...KPI_LABEL, color: hasDebt ? t.red : undefined }}>Дебіторка{hasDebt ? " ⚠" : ""}</span>
              <div style={{ ...KPI_NUM, color: hasDebt ? t.red : kpi.debt===0 ? "#374151" : t.text, fontWeight:900 }}><AnimNum value={kpi.debt} fmt={fmt}/></div>
            </div>
            <div style={{ paddingTop:10, borderTop:`1px solid ${t.border}`, marginTop:10 }}>
              <span style={{ fontSize:10, color: hasDebt ? t.red : "#9CA3AF" }}>
                {debtCol ? (hasDebt ? "Загальна дебіторська заборгованість" : "Заборгованість відсутня") : "Дані відсутні"}
              </span>
            </div>
          </div>
        );
      })()}

      {/* 5 — Logistics */}
      <div className="kpi-card" style={{ ...KPI_CARD_BASE, ...cardBg, border:`1px solid ${t.border}` }}>
        <div>
          <span style={{ ...KPI_LABEL }}>Логістика</span>
          <div style={{ ...KPI_NUM, color:t.amb }}><AnimNum value={kpi.logistics} fmt={fmt}/></div>
          <LflBadge current={kpi.logistics} previous={prevKpi?.logistics??null} fmt={fmt} t={t}/>
        </div>
        <div style={{ paddingTop:10, borderTop:`1px solid ${t.border}`, marginTop:10 }}>
          <span style={{ fontSize:10, color:"#9CA3AF" }}>Доставка + Комісія</span>
        </div>
      </div>

    </div>
  );
});

/* ─── sidebar filter button ─────────────────────────────────── */
const SidebarFilterBtn = memo(function SidebarFilterBtn({
  label, active, onClick, t, logo, compact,
}: {
  label:string; active:boolean; onClick:()=>void; t:T; logo?:React.ReactNode; compact?:boolean;
}) {
  const [hovered, setHovered] = useState(false);
  return (
    <button
      onClick={onClick}
      onMouseEnter={()=>setHovered(true)}
      onMouseLeave={()=>setHovered(false)}
      style={{
        display:"flex", alignItems:"center", gap:7, width:"100%",
        padding: compact ? "5px 10px" : "7px 10px",
        borderRadius:6, border:"none",
        borderLeft: active ? `2px solid ${t.blue}` : `2px solid transparent`,
        background: active ? `${t.blue}12` : hovered ? "rgba(0,0,0,0.04)" : "transparent",
        color: active ? t.blue : (compact ? t.dim : (t.dark ? "rgba(255,255,255,0.8)" : "#111111")),
        fontSize: compact ? 11 : 12,
        fontWeight: active ? 700 : (compact ? 400 : 500),
        cursor:"pointer", textAlign:"left",
        transition:"background 0.15s ease, color 0.15s ease",
      }}>
      {logo && <span style={{ flexShrink:0, opacity: active ? 1 : 0.55, display:"flex", alignItems:"center" }}>{logo}</span>}
      <span style={{ flex:1, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{label}</span>
      {active && <div style={{ width:5, height:5, borderRadius:"50%", background:t.blue, flexShrink:0 }}/>}
    </button>
  );
});

/* ── Collapsible sidebar section ─────────────────────────────── */
function SidebarSection({ icon, label, open, onToggle, children, st }: {
  icon: React.ReactNode; label: string; open: boolean;
  onToggle: ()=>void; children: React.ReactNode; st: T;
}) {
  return (
    <div style={{ display:"flex", flexDirection:"column" }}>
      <button className="sidebar-section-header" onClick={onToggle}>
        <span style={{ color:st.dim, flexShrink:0, display:"flex", alignItems:"center" }}>{icon}</span>
        <span style={{
          fontSize:10, fontWeight:700, letterSpacing:"0.08em",
          textTransform:"uppercase" as const, color:st.sub, flex:1, textAlign:"left" as const,
        }}>{label}</span>
        <ChevronDown size={12} style={{
          color:st.dim, flexShrink:0,
          transform: open ? "rotate(0deg)" : "rotate(-90deg)",
          transition:"transform 0.22s cubic-bezier(0.22,1,0.36,1)",
        }}/>
      </button>
      <div className="sidebar-section-body" data-open={String(open)}
        style={{ display:"flex", flexDirection:"column", gap:2 }}>
        {children}
      </div>
    </div>
  );
}

/* ─── MiniSparkline — tiny SVG sparkline for brand grid ─────────── */
function MiniSparkline({ data, color }: { data: number[]; color: string }) {
  if (data.length < 2) return null;
  const minV = Math.min(...data);
  const maxV = Math.max(...data);
  const range = maxV - minV || 1;
  const W = 40, H = 14;
  const pts = data.map((v, i) => {
    const x = (i / (data.length - 1)) * W;
    const y = H - ((v - minV) / range) * (H - 2) - 1;
    return `${x.toFixed(1)},${y.toFixed(1)}`;
  }).join(" ");
  const lastV = data[data.length - 1];
  const lastX = W;
  const lastY = H - ((lastV - minV) / range) * (H - 2) - 1;
  return (
    <svg width={W} height={H} style={{ flexShrink:0, overflow:"visible" }}>
      <polyline points={pts} fill="none" stroke={color} strokeWidth={1.5} strokeLinejoin="round" strokeLinecap="round" opacity={0.8}/>
      <circle cx={lastX} cy={lastY} r={2.5} fill={color}/>
    </svg>
  );
}

/* ─── BrandGridCell — 2-col icon grid ──────────────────────────── */
interface BrandTrend { pct:number|null; sparkData:number[]; badge:"rising"|"stable"|"risk"|"none"; curr:number }
function BrandGridCell({ brand, active, onClick, t, logo, isTop, trend }: { brand:string; active:boolean; onClick:()=>void; t:T; logo?: React.ReactNode; isTop?: boolean; trend?: BrandTrend }) {
  const [hov, setHov] = React.useState(false);
  const short = brand.length > 9 ? brand.slice(0,8)+"…" : brand;
  return (
    <button
      onClick={onClick}
      onMouseEnter={()=>setHov(true)}
      onMouseLeave={()=>setHov(false)}
      style={{
        position:"relative", display:"flex", flexDirection:"column", alignItems:"center",
        padding:"7px 4px 5px", borderRadius:6, border:"none",
        background: active ? `${t.blue}15` : hov ? "rgba(0,0,0,0.04)" : "transparent",
        cursor:"pointer", gap:3,
        outline: active ? `2px solid ${t.blue}` : isTop ? `2px solid rgba(184,134,11,0.55)` : "2px solid transparent",
        outlineOffset:-1,
        transition:"all 0.15s ease",
      }}
    >
      {isTop && (
        <span style={{ position:"absolute", top:2, right:3, fontSize:8, lineHeight:1, pointerEvents:"none" }} title="Top Performer">🏆</span>
      )}
      {logo ?? <BrandLogo brand={brand} size={22}/>}
      <span style={{ fontSize:9, fontWeight:active?700:500, color:active?t.blue:t.sub, textAlign:"center", lineHeight:1.2 }}>{short}</span>
      {trend && trend.sparkData.length >= 2 && (
        <MiniSparkline
          data={trend.sparkData}
          color={trend.badge==="rising"?"#16A34A":trend.badge==="risk"?"#FF4D4D":"#9CA3AF"}
        />
      )}
      {trend && trend.badge !== "none" && (
        <span style={{
          fontSize:7.5, fontWeight:700, letterSpacing:"0.03em",
          padding:"1px 5px", borderRadius:3, lineHeight:1.6,
          background:trend.badge==="rising"?"rgba(22,163,74,0.12)":trend.badge==="risk"?"rgba(255,77,77,0.1)":"rgba(156,163,175,0.12)",
          color:trend.badge==="rising"?"#16A34A":trend.badge==="risk"?"#FF4D4D":"#6B7280",
        }}>
          {trend.badge==="rising"?`↑ +${trend.pct!.toFixed(0)}%`:trend.badge==="risk"?`↓ Ризик`:`→ Стаб.`}
        </span>
      )}
    </button>
  );
}

/* ─── HubberSidebarPanel — sidebar card + portal modal ─────────── */
function HubberSidebarPanel({
  data, setData, selYear, setSelYear, t, mktBreakdown
}: {
  data: HubberQuick|null;
  setData: React.Dispatch<React.SetStateAction<HubberQuick|null>>;
  selYear: string; setSelYear: (y:string)=>void;
  t: T;
  mktBreakdown?: { name:string; net:number; pct:number }[];
}) {
  const inputRef = React.useRef<HTMLInputElement>(null);
  const [loading, setLoading] = React.useState(false);
  const [err, setErr] = React.useState<string|null>(null);
  const [modalOpen, setModalOpen] = React.useState(false);
  const [cmpA, setCmpA] = React.useState("");
  const [cmpB, setCmpB] = React.useState("");

  async function handleFile(f: File) {
    setLoading(true); setErr(null);
    try {
      const q = await parseHubberQuick(f);
      setData(q);
      const lastY = q.years[q.years.length-1]||"2025";
      const prevY = q.years[q.years.length-2]||"";
      setSelYear(lastY);
      setCmpA(lastY);
      setCmpB(prevY);
    } catch(e) { setErr(e instanceof Error ? e.message : "Помилка"); }
    finally { setLoading(false); }
  }

  /* ── derived stats ── */
  const cardBg = t.dark ? "rgba(255,255,255,0.04)" : "rgba(0,82,255,0.03)";
  const cardBorder = t.dark ? "rgba(255,255,255,0.1)" : "#DCDCD2";
  const yearTotal = data ? (data.yearTotals[selYear] ?? Object.values(data.values[selYear]??{}).reduce((a,b)=>a+b,0)) : 0;
  const prevTotal = data ? (data.yearTotals[String(+selYear-1)] ?? 0) : 0;
  const delta = prevTotal > 0 ? ((yearTotal - prevTotal) / prevTotal * 100) : null;

  const { grandTotal, bestYear, worstYear } = React.useMemo(()=>{
    if (!data) return { grandTotal:0, bestYear:"", worstYear:"" };
    const activeYears = data.years.filter(y=>(data.yearTotals[y]??0)>0);
    const grand = activeYears.reduce((s,y)=>s+(data.yearTotals[y]??0),0);
    const best = activeYears.reduce((b,y)=>(data.yearTotals[y]??0)>(data.yearTotals[b]??0)?y:b, activeYears[0]??"");
    const worst = activeYears.reduce((w,y)=>(data.yearTotals[y]??0)<(data.yearTotals[w]??0)?y:w, activeYears[0]??"");
    return { grandTotal:grand, bestYear:best, worstYear:worst };
  }, [data]);

  const UA_SHORT = ["Січ","Лют","Бер","Кві","Тра","Чер","Лип","Сер","Вер","Жов","Лис","Гру"];

  /* single-year trend data */
  const trendData = React.useMemo(()=>{
    if (!data) return [];
    return data.months.map((m,i)=>({ m: UA_SHORT[i]??m.slice(0,3), v: data.values[selYear]?.[m]??0 }));
  }, [data, selYear]);

  /* compare data — A vs B */
  const cmpData = React.useMemo(()=>{
    if (!data || !cmpA || !cmpB) return [];
    return data.months.map((m,i)=>({
      m: UA_SHORT[i]??m.slice(0,3),
      a: data.values[cmpA]?.[m]??0,
      b: data.values[cmpB]?.[m]??0,
    }));
  }, [data, cmpA, cmpB]);

  const isComparing = cmpA && cmpB && cmpA !== cmpB;

  /* ── visible years — hide future years with zero data ── */
  const displayYears = React.useMemo(()=>
    data ? data.years.filter(y=>(data.yearTotals[y]??Object.values(data.values[y]??{}).reduce((s:number,v:number)=>s+v,0))>0) : []
  , [data]);

  /* all-years bar data for neon BarChart */
  const yearBarData = React.useMemo(()=>{
    if (!data) return [];
    return displayYears.map(y=>({
      year: y,
      total: data.yearTotals[y] ?? Object.values(data.values[y]??{}).reduce((a: number,b: number)=>a+b,0),
    })).filter(d=>d.total>0);
  }, [data, displayYears]);

  /* ── heatmap max cell ── */
  const maxCellVal = React.useMemo(()=>{
    if (!data) return 1;
    let mx = 0;
    for (const y of displayYears) for (const m of data.months) { const v=data.values[y]?.[m]??0; if(v>mx) mx=v; }
    return mx || 1;
  }, [data, displayYears]);

  /* ── month drill-down popup ── */
  const [drillMonth, setDrillMonth] = React.useState<string|null>(null);

  /* ── Portal modal ── */
  const modal = (data && modalOpen) ? createPortal(
    <div
      onClick={e=>{ if(e.target===e.currentTarget) setModalOpen(false); }}
      style={{ position:"fixed", inset:0, zIndex:99999, background:"rgba(5,5,20,0.6)", backdropFilter:"blur(6px)", WebkitBackdropFilter:"blur(6px)", display:"flex", alignItems:"center", justifyContent:"center", padding:"3vh 3vw" }}
    >
      <div style={{ background:"#F0F0E8", borderRadius:16, width:"80vw", height:"80vh", display:"flex", flexDirection:"column", boxShadow:"0 40px 100px rgba(0,0,0,0.35)", overflow:"hidden" }}>

        {/* ── Header ── */}
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"16px 28px 14px", borderBottom:"1px solid #DCDCD2", flexShrink:0 }}>
          <div>
            <div style={{ fontSize:20, fontWeight:800, color:"#0A0A0A", letterSpacing:"-0.03em" }}>Дохід за роки (2017–2026)</div>
            <div style={{ fontSize:11, color:"#6B7280", marginTop:2 }}>Аркуш «Дохід» · {data.fileName}</div>
          </div>
          <button onClick={()=>setModalOpen(false)} style={{ background:"#0A0A0A", border:"none", borderRadius:10, cursor:"pointer", color:"#fff", width:36, height:36, display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }} title="Закрити">
            <X size={16}/>
          </button>
        </div>

        {/* ── Stats bar ── */}
        <div style={{ display:"flex", gap:0, borderBottom:"1px solid #DCDCD2", flexShrink:0, background:"#FFFFFF" }}>
          {[
            { label:"Загальний дохід", value:fmt(grandTotal), color:"#0052FF" },
            { label:"🏆 Рекорд", value:`${bestYear} · ${fmtK(data.yearTotals[bestYear]??0)}`, color:"#B8860B" },
            { label:"🌱 Старт", value:`${worstYear} · ${fmtK(data.yearTotals[worstYear]??0)}`, color:"#6B7280" },
          ].map((s,i)=>(
            <div key={i} style={{ flex:1, padding:"10px 20px", borderRight:"1px solid #DCDCD2" }}>
              <div style={{ fontSize:9, fontWeight:700, letterSpacing:"0.07em", textTransform:"uppercase" as const, color:"#9CA3AF", marginBottom:3 }}>{s.label}</div>
              <div style={{ fontSize:14, fontWeight:800, color:s.color, letterSpacing:"-0.02em" }}>{s.value}</div>
            </div>
          ))}
          {/* Marketplace contribution column */}
          {mktBreakdown && mktBreakdown.length > 0 && (
            <div style={{ minWidth:200, padding:"10px 20px", borderLeft:"1px solid #DCDCD2" }}>
              <div style={{ fontSize:9, fontWeight:700, letterSpacing:"0.07em", textTransform:"uppercase" as const, color:"#9CA3AF", marginBottom:6 }}>🛒 Маркетплейси (усього)</div>
              <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                {mktBreakdown.slice(0,4).map((m,i)=>(
                  <div key={m.name} style={{ display:"flex", alignItems:"center", gap:6 }}>
                    <div style={{ width:`${Math.max(4,m.pct*0.9)}px`, height:5, borderRadius:2, background:i===0?"#0052FF":i===1?"#3376FF":i===2?"#F59E0B":"#9CA3AF", flexShrink:0 }}/>
                    <span style={{ fontSize:10, fontWeight:i===0?700:500, color:i===0?"#0052FF":"#374151" }}>{m.name}</span>
                    <span style={{ marginLeft:"auto", fontSize:10, fontWeight:700, color:"#374151" }}>{m.pct.toFixed(0)}%</span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>

        {/* ── Year pills row + selected total ── */}
        <div style={{ display:"flex", alignItems:"center", gap:5, padding:"10px 20px 9px", borderBottom:"1px solid #DCDCD2", flexShrink:0, flexWrap:"wrap" }}>
          {displayYears.map(y=>(
            <button key={y} onClick={()=>setSelYear(y)} style={{
              padding:"3px 11px", borderRadius:12, border:"none", cursor:"pointer", fontSize:11, fontWeight:700,
              background: selYear===y ? hubberYearColor(y, displayYears) : "#FFFFFF",
              color: selYear===y ? "#fff" : "#374151",
              boxShadow: selYear===y ? `0 2px 8px ${hubberYearColor(y, displayYears)}50` : "0 1px 3px rgba(0,0,0,0.07)",
              transition:"all 0.14s ease",
            }}>{y}</button>
          ))}
          <div style={{ marginLeft:"auto", display:"flex", alignItems:"center", gap:6 }}>
            <span style={{ fontSize:18, fontWeight:800, color:hubberYearColor(selYear, displayYears), letterSpacing:"-0.03em" }}>{fmtK(yearTotal)}</span>
            {delta !== null && <span style={{ fontSize:11, fontWeight:700, color:delta>=0?"#16A34A":"#FF4D4D" }}>{delta>=0?"▲":"▼"}{Math.abs(delta).toFixed(0)}%</span>}
          </div>
        </div>

        {/* ── Body: chart panel + table ── */}
        <div style={{ display:"flex", flex:1, overflow:"hidden" }}>

          {/* Chart panel */}
          <div style={{ width:310, flexShrink:0, borderRight:"1px solid #DCDCD2", padding:"14px 12px 14px 16px", display:"flex", flexDirection:"column", gap:10 }}>
            {/* Compare dropdowns */}
            <div style={{ display:"flex", flexDirection:"column", gap:6 }}>
              <div style={{ fontSize:10, fontWeight:700, color:"#374151", letterSpacing:"0.05em", textTransform:"uppercase" as const }}>Порівняти роки</div>
              <div style={{ display:"flex", gap:6 }}>
                {[{val:cmpA, set:setCmpA, label:"Рік A"},{val:cmpB, set:setCmpB, label:"Рік B"}].map(({val,set,label})=>(
                  <select key={label} value={val} onChange={e=>set(e.target.value)}
                    style={{ flex:1, padding:"4px 6px", borderRadius:6, border:"1px solid #DCDCD2", fontSize:11, fontWeight:600, color:"#374151", background:"#fff", cursor:"pointer", outline:"none" }}>
                    <option value="">{label}</option>
                    {displayYears.map(y=><option key={y} value={y}>{y}</option>)}
                  </select>
                ))}
              </div>
              {isComparing && (
                <div style={{ display:"flex", gap:12, alignItems:"center" }}>
                  <span style={{ fontSize:10, color:hubberYearColor(cmpA, displayYears), fontWeight:700 }}>━ {cmpA}</span>
                  <span style={{ fontSize:10, color:hubberYearColor(cmpB, displayYears), fontWeight:700 }}>━ {cmpB}</span>
                  <span style={{ fontSize:10, color:"#9CA3AF" }}>
                    Δ {(((data.yearTotals[cmpB]??0)-(data.yearTotals[cmpA]??0))/(data.yearTotals[cmpA]??1)*100).toFixed(0)}%
                  </span>
                </div>
              )}
            </div>

            {/* Chart */}
            <div style={{ flex:1, minHeight:0 }}>
              <ResponsiveContainer width="100%" height="100%">
                {isComparing ? (
                  <LineChart data={cmpData} margin={{ top:4, right:8, left:-12, bottom:0 }}>
                    <CartesianGrid strokeDasharray="1 0" stroke="rgba(0,0,0,0.05)" vertical={false}/>
                    <XAxis dataKey="m" tick={{ fontSize:9, fill:"#6B7280" }} axisLine={false} tickLine={false}/>
                    <YAxis tick={{ fontSize:9, fill:"#9CA3AF" }} axisLine={false} tickLine={false} tickFormatter={(v:number)=>v>=1000?`${(v/1000).toFixed(0)}k`:String(v)}/>
                    <Tooltip contentStyle={{ background:"#fff", border:"1px solid #DCDCD2", borderRadius:8, fontSize:11 }} formatter={(v:number,name:string)=>[fmtK(v), name===`a`?cmpA:cmpB]}/>
                    <Legend formatter={(value:string)=>value===`a`?cmpA:cmpB} wrapperStyle={{ fontSize:11 }}/>
                    <Line type="monotone" dataKey="a" name="a" stroke={hubberYearColor(cmpA, displayYears)} strokeWidth={2.5} dot={{ r:3, fill:hubberYearColor(cmpA, displayYears), strokeWidth:0 }} activeDot={{ r:5 }} isAnimationActive={true} animationDuration={500}/>
                    <Line type="monotone" dataKey="b" name="b" stroke={hubberYearColor(cmpB, displayYears)} strokeWidth={2.5} dot={{ r:3, fill:hubberYearColor(cmpB, displayYears), strokeWidth:0 }} activeDot={{ r:5 }} isAnimationActive={true} animationDuration={500}/>
                  </LineChart>
                ) : (
                  <BarChart data={yearBarData} margin={{ top:4, right:8, left:-12, bottom:0 }}>
                    <CartesianGrid strokeDasharray="1 0" stroke="rgba(0,0,0,0.05)" vertical={false}/>
                    <XAxis dataKey="year" tick={{ fontSize:9, fill:"#6B7280" }} axisLine={false} tickLine={false}/>
                    <YAxis tick={{ fontSize:9, fill:"#9CA3AF" }} axisLine={false} tickLine={false} tickFormatter={(v:number)=>v>=1000?`${(v/1000).toFixed(0)}k`:String(v)}/>
                    <Tooltip contentStyle={{ background:"#fff", border:"1px solid #DCDCD2", borderRadius:8, fontSize:11 }} formatter={(v:number)=>[fmtK(v),"Дохід"]} labelFormatter={(l:string)=>`${l} рік`}/>
                    <Bar dataKey="total" radius={[4,4,0,0]} isAnimationActive={true} animationDuration={600}>
                      {yearBarData.map((entry, idx) => (
                        <Cell key={entry.year} fill={hubberYearColor(entry.year, displayYears)} />
                      ))}
                    </Bar>
                  </BarChart>
                )}
              </ResponsiveContainer>
              {/* Neon color legend */}
              {!isComparing && (
                <div style={{ display:"flex", flexWrap:"wrap", gap:"4px 10px", marginTop:6, justifyContent:"center" }}>
                  {yearBarData.map(d=>(
                    <div key={d.year} style={{ display:"flex", alignItems:"center", gap:3 }}>
                      <div style={{ width:8, height:8, borderRadius:2, background:hubberYearColor(d.year, displayYears), flexShrink:0 }}/>
                      <span style={{ fontSize:9, fontWeight:600, color:"#6B7280" }}>{d.year}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>

          {/* Scrollable table */}
          <div style={{ flex:1, overflowY:"auto", overflowX:"auto" }}>
            <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11.5 }}>
              <thead style={{ position:"sticky", top:0, zIndex:2 }}>
                <tr style={{ background:"#F0F0E8" }}>
                  <th style={{ textAlign:"left", padding:"9px 16px", borderBottom:"2px solid #DCDCD2", color:"#374151", fontWeight:700, minWidth:80, whiteSpace:"nowrap" as const }}>Місяць</th>
                  {displayYears.map(y=>{
                    const isBest = y===bestYear;
                    const isSel = y===selYear;
                    return (
                      <th key={y} onClick={()=>setSelYear(y)} style={{
                        padding:"9px 11px", borderBottom:"2px solid #DCDCD2",
                        color: isSel?hubberYearColor(y, displayYears):isBest?"#B8860B":"#374151",
                        fontWeight: isSel||isBest?800:600,
                        textAlign:"right" as const, minWidth:72, cursor:"pointer",
                        background: isBest?"rgba(184,134,11,0.07)":"transparent",
                        whiteSpace:"nowrap" as const,
                      }}>
                        {isBest?"🏆 ":""}{y}
                      </th>
                    );
                  })}
                  <th style={{ padding:"9px 11px", borderBottom:"2px solid #DCDCD2", color:"#0052FF", fontWeight:700, textAlign:"right" as const, minWidth:72, whiteSpace:"nowrap" as const }}>
                    LFL %<br/><span style={{ fontSize:9, fontWeight:400, color:"#6B7280" }}>{selYear} / {String(+selYear-1)}</span>
                  </th>
                </tr>
              </thead>
              <tbody>
                {data.months.map((m,i)=>(
                  <tr key={m} style={{ background: i%2===0?"#FFFFFF":"#F9F9F4" }}>
                    <td
                      onClick={()=>setDrillMonth(m)}
                      style={{ padding:"7px 16px", color:"#374151", fontWeight:500, borderBottom:"1px solid #F0F0E8", whiteSpace:"nowrap" as const, cursor:"pointer", userSelect:"none" as const }}
                      title="Клікніть для деталей"
                    >
                      <span style={{ borderBottom:"1px dashed #9CA3AF" }}>{m}</span>
                    </td>
                    {displayYears.map(y=>{
                      const val = data.values[y]?.[m]??0;
                      const isSel = y===selYear;
                      const isBest = y===bestYear;
                      const isCmpA = y===cmpA && isComparing;
                      const isCmpB = y===cmpB && isComparing;
                      const heat = val > 0 ? Math.min(1, val/maxCellVal) : 0;
                      const heatBg = val > 0 ? `rgba(138,154,91,${(heat*0.22).toFixed(3)})` : "transparent";
                      const cellBg = isBest
                        ? (i%2===0?"rgba(184,134,11,0.10)":"rgba(184,134,11,0.15)")
                        : isSel
                          ? (i%2===0?"rgba(0,82,255,0.05)":"rgba(0,82,255,0.09)")
                          : heatBg;
                      const cellColor = val===0?"#C0C0B8":isSel?hubberYearColor(y, displayYears):isCmpA?hubberYearColor(cmpA, displayYears):isCmpB?hubberYearColor(cmpB, displayYears):isBest?"#8B6914":"#0A0A0A";
                      return (
                        <td key={y} style={{ padding:"7px 11px", textAlign:"right" as const, borderBottom:"1px solid #F0F0E8", fontWeight:val>0?(isSel||isBest?700:500):400, color:cellColor, background:cellBg }}>
                          {val>0 ? fmtK(val) : "—"}
                        </td>
                      );
                    })}
                    {(()=>{
                      const prevY = String(+selYear-1);
                      const curr = data.values[selYear]?.[m]??0;
                      const prev = data.values[prevY]?.[m]??0;
                      if (prev===0) return <td key="lfl" style={{ padding:"7px 11px", textAlign:"right" as const, borderBottom:"1px solid #F0F0E8", color:"#C0C0B8" }}>—</td>;
                      const pct = ((curr-prev)/prev)*100;
                      return <td key="lfl" style={{ padding:"7px 11px", textAlign:"right" as const, borderBottom:"1px solid #F0F0E8", fontWeight:700, color:pct>=0?"#16A34A":"#FF4D4D", background:i%2===0?"rgba(0,82,255,0.02)":"rgba(0,82,255,0.04)" }}>
                        {pct>=0?"+":""}{pct.toFixed(0)}%
                      </td>;
                    })()}
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr style={{ background:"#F0F0E8", borderTop:"2px solid #DCDCD2" }}>
                  <td style={{ padding:"9px 16px", fontWeight:800, color:"#0052FF", fontSize:12, whiteSpace:"nowrap" as const }}>Всього</td>
                  {displayYears.map(y=>{
                    const total = data.yearTotals[y]??Object.values(data.values[y]??{}).reduce((a,b)=>a+b,0);
                    const isBest = y===bestYear;
                    const isSel = y===selYear;
                    return (
                      <td key={y} style={{ padding:"9px 11px", textAlign:"right" as const, fontWeight:800, fontSize:12, color:isBest?"#B8860B":isSel?"#0052FF":"#374151", background:isBest?"rgba(184,134,11,0.1)":isSel?"rgba(0,82,255,0.08)":"transparent" }}>
                        {total>0?fmtK(total):"—"}
                      </td>
                    );
                  })}
                  {(()=>{
                    const prevY = String(+selYear-1);
                    const curr = data.yearTotals[selYear]??Object.values(data.values[selYear]??{}).reduce((a,b)=>a+b,0);
                    const prev = data.yearTotals[prevY]??Object.values(data.values[prevY]??{}).reduce((a,b)=>a+b,0);
                    if (prev===0) return <td key="lfl" style={{ padding:"9px 11px", textAlign:"right" as const, fontWeight:800, fontSize:12, color:"#9CA3AF" }}>—</td>;
                    const pct = ((curr-prev)/prev)*100;
                    return <td key="lfl" style={{ padding:"9px 11px", textAlign:"right" as const, fontWeight:800, fontSize:12, color:pct>=0?"#16A34A":"#FF4D4D" }}>
                      {pct>=0?"+":""}{pct.toFixed(0)}%
                    </td>;
                  })()}
                </tr>
              </tfoot>
            </table>
          </div>

        </div>
      </div>
    </div>,
    document.body
  ) : null;

  /* ── Month drill-down popup ── */
  const drillPopup = (data && drillMonth) ? createPortal(
    <div
      onClick={e=>{ if(e.target===e.currentTarget) setDrillMonth(null); }}
      style={{ position:"fixed", inset:0, zIndex:999999, background:"rgba(5,5,20,0.5)", backdropFilter:"blur(4px)", WebkitBackdropFilter:"blur(4px)", display:"flex", alignItems:"center", justifyContent:"center", padding:"4vh 4vw" }}
    >
      <div style={{ background:"#F0F0E8", borderRadius:14, width:"min(560px,90vw)", maxHeight:"70vh", display:"flex", flexDirection:"column", boxShadow:"0 32px 80px rgba(0,0,0,0.32)", overflow:"hidden" }}>
        {/* header */}
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding:"14px 22px 12px", borderBottom:"1px solid #DCDCD2", flexShrink:0 }}>
          <div>
            <div style={{ fontSize:17, fontWeight:800, color:"#0A0A0A", letterSpacing:"-0.03em" }}>📅 {drillMonth} — усі роки</div>
            <div style={{ fontSize:10, color:"#6B7280", marginTop:2 }}>Порівняння цього місяця за роками</div>
          </div>
          <button onClick={()=>setDrillMonth(null)} style={{ background:"#0A0A0A", border:"none", borderRadius:8, cursor:"pointer", color:"#fff", width:30, height:30, display:"flex", alignItems:"center", justifyContent:"center" }}>
            <X size={13}/>
          </button>
        </div>
        {/* table */}
        <div style={{ overflowY:"auto", flex:1, padding:"14px 22px 18px" }}>
          {(() => {
            const rows = displayYears.map(y=>({ y, v:data.values[y]?.[drillMonth]??0 })).filter(r=>r.v>0);
            const maxV = rows.reduce((m,r)=>Math.max(m,r.v),0)||1;
            if (rows.length===0) return <div style={{ color:"#9CA3AF", fontSize:13, textAlign:"center", padding:"24px 0" }}>Немає даних за цей місяць</div>;
            return (
              <div style={{ display:"flex", flexDirection:"column", gap:7 }}>
                {rows.map(({y,v})=>{
                  const pct = v/maxV*100;
                  const isB = y===bestYear;
                  return (
                    <div key={y} style={{ display:"flex", alignItems:"center", gap:10 }}>
                      <span style={{ fontSize:12, fontWeight:isB?800:600, color:hubberYearColor(y, displayYears), width:38, flexShrink:0, textAlign:"right" }}>{y}</span>
                      <div style={{ flex:1, background:"rgba(0,0,0,0.06)", borderRadius:3, height:14, overflow:"hidden" }}>
                        <div style={{ width:`${pct}%`, height:"100%", background:hubberYearColor(y, displayYears), borderRadius:3, transition:"width 0.4s ease" }}/>
                      </div>
                      <span style={{ fontSize:12, fontWeight:700, color:hubberYearColor(y, displayYears), width:70, textAlign:"right", flexShrink:0 }}>{fmtK(v)}</span>
                      {y===String(+y) && (() => {
                        const prevV = data.values[String(+y-1)]?.[drillMonth]??0;
                        if (prevV===0) return null;
                        const lfl = ((v-prevV)/prevV*100);
                        return <span style={{ fontSize:10, fontWeight:700, color:lfl>=0?"#16A34A":"#FF4D4D", width:48, textAlign:"right", flexShrink:0 }}>{lfl>=0?"+":""}{lfl.toFixed(0)}%</span>;
                      })()}
                    </div>
                  );
                })}
              </div>
            );
          })()}
        </div>
      </div>
    </div>,
    document.body
  ) : null;

  return (
    <>
      {/* ── Sidebar card ── */}
      <div style={{ marginTop:10, padding:"10px 10px 11px", borderRadius:8, background:cardBg, border:`1px ${data?"solid":"dashed"} ${cardBorder}` }}>
        <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:data?6:8 }}>
          <div style={{ fontSize:10, fontWeight:700, letterSpacing:"0.06em", textTransform:"uppercase" as const, color:t.sub }}>
            📊 Дохід за роки
          </div>
          {data && (
            <button onClick={()=>{setData(null);setModalOpen(false);}} style={{ background:"none", border:"none", cursor:"pointer", color:t.dim, padding:"1px 2px", lineHeight:1 }} title="Очистити">
              <X size={10}/>
            </button>
          )}
        </div>

        {/* shared file input — used in both states */}
        <input ref={inputRef} type="file" accept=".xls,.xlsx,.ods" style={{display:"none"}}
          onChange={e=>{const f=e.target.files?.[0];if(f)handleFile(f);e.target.value="";}}/>

        {!data ? (
          <>
            {err && <div style={{ fontSize:9, color:t.dark?"#E29578":"#FF4D4D", marginBottom:6 }}>{err}</div>}
            <button
              onClick={()=>inputRef.current?.click()}
              disabled={loading}
              style={{
                width:"100%", padding:"5px 8px", borderRadius:6,
                background:"transparent", border:`1px solid ${t.dark?"rgba(255,255,255,0.15)":t.border}`,
                color:t.sub, fontSize:10, fontWeight:600, cursor:loading?"wait":"pointer",
                display:"flex", alignItems:"center", justifyContent:"center", gap:5,
                transition:"background 0.15s ease",
              }}
            >
              <Upload size={10}/> {loading ? "Завантаження…" : "Завантажити дохід.xls"}
            </button>
          </>
        ) : (
          <>
            <div style={{ marginBottom:8 }}>
              <div style={{ fontSize:16, fontWeight:800, color:t.blue, letterSpacing:"-0.03em", lineHeight:1 }}>{fmtK(yearTotal)}</div>
              <div style={{ display:"flex", alignItems:"center", gap:4, marginTop:2 }}>
                <span style={{ fontSize:9, color:t.dim }}>{selYear} · Хаббер</span>
                {delta !== null && (
                  <span style={{ fontSize:9, fontWeight:700, color:delta>=0?"#16A34A":"#FF4D4D" }}>
                    {delta>=0?"▲":"▼"}{Math.abs(delta).toFixed(0)}%
                  </span>
                )}
              </div>
            </div>
            <button
              onClick={()=>setModalOpen(true)}
              onMouseEnter={e=>(e.currentTarget.style.opacity="0.85")}
              onMouseLeave={e=>(e.currentTarget.style.opacity="1")}
              style={{
                width:"100%", padding:"6px 10px", borderRadius:6,
                background:t.blue, border:"none",
                color:"#fff", fontSize:10, fontWeight:700, cursor:"pointer",
                display:"flex", alignItems:"center", justifyContent:"center", gap:5,
                transition:"opacity 0.15s ease",
              }}
            >
              Відкрити повний звіт →
            </button>
            {/* Replace file / clear cache */}
            <button
              onClick={()=>inputRef.current?.click()}
              disabled={loading}
              title="Замінити збережені дані новим файлом"
              style={{
                width:"100%", marginTop:5, padding:"4px 8px", borderRadius:5,
                background:"transparent", border:`1px solid ${t.dark?"rgba(255,255,255,0.1)":t.border}`,
                color:t.dim, fontSize:9, fontWeight:600, cursor:loading?"wait":"pointer",
                display:"flex", alignItems:"center", justifyContent:"center", gap:4,
                opacity:0.75, transition:"opacity 0.15s ease",
              }}
              onMouseEnter={e=>(e.currentTarget.style.opacity="1")}
              onMouseLeave={e=>(e.currentTarget.style.opacity="0.75")}
            >
              <Upload size={9}/> {loading ? "Завантаження…" : "Новий файл / оновити"}
            </button>
          </>
        )}
      </div>

      {/* Portal modal — rendered at document.body, outside stacking context */}
      {modal}
      {drillPopup}
    </>
  );
}

/* Reusable shimmer skeleton bar */
function SkeletonBar({ w = "100%", h = 14, r = 6 }: { w?: string|number; h?: number; r?: number }) {
  return <div className="orbit-skel" style={{ width: w, height: h, borderRadius: r, flexShrink: 0 }}/>;
}

function MktLogo({ brand, size = 16 }: { brand: string; size?: number }) {
  const b = brand.toLowerCase();
  const s = size;
  if (b === "каста" || b === "kasta") return (
    <svg width={s} height={s} viewBox="0 0 20 20" style={{flexShrink:0}}>
      <circle cx="10" cy="10" r="10" fill="#00B0A8"/>
      <text x="10" y="14.5" textAnchor="middle" fill="white" fontSize="11" fontWeight="bold" fontFamily="sans-serif">К</text>
    </svg>
  );
  if (b === "розетка" || b === "rozetka") return (
    <svg width={s} height={s} viewBox="0 0 20 20" style={{flexShrink:0}}>
      <circle cx="10" cy="10" r="10" fill="#00A046"/>
      <circle cx="7.2" cy="8.2" r="1.3" fill="white"/>
      <circle cx="12.8" cy="8.2" r="1.3" fill="white"/>
      <path d="M6.2 12 Q10 16 13.8 12" stroke="white" strokeWidth="1.5" fill="none" strokeLinecap="round"/>
    </svg>
  );
  if (b === "хаббер" || b === "hubber") return (
    <svg width={s} height={s} viewBox="0 0 20 20" style={{flexShrink:0}}>
      <rect width="20" height="20" rx="5" fill="#1B4F8B"/>
      <text x="10" y="14.5" textAnchor="middle" fill="white" fontSize="11" fontWeight="bold" fontFamily="sans-serif">H</text>
    </svg>
  );
  if (b === "шафа" || b === "shafa") return (
    <svg width={s} height={s} viewBox="0 0 20 20" style={{flexShrink:0}}>
      <circle cx="10" cy="10" r="10" fill="#8BA89E"/>
      <path d="M10 15.5 C10 15.5 4.5 11.8 4.5 8.2 A3.6 3.6 0 0 1 10 6.8 A3.6 3.6 0 0 1 15.5 8.2 C15.5 11.8 10 15.5 10 15.5Z" fill="white"/>
    </svg>
  );
  return (
    <div style={{ width:s, height:s, borderRadius:"50%", background:"#0052FF", display:"flex", alignItems:"center", justifyContent:"center", fontSize:Math.round(s*0.55), fontWeight:700, color:"#ffffff", flexShrink:0 }}>
      {brand.charAt(0)}
    </div>
  );
}

function BrandLogo({ brand, size = 18 }: { brand: string; size?: number }) {
  const b = brand.toLowerCase().replace(/[\s\-_]/g, "");
  const s = size;

  /* ── Білий Халат — white coat on soft blue ── */
  if (b.includes("білий") || b.includes("halat") || b.includes("халат") || b.includes("bilyi")) return (
    <svg width={s} height={s} viewBox="0 0 24 24" style={{flexShrink:0}}>
      <circle cx="12" cy="12" r="12" fill="#3A7BD5"/>
      {/* coat collar + body */}
      <path d="M9 5 L9 8 L12 10 L15 8 L15 5 L13.5 5 L12 7 L10.5 5 Z" fill="white"/>
      <rect x="8" y="9.5" width="8" height="9" rx="1.5" fill="white"/>
      {/* pocket */}
      <rect x="13" y="13" width="2.5" height="2.5" rx="0.5" fill="#3A7BD5"/>
      {/* buttons */}
      <circle cx="12" cy="11.5" r="0.6" fill="#3A7BD5"/>
      <circle cx="12" cy="13.5" r="0.6" fill="#3A7BD5"/>
    </svg>
  );

  /* ── Kalyna / Калина — red berry cluster ── */
  if (b.includes("калин") || b.includes("kalyn") || b.includes("kalina")) return (
    <svg width={s} height={s} viewBox="0 0 24 24" style={{flexShrink:0}}>
      <circle cx="12" cy="12" r="12" fill="#1A1A2E"/>
      {/* stem */}
      <line x1="12" y1="18" x2="12" y2="11" stroke="#4A7C35" strokeWidth="1.5" strokeLinecap="round"/>
      <line x1="12" y1="14" x2="9" y2="11.5" stroke="#4A7C35" strokeWidth="1.2" strokeLinecap="round"/>
      <line x1="12" y1="13" x2="15" y2="10.5" stroke="#4A7C35" strokeWidth="1.2" strokeLinecap="round"/>
      {/* berries */}
      <circle cx="12" cy="8.5" r="2.2" fill="#E8192C"/>
      <circle cx="7.8" cy="10.2" r="1.8" fill="#E8192C"/>
      <circle cx="16.2" cy="9.2" r="1.8" fill="#E8192C"/>
      <circle cx="9.5" cy="7.5" r="1.6" fill="#C0111F"/>
      <circle cx="14.5" cy="7.8" r="1.6" fill="#C0111F"/>
      {/* berry highlights */}
      <circle cx="11.4" cy="7.8" r="0.6" fill="rgba(255,255,255,0.45)"/>
      <circle cx="7.3" cy="9.6" r="0.5" fill="rgba(255,255,255,0.4)"/>
      <circle cx="15.7" cy="8.6" r="0.5" fill="rgba(255,255,255,0.4)"/>
    </svg>
  );

  /* ── Elita / Еліта — red flower ── */
  if (b.includes("еліт") || b.includes("elit") || b.includes("elita")) return (
    <svg width={s} height={s} viewBox="0 0 24 24" style={{flexShrink:0}}>
      <circle cx="12" cy="12" r="12" fill="#FFF0F0"/>
      {/* petals */}
      {[0,60,120,180,240,300].map(deg=>{
        const rad = (deg * Math.PI) / 180;
        const cx = 12 + 4.2 * Math.sin(rad);
        const cy = 12 - 4.2 * Math.cos(rad);
        return <ellipse key={deg} cx={cx} cy={cy} rx="2.4" ry="3.2"
          transform={`rotate(${deg},${cx},${cy})`} fill="#E8192C" opacity="0.92"/>;
      })}
      {/* centre */}
      <circle cx="12" cy="12" r="3" fill="#FFD700"/>
      <circle cx="12" cy="12" r="1.5" fill="#E8A800"/>
    </svg>
  );

  /* ── Afina / Афіна — green daisy ── */
  if (b.includes("афін") || b.includes("afin") || b.includes("afina")) return (
    <svg width={s} height={s} viewBox="0 0 24 24" style={{flexShrink:0}}>
      <circle cx="12" cy="12" r="12" fill="#E8F5E9"/>
      {/* petals */}
      {[0,40,80,120,160,200,240,280,320].map(deg=>{
        const rad = (deg * Math.PI) / 180;
        const cx = 12 + 4 * Math.sin(rad);
        const cy = 12 - 4 * Math.cos(rad);
        return <ellipse key={deg} cx={cx} cy={cy} rx="1.8" ry="3.0"
          transform={`rotate(${deg},${cx},${cy})`} fill="#2E7D32" opacity="0.85"/>;
      })}
      {/* centre */}
      <circle cx="12" cy="12" r="3" fill="#FFD700"/>
      <circle cx="12" cy="12" r="1.6" fill="#F9A825"/>
    </svg>
  );

  /* ── Artmon / EL-ARTMON — black/dark rounded rect with "A" ── */
  if (b.includes("artmon") || b.includes("артмон")) return (
    <svg width={s} height={s} viewBox="0 0 24 24" style={{flexShrink:0}}>
      <rect width="24" height="24" rx="5" fill="#111111"/>
      {/* stylised A */}
      <path d="M12 5.5 L16.5 18 H14.5 L13.4 14.5 H10.6 L9.5 18 H7.5 Z M12 9 L11.2 13 H12.8 Z" fill="white"/>
    </svg>
  );

  /* ── Erka / Ерка — red rounded rect with ERKA ── */
  if (b.includes("ерка") || b.includes("erka")) return (
    <svg width={s} height={s} viewBox="0 0 24 24" style={{flexShrink:0}}>
      <rect width="24" height="24" rx="5" fill="#D32F2F"/>
      <text x="12" y="16" textAnchor="middle" fill="white" fontSize="9" fontWeight="800"
        fontFamily="Arial,sans-serif" letterSpacing="0.5">ERKA</text>
    </svg>
  );

  /* ── Iglen / IGLEN — swan + UA flag colours ── */
  if (b.includes("iglen") || b.includes("іглен") || b.includes("iгlen")) return (
    <svg width={s} height={s} viewBox="0 0 24 24" style={{flexShrink:0}}>
      {/* upper half blue (sky), lower half yellow (wheat) */}
      <clipPath id="iglen-clip"><circle cx="12" cy="12" r="12"/></clipPath>
      <g clipPath="url(#iglen-clip)">
        <rect x="0" y="0" width="24" height="12" fill="#005BBB"/>
        <rect x="0" y="12" width="24" height="12" fill="#FFD500"/>
      </g>
      {/* swan body */}
      <ellipse cx="13" cy="14.5" rx="5.5" ry="3" fill="white" opacity="0.95"/>
      {/* swan neck */}
      <path d="M9 14 C8 12 9 9 12 8.5 C13 8.3 13.5 9 12.5 9.5 C11 10 10 12 10.5 14Z" fill="white" opacity="0.95"/>
      {/* swan head */}
      <ellipse cx="12.5" cy="8" rx="1.4" ry="1.1" fill="white" opacity="0.95"/>
      {/* beak */}
      <path d="M13.8 8 L15.4 8.2 L13.6 8.5Z" fill="#FFA000"/>
      {/* eye */}
      <circle cx="12.3" cy="7.7" r="0.4" fill="#1A1A2E"/>
    </svg>
  );

  /* ── Brozell / Brозель ── */
  if (b.includes("brozell") || b.includes("брозел") || b.includes("brozel")) return (
    <svg width={s} height={s} viewBox="0 0 24 24" style={{flexShrink:0}}>
      <circle cx="12" cy="12" r="12" fill="#6A1B9A"/>
      <text x="12" y="16" textAnchor="middle" fill="white" fontSize="8.5" fontWeight="800"
        fontFamily="Arial,sans-serif" letterSpacing="0.3">BRZ</text>
    </svg>
  );

  /* ── generic fallback: first letter initial ── */
  const hue = Math.abs(brand.split("").reduce((a,c)=>a+c.charCodeAt(0),0)) % 360;
  return (
    <svg width={s} height={s} viewBox="0 0 24 24" style={{flexShrink:0}}>
      <circle cx="12" cy="12" r="12" fill={`hsl(${hue},55%,42%)`}/>
      <text x="12" y="16.5" textAnchor="middle" fill="white" fontSize="11" fontWeight="700"
        fontFamily="sans-serif">{brand.charAt(0).toUpperCase()}</text>
    </svg>
  );
}

function Chip({ label, active, onClick, t, variant = "default", logo }: {
  label: string;
  active: boolean;
  onClick: () => void;
  t: T;
  variant?: "default" | "meta";
  logo?: React.ReactNode;
}) {
  const [hovered, setHovered] = useState(false);

  const metaInactiveColor  = t.dark ? "rgba(100,160,255,0.10)" : "rgba(30,100,220,0.07)";
  const metaBorderInactive = t.dark ? "rgba(100,160,255,0.22)" : "rgba(30,100,220,0.20)";
  const metaTextInactive   = t.dark ? "#7ab4ff" : "#3a7bd5";

  let bg: string, border: string, color: string, shadow: string;
  if (active) {
    bg     = t.blue;
    border = t.blue;
    color  = "#ffffff";
    shadow = `0 2px 10px ${t.blue}55`;
  } else if (variant === "meta") {
    bg     = hovered ? metaInactiveColor + "cc" : metaInactiveColor;
    border = metaBorderInactive;
    color  = metaTextInactive;
    shadow = "none";
  } else {
    bg     = hovered
      ? (t.dark ? "rgba(255,255,255,0.07)" : "rgba(0,0,0,0.05)")
      : "transparent";
    border = t.dark ? "rgba(255,255,255,0.16)" : "rgba(0,0,0,0.14)";
    color  = t.sub;
    shadow = "none";
  }

  return (
    <button
      onClick={onClick}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{
        padding: logo ? "5px 13px 5px 9px" : "5px 14px",
        borderRadius: 9,
        fontSize: 12,
        fontWeight: active ? 700 : 500,
        cursor: "pointer",
        transition: "all 0.15s ease",
        background: bg,
        border: `1px solid ${border}`,
        color,
        boxShadow: shadow,
        transform: hovered && !active ? "translateY(-1px)" : "none",
        whiteSpace: "nowrap",
        flexShrink: 0,
        letterSpacing: active ? "0.01em" : "normal",
        display: "flex",
        alignItems: "center",
        gap: logo ? 6 : 0,
      }}
    >
      {logo}
      {label}
    </button>
  );
}

function TipBox({ active, payload, label, t }: { active?:boolean; payload?:{value:number;name:string;color:string}[]; label?:string; t:T }) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background:t.dark?"rgba(4,6,14,0.97)":"#fff", border:`1px solid ${t.border}`, borderRadius:10, padding:"10px 16px" }}>
      <p style={{ color:t.dim, fontSize:11, marginBottom:6 }}>{label}</p>
      {payload.map((p,i)=><p key={i} style={{ color:p.color, fontSize:13, fontWeight:600, margin:"2px 0" }}>{p.name}: {fmt(p.value)}</p>)}
    </div>
  );
}

/* ─── main component ─────────────────────────────────────────── */
/* ─── Rejection-reason tooltip advice lookup ──────────────────── */
const REASON_ADVICE: { match: RegExp; emoji: string; cat: string; text: string }[] = [
  { match: /причина не вказана|не вказан|не зазначен/i,      emoji:"⚠️", cat:"КРИТИЧНО",   text:"Впровадьте обов'язкове поле «Причина» в CRM. Це дозволить контролювати логістичні витрати." },
  { match: /не той розмір|не той колір|відправили не той/i,   emoji:"📦", cat:"СКЛАД",      text:"Помилка комплектації. Потрібне сканування штрих-кодів перед відправкою." },
  { match: /не підійшов розмір|не підійшов/i,                 emoji:"📏", cat:"КОНТЕНТ",    text:"Додайте в картку товару «живі» заміри виробу та параметри моделі на фото." },
  { match: /причина невідома|невідома причина/i,              emoji:"📞", cat:"СЕРВІС",     text:"Менеджер не з'ясував деталі. Впровадьте скрипт обов'язкового опитування при відмові." },
  { match: /форма чашечки|об.єм чашечки|чашечк/i,            emoji:"👙", cat:"АСОРТИМЕНТ", text:"Специфічний фасон. Додайте відео-огляд товару для кращого розуміння форми." },
];
function getReasonAdvice(reason: string) {
  return REASON_ADVICE.find(a => a.match.test(reason)) ?? { emoji:"💬", cat:"АНАЛІЗ", text:"Відстежуйте цю причину окремо та збирайте зворотний зв'язок від клієнтів." };
}

/* ─── Persistence helpers ──────────────────────────────────────── */
function readFiltersCache(): { brandFilter:string; monthFilter:string; companyFilter:string; yearFilter:string; hubberQYear:string; darkMode:boolean } {
  try {
    const s = localStorage.getItem(FILTERS_KEY);
    if (s) return { brandFilter:"All", monthFilter:"All", companyFilter:"All", yearFilter:"All", hubberQYear:"2025", darkMode:true, ...JSON.parse(s) };
  } catch {}
  return { brandFilter:"All", monthFilter:"All", companyFilter:"All", yearFilter:"All", hubberQYear:"2025", darkMode:true };
}
function readHubberCache(): HubberQuick|null {
  try {
    const s = localStorage.getItem(HUBBER_KEY);
    if (s) return JSON.parse(s) as HubberQuick;
  } catch {}
  return null;
}

export default function Dashboard() {
  const [fileData,    setFileData]    = useState<FileData|null>(null);
  const [uploadedAt,  setUploadedAt]  = useState<Date|null>(null);
  const [isDragging,  setIsDragging]  = useState(false);
  const [parseError,  setParseError]  = useState<string|null>(null);
  const _fc = React.useMemo(readFiltersCache, []);
  const [brandFilter,   setBrandFilter]   = useState(()=>_fc.brandFilter);
  const [monthFilter,   setMonthFilter]   = useState(()=>_fc.monthFilter);
  const [companyFilter, setCompanyFilter] = useState(()=>_fc.companyFilter);
  const [search,        setSearch]        = useState("");
  const [productSearch, setProductSearch] = useState("");
  const [productSort,   setProductSort]   = useState<{col:"qty"|"net"|"name"; dir:"asc"|"desc"}>({col:"qty",dir:"desc"});
  const [yearFilter,    setYearFilter]    = useState<string>(()=>_fc.yearFilter);
  const [mktOpen,     setMktOpen]     = useState(true);
  const [dateOpen,    setDateOpen]    = useState(true);
  const [brandsOpen,  setBrandsOpen]  = useState(false);
  const [hubberQuick, setHubberQuick] = useState<HubberQuick|null>(readHubberCache);
  const [hubberQYear, setHubberQYear] = useState(()=>_fc.hubberQYear);
  const [darkMode,    setDarkMode]    = useState(()=>_fc.darkMode);
  const [rejTooltip,  setRejTooltip]  = useState<{ reason:string; rect:DOMRect } | null>(null);
  const [cityFilter,  setCityFilter]  = useState<string|null>(null);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);
  const t = darkMode ? DK : LT;

  /* ── Memoize the glass base style — only recomputes when theme changes ── */
  const glassBase = useMemo(() => glass(t), [t]);

  /* ── localStorage restore — stamp rows synchronously before setFileData ── */
  useEffect(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed: FileData = JSON.parse(saved);
        // Reject any cache saved by a different code version (e.g. HMR contamination)
        if (parsed.cacheVersion !== STORAGE_KEY) return;
        stampRows(parsed.rows, parsed.cols);
        setFileData(parsed);
      }
    } catch { /* ignore */ }
  }, []);

  /* ── localStorage save — main data ── */
  useEffect(() => {
    try {
      if (fileData) localStorage.setItem(STORAGE_KEY, JSON.stringify(fileData));
      else localStorage.removeItem(STORAGE_KEY);
    } catch { /* quota exceeded — ignore */ }
  }, [fileData]);

  /* ── localStorage save — Hubber archive ── */
  useEffect(() => {
    try {
      if (hubberQuick) localStorage.setItem(HUBBER_KEY, JSON.stringify(hubberQuick));
      else localStorage.removeItem(HUBBER_KEY);
    } catch { /* quota exceeded — ignore */ }
  }, [hubberQuick]);

  /* ── localStorage save — filters & prefs ── */
  useEffect(() => {
    try {
      localStorage.setItem(FILTERS_KEY, JSON.stringify({ brandFilter, monthFilter, companyFilter, yearFilter, hubberQYear, darkMode }));
    } catch { /* ignore */ }
  }, [brandFilter, monthFilter, companyFilter, yearFilter, hubberQYear, darkMode]);

  /* ── file processing ── */
  function processFile(file: File) {
    setParseError(null); setFileData(null); setBrandFilter("All"); setMonthFilter("All"); setYearFilter("All"); setSearch("");
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target?.result as ArrayBuffer), { type:"array", cellDates:true });
        const isMultiSheet = wb.SheetNames.length > 1;
        const allRows: Row[] = [];
        const colSet = new Set<string>();

        for (const sheetName of wb.SheetNames) {
          const sh = wb.Sheets[sheetName];
          const rawRows = XLSX.utils.sheet_to_json<Row>(sh, { raw:true,  cellDates:true });
          const fmtRows = XLSX.utils.sheet_to_json<Row>(sh, { raw:false });
          if (!rawRows.length) continue;

          // Collect keys from EVERY row (not just first) — some sheets have columns
          // that are empty in row 1 (e.g. "комісія банку" in Sl-Artmon starts empty).
          // Keys are NORMALIZED to lowercase+trimmed so that "Дата"/"ДАТА"/"дата" all
          // merge into the same column key — preventing silent "No Date" misses.
          rawRows.forEach((r, i) => {
            const fmt = (fmtRows[i] ?? {}) as Row;
            // Collect all original keys from both passes (same sheet → same keys, but
            // collect from both to be safe in edge cases where a cell is empty in one pass)
            const origKeys = new Set([...Object.keys(r), ...Object.keys(fmt)]);

            const merged: Row = {};
            for (const origKey of origKeys) {
              // Normalize to lowercase+trimmed so "Дата"/"ДАТА"/"дата" all map to "дата"
              const normKey = origKey.trim().toLowerCase();
              const rawVal  = r[origKey];
              const fmtVal  = fmt[origKey];
              // Prefer raw numeric for financial values; prefer formatted string for dates
              merged[normKey] = typeof rawVal === "number" && Math.abs(rawVal) <= 9_999_999
                ? rawVal
                : (fmtVal ?? rawVal);
              colSet.add(normKey);
            }
            if (isMultiSheet) merged["_sheet_"] = sheetName;
            allRows.push(merged);
          });
        }

        if (!allRows.length) { setParseError("Файл порожній або не вдалося зчитати рядки."); return; }
        const columns = Array.from(colSet);
        const cols = detectCols(columns, allRows);

        // ── PRE-STAMP every row with calculated fields ──────────────────
        stampRows(allRows, cols);

        // ── AUDIT: totals by marketplace (console.table) ────────────────
        const totalsByMarketplace: Record<string, { rows: number; net_income: number }> = {};
        let grandNet = 0;
        for (const r of allRows) {
          const mkt = String(r._mkt ?? "");
          if (!mkt) continue; // skip rows with no магазин
          if (!totalsByMarketplace[mkt]) totalsByMarketplace[mkt] = { rows: 0, net_income: 0 };
          totalsByMarketplace[mkt].rows++;
          totalsByMarketplace[mkt].net_income += r._net as number;
          grandNet += r._net as number;
        }
        // Add grand total row
        totalsByMarketplace["── GRAND TOTAL ──"] = { rows: allRows.filter(r=>r._mkt).length, net_income: grandNet };
        console.table(totalsByMarketplace);

        setFileData({ fileName:file.name, columns, rows:allRows, cols, isMultiSheet, sheetNames:wb.SheetNames, cacheVersion:STORAGE_KEY });
        setUploadedAt(new Date());
      } catch (err) {
        setParseError("Не вдалося обробити дані. Будь ласка, перевірте формат та структуру файлу.");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  const handleDrop = useCallback((e: React.DragEvent) => { e.preventDefault(); setIsDragging(false); const f=e.dataTransfer.files?.[0]; if(f) processFile(f); }, []);
  const handleChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => { const f=e.target.files?.[0]; if(f) processFile(f); if(fileRef.current) fileRef.current.value=""; }, []);
  const clear = useCallback(() => { setFileData(null); setParseError(null); setBrandFilter("All"); setMonthFilter("All"); setYearFilter("All"); setCompanyFilter("All"); setSearch(""); }, []);

  /* ── normalize marketplace name (case) ── */
  function normBrand(v: unknown): string {
    const s = String(v ?? "").trim();
    if (!s) return "";
    return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
  }

  /* ── row marketplace helper — always uses `магазин` column if available ── */
  function rowBrand(row: Row): string {
    if (!fileData) return "";
    // Priority: магазин column (marketplace) over sheet name
    if (fileData.cols.brand) return normBrand(row[fileData.cols.brand]);
    if (fileData.isMultiSheet) return String(row["_sheet_"] ?? "");
    return "";
  }

  /* ── marketplace values — from pre-stamped _mkt ── */
  const brands = useMemo(()=>{
    if (!fileData) return [];
    const s = new Set(fileData.rows.map(r=>String(r._mkt??"")));
    s.delete("");
    return Array.from(s).sort();
  },[fileData]);

  // Preset brands shown first
  const brandChips = useMemo(()=>{
    const inData = new Set(brands);
    const pre = PRESET_BRANDS.filter(b=>inData.has(b));
    const rest = brands.filter(b=>!PRESET_BRANDS.includes(b));
    return [...pre, ...rest];
  },[brands]);

  /* ── current month key e.g. "2026-05" ── */
  const CURRENT_MONTH_KEY = useMemo(()=>toMonthKey(new Date()),[]);

  /* ── company (brand) values — from _sheet_ column ── */
  const companies = useMemo(()=>{
    if (!fileData?.isMultiSheet) return [];
    const s = new Set<string>();
    fileData.rows.forEach(r=>{ const v=String(r["_sheet_"]??"").trim(); if(v) s.add(v); });
    return Array.from(s).sort();
  },[fileData]);

  /* ── month filter chips — ALL months chronologically + "No Date" last ── */
  const months = useMemo(()=>{
    if (!fileData) return [];
    const s = new Set<string>();
    let hasNoDate = false;
    // Only look at rows with a marketplace (same set as filtered)
    fileData.rows.filter(r => r._mkt).forEach(r=>{
      const k = String(r._monthKey ?? "");
      if (k === "No Date") { hasNoDate = true; return; }
      if (k) s.add(k);
    });
    const sorted = Array.from(s).sort(); // chronological (lexicographic works for YYYY-MM)
    if (hasNoDate) sorted.push("No Date"); // always last
    return sorted;
  },[fileData]);

  /* ── years derived from months (for hierarchical date filter) ── */
  const years = useMemo(()=>{
    const s = new Set<string>();
    months.filter(m=>m!=="No Date").forEach(m=>s.add(m.slice(0,4)));
    return Array.from(s).sort();
  },[months]);

  /* ── months visible in the month row — filtered by yearFilter ── */
  const visibleMonths = useMemo(()=>{
    return months.filter(m=>m!=="No Date" && (yearFilter==="All" || m.startsWith(yearFilter)));
  },[months, yearFilter]);

  /* ── filtered rows — uses pre-stamped _mkt / _sheet_ / _net ── */
  const filtered = useMemo(()=>{
    if (!fileData) return [];
    return fileData.rows.filter(r=>{
      // Only count rows that have a known marketplace (same as CRITICAL CHECK)
      if (!r._mkt) return false;
      // Marketplace filter (UI chip) — uses pre-normalized _mkt
      if (brandFilter!=="All" && r._mkt !== brandFilter) return false;
      // Brand/company filter (UI chip — _sheet_)
      if (companyFilter!=="All" && String(r["_sheet_"]??"").trim()!==companyFilter) return false;
      // Month filter (UI chip) — uses pre-stamped _monthKey
      if (monthFilter!=="All" && String(r._monthKey ?? "No Date") !== monthFilter) return false;
      return true;
    });
  },[fileData, brandFilter, companyFilter, monthFilter]);

  /* ── KPIs — pure reduce on pre-stamped fields ── */
  const kpi = useMemo(()=>{
    if (!fileData) return null;
    const c = fileData.cols;
    let net=0, del=0, com=0, debt=0, refs=0, grossIncome=0;
    for (const r of filtered) {
      net         += r._net   as number;
      grossIncome += r._gross as number;
      del         += r._ship  as number;
      com         += r._fee   as number;
      debt        += toNum(c.debt ? r[c.debt] : null);
      if (isRefusal(r, c)) refs++;
    }
    return { net, del, com, debt, grossIncome, logistics:del+com, orders:filtered.length, refs, successOrders:filtered.length-refs, returnRate:filtered.length>0?(refs/filtered.length)*100:0 };
  },[filtered, fileData]);

  /* ── LFL: prev-month KPI (same brand/company filters, prior month) ── */
  const prevKpi = useMemo(()=>{
    if (!fileData || monthFilter==="All" || monthFilter==="No Date") return null;
    const prevMk = prevMonthKey(monthFilter);
    const c = fileData.cols;
    let net=0, logistics=0, orders=0;
    for (const r of fileData.rows) {
      if (!r._mkt) continue;
      if (brandFilter!=="All"   && r._mkt!==brandFilter) continue;
      if (companyFilter!=="All" && String(r["_sheet_"]??"").trim()!==companyFilter) continue;
      if (String(r._monthKey??"No Date")!==prevMk) continue;
      net       += r._net   as number;
      logistics += (r._fee  as number) + (r._ship as number);
      orders++;
    }
    if (orders===0) return null;
    return { net, logistics, orders };
  },[fileData, monthFilter, brandFilter, companyFilter]);

  /* ── chart ─────────────────────────────────────────────────────
     Source: same `filtered` array as KPI (no re-calculation).
     Every bar = SUM(income − commission − shipping) = r._net.
     "No Date" rows get a "Без дати" bar at the end.
     Sum of ALL bars MUST equal kpi.net. ─────────────────────── */
  const chartData = useMemo(()=>{
    if (!filtered.length) return [];

    // Group all filtered rows by _monthKey; "No Date" rows go to "No Date" bucket
    const bucketMap = new Map<string, { net: number; logistics: number; rows: number }>();
    for (const r of filtered) {
      const k = String(r._monthKey ?? "No Date");
      if (!bucketMap.has(k)) bucketMap.set(k, { net: 0, logistics: 0, rows: 0 });
      const b = bucketMap.get(k)!;
      b.net       += r._net   as number;
      b.logistics += (r._fee  as number) + (r._ship as number);
      b.rows++;
    }

    // Sort: real months ascending, "No Date" always last
    const sortedKeys = Array.from(bucketMap.keys()).sort((a, b) => {
      if (a === "No Date") return 1;
      if (b === "No Date") return -1;
      return a.localeCompare(b);
    });

    const buckets = sortedKeys.map(k => {
      const { net, logistics, rows } = bucketMap.get(k)!;
      const isFuture = k !== "No Date" && k > CURRENT_MONTH_KEY;
      const isNoDate = k === "No Date";
      const label    = isNoDate ? "Без дати" : toMonthShort(k);
      return { key: k, label, netIncome: net, logistics, isFuture, isNoDate, rows };
    });

    // MANDATORY: print Month | Net Income table so user can verify every bucket
    const tableRows = buckets.map(b => ({ Місяць: b.label, "Net Income": +b.netIncome.toFixed(2), "Рядки": b.rows }));
    console.table(tableRows);
    const chartTotal = buckets.reduce((s, b) => s + b.netIncome, 0);
    console.log("Chart total (all bars):", +chartTotal.toFixed(2), "| KPI net:", +(filtered.reduce((s,r)=>s+(r._net as number),0)).toFixed(2));

    return buckets;
  },[filtered, CURRENT_MONTH_KEY]);

  /* ── Sync Error: chart sum must equal KPI net (The Mismatch Rule) ── */
  const syncError = useMemo(()=>{
    if (!kpi || !chartData.length) return false;
    const chartSum = chartData.reduce((s,b) => s + b.netIncome, 0);
    return Math.abs(chartSum - kpi.net) > 0.5; // tolerance: 50 kopecks
  },[kpi, chartData]);

  /* ── top products ── */
  const topProducts = useMemo(()=>{
    const map=new Map<string,{rev:number;orders:number;qty:number;net:number}>();
    for (const r of filtered) {
      const raw = getRowProduct(r, fileData?.cols.product ?? null);
      const key = normalizeProductKey(raw);
      if (!key) continue;
      if(!map.has(key)) map.set(key,{rev:0,orders:0,qty:0,net:0});
      const e=map.get(key)!;
      e.orders++; e.rev+=(r._gross as number);
      e.qty+=parseNum(fileData?.cols.quantity?r[fileData.cols.quantity]:null)||1;
      e.net+=(r._net as number);
    }
    return Array.from(map.entries()).map(([name,v])=>({name,...v})).sort((a,b)=>b.net-a.net).slice(0,10);
  },[filtered, fileData]);

  /* ── top-3 products by gross sales (сума замовлення) ── */
  const topRevProducts = useMemo(()=>{
    const map = new Map<string,number>();
    for (const r of filtered) {
      const raw = getRowProduct(r, fileData?.cols.product ?? null);
      const key = normalizeProductKey(raw);
      if (!key) continue;
      map.set(key, (map.get(key) ?? 0) + (r._gross as number));
    }
    return Array.from(map.entries())
      .map(([name, rev]) => ({ name, rev }))
      .sort((a,b) => b.rev - a.rev)
      .slice(0, 3);
  },[filtered, fileData]);

  /* ── ALL products — full analytics table (filter-aware, sortable) ── */
  const allProducts = useMemo(()=>{
    if (!fileData) return [];
    const map = new Map<string,{qty:number;net:number;rows:number;refs:number}>();
    for (const r of filtered) {
      const raw = getRowProduct(r, fileData.cols.product ?? null);
      const key = normalizeProductKey(raw);
      if (!key) continue;
      if (!map.has(key)) map.set(key, {qty:0, net:0, rows:0, refs:0});
      const e = map.get(key)!;
      e.rows++;
      e.qty += parseNum(fileData.cols.quantity ? r[fileData.cols.quantity] : null) || 1;
      e.net += (r._net as number);
      if (isRefusal(r, fileData.cols)) e.refs++;
    }
    const entries = Array.from(map.entries()).map(([name,v])=>({name,...v}));
    // Apply search
    const q = productSearch.toLowerCase().trim();
    const searched = q ? entries.filter(p => p.name.toLowerCase().includes(q)) : entries;
    // Apply sort
    return searched.sort((a,b)=>{
      const mul = productSort.dir==="desc" ? -1 : 1;
      if (productSort.col==="name") return mul * a.name.localeCompare(b.name, "uk");
      if (productSort.col==="qty")  return mul * (a.qty - b.qty);
      return mul * (a.net - b.net);
    });
  }, [filtered, fileData, productSearch, productSort]);

  /* ── top customers ── */
  const topCustomers = useMemo(()=>{
    const map=new Map<string,{orders:number;net:number}>();
    for (const r of filtered) {
      const key = getRowCustomer(r, fileData?.cols.customer ?? null);
      if (!key) continue;
      if(!map.has(key)) map.set(key,{orders:0,net:0});
      const e=map.get(key)!;
      e.orders++;
      e.net += (r._net as number);
    }
    return Array.from(map.entries()).map(([name,v])=>({name,...v})).sort((a,b)=>b.net-a.net).slice(0,20);
  },[filtered, fileData]);

  /* ── customer insights — retention + VIP ── */
  const customerInsights = useMemo(()=>{
    if (!fileData) return null;
    const c = fileData.cols;
    // Use phone as deduplication key when it's a separate col from customer; otherwise customer key
    const map = new Map<string, { orders:number; net:number; phone:string; displayName:string }>();
    for (const r of filtered) {
      const custKey = getRowCustomer(r, c.customer ?? null);
      if (!custKey) continue;
      // Prefer phone col for dedup if available and different from customer col
      const phoneVal = getRowPhone(r, c.phone !== c.customer ? c.phone : null);
      // Determine deduplification key: phone number wins (more unique)
      const dedupeKey = phoneVal && isPhoneString(phoneVal) ? phoneVal : custKey;
      if (!map.has(dedupeKey)) map.set(dedupeKey, { orders:0, net:0, phone: phoneVal, displayName: isPhoneString(custKey) ? "" : custKey });
      const e = map.get(dedupeKey)!;
      e.orders++;
      e.net += r._net as number;
      if (!e.phone && phoneVal) e.phone = phoneVal;
    }
    const all = Array.from(map.entries()).map(([key, v]) => ({ key, ...v }));
    const newC       = all.filter(x => x.orders === 1).length;
    const returning  = all.filter(x => x.orders >= 2).length;
    const total      = all.length;
    const vip        = [...all].sort((a,b) => b.net - a.net).slice(0, 10);
    return { newC, returning, total, vip };
  },[filtered, fileData]);

  /* ── top cities — if city col detected ── */
  const topCities = useMemo(()=>{
    if (!fileData?.cols.city) return [];
    const map = new Map<string, number>();
    for (const r of filtered) {
      const v = String(r[fileData.cols.city!] ?? "").trim();
      if (!v || v.length > 80) continue;
      map.set(v, (map.get(v) ?? 0) + 1);
    }
    return Array.from(map.entries())
      .map(([city, count]) => ({ city, count }))
      .sort((a,b) => b.count - a.count)
      .slice(0, 5);
  },[filtered, fileData]);

  /* ── rejection reasons — top reasons from filtered refusal rows ── */
  const rejectionReasons = useMemo(()=>{
    if (!fileData) return [];
    const c = fileData.cols;
    const map = new Map<string,number>();
    for (const r of filtered) {
      if (!isRefusal(r, c)) continue;
      // city filter: if active, only include rows matching the selected city
      if (cityFilter && c.city) {
        const norm = extractCity(String(r[c.city] ?? "").trim());
        if (norm !== cityFilter) continue;
      }
      // Try reason col first, then status col, then fall back to generic label
      let reason = "";
      if (c.reason) {
        const v = String(r[c.reason]??"").trim();
        if (v && v.length > 1 && v.length < 120) reason = v;
      }
      if (!reason && c.status) {
        const v = String(r[c.status]??"").trim();
        if (v && v.length > 1 && v.length < 120) reason = v;
      }
      if (!reason) reason = "Причина не вказана";
      map.set(reason, (map.get(reason)??0) + 1);
    }
    return Array.from(map.entries())
      .map(([reason, count]) => ({ reason, count }))
      .sort((a,b) => b.count - a.count)
      .slice(0, 10);
  },[filtered, fileData, cityFilter]);

  /* ── brand MoM trend + sparkline ── */
  const brandTrend = useMemo(()=>{
    const result = new Map<string, BrandTrend>();
    if (!fileData) return result;
    // collect all month keys
    const allMks = Array.from(new Set(
      fileData.rows.map(r=>String(r._monthKey??"")).filter(k=>k && k!=="No Date")
    )).sort();
    const lastMk  = allMks[allMks.length-1] ?? "";
    const prevMk  = lastMk ? prevMonthKey(lastMk) : "";
    const spark6  = allMks.slice(-6);
    for (const brand of brandChips) {
      const byMonth = new Map<string,number>();
      for (const r of fileData.rows) {
        if (r._mkt !== brand) continue;
        const mk = String(r._monthKey ?? "");
        if (!mk || mk === "No Date") continue;
        byMonth.set(mk, (byMonth.get(mk)??0) + (r._net as number));
      }
      const curr = byMonth.get(lastMk) ?? 0;
      const prev = byMonth.get(prevMk) ?? 0;
      const pct  = prev !== 0 ? ((curr-prev)/Math.abs(prev))*100 : null;
      const sparkData = spark6.map(mk => byMonth.get(mk) ?? 0);
      const badge: BrandTrend["badge"] =
        curr < 0 || (pct !== null && pct < -5) ? "risk" :
        pct !== null && pct > 15 ? "rising" :
        pct !== null ? "stable" : "none";
      result.set(brand, { pct, sparkData, badge, curr });
    }
    return result;
  }, [fileData, brandChips]);

  /* ── bar chart show-trend toggle ── */
  const [showBarTrend, setShowBarTrend] = useState(false);

  /* ── 2026 projection from Hubber YTD ── */
  const hubberProj2026 = useMemo(()=>{
    if (!hubberQuick) return null;
    const vals = hubberQuick.values["2026"] ?? {};
    const entries = Object.values(vals).filter((v): v is number => typeof v === "number" && v > 0);
    if (entries.length === 0) return null;
    const ytd = entries.reduce((s,v)=>s+v,0);
    const monthsIn = entries.length;
    const projected = (ytd / monthsIn) * 12;
    const bestY = Object.entries(hubberQuick.yearTotals).reduce((b,[y,v])=>v>(b[1]??0)?[y,v]:b, ["",0] as [string,number]);
    return { ytd, monthsIn, projected, bestYear:bestY[0], bestTotal:bestY[1] as number };
  }, [hubberQuick]);

  /* ── Hubber LFL: same month vs same month of previous year ── */
  const hubberLfl = useMemo(()=>{
    if (!hubberQuick || monthFilter==="All" || monthFilter==="No Date") return null;
    const [yearStr, mStr] = monthFilter.split("-");
    const prevYearStr = String(+yearStr - 1);
    const monthIdx = parseInt(mStr, 10) - 1;
    const monthName = hubberQuick.months[monthIdx];
    if (!monthName) return null;
    const curr = hubberQuick.values[yearStr]?.[monthName] ?? 0;
    const prev = hubberQuick.values[prevYearStr]?.[monthName] ?? 0;
    if (prev === 0) return null;
    return { curr, prev, pct:((curr-prev)/prev)*100, year:yearStr, prevYear:prevYearStr, monthName };
  }, [hubberQuick, monthFilter]);

  /* ── Chart data extended with prev-year Hubber overlay ── */
  const chartDataWithPrevYear = useMemo(()=>{
    if (!hubberQuick || !chartData.length) return chartData;
    return chartData.map(b => {
      if (!b.key || b.key==="No Date") return { ...b, prevYear:null as number|null };
      const [yearStr, mStr] = b.key.split("-");
      const prevYearStr = String(+yearStr - 1);
      const monthIdx = parseInt(mStr, 10) - 1;
      const monthName = hubberQuick.months[monthIdx];
      const v = monthName ? (hubberQuick.values[prevYearStr]?.[monthName] ?? null) : null;
      return { ...b, prevYear: v !== null && v > 0 ? v : null };
    });
  }, [chartData, hubberQuick]);

  /* ── return optimisation advice — shown next to Топ Причин Повернень ── */
  const returnAdvice = useMemo(()=>{
    if (!kpi || rejectionReasons.length === 0) return [];
    const total = kpi.refs;
    const items: { icon:string; title:string; msg:string; badge:string }[] = [];

    // #1 top reason = "not specified"
    const top = rejectionReasons[0];
    if (top && /причина не вказана|не вказан|unknown|не відомо|не зазначен/i.test(top.reason)) {
      const pct = total > 0 ? (top.count / total * 100) : 0;
      items.push({
        icon:"⚠️", badge:`${pct.toFixed(0)}% невідомо`,
        title:"Висока втрата даних",
        msg:"Зробіть поле «Причина повернення» обов'язковим у CRM для контролю логістичних витрат.",
      });
    }

    // #2 size-related reasons (any rank)
    const sizeRows = rejectionReasons.filter(r=>/розмір|size|не той|wrong size/i.test(r.reason));
    if (sizeRows.length > 0) {
      const cnt = sizeRows.reduce((s,r)=>s+r.count,0);
      const pct = total > 0 ? (cnt / total * 100) : 0;
      if (pct > 0) items.push({
        icon:"📦", badge:`${pct.toFixed(0)}% розмір`,
        title:"Помилка пакування",
        msg:"Впровадьте сканування штрих-кодів на складі перед відправкою.",
      });
    }

    return items;
  }, [kpi, rejectionReasons]);

  /* ── table ── */
  const tableRows = useMemo(()=>{
    const q=search.toLowerCase().trim();
    const base=q?filtered.filter(r=>fileData!.columns.some(c=>String(r[c]??"").toLowerCase().includes(q))):filtered;
    return base.slice(0,100);
  },[filtered,search,fileData]);

  /* ── bar: net income by marketplace ── */
  const marketplaceBar = useMemo(()=>{
    if (!fileData) return [];
    const map = new Map<string,{net:number;orders:number;refs:number}>();
    for (const r of filtered) {
      const brand = rowBrand(r) || "Інше";
      if (!map.has(brand)) map.set(brand,{net:0,orders:0,refs:0});
      const e=map.get(brand)!;
      e.orders++;
      if (isRefusal(r,fileData.cols)) e.refs++;
      e.net += r._net as number;
    }
    return Array.from(map.entries()).map(([name,v])=>({name,...v})).sort((a,b)=>b.net-a.net);
  },[filtered,fileData]);

  /* ── marketplace bar with MoM % overlay ── */
  const marketplaceBarWithMoM = useMemo(()=>{
    if (!fileData) return marketplaceBar.map(e=>({...e, momPct:null as number|null}));
    let currMk: string, prevMk: string;
    if (monthFilter!=="All" && monthFilter!=="No Date") {
      currMk = monthFilter; prevMk = prevMonthKey(monthFilter);
    } else {
      const mks = Array.from(new Set(filtered.map(r=>String(r._monthKey??"")).filter(k=>k&&k!=="No Date"))).sort();
      currMk = mks[mks.length-1]??""; prevMk = currMk ? prevMonthKey(currMk) : "";
    }
    if (!prevMk) return marketplaceBar.map(e=>({...e, momPct:null as number|null}));
    const prevMap = new Map<string,number>();
    for (const r of fileData.rows) {
      if (!r._mkt || String(r._monthKey)!==prevMk) continue;
      if (companyFilter!=="All" && String(r["_sheet_"]??"").trim()!==companyFilter) continue;
      prevMap.set(String(r._mkt), (prevMap.get(String(r._mkt))??0)+(r._net as number));
    }
    return marketplaceBar.map(e=>{
      const prev = prevMap.get(e.name) ?? 0;
      const momPct = prev!==0 ? ((e.net-prev)/Math.abs(prev))*100 : null;
      return {...e, momPct};
    });
  }, [fileData, marketplaceBar, showBarTrend, monthFilter, filtered, companyFilter]);

  /* ── marketplace breakdown for archive modal ── */
  const mktBreakdownForArchive = useMemo(()=>{
    if (!fileData || marketplaceBar.length===0) return [];
    const total = marketplaceBar.filter(e=>e.net>0).reduce((s,e)=>s+e.net,0) || 1;
    return marketplaceBar
      .filter(e=>e.net>0)
      .map(e=>({ name:e.name, net:e.net, pct:(e.net/total)*100 }))
      .sort((a,b)=>b.net-a.net);
  }, [fileData, marketplaceBar]);

  /* ── pie: success vs refusals ── */
  const pieData = useMemo(()=>{
    if (!kpi||kpi.orders===0) return [];
    return [
      { name:"Успішні", value:kpi.successOrders },
      { name:"Відмови", value:kpi.refs },
    ];
  },[kpi]);

  /* ── city top-5 for geography module ── */
  const cityTop = useMemo(()=>{
    if (!fileData?.cols.city) return [];
    const orderMap = new Map<string,number>();
    const salesMap = new Map<string,number>(); // gross sales value (сума замовлення)
    for (const r of filtered) {
      const raw = String(r[fileData.cols.city!] ?? "").trim();
      if (!raw || raw==="-" || raw.toLowerCase()==="немає") continue;
      const key = extractCity(raw);
      if (!key || key.length < 2) continue;
      // Gross sales from revenue col (сума замовлення); fall back to _net
      const gross = fileData.cols.revenue ? toNum(r[fileData.cols.revenue]) : (r._net as number);
      orderMap.set(key, (orderMap.get(key) ?? 0) + 1);
      salesMap.set(key, (salesMap.get(key) ?? 0) + gross);
    }
    const totalOrders = Array.from(orderMap.values()).reduce((s,v)=>s+v,0) || 1;
    const totalSales  = Array.from(salesMap.values()).reduce((s,v)=>s+v,0) || 1;
    const sorted = Array.from(orderMap.entries())
      .map(([name, orders]) => {
        const revenue  = salesMap.get(name) ?? 0;
        const avgCheck = orders > 0 ? revenue / orders : 0;
        return {
          name, orders,
          pct:     (orders  / totalOrders) * 100,
          salesPct:(revenue / totalSales)  * 100,
          revenue, avgCheck,
        };
      })
      .sort((a,b) => b.salesPct - a.salesPct); // sort by % of total sales, descending
    return sorted.slice(0, 5);
  },[filtered, fileData]);

  /* ── daily trend: net income per day for the most relevant month ──
     Priority: 1) explicit monthFilter  2) most recent month in data  3) empty */
  const { dailyTrend, dailyTrendMonthLabel } = useMemo(()=>{
    if (!fileData?.cols.date) return { dailyTrend: [], dailyTrendMonthLabel: "" };

    // Determine which month to show
    let mKey: string;
    if (monthFilter && monthFilter !== "All" && monthFilter !== "No Date") {
      mKey = monthFilter;
    } else {
      // Pick the most recent month that has any filtered rows
      let latestKey = "";
      for (const r of filtered) {
        const d = parseDate(r[fileData.cols.date!]);
        if (!d) continue;
        const k = toMonthKey(d);
        if (!latestKey || k > latestKey) latestKey = k;
      }
      if (!latestKey) return { dailyTrend: [], dailyTrendMonthLabel: "" };
      mKey = latestKey;
    }

    const [y, m] = mKey.split("-").map(Number);
    const daysInMonth = new Date(y, m, 0).getDate();
    const map = new Map<number,number>();
    for (const r of filtered) {
      const d = parseDate(r[fileData.cols.date!]);
      if (!d || toMonthKey(d) !== mKey) continue;
      const day = d.getDate();
      map.set(day, (map.get(day) || 0) + rowNet(r, fileData.cols));
    }
    // Only emit days up to today if this is the current month, otherwise full month
    const nowKey = toMonthKey(new Date());
    const maxDay = mKey === nowKey ? new Date().getDate() : daysInMonth;
    const result: Array<{day:string; net:number}> = [];
    for (let i = 1; i <= maxDay; i++) result.push({ day: String(i), net: map.get(i) || 0 });

    const label = `${MUK[m - 1]} ${y}`;
    return { dailyTrend: result, dailyTrendMonthLabel: label };
  },[filtered, fileData, monthFilter]);

  /* ─── render ──────────────────────────────────────────────── */
  return (
    <div style={{ background:t.bg, minHeight:"100vh", fontFamily:"'Inter',-apple-system,BlinkMacSystemFont,sans-serif", letterSpacing:"-0.01em" }}>

      {/* navbar */}
      <div style={{ background:t.nav, borderBottom:`1px solid ${t.border}`, padding:"0 28px", display:"flex", alignItems:"center", justifyContent:"space-between", height:56, position:"sticky", top:0, zIndex:100 }}>
        {/* Brand */}
        <div style={{ display:"flex", alignItems:"center", gap:10 }}>
          {/* Mobile hamburger toggle */}
          {fileData && (
            <button className="sidebar-toggle-btn" onClick={()=>setSidebarOpen(v=>!v)} aria-label="Toggle sidebar">
              <Menu size={18}/>
            </button>
          )}
          {/* Base-style mark: solid electric blue circle */}
          <div style={{ width:28, height:28, borderRadius:6, background:"#0052FF", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>
            <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
              <circle cx="7" cy="7" r="5.5" stroke="#ffffff" strokeWidth="1.4"/>
              <circle cx="7" cy="7" r="2" fill="#ffffff"/>
            </svg>
          </div>
          <div style={{ display:"flex", alignItems:"baseline", gap:6 }}>
            <span style={{ color:t.text, fontSize:15, fontWeight:900, letterSpacing:"-0.03em" }}>SOLANA</span>
            <span style={{ color:t.dim, fontSize:13, fontWeight:400 }}>//</span>
            <span style={{ color:"#0052FF", fontSize:13, fontWeight:700, letterSpacing:"-0.02em" }}>CORE</span>
          </div>
        </div>
        {/* Actions */}
        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
          {/* Persistence badge — visible when any data is loaded from/into storage */}
          {(fileData || hubberQuick) && (
            <div title={[fileData?"Аналітика збережена":"", hubberQuick?"Hubber архів збережено":""].filter(Boolean).join(" · ")}
              style={{ display:"flex", alignItems:"center", gap:5, padding:"4px 9px", borderRadius:5, background: t.dark?"rgba(22,163,74,0.12)":"rgba(22,163,74,0.09)", border:`1px solid rgba(22,163,74,0.25)` }}>
              <HardDrive size={11} style={{ color:"#16A34A" }}/>
              <span style={{ fontSize:10, fontWeight:700, color:"#16A34A", letterSpacing:"0.03em" }}>Дані збережено</span>
            </div>
          )}
          {fileData && (
            <button onClick={clear} style={{ padding:"5px 12px", borderRadius:6, background:"transparent", border:`1px solid ${t.border}`, color:t.sub, fontSize:11, fontWeight:500, cursor:"pointer", display:"flex", alignItems:"center", gap:4 }}><X size={11} /> Скинути</button>
          )}
          <button onClick={()=>fileRef.current?.click()} style={{ padding:"6px 16px", borderRadius:6, background:"#0052FF", color:"#ffffff", fontSize:12, fontWeight:600, cursor:"pointer", border:"none", display:"flex", alignItems:"center", gap:6, letterSpacing:"-0.01em" }}>
            <Upload size={12} /> Генерація звіту
          </button>
          <button onClick={()=>setDarkMode(!darkMode)} style={{ width:32, height:32, borderRadius:6, background:"transparent", border:`1px solid ${t.border}`, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", color:t.dim }}>
            {darkMode?<Sun size={13}/>:<Moon size={13}/>}
          </button>
        </div>
      </div>

      {/* ── Body: sidebar (fixed) + main content ─────────────── */}
      <div>

        {/* Mobile overlay */}
        {fileData && sidebarOpen && (
          <div className={`sidebar-overlay${sidebarOpen ? " sidebar-open" : ""}`} onClick={()=>setSidebarOpen(false)}/>
        )}

        {/* ── LEFT SIDEBAR — position:fixed via CSS, never shifts content ── */}
        {fileData && (()=>{
          const st: T = t;
          return (
          <nav className={`orbit-sidebar${sidebarOpen ? " sidebar-open" : ""}`} style={{
            background: st.bg,
            borderRight: `1px solid ${st.border}`,
            padding: "14px 8px 36px",
            display:"flex", flexDirection:"column", gap:4,
          }}>

            {/* Status pill */}
            <div style={{
              padding:"9px 12px", borderRadius:6, marginBottom:8,
              background:"rgba(0,82,255,0.05)", border:"1px solid rgba(0,82,255,0.14)",
            }}>
              <p style={{ fontSize:11, fontWeight:700, color:"#0052FF", margin:0 }}>● Аналіз активний</p>
              <p style={{ fontSize:10, color:st.dim, margin:"2px 0 0" }}>Solana // Core</p>
            </div>

            {/* Marketplace — collapsible + 2-col icon grid */}
            {brandChips.length>0 && (
              <SidebarSection
                icon={<Store size={12}/>}
                label="Маркетплейс"
                open={mktOpen}
                onToggle={()=>setMktOpen(v=>!v)}
                st={st}
              >
                <SidebarFilterBtn label="Всі маркетплейси" active={brandFilter==="All"} onClick={()=>setBrandFilter("All")} t={st}/>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(2, 1fr)", gap:4, marginTop:2 }}>
                  {brandChips.map(b=>(
                    <BrandGridCell key={b} brand={b} active={brandFilter===b} onClick={()=>setBrandFilter(b)} t={st} logo={<MktLogo brand={b} size={22}/>} isTop={b===(marketplaceBar[0]?.name??"")} trend={brandTrend.get(b)}/>
                  ))}
                </div>
              </SidebarSection>
            )}

            {/* Divider */}
            {brandChips.length>0 && companies.length>0 && <div style={{ height:1, background:st.border, margin:"4px 10px" }}/>}

            {/* Collections — collapsible + 2-col icon grid */}
            {companies.length>0 && (
              <SidebarSection
                icon={<Building2 size={12}/>}
                label="Колекції SOLANA"
                open={brandsOpen}
                onToggle={()=>setBrandsOpen(v=>!v)}
                st={st}
              >
                <SidebarFilterBtn label="Всі бренди" active={companyFilter==="All"} onClick={()=>setCompanyFilter("All")} t={st}/>
                <div style={{ display:"grid", gridTemplateColumns:"repeat(2, 1fr)", gap:4, marginTop:2 }}>
                  {companies.map(c=>(
                    <BrandGridCell key={c} brand={c} active={companyFilter===c} onClick={()=>setCompanyFilter(c)} t={st}/>
                  ))}
                </div>
              </SidebarSection>
            )}

            {/* Divider */}
            {months.length>0 && <div style={{ height:1, background:st.border, margin:"4px 10px" }}/>}

            {/* Date — collapsible (moved to bottom) */}
            {months.length>0 && (
              <SidebarSection
                icon={<CalendarDays size={12}/>}
                label="Дата"
                open={dateOpen}
                onToggle={()=>setDateOpen(v=>!v)}
                st={st}
              >
                {years.length>1 && (
                  <div style={{ padding:"2px 10px 8px" }}>
                    <span style={{ fontSize:9, color:st.dim, display:"block", marginBottom:4, fontWeight:600, letterSpacing:"0.06em", textTransform:"uppercase" }}>Рік:</span>
                    <div style={{ display:"flex", gap:4, flexWrap:"wrap" }}>
                      <Chip label="Всі" active={yearFilter==="All"} onClick={()=>setYearFilter("All")} t={st} variant="meta"/>
                      {years.map(y=>(
                        <Chip key={y} label={y} active={yearFilter===y} onClick={()=>{setYearFilter(y); if(monthFilter!=="All"&&monthFilter!=="No Date"&&!monthFilter.startsWith(y)) setMonthFilter("All");}} t={st}/>
                      ))}
                    </div>
                  </div>
                )}
                <SidebarFilterBtn label="Всі місяці" active={monthFilter==="All"} onClick={()=>setMonthFilter("All")} t={st}/>
                {visibleMonths.map(m=>(
                  <SidebarFilterBtn key={m} label={toMonthLabel(m)} active={monthFilter===m} onClick={()=>setMonthFilter(m)} t={st}/>
                ))}
                {months.includes("No Date") && (
                  <SidebarFilterBtn label="Без дати" active={monthFilter==="No Date"} onClick={()=>setMonthFilter("No Date")} t={st}/>
                )}
              </SidebarSection>
            )}

            {/* Reset — minimal/discreet */}
            {(brandFilter!=="All"||monthFilter!=="All"||companyFilter!=="All"||yearFilter!=="All") && (
              <button
                onClick={()=>{setBrandFilter("All");setMonthFilter("All");setCompanyFilter("All");setYearFilter("All");}}
                style={{
                  marginTop:10, padding:"5px 10px", borderRadius:5,
                  background:"transparent", border:`1px solid ${st.border}`,
                  color:st.dim, fontSize:9.5, fontWeight:600, cursor:"pointer",
                  display:"flex", alignItems:"center", justifyContent:"center", gap:4,
                  transition:"background 0.15s ease", letterSpacing:"0.03em",
                }}>
                <RefreshCw size={9}/> Скинути фільтри
              </button>
            )}

            {/* Hubber archive quick panel */}
            <HubberSidebarPanel
              data={hubberQuick}
              setData={setHubberQuick}
              selYear={hubberQYear}
              setSelYear={setHubberQYear}
              t={st}
              mktBreakdown={mktBreakdownForArchive}
            />

          </nav>
          );
        })()}

        {/* ── MAIN CONTENT — offset by 252px when sidebar is fixed ── */}
        <div className={`orbit-content${fileData ? " orbit-main-offset" : ""}`} style={{ minWidth:0, padding:"14px 16px 40px" }}>

          {/* error */}
        {parseError && (
          <div style={{ ...glass(t,t.red), padding:"12px 16px", display:"flex", alignItems:"center", gap:10, marginBottom:14 }}>
            <span style={{ color:t.red, fontSize:13, flex:1 }}>{parseError}</span>
            <button onClick={()=>setParseError(null)} style={{ background:"none", border:"none", cursor:"pointer", color:t.dim }}><X size={12}/></button>
          </div>
        )}

        {/* upload zone — shown when no data */}
        {!fileData && (
          <div
            onDragOver={e=>{e.preventDefault();setIsDragging(true);}}
            onDragLeave={()=>setIsDragging(false)}
            onDrop={handleDrop}
            onClick={()=>fileRef.current?.click()}
            style={{ border:`1px dashed ${isDragging?"#0052FF":t.border}`, borderRadius:8, padding:"90px 40px", display:"flex", flexDirection:"column", alignItems:"center", gap:20, cursor:"pointer", transition:"all 0.15s", background:isDragging?"rgba(0,82,255,0.03)":"transparent" }}
          >
            <div style={{ width:64, height:64, borderRadius:8, background:"rgba(0,82,255,0.06)", border:"1px solid rgba(0,82,255,0.14)", display:"flex", alignItems:"center", justifyContent:"center" }}>
              <Upload size={26} style={{ color:"#0052FF" }}/>
            </div>
            <div style={{ textAlign:"center" }}>
              <p style={{ color:t.text, fontSize:18, fontWeight:800, margin:0, letterSpacing:"-0.03em" }}>Завантажте дані звітності</p>
              <p style={{ color:t.dim, fontSize:13, marginTop:6, fontWeight:400 }}>Підтримуються стандартні формати звітності</p>
            </div>
            <div style={{ padding:"9px 28px", background:"#0052FF", borderRadius:6, color:"#ffffff", fontSize:13, fontWeight:600, letterSpacing:"-0.01em" }}>Завантажити дані</div>
          </div>
        )}

        {/* dashboard */}
        {fileData && kpi && (
          <div style={{ display:"flex", flexDirection:"column", gap:10 }}>

            {/* ── BASE-PROTOCOL HERO ──────────────────────────────────── */}
            <div className="orbit-fadein" style={{ marginBottom:4, paddingBottom:18, borderBottom:`1px solid ${t.border}` }}>
              <div style={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", gap:16, flexWrap:"wrap" }}>
                <div>
                  <p style={{ margin:"0 0 4px", fontSize:11, fontWeight:700, letterSpacing:"0.08em", textTransform:"uppercase", color:"#0052FF" }}>
                    Orbit Analytics · Powered by Base Protocol
                  </p>
                  <h1 style={{ margin:0, fontSize:"clamp(22px, 2.2vw, 32px)", fontWeight:900, letterSpacing:"-0.045em", lineHeight:1.08, color:t.text }}>
                    A global business,{" "}
                    <span style={{ color:"#0052FF" }}>built on data</span>
                  </h1>
                  <p style={{ margin:"6px 0 0", fontSize:13, color:t.dim, fontWeight:400, letterSpacing:"-0.01em" }}>
                    Реальна аналітика · {kpi.orders.toLocaleString("uk-UA")} замовлень · {fmt(kpi.grossIncome)} ₴ виручки
                  </p>
                </div>
                {/* live status chip */}
                <div style={{ display:"flex", alignItems:"center", gap:6, padding:"6px 14px", borderRadius:20, background:"rgba(0,82,255,0.06)", border:"1px solid rgba(0,82,255,0.16)", flexShrink:0, alignSelf:"center" }}>
                  <span style={{ width:7, height:7, borderRadius:"50%", background:"#0052FF", boxShadow:"0 0 0 3px rgba(0,82,255,0.18)", flexShrink:0 }}/>
                  <span style={{ fontSize:11, fontWeight:700, color:"#0052FF", letterSpacing:"0.04em", textTransform:"uppercase" as const }}>Live</span>
                </div>
              </div>
            </div>

            {/* KPI row — memoized, only re-renders when kpi data changes */}
            <KpiRow
              kpi={kpi}
              prevKpi={prevKpi ?? null}
              hubberLfl={hubberLfl ?? null}
              filteredCount={filtered.length}
              syncError={syncError ?? false}
              debtCol={fileData.cols.debt ?? null}
              t={t}
              fmt={fmt}
            />

            {/* ── Annual revenue projection card ── */}
            {hubberProj2026 && (()=>{
              const fmtWhole = (n: number) => Math.round(n).toLocaleString("uk-UA").replace(/,/g," ");
              const vsRec = hubberProj2026.bestTotal > 0
                ? ((hubberProj2026.projected - hubberProj2026.bestTotal) / hubberProj2026.bestTotal * 100)
                : null;
              return (
                <div className="orbit-fadein" style={{ ...glassBase, padding:"20px 28px 18px", display:"flex", flexDirection:"column", alignItems:"center", textAlign:"center", gap:6, animationDelay:"60ms", borderLeft:"3px solid #0052FF" }}>
                  <div style={{ fontSize:10, fontWeight:700, letterSpacing:"0.07em", textTransform:"uppercase" as const, color:t.dim }}>📊 Дохід за рік</div>
                  <div style={{ fontSize:36, fontWeight:900, color:"#0052FF", letterSpacing:"-0.04em", lineHeight:1, margin:"4px 0 0" }}>
                    {fmtWhole(hubberProj2026.projected)} ₴
                  </div>
                  <div style={{ fontSize:11, color:"#6B7280", marginTop:2 }}>Базується на результатах за {hubberProj2026.monthsIn} міс. (YTD)</div>

                  {/* Secondary row */}
                  <div style={{ display:"flex", gap:20, marginTop:10, paddingTop:10, borderTop:`1px solid ${t.border}`, width:"100%", justifyContent:"center", flexWrap:"wrap" }}>
                    <div style={{ textAlign:"center" }}>
                      <div style={{ fontSize:9, fontWeight:700, letterSpacing:"0.05em", textTransform:"uppercase" as const, color:t.dim, marginBottom:3 }}>YTD факт</div>
                      <div style={{ fontSize:14, fontWeight:800, color:t.text, letterSpacing:"-0.02em" }}>{fmtWhole(hubberProj2026.ytd)} ₴</div>
                    </div>
                    {vsRec !== null && (
                      <>
                        <div style={{ width:1, background:t.border }}/>
                        <div style={{ textAlign:"center" }}>
                          <div style={{ fontSize:9, fontWeight:700, letterSpacing:"0.05em", textTransform:"uppercase" as const, color:t.dim, marginBottom:3 }}>vs рекорд {hubberProj2026.bestYear}</div>
                          <div style={{ fontSize:14, fontWeight:800, letterSpacing:"-0.02em", color:vsRec>=0?"#0052FF":t.red }}>
                            {vsRec>=0?"+":""}{vsRec.toFixed(0)}%
                          </div>
                        </div>
                      </>
                    )}
                  </div>
                </div>
              );
            })()}

            {/* area chart — monthly trend — locked 360px */}
            <ChartErrorBoundary t={t} label="Динаміка по місяцях">
            {(()=>{
              const pastBars    = chartData.filter(d=>!d.isFuture && !d.isNoDate);
              const noDateBar   = chartData.find(d=> d.isNoDate);
              const futureBars  = chartData.filter(d=> d.isFuture);
              const pastCount   = pastBars.length;
              const futureCount = futureBars.length;
              const noDateCount = noDateBar?.rows ?? 0;
              const refLabel    = pastBars.at(-1)?.label;
              return (
              <div className="analytics-card analytics-card--trend orbit-fadein" style={{ ...glassBase, padding:"20px 20px 12px", animationDelay:"80ms" }}>
                {chartData.length < 1 && (
                  <div style={{ display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", minHeight:200, gap:8 }}>
                    <span style={{ fontSize:28, opacity:0.18 }}>📈</span>
                    <span style={{ fontSize:13, fontWeight:600, color:t.sub }}>Динаміка по місяцях</span>
                    <span style={{ fontSize:11, color:t.dim }}>Немає даних для обраного фільтра</span>
                  </div>
                )}
                {chartData.length >= 1 && (<>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:16 }}>
                  <div>
                    <p style={{ color:t.text, fontSize:14, fontWeight:700, margin:0 }}>Динаміка по місяцях</p>
                    <p style={{ color:t.dim, fontSize:11, margin:"3px 0 0" }}>
                      {pastCount} міс. фактичних{futureCount>0?` · ${futureCount} прогноз`:""}{noDateCount>0?` · ${noDateCount} без дати`:""}
                    </p>
                  </div>
                  <div style={{ display:"flex", gap:14, flexWrap:"wrap" }}>
                    {[[t.em,"Дохід"],[t.blue,"Логістика"]].map(([c,n])=>(
                      <div key={n} style={{ display:"flex", alignItems:"center", gap:5 }}>
                        <div style={{ width:10, height:3, borderRadius:2, background:c }}/>
                        <span style={{ fontSize:11, color:t.sub }}>{n}</span>
                      </div>
                    ))}
                    {hubberQuick && (
                      <div style={{ display:"flex", alignItems:"center", gap:5 }}>
                        <div style={{ width:14, height:0, borderTop:"1.5px dashed #9CA3AF" }}/>
                        <span style={{ fontSize:11, color:t.sub }}>Хаббер мин. рік</span>
                      </div>
                    )}
                  </div>
                </div>
                <ResponsiveContainer width="100%" height={230}>
                  <AreaChart data={chartDataWithPrevYear} margin={{ top:4, right:8, left:8, bottom:4 }}>
                    <defs>
                      <linearGradient id="gE" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor={t.em} stopOpacity={t.dark?0.22:0.16}/>
                        <stop offset="100%" stopColor={t.em} stopOpacity={0}/>
                      </linearGradient>
                      <linearGradient id="gB" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor={t.blue} stopOpacity={t.dark?0.18:0.12}/>
                        <stop offset="100%" stopColor={t.blue} stopOpacity={0}/>
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="1 0" stroke={t.dark?"rgba(255,255,255,0.04)":"rgba(0,0,0,0.05)"} vertical={false}/>
                    <XAxis dataKey="label" tick={{ fontSize:11, fill:t.sub }} tickLine={false} axisLine={false} interval="preserveStartEnd"/>
                    <YAxis tickFormatter={v=>fmt(v)} tick={{ fontSize:10, fill:t.dim }} tickLine={false} axisLine={false} width={96} domain={["auto","auto"]}/>
                    <Tooltip content={<TipBox t={t}/>} cursor={{ stroke:t.dark?"rgba(255,255,255,0.06)":"rgba(0,0,0,0.06)", strokeWidth:1 }}/>
                    {refLabel && futureCount>0 && (
                      <ReferenceLine x={refLabel} stroke={t.dim} strokeDasharray="5 3" strokeWidth={1.5}
                        label={{ value:"прогноз →", position:"insideTopRight", fontSize:9, fill:t.dim, fontWeight:600 }}/>
                    )}
                    <Area isAnimationActive={true} animationDuration={500} animationEasing="ease-out" type="monotone" dataKey="netIncome" name="Дохід" stroke={t.em} strokeWidth={2.5} fill="url(#gE)"
                      dot={chartData.length<=36 ? (p: Record<string,unknown>) => {
                        const isND = (chartData[p.index as number] ?? {}).isNoDate;
                        return <circle key={p.index as number} cx={p.cx as number} cy={p.cy as number} r={isND?5:3} fill={isND?"#f59e0b":t.em} stroke="none"/>;
                      } : false}
                      activeDot={{ r:5, fill:t.em, strokeWidth:0 }}/>
                    <Area isAnimationActive={true} animationDuration={500} animationEasing="ease-out" type="monotone" dataKey="logistics" name="Логістика" stroke={t.blue} strokeWidth={2.5} fill="url(#gB)"
                      dot={chartData.length<=36?{r:3,fill:t.blue,strokeWidth:0}:false}
                      activeDot={{ r:5, fill:t.blue, strokeWidth:0 }}/>
                    {hubberQuick && (
                      <Line isAnimationActive={false} type="monotone" dataKey="prevYear" name="Хаббер мин. рік"
                        stroke="#9CA3AF" strokeWidth={1.5} strokeDasharray="5 3"
                        dot={false} activeDot={{ r:3, fill:"#9CA3AF", strokeWidth:0 }} connectNulls/>
                    )}
                  </AreaChart>
                </ResponsiveContainer>
                </>)}
              </div>
              );
            })()}
            </ChartErrorBoundary>

            {/* ── Row 2: Bar chart by marketplace + Donut + Pie chart ── */}
            <ChartErrorBoundary t={t} label="Аналітика маркетплейсів">
            <div style={{ display:"grid", gridTemplateColumns:"1fr 260px 260px", gap:10 }}>

              {/* Bar chart — net income by marketplace */}
              <div className="analytics-card analytics-card--bar orbit-fadein" style={{ ...glassBase, padding:"20px 20px 12px", animationDelay:"140ms" }}>
                {marketplaceBar.length > 0 ? (<>
                  <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:14 }}>
                    <p style={{ color:t.text, fontSize:14, fontWeight:700, margin:0 }}>Чистий Дохід по Маркетплейсах</p>
                    <button
                      onClick={()=>setShowBarTrend(v=>!v)}
                      style={{ padding:"3px 10px", borderRadius:6, border:`1px solid ${showBarTrend?t.blue:t.border}`, background:showBarTrend?`${t.blue}14`:"transparent", color:showBarTrend?t.blue:t.sub, fontSize:10, fontWeight:700, cursor:"pointer", transition:"all 0.14s ease" }}
                    >
                      {showBarTrend?"✓ Тренд":"Показати тренд"}
                    </button>
                  </div>
                  {(()=>{
                    const totalNet = marketplaceBarWithMoM.filter(e=>e.net>0).reduce((s,e)=>s+e.net,0)||1;
                    return (
                      <ResponsiveContainer width="100%" height={220}>
                        <BarChart data={marketplaceBarWithMoM} margin={{ top:showBarTrend?28:18, right:8, left:8, bottom:4 }}>
                          <CartesianGrid strokeDasharray="1 0" stroke={t.dark?"rgba(255,255,255,0.04)":"rgba(0,0,0,0.05)"} vertical={false}/>
                          <XAxis dataKey="name" tick={{ fontSize:11, fill:t.sub }} tickLine={false} axisLine={false}/>
                          <YAxis tickFormatter={v=>fmt(v)} tick={{ fontSize:10, fill:t.dim }} tickLine={false} axisLine={false} width={90}/>
                          <Tooltip formatter={(v:number)=>fmt(v)} contentStyle={{ background:t.dark?"rgba(4,6,14,0.97)":"#fff", border:`1px solid ${t.border}`, borderRadius:8, fontSize:12 }}/>
                          <Bar isAnimationActive={true} animationDuration={500} animationEasing="ease-out" dataKey="net" name="Дохід" radius={[6,6,0,0]}
                            label={(props: Record<string,unknown>) => {
                              const entry = marketplaceBarWithMoM[props.index as number];
                              if (!entry || entry.net <= 0) return <text key={props.index as number}/>;
                              const sharePct = ((entry.net / totalNet) * 100).toFixed(0);
                              const cx = props.x as number + (props.width as number) / 2;
                              const cy = props.y as number - 6;
                              if (showBarTrend && entry.momPct != null) {
                                const momColor = entry.momPct >= 0 ? "#16A34A" : "#FF4D4D";
                                return (
                                  <g key={props.index as number}>
                                    <text x={cx} y={cy-10} textAnchor="middle" fontSize={9} fontWeight={700} fill={t.dim}>{sharePct}%</text>
                                    <text x={cx} y={cy}    textAnchor="middle" fontSize={9} fontWeight={700} fill={momColor}>{entry.momPct>=0?"+":""}{entry.momPct.toFixed(0)}% MoM</text>
                                  </g>
                                );
                              }
                              return <text key={props.index as number} x={cx} y={cy} textAnchor="middle" fontSize={9} fontWeight={700} fill={t.sub}>{sharePct}%</text>;
                            }}
                          >
                            {marketplaceBarWithMoM.map((entry,i)=>(
                              <Cell key={i} fill={entry.net<0 ? t.red : BASE_BLUE[i] ?? BASE_BLUE[BASE_BLUE.length-1]}/>
                            ))}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    );
                  })()}
                  {/* Insight box */}
                  {marketplaceBar.length >= 2 && (()=>{
                    const leader = marketplaceBar[0];
                    const second = marketplaceBar[1];
                    const totalPos = marketplaceBar.filter(e=>e.net>0).reduce((s,e)=>s+e.net,0)||1;
                    const leaderPct = ((leader.net/totalPos)*100).toFixed(0);
                    return (
                      <div style={{ marginTop:12, padding:"9px 12px", borderRadius:8, background:t.dark?"rgba(138,154,91,0.10)":"rgba(138,154,91,0.07)", border:"1px solid rgba(138,154,91,0.22)", display:"flex", alignItems:"flex-start", gap:7 }}>
                        <span style={{ fontSize:12, flexShrink:0, marginTop:1 }}>💡</span>
                        <span style={{ fontSize:10, color:t.text, lineHeight:1.55 }}>
                          <strong style={{ color:"#004080" }}>{leader.name}</strong> генерує основний потік готівки ({leaderPct}%).{" "}
                          {second && <span>Рентабельність на <strong style={{ color:"#1E90FF" }}>{second.name}</strong> вища завдяки нижчій вартості логістики.</span>}
                        </span>
                      </div>
                    );
                  })()}
                </>) : (
                  <div style={{ display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", minHeight:180, gap:8 }}>
                    <span style={{ fontSize:26, opacity:0.18 }}>📊</span>
                    <span style={{ fontSize:12, fontWeight:600, color:t.sub }}>Чистий Дохід по Маркетплейсах</span>
                    <span style={{ fontSize:11, color:t.dim }}>Немає даних для обраного фільтра</span>
                  </div>
                )}
              </div>

              {/* Donut chart — marketplace share */}
              <div className="analytics-card orbit-fadein" style={{ ...glassBase, padding:"20px", display:"flex", flexDirection:"column", gap:10, animationDelay:"160ms" }}>
                {marketplaceBar.filter(e=>e.net>0).length > 0 ? (()=>{
                  const donutData   = marketplaceBar.filter(e=>e.net>0);
                  const donutTotal  = donutData.reduce((s,e)=>s+e.net,0)||1;
                  return (<>
                    <p style={{ color:t.text, fontSize:13, fontWeight:700, margin:0, letterSpacing:"-0.02em" }}>Частка ринку</p>
                    <ResponsiveContainer width="100%" height={160}>
                      <PieChart>
                        <Pie
                          isAnimationActive={true} animationDuration={600} animationEasing="ease-out"
                          data={donutData} cx="50%" cy="50%"
                          innerRadius={46} outerRadius={68}
                          paddingAngle={3} dataKey="net"
                        >
                          {donutData.map((_,i)=>(
                            <Cell key={i} fill={BASE_BLUE[i] ?? BASE_BLUE[BASE_BLUE.length-1]}/>
                          ))}
                        </Pie>
                        <Tooltip
                          formatter={(v:number, name:string)=>[`${fmt(v)} (${((v/donutTotal)*100).toFixed(1)}%)`, name]}
                          contentStyle={{ background:t.dark?"rgba(4,6,14,0.97)":"#fff", border:`1px solid ${t.border}`, borderRadius:8, fontSize:11 }}
                        />
                      </PieChart>
                    </ResponsiveContainer>
                    {/* Legend */}
                    <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
                      {donutData.map((entry, i) => {
                        const pct = ((entry.net/donutTotal)*100).toFixed(1);
                        const clr = BASE_BLUE[i] ?? BASE_BLUE[BASE_BLUE.length-1];
                        return (
                          <div key={entry.name} style={{ display:"flex", alignItems:"center", justifyContent:"space-between", gap:6 }}>
                            <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                              <div style={{ width:8, height:8, borderRadius:99, background:clr, flexShrink:0 }}/>
                              <span style={{ fontSize:11, color:i===0?"#004080":t.text, fontWeight:i===0?700:500 }}>{entry.name}</span>
                            </div>
                            <span style={{ fontSize:11, fontWeight:700, color:i===0?"#004080":"#9CA3AF" }}>{pct}%</span>
                          </div>
                        );
                      })}
                    </div>
                  </>);
                })() : (
                  <div style={{ display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", flex:1, gap:8 }}>
                    <span style={{ fontSize:24, opacity:0.18 }}>🍩</span>
                    <span style={{ fontSize:11, color:t.dim }}>Частка ринку</span>
                  </div>
                )}
              </div>

              {/* Pie chart — success vs refusals */}
              <div className="analytics-card analytics-card--pie orbit-fadein" style={{ ...glassBase, padding:"20px", display:"flex", flexDirection:"column", gap:12, animationDelay:"200ms" }}>
                {pieData.length > 0 ? (<>
                  <p style={{ color:t.text, fontSize:14, fontWeight:700, margin:0 }}>Успішні / Відмови</p>
                  <ResponsiveContainer width="100%" height={180}>
                    <PieChart>
                    <Pie isAnimationActive={true} animationDuration={500} animationEasing="ease-out" animationBegin={100} data={pieData} cx="50%" cy="50%" innerRadius={52} outerRadius={76} paddingAngle={3} dataKey="value">
                      <Cell fill={t.em}/>
                      <Cell fill={t.red}/>
                    </Pie>
                    <Tooltip formatter={(v:number,n:string)=>[v,n]} contentStyle={{ background:t.dark?"rgba(4,6,14,0.97)":"#fff", border:`1px solid ${t.border}`, borderRadius:8, fontSize:12 }}/>
                    <Legend iconType="circle" iconSize={8} wrapperStyle={{ fontSize:12, color:t.sub }}/>
                  </PieChart>
                  </ResponsiveContainer>
                  <div style={{ display:"flex", gap:16 }}>
                    <div style={{ textAlign:"center" }}>
                      <div style={{ fontSize:22, fontWeight:900, color:t.em }}>{kpi!.successOrders}</div>
                      <div style={{ fontSize:10, color:t.dim }}>Успішні</div>
                    </div>
                    <div style={{ width:1, background:t.border }}/>
                    <div style={{ textAlign:"center" }}>
                      <div style={{ fontSize:22, fontWeight:900, color:t.red }}>{kpi!.refs}</div>
                      <div style={{ fontSize:10, color:t.dim }}>Відмови</div>
                    </div>
                    <div style={{ width:1, background:t.border }}/>
                    <div style={{ textAlign:"center" }}>
                      <div style={{ fontSize:22, fontWeight:900, color:t.amb }}>{kpi!.returnRate.toFixed(1)}%</div>
                      <div style={{ fontSize:10, color:t.dim }}>% Відмов</div>
                    </div>
                  </div>
                </>) : (
                  <div style={{ display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", minHeight:180, gap:8 }}>
                    <span style={{ fontSize:26, opacity:0.18 }}>🥧</span>
                    <span style={{ fontSize:12, fontWeight:600, color:t.sub }}>Успішні / Відмови</span>
                    <span style={{ fontSize:11, color:t.dim }}>Немає даних для обраного фільтра</span>
                  </div>
                )}
              </div>
            </div>
            </ChartErrorBoundary>

            {/* ── City geography module ── */}
            {cityTop.length > 0 && (()=>{
              const kyivEntry   = cityTop.find(c => c.name.toLowerCase().startsWith("київ") || c.name.toLowerCase().startsWith("kyiv"));
              const kyivPct     = kyivEntry?.salesPct ?? 0;
              const top2pct     = cityTop.slice(0,2).reduce((s,c)=>s+c.salesPct, 0);
              const top2names   = cityTop.slice(0,2).map(c=>c.name).join(" та ");
              return (
                <div className="orbit-fadein" style={{ ...glassBase, padding:"20px 22px 18px", animationDelay:"240ms" }}>
                  {/* Header */}
                  <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:14 }}>
                    <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                      <span style={{ fontSize:16 }}>📍</span>
                      <p style={{ color:t.text, fontSize:14, fontWeight:700, margin:0, letterSpacing:"-0.02em" }}>Географія та середній чек</p>
                    </div>
                    <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                      {cityFilter && (
                        <button onClick={()=>setCityFilter(null)} style={{ fontSize:10, fontWeight:700, padding:"2px 8px", borderRadius:5, border:`1px solid ${t.blue}44`, background:`${t.blue}14`, color:t.blue, cursor:"pointer" }}>✕ {cityFilter}</button>
                      )}
                      <span style={{ fontSize:10, fontWeight:700, letterSpacing:"0.06em", textTransform:"uppercase" as const, color:"#9CA3AF" }}>Топ {cityTop.length}</span>
                    </div>
                  </div>

                  {/* ComposedChart — bars = orders volume (with salesPct label), line = avg check */}
                  {(()=>{
                    const totalOrds  = cityTop.reduce((s,c)=>s+c.orders,0) || 1;
                    const chartData  = cityTop.map(c => ({
                      name:     c.name,
                      замовлень: c.orders,
                      pctLabel:  parseFloat(c.salesPct.toFixed(1)),
                      середній:  Math.round(c.avgCheck),
                    }));
                    const maxAvgLine = Math.max(...cityTop.map(c=>c.avgCheck)) || 1;
                    return (
                      <ResponsiveContainer width="100%" height={155}>
                        <ComposedChart data={chartData} margin={{ top:18, right:12, left:-10, bottom:2 }}>
                          <CartesianGrid strokeDasharray="1 0" stroke={t.dark?"rgba(255,255,255,0.04)":"rgba(0,0,0,0.05)"} vertical={false}/>
                          <XAxis dataKey="name" tick={{ fontSize:10, fill:t.sub }} tickLine={false} axisLine={false}/>
                          <YAxis yAxisId="left"  tick={{ fontSize:9, fill:t.dim }} tickLine={false} axisLine={false} width={30} tickFormatter={(v:number)=>String(v)}/>
                          <YAxis yAxisId="right" orientation="right" tick={{ fontSize:9, fill:"#1E90FF" }} tickLine={false} axisLine={false} width={50} domain={[0, maxAvgLine*1.35]} tickFormatter={(v:number)=>v>0?`${Math.round(v/100)*100} ₴`:""} />
                          <Tooltip
                            contentStyle={{ background:t.dark?"rgba(4,6,14,0.97)":"#fff", border:`1px solid ${t.border}`, borderRadius:8, fontSize:11 }}
                            formatter={(v:number, name:string) =>
                              name==="замовлень"
                                ? [`${v} зам. (${((v/totalOrds)*100).toFixed(1)}%)`, "Замовлення"]
                                : [`${v.toLocaleString("uk-UA")} ₴`, "Сер. чек"]
                            }
                          />
                          <Bar yAxisId="left" dataKey="замовлень" name="замовлень" radius={[4,4,0,0]} maxBarSize={40}
                            label={(props: Record<string,unknown>) => {
                              const entry = chartData[props.index as number];
                              if (!entry) return <text/>;
                              const cx = props.x as number + (props.width as number) / 2;
                              const cy = props.y as number - 5;
                              return <text key={props.index as number} x={cx} y={cy} textAnchor="middle" fontSize={9} fontWeight={700} fill={props.index === 0 ? "#004080" : "#6B7280"}>{entry.pctLabel}%</text>;
                            }}
                          >
                            {chartData.map((_,i)=>(
                              <Cell key={i} fill={BASE_BLUE[i] ?? BASE_BLUE[BASE_BLUE.length-1]}/>
                            ))}
                          </Bar>
                          <Line yAxisId="right" type="monotone" dataKey="середній" name="середній" stroke="#1E90FF" strokeWidth={2.5} dot={{ r:4, fill:"#1E90FF", strokeWidth:0 }} activeDot={{ r:5, fill:"#004080" }}/>
                        </ComposedChart>
                      </ResponsiveContainer>
                    );
                  })()}

                  {/* Column headers */}
                  <div style={{ display:"grid", gridTemplateColumns:"22px 1fr 56px 80px", gap:8, marginBottom:6, paddingBottom:6, borderBottom:`1px solid ${t.border}` }}>
                    <span/>
                    <span style={{ fontSize:9, fontWeight:700, letterSpacing:"0.06em", textTransform:"uppercase" as const, color:"#9CA3AF" }}>Місто</span>
                    <span style={{ fontSize:9, fontWeight:700, letterSpacing:"0.06em", textTransform:"uppercase" as const, color:"#9CA3AF", textAlign:"right" }}>Частка %</span>
                    <span style={{ fontSize:9, fontWeight:700, letterSpacing:"0.06em", textTransform:"uppercase" as const, color:"#9CA3AF", textAlign:"right" }}>Сер. чек ₴</span>
                  </div>

                  {/* Clickable bar rows */}
                  <div style={{ display:"flex", flexDirection:"column", gap:5 }}>
                    {cityTop.map((city, i) => {
                      const maxSalesPct = cityTop[0]?.salesPct ?? 1;
                      const barW      = (city.salesPct / maxSalesPct) * 100;
                      const isKyiv    = city.name.toLowerCase().startsWith("київ") || city.name.toLowerCase().startsWith("kyiv");
                      const active    = cityFilter === city.name;
                      const maxAvgChk = Math.max(...cityTop.map(c=>c.avgCheck));
                      const isBest    = city.avgCheck === maxAvgChk && maxAvgChk > 0;
                      const fmtChk    = (n:number) => Math.round(n).toLocaleString("uk-UA").replace(/,/g," ");
                      return (
                        <div key={city.name}>
                          <div
                            onClick={()=>setCityFilter(active ? null : city.name)}
                            style={{
                              display:"grid", gridTemplateColumns:"22px 1fr 56px 80px", alignItems:"center", gap:8,
                              cursor:"pointer", borderRadius:8, padding:"8px 8px", margin:"0 -8px",
                              background: active ? (t.dark?"rgba(0,64,128,0.18)":"rgba(0,64,128,0.06)") : "transparent",
                              border: active ? "1px solid rgba(0,64,128,0.25)" : "1px solid transparent",
                              transition:"background 0.14s, border-color 0.14s",
                            }}
                          >
                            {/* Rank */}
                            <span style={{ fontSize:11, fontWeight:700, color:active?"#004080":i===0?"#004080":"#9CA3AF", textAlign:"right" }}>{i+1}</span>
                            {/* City name + bar */}
                            <div>
                              <div style={{ display:"flex", alignItems:"center", gap:5, marginBottom:4 }}>
                                <span style={{ fontSize:12, fontWeight:700, color:active?"#004080":isKyiv?"#004080":t.text }}>{city.name}</span>
                                {isKyiv && kyivPct>30 && <span style={{ fontSize:8, fontWeight:700, padding:"1px 5px", borderRadius:4, background:"#DBEAFE", color:"#1E40AF" }}>Хаб</span>}
                                {active && <span style={{ fontSize:8, fontWeight:700, padding:"1px 5px", borderRadius:4, background:"rgba(0,64,128,0.10)", color:"#004080" }}>↓ Повернення</span>}
                              </div>
                              <div style={{ height:3, borderRadius:99, background:t.dark?"rgba(255,255,255,0.06)":"rgba(0,64,128,0.08)", overflow:"hidden" }}>
                                <div style={{ width:`${barW}%`, height:"100%", borderRadius:99, background: BASE_BLUE[i] ?? BASE_BLUE[BASE_BLUE.length-1], transition:"width 0.6s ease", opacity: active ? 1 : 0.85 }}/>
                              </div>
                            </div>
                            {/* Sales share % */}
                            <span style={{ fontSize:13, fontWeight:700, color:active?"#004080":i===0?"#004080":"#6B7280", textAlign:"right" }}>{city.salesPct.toFixed(1)}%</span>
                            {/* Avg check */}
                            <div style={{ textAlign:"right" }}>
                              <span style={{ fontSize:12, fontWeight:700, color:isBest?"#004080":t.text }}>{fmtChk(city.avgCheck)} ₴</span>
                              {isBest && <span style={{ marginLeft:3, fontSize:10 }}>⭐</span>}
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>

                  {/* Insight + city filter hint */}
                  <div style={{ marginTop:14, padding:"10px 14px", borderRadius:8, background:t.dark?"rgba(0,64,128,0.12)":"rgba(0,64,128,0.04)", border:"1px solid rgba(0,64,128,0.14)", display:"flex", alignItems:"flex-start", gap:8 }}>
                    <span style={{ fontSize:13, flexShrink:0, marginTop:1 }}>💡</span>
                    <span style={{ fontSize:11, color:t.text, lineHeight:1.55 }}>
                      {top2pct >= 40
                        ? <><strong style={{ color:"#004080" }}>{top2names}</strong>{` генерують ${top2pct.toFixed(0)}% продажів. Рекомендовано оптимізувати складські залишки під ці регіони.`}</>
                        : "Натисніть на місто, щоб побачити причини повернень саме звідти."}
                    </span>
                  </div>
                  {cityFilter && (
                    <div style={{ marginTop:8, padding:"8px 12px", borderRadius:7, background:t.dark?"rgba(0,82,255,0.10)":"rgba(0,82,255,0.06)", border:`1px solid ${t.blue}33`, display:"flex", alignItems:"center", gap:6 }}>
                      <span style={{ fontSize:11 }}>🔍</span>
                      <span style={{ fontSize:11, color:t.blue, fontWeight:600 }}>Топ причин повернень для <strong>{cityFilter}</strong> — дивіться нижче</span>
                    </div>
                  )}
                </div>
              );
            })()}

            {/* ── Rejection reasons + Advice — always in DOM, CSS-animated height to prevent jump ── */}
            <div className="orbit-rejection-section" data-empty={rejectionReasons.length===0 ? "true" : "false"}>
            <div style={{ display:"grid", gridTemplateColumns:rejectionReasons.length>0?"1fr 268px":"1fr", gap:10, alignItems:"start" }}>
              <div style={{ ...glassBase, padding:"20px 20px 14px" }}>
                <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:16 }}>
                  <div>
                    <p style={{ color:t.text, fontSize:14, fontWeight:700, margin:0, letterSpacing:"-0.02em" }}>
                      Топ Причин Повернень
                      {cityFilter && <span style={{ fontSize:11, fontWeight:600, color:t.blue, marginLeft:7 }}>· {cityFilter}</span>}
                    </p>
                    {cityFilter && (
                      <div style={{ display:"flex", alignItems:"center", gap:5, marginTop:3 }}>
                        <span style={{ fontSize:10, color:"#9CA3AF" }}>Фільтр по місту:</span>
                        <button
                          onClick={()=>setCityFilter(null)}
                          style={{ fontSize:10, fontWeight:700, padding:"1px 7px", borderRadius:4, border:`1px solid ${t.blue}44`, background:`${t.blue}10`, color:t.blue, cursor:"pointer" }}
                        >✕ Скинути</button>
                      </div>
                    )}
                  </div>
                  <div style={{ padding:"3px 10px", borderRadius:6, background:`${t.red}18`, border:`1px solid ${t.red}44`, fontSize:10, fontWeight:700, color:t.red }}>
                    {kpi!.returnRate.toFixed(1)}% відмов
                  </div>
                </div>
                <div style={{ display:"flex", flexDirection:"column", gap:9 }}>
                  {rejectionReasons.map((item, i) => {
                    const maxCount = rejectionReasons[0].count;
                    const pct = (item.count / maxCount) * 100;
                    const ofTotal = kpi!.refs > 0 ? (item.count / kpi!.refs) * 100 : 0;
                    const isHovered = rejTooltip?.reason === item.reason;
                    return (
                      <div
                        key={i}
                        style={{ display:"flex", alignItems:"center", gap:10, cursor:"pointer", borderRadius:6, padding:"3px 4px", margin:"-3px -4px", transition:"background 0.12s ease", background: isHovered?(t.dark?"rgba(0,82,255,0.1)":"rgba(0,82,255,0.05)"):"transparent" }}
                        onMouseEnter={e => setRejTooltip({ reason:item.reason, rect:e.currentTarget.getBoundingClientRect() })}
                        onMouseLeave={() => setRejTooltip(null)}
                      >
                        <span style={{ fontSize:9, fontWeight:700, color: isHovered?t.blue:t.dim, width:18, textAlign:"right", flexShrink:0, transition:"color 0.12s" }}>{i+1}</span>
                        <div style={{ flex:1, minWidth:0 }}>
                          <div style={{ display:"flex", justifyContent:"space-between", marginBottom:3 }}>
                            <span style={{ fontSize:11, color: isHovered?t.blue:t.text, fontWeight: isHovered?700:500, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", maxWidth:"70%", transition:"color 0.12s" }}>{item.reason}</span>
                            <span style={{ fontSize:11, fontWeight:700, color:t.red, flexShrink:0, marginLeft:8 }}>{item.count} <span style={{ fontSize:9, color:t.dim, fontWeight:400 }}>({ofTotal.toFixed(0)}%)</span></span>
                          </div>
                          <div style={{ height:6, borderRadius:3, background:t.dark?"rgba(255,255,255,0.06)":"rgba(0,0,0,0.07)", overflow:"hidden" }}>
                            <div style={{
                              height:"100%", borderRadius:3,
                              width:`${pct}%`,
                              background: isHovered ? t.blue : (i===0 ? t.red : `${t.red}${Math.round(80 - i * 6).toString(16).padStart(2,"0")}`),
                              transition:"width 0.4s ease, background 0.15s ease",
                            }}/>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* 💡 Порада щодо оптимізації — advice card, right column */}
              {rejectionReasons.length > 0 && (
                <div style={{ ...glassBase, padding:"18px 18px 16px", display:"flex", flexDirection:"column", gap:11, background: t.dark?"rgba(234,179,8,0.07)":"#FEFCE8", border:`1px solid ${t.dark?"rgba(234,179,8,0.22)":"#FDE68A"}` }}>
                  <div style={{ display:"flex", alignItems:"center", gap:7, marginBottom:2 }}>
                    <span style={{ fontSize:15 }}>💡</span>
                    <span style={{ fontSize:12, fontWeight:800, color:t.dark?"#FDE68A":"#1C1917", letterSpacing:"-0.02em" }}>Порада щодо оптимізації</span>
                  </div>

                  {returnAdvice.length > 0 ? returnAdvice.map((adv,i)=>(
                    <div key={i} style={{ borderRadius:8, padding:"10px 12px", background:t.dark?"rgba(234,179,8,0.1)":"rgba(253,230,138,0.45)", border:`1px solid ${t.dark?"rgba(234,179,8,0.18)":"#FCD34D"}` }}>
                      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:5 }}>
                        <div style={{ display:"flex", alignItems:"center", gap:5 }}>
                          <span style={{ fontSize:12 }}>{adv.icon}</span>
                          <span style={{ fontSize:11, fontWeight:800, color:t.dark?"#FDE68A":"#92400E" }}>{adv.title}</span>
                        </div>
                        <span style={{ fontSize:9, fontWeight:700, padding:"2px 6px", borderRadius:4, background:t.dark?"rgba(234,179,8,0.2)":"#FEF08A", color:t.dark?"#FDE68A":"#78350F", flexShrink:0, marginLeft:6 }}>{adv.badge}</span>
                      </div>
                      <p style={{ margin:0, fontSize:10, color:t.dark?"rgba(253,230,138,0.82)":"#44403C", lineHeight:1.55 }}>{adv.msg}</p>
                    </div>
                  )) : (
                    <div style={{ display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", flex:1, gap:8, padding:"18px 0" }}>
                      <span style={{ fontSize:22, opacity:0.3 }}>✅</span>
                      <span style={{ fontSize:11, color:t.dim, textAlign:"center", lineHeight:1.4 }}>Критичних відхилень<br/>не виявлено</span>
                    </div>
                  )}
                </div>
              )}
            </div>{/* /inner grid */}
            </div>{/* /orbit-rejection-section */}

            {/* Daily trend — locked 280px */}
            <ChartErrorBoundary t={t} label="Щоденний тренд">
            <div className="analytics-card analytics-card--daily orbit-fadein" style={{ ...glassBase, padding:"20px 20px 12px", animationDelay:"260ms" }}>
              {dailyTrend.length > 0 ? (<>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
                  <div>
                    <p style={{ color:t.text, fontSize:14, fontWeight:700, margin:0 }}>Щоденний Чистий Дохід — {dailyTrendMonthLabel}</p>
                    <p style={{ color:t.dim, fontSize:11, margin:"3px 0 0" }}>тільки фактичні дані · без прогнозів</p>
                  </div>
                </div>
                <ResponsiveContainer width="100%" height={200}>
                  <AreaChart data={dailyTrend} margin={{ top:4, right:8, left:8, bottom:4 }}>
                    <defs>
                      <linearGradient id="gDay" x1="0" y1="0" x2="0" y2="1">
                        <stop offset="0%" stopColor={t.blue} stopOpacity={t.dark?0.2:0.12}/>
                        <stop offset="100%" stopColor={t.blue} stopOpacity={0}/>
                      </linearGradient>
                    </defs>
                    <CartesianGrid strokeDasharray="1 0" stroke={t.dark?"rgba(255,255,255,0.04)":"rgba(0,0,0,0.05)"} vertical={false}/>
                    <XAxis dataKey="day" tick={{ fontSize:10, fill:t.sub }} tickLine={false} axisLine={false} label={{ value:"День місяця", position:"insideBottom", offset:-2, fill:t.dim, fontSize:9 }}/>
                    <YAxis tickFormatter={v=>fmt(v)} tick={{ fontSize:10, fill:t.dim }} tickLine={false} axisLine={false} width={90}/>
                    <Tooltip formatter={(v:number)=>fmt(v)} contentStyle={{ background:t.dark?"rgba(4,6,14,0.97)":"rgba(255,255,255,0.97)", border:`1px solid ${t.border}`, borderRadius:10, fontSize:12, boxShadow:"0 4px 24px rgba(0,0,0,0.08)" }}/>
                    <Area isAnimationActive={true} animationDuration={500} animationEasing="ease-out" type="monotone" dataKey="net" name="Дохід" stroke={t.blue} strokeWidth={2.5} fill="url(#gDay)" dot={{ r:3, fill:t.blue, strokeWidth:0 }} activeDot={{ r:5, fill:t.blue, strokeWidth:0 }}/>
                  </AreaChart>
                </ResponsiveContainer>
              </>) : (
                <div style={{ display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", height:"100%", minHeight:200, gap:8 }}>
                  <span style={{ fontSize:28, opacity:0.18 }}>📅</span>
                  <span style={{ fontSize:13, fontWeight:600, color:t.sub }}>Щоденний тренд</span>
                  <span style={{ fontSize:11, color:t.dim, textAlign:"center", maxWidth:260 }}>
                    Немає щоденних даних для обраного фільтра
                  </span>
                </div>
              )}
            </div>
            </ChartErrorBoundary>

            {/* ── Детальна аналітика ── */}
            <div style={{ display:"flex", alignItems:"center", gap:8, padding:"6px 0 2px" }}>
              <div style={{ width:4, height:16, borderRadius:2, background:t.blue }}/>
              <span style={{ color:t.text, fontSize:14, fontWeight:700 }}>Детальна аналітика</span>
            </div>

            {/* 💰 Top-3 products by gross revenue (сума замовлення) */}
            {topRevProducts.length > 0 && (
              <div className="orbit-fadein" style={{ ...glassBase, padding:"16px 22px 14px", animationDelay:"260ms" }}>
                <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:14 }}>
                  <div style={{ display:"flex", alignItems:"center", gap:7 }}>
                    <span style={{ fontSize:15 }}>💰</span>
                    <p style={{ color:t.text, fontSize:14, fontWeight:700, margin:0, letterSpacing:"-0.02em" }}>Топ товарів за виручкою</p>
                  </div>
                  <span style={{ fontSize:10, fontWeight:700, letterSpacing:"0.06em", textTransform:"uppercase" as const, color:"#9CA3AF" }}>Топ 3</span>
                </div>
                <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                  {(()=>{
                    const maxRev = topRevProducts[0]?.rev ?? 1;
                    const fmtRev = (n:number) => Math.round(n).toLocaleString("uk-UA").replace(/,/g," ");
                    const MEDALS = ["🥇","🥈","🥉"];
                    return topRevProducts.map((p, i) => (
                      <div key={p.name} style={{ display:"flex", alignItems:"center", gap:10 }}>
                        <span style={{ fontSize:14, flexShrink:0 }}>{MEDALS[i]}</span>
                        <div style={{ flex:1, minWidth:0 }}>
                          <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", gap:8, marginBottom:4 }}>
                            <span style={{ fontSize:12, fontWeight:700, color:t.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }} title={p.name}>{p.name}</span>
                            <span style={{ fontSize:13, fontWeight:700, color:i===0?"#004080":t.text, flexShrink:0 }}>{fmtRev(p.rev)} ₴</span>
                          </div>
                          <div style={{ height:4, borderRadius:99, background:t.dark?"rgba(255,255,255,0.06)":"rgba(0,64,128,0.08)", overflow:"hidden" }}>
                            <div style={{ width:`${(p.rev/maxRev)*100}%`, height:"100%", borderRadius:99, background: BASE_BLUE[i] ?? BASE_BLUE[BASE_BLUE.length-1], transition:"width 0.6s ease" }}/>
                          </div>
                        </div>
                      </div>
                    ));
                  })()}
                </div>
              </div>
            )}

            {/* top products + top customers */}
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>

              <div style={{ ...glassBase, overflow:"hidden" }}>
                <div style={{ padding:"14px 18px 10px", borderBottom:`1px solid ${t.border}` }}>
                  <p style={{ color:t.text, fontSize:13, fontWeight:600, margin:0 }}>Топ-10 Товарів за Чистим Доходом</p>
                </div>
                {topProducts.length>0 ? (
                  <div style={{ overflowX:"auto" }}>
                    <table style={{ width:"100%", fontSize:11, borderCollapse:"collapse" }}>
                      <thead><tr style={{ background:t.dark?"rgba(4,6,14,0.9)":t.in }}>
                        {["#","Назва, Модель","К-сть","Дохід ₴"].map(h=>(
                          <th key={h} style={{ padding:"8px 14px", textAlign:"left", fontWeight:600, fontSize:9, letterSpacing:"0.07em", textTransform:"uppercase", color:t.dim, borderBottom:`1px solid ${t.border}` }}>{h}</th>
                        ))}
                      </tr></thead>
                      <tbody>{topProducts.map((p,i)=>(
                        <tr key={i} style={{ borderBottom:`1px solid ${t.dark?"rgba(255,255,255,0.035)":"rgba(0,0,0,0.04)"}`, background:i%2===0?"transparent":(t.dark?"rgba(255,255,255,0.015)":"rgba(0,0,0,0.015)") }}>
                          <td style={{ padding:"8px 14px", color:t.dim, fontWeight:700, verticalAlign:"top" }}>{i+1}</td>
                          <td title={p.name} style={{ padding:"8px 14px", color:t.text, fontWeight:500, maxWidth:200, wordBreak:"break-word", whiteSpace:"normal", lineHeight:1.4 }}>{p.name}</td>
                          <td style={{ padding:"8px 14px", color:t.sub, verticalAlign:"top", whiteSpace:"nowrap" }}>{p.qty%1===0?p.qty.toFixed(0):p.qty.toFixed(1)}</td>
                          <td style={{ padding:"8px 14px", fontWeight:600, color:p.net>=0?t.em:t.red, verticalAlign:"top", whiteSpace:"nowrap" }}>{fmt(p.net)}</td>
                        </tr>
                      ))}</tbody>
                    </table>
                  </div>
                ) : (
                  <p style={{ color:t.dim, fontSize:12, textAlign:"center", padding:"22px 0" }}>Немає даних для поточного фільтра</p>
                )}
              </div>

              <div style={{ ...glassBase, overflow:"hidden" }}>
                <div style={{ padding:"14px 18px 10px", borderBottom:`1px solid ${t.border}` }}>
                  <p style={{ color:t.text, fontSize:13, fontWeight:600, margin:0 }}>Топ Покупців за Чистим Доходом</p>
                  <p style={{ color:t.dim, fontSize:10, margin:"2px 0 0" }}>Топ 20</p>
                </div>
                {topCustomers.length>0 ? (
                  <div style={{ overflowX:"auto", maxHeight:340, overflowY:"auto" }}>
                    <table style={{ width:"100%", fontSize:11, borderCollapse:"collapse" }}>
                      <thead style={{ position:"sticky", top:0, background:t.dark?"rgba(4,6,14,0.98)":t.in }}>
                        <tr>
                          {["#","ПІБ / Телефон","Замовлень","Витрачено ₴"].map(h=>(
                            <th key={h} style={{ padding:"8px 14px", textAlign:"left", fontWeight:600, fontSize:9, letterSpacing:"0.07em", textTransform:"uppercase", color:t.dim, borderBottom:`1px solid ${t.border}` }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>{topCustomers.map((c,i)=>(
                        <tr key={i} style={{ borderBottom:`1px solid ${t.dark?"rgba(255,255,255,0.035)":"rgba(0,0,0,0.04)"}`, background:i%2===0?"transparent":(t.dark?"rgba(255,255,255,0.015)":"rgba(0,0,0,0.015)") }}>
                          <td style={{ padding:"8px 14px", color:t.dim, fontWeight:700, verticalAlign:"top" }}>{i+1}</td>
                          <td title={c.name} style={{ padding:"8px 14px", color:t.text, fontWeight:500, maxWidth:180, wordBreak:"break-word", whiteSpace:"normal", lineHeight:1.4 }}>{c.name}</td>
                          <td style={{ padding:"8px 14px", color:t.sub, verticalAlign:"top" }}>{c.orders}</td>
                          <td style={{ padding:"8px 14px", fontWeight:600, color:c.net>=0?t.em:t.red, verticalAlign:"top", whiteSpace:"nowrap" }}>{fmt(c.net)}</td>
                        </tr>
                      ))}</tbody>
                    </table>
                  </div>
                ) : (
                  <p style={{ color:t.dim, fontSize:12, textAlign:"center", padding:"22px 0" }}>Немає даних для поточного фільтра</p>
                )}
              </div>
            </div>

            {/* ── Product analytics full table ── */}
            {(()=>{
              const totalNet = allProducts.reduce((s,p)=>s+p.net, 0);
              const maxNet   = Math.max(...allProducts.map(p=>Math.abs(p.net)), 1);

              const SortTh = ({col,label,align="left"}:{col:"qty"|"net"|"name";label:string;align?:string})=>{
                const active = productSort.col===col;
                const toggle = ()=>setProductSort(s=>s.col===col?{col,dir:s.dir==="desc"?"asc":"desc"}:{col,dir:"desc"});
                return (
                  <th onClick={toggle} style={{ padding:"9px 14px", textAlign:align as "left"|"right", fontWeight:600, fontSize:9, letterSpacing:"0.07em", textTransform:"uppercase", color:active?t.blue:t.dim, borderBottom:`1px solid ${t.border}`, cursor:"pointer", whiteSpace:"nowrap", userSelect:"none" }}>
                    {label}{active ? (productSort.dir==="desc"?" ↓":" ↑") : " ↕"}
                  </th>
                );
              };

              return (
                <div style={{ ...glassBase, overflow:"hidden" }}>
                  {/* header + search */}
                  <div style={{ padding:"14px 18px 12px", borderBottom:`1px solid ${t.border}`, display:"flex", alignItems:"center", justifyContent:"space-between", gap:12, flexWrap:"wrap" }}>
                    <div>
                      <p style={{ color:t.text, fontSize:13, fontWeight:700, margin:0 }}>
                        Аналітика Товарів
                        <span style={{ marginLeft:10, fontSize:11, fontWeight:400, color:t.dim }}>
                          {allProducts.length} позицій
                        </span>
                      </p>
                      <p style={{ color:t.dim, fontSize:11, margin:"2px 0 0" }}>
                        Реагує на всі активні фільтри
                      </p>
                    </div>
                    <div style={{ display:"flex", gap:8, alignItems:"center" }}>
                      <div style={{ position:"relative", display:"flex", alignItems:"center" }}>
                        <Search size={12} style={{ position:"absolute", left:10, color:t.dim, pointerEvents:"none" }}/>
                        <input
                          value={productSearch}
                          onChange={e=>setProductSearch(e.target.value)}
                          placeholder="Пошук моделі…"
                          style={{ background:t.in, border:`1px solid ${t.border}`, borderRadius:8, padding:"6px 28px 6px 28px", color:t.text, fontSize:12, outline:"none", width:200 }}
                        />
                        {productSearch && (
                          <button onClick={()=>setProductSearch("")} style={{ position:"absolute", right:8, background:"none", border:"none", cursor:"pointer", color:t.dim, display:"flex" }}>
                            <X size={10}/>
                          </button>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* table */}
                  {allProducts.length>0 ? (
                    <div style={{ overflowX:"auto", maxHeight:440, overflowY:"auto" }}>
                      <table style={{ width:"100%", fontSize:11, borderCollapse:"collapse" }}>
                        <thead style={{ position:"sticky", top:0, zIndex:10, background:t.dark?"rgba(4,6,14,0.98)":t.in }}>
                          <tr>
                            <th style={{ padding:"9px 14px", textAlign:"left", fontWeight:600, fontSize:9, letterSpacing:"0.07em", textTransform:"uppercase", color:t.dim, borderBottom:`1px solid ${t.border}`, whiteSpace:"nowrap", width:36 }}>#</th>
                            <SortTh col="name" label="Назва, Модель"/>
                            <SortTh col="qty"  label="Кількість" align="right"/>
                            <SortTh col="net"  label="Чистий дохід" align="right"/>
                            <th style={{ padding:"9px 14px", borderBottom:`1px solid ${t.border}`, width:"22%", minWidth:120 }}/>
                          </tr>
                        </thead>
                        <tbody>
                          {allProducts.map((p,i)=>{
                            const pct      = totalNet !== 0 ? Math.round(Math.abs(p.net)/Math.abs(totalNet)*100) : 0;
                            const barPct   = Math.round(Math.abs(p.net)/maxNet*100);
                            const barCol   = p.net>=0 ? t.em : t.red;
                            const refRate  = p.rows > 0 ? (p.refs / p.rows) * 100 : 0;
                            const highRef  = refRate > 30 && p.refs > 0;
                            return (
                              <tr key={i} style={{ borderBottom:`1px solid ${t.dark?"rgba(255,255,255,0.03)":"rgba(0,0,0,0.04)"}`, background: highRef ? (t.dark?"rgba(226,149,120,0.05)":"rgba(226,149,120,0.04)") : i%2===0?"transparent":(t.dark?"rgba(255,255,255,0.015)":"rgba(0,0,0,0.012)") }}>
                                <td style={{ padding:"8px 14px", color:t.dim, fontSize:10, fontWeight:600 }}>{i+1}</td>
                                <td style={{ padding:"8px 14px", minWidth:160 }}>
                                  <div style={{ display:"flex", alignItems:"flex-start", gap:5, flexWrap:"wrap" }}>
                                    <span title={p.name} style={{ color:t.text, fontWeight:500, wordBreak:"break-word", whiteSpace:"normal", lineHeight:1.4 }}>{p.name}</span>
                                    {highRef && (
                                      <span title={`Відмов: ${p.refs} з ${p.rows} (${refRate.toFixed(0)}%)`} style={{ flexShrink:0, display:"inline-flex", alignItems:"center", gap:3, padding:"1px 6px", borderRadius:5, background:"rgba(226,149,120,0.18)", border:`1px solid ${t.red}70`, fontSize:9, fontWeight:700, color:t.red, whiteSpace:"nowrap" }}>
                                        ⚠ {refRate.toFixed(0)}% відмов
                                      </span>
                                    )}
                                  </div>
                                </td>
                                <td style={{ padding:"8px 14px", color:t.sub, textAlign:"right", fontVariantNumeric:"tabular-nums" }}>{p.qty%1===0?p.qty.toFixed(0):p.qty.toFixed(1)}</td>
                                <td style={{ padding:"8px 14px", textAlign:"right", fontWeight:700, color:p.net>=0?t.em:t.red, fontVariantNumeric:"tabular-nums", whiteSpace:"nowrap" }}>
                                  {fmt(p.net)}
                                  <span style={{ marginLeft:6, fontSize:9, fontWeight:400, color:t.dim }}>{pct}%</span>
                                </td>
                                <td style={{ padding:"8px 14px 8px 10px" }}>
                                  <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                                    <div style={{ flex:1, height:5, borderRadius:3, background:t.dark?"rgba(255,255,255,0.07)":"rgba(0,0,0,0.07)", overflow:"hidden" }}>
                                      <div style={{ width:`${barPct}%`, height:"100%", borderRadius:3, background:barCol, transition:"width 0.3s ease" }}/>
                                    </div>
                                  </div>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    <p style={{ color:t.dim, fontSize:12, textAlign:"center", padding:"28px 0" }}>
                      {productSearch ? `Нічого не знайдено за «${productSearch}»` : "Немає даних для поточного фільтра"}
                    </p>
                  )}
                </div>
              );
            })()}

            {/* ── АНАЛІТИКА КЛІЄНТІВ ─────────────────────────────────── */}
            {customerInsights && customerInsights.total > 0 && (
              <>
                {/* Section header */}
                <div style={{ display:"flex", alignItems:"center", gap:8, padding:"6px 0 2px" }}>
                  <div style={{ width:4, height:16, borderRadius:2, background:t.blue }}/>
                  <span style={{ color:t.text, fontSize:14, fontWeight:700 }}>Аналітика клієнтів</span>
                  <span style={{ fontSize:11, color:t.dim }}>· {customerInsights.total.toLocaleString()} унікальних</span>
                </div>

                {/* ── Row 1: Metric cards ── */}
                <div style={{ display:"grid", gridTemplateColumns:`repeat(${topCities.length>0?4:3},1fr)`, gap:14 }}>

                  {/* AOV card */}
                  <div style={{ ...glassBase, border:`1px solid ${t.blue}44`, padding:"20px 20px", display:"flex", flexDirection:"column", gap:7 }}>
                    <span style={{ fontSize:10, fontWeight:700, letterSpacing:"0.1em", textTransform:"uppercase", color:t.dim }}>Середній чек (AOV)</span>
                    <span style={{ fontSize:28, fontWeight:900, color:t.blue, lineHeight:1 }}>
                      {kpi && kpi.orders > 0 ? fmt(kpi.net / kpi.orders) : "—"}
                    </span>
                    <span style={{ fontSize:11, color:t.dim }}>
                      {kpi && kpi.orders > 0 ? `${fmt(kpi.net)} ÷ ${kpi.orders.toLocaleString()} замовлень` : "Немає даних"}
                    </span>
                  </div>

                  {/* Retention card */}
                  <div style={{ ...glassBase, border:`1px solid ${t.em}44`, padding:"20px 20px", display:"flex", flexDirection:"column", gap:7 }}>
                    <span style={{ fontSize:10, fontWeight:700, letterSpacing:"0.1em", textTransform:"uppercase", color:t.dim }}>Утримання клієнтів</span>
                    <div style={{ display:"flex", alignItems:"flex-end", gap:10 }}>
                      <div>
                        <div style={{ fontSize:22, fontWeight:900, color:t.em, lineHeight:1 }}>{customerInsights.returning}</div>
                        <div style={{ fontSize:10, color:t.dim, marginTop:2 }}>Постійні (2+ замовлень)</div>
                      </div>
                      <div style={{ width:1, background:t.border, alignSelf:"stretch" }}/>
                      <div>
                        <div style={{ fontSize:22, fontWeight:900, color:t.sub, lineHeight:1 }}>{customerInsights.newC}</div>
                        <div style={{ fontSize:10, color:t.dim, marginTop:2 }}>Нові (1 замовлення)</div>
                      </div>
                    </div>
                    {/* Mini retention bar */}
                    {customerInsights.total > 0 && (
                      <div>
                        <div style={{ display:"flex", justifyContent:"space-between", marginBottom:3 }}>
                          <span style={{ fontSize:9, color:t.em, fontWeight:700 }}>
                            {((customerInsights.returning/customerInsights.total)*100).toFixed(0)}% постійних
                          </span>
                          <span style={{ fontSize:9, color:t.dim }}>
                            {((customerInsights.newC/customerInsights.total)*100).toFixed(0)}% нових
                          </span>
                        </div>
                        <div style={{ height:6, borderRadius:3, background:t.dark?"rgba(255,255,255,0.07)":"rgba(0,0,0,0.07)", overflow:"hidden", display:"flex" }}>
                          <div style={{ width:`${(customerInsights.returning/customerInsights.total)*100}%`, background:t.em, borderRadius:"3px 0 0 3px", transition:"width 0.4s ease" }}/>
                          <div style={{ flex:1, background:t.dark?"rgba(255,255,255,0.12)":"rgba(0,0,0,0.1)" }}/>
                        </div>
                      </div>
                    )}
                  </div>

                  {/* Avg orders per returning customer */}
                  <div style={{ ...glassBase, border:`1px solid ${t.amb}44`, padding:"20px 20px", display:"flex", flexDirection:"column", gap:7 }}>
                    <span style={{ fontSize:10, fontWeight:700, letterSpacing:"0.1em", textTransform:"uppercase", color:t.dim }}>Глибина покупок</span>
                    <span style={{ fontSize:28, fontWeight:900, color:t.amb, lineHeight:1 }}>
                      {customerInsights.returning > 0
                        ? (customerInsights.vip.filter(v=>v.orders>=2).reduce((s,v)=>s+v.orders,0)/customerInsights.returning).toFixed(1)
                        : "—"}
                    </span>
                    <span style={{ fontSize:11, color:t.dim }}>Сер. замовлень / постійний</span>
                  </div>

                  {/* Top city (if detected) */}
                  {topCities.length>0 && (
                    <div style={{ ...glassBase, border:`1px solid ${t.blue}33`, padding:"20px 20px", display:"flex", flexDirection:"column", gap:8 }}>
                      <span style={{ fontSize:10, fontWeight:700, letterSpacing:"0.1em", textTransform:"uppercase", color:t.dim }}>
                        Топ-5 міст <span style={{ fontSize:9, fontWeight:400, textTransform:"none" }}>· {fileData.cols.city}</span>
                      </span>
                      {topCities.map((item, i) => {
                        const maxC = topCities[0].count;
                        return (
                          <div key={i} style={{ display:"flex", alignItems:"center", gap:8 }}>
                            <span style={{ fontSize:9, color:t.dim, width:12, textAlign:"right", flexShrink:0 }}>{i+1}</span>
                            <div style={{ flex:1, minWidth:0 }}>
                              <div style={{ display:"flex", justifyContent:"space-between", marginBottom:2 }}>
                                <span style={{ fontSize:11, color:t.text, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{String(item.city || item)}</span>
                                <span style={{ fontSize:11, fontWeight:700, color:t.blue, flexShrink:0, marginLeft:6 }}>{item.count}</span>
                              </div>
                              <div style={{ height:4, borderRadius:2, background:t.dark?"rgba(255,255,255,0.06)":"rgba(0,0,0,0.07)", overflow:"hidden" }}>
                                <div style={{ width:`${(item.count/maxC)*100}%`, height:"100%", borderRadius:2, background:t.blue }}/>
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>

                {/* ── VIP Table ── */}
                {customerInsights.vip.length > 0 && (
                  <div style={{ ...glassBase, overflow:"hidden" }}>
                    <div style={{ padding:"14px 18px 10px", borderBottom:`1px solid ${t.border}`, display:"flex", alignItems:"center", gap:10 }}>
                      <div>
                        <p style={{ color:t.text, fontSize:13, fontWeight:700, margin:0 }}>VIP Клієнти — Топ 10</p>
                        <p style={{ color:t.dim, fontSize:10, margin:"2px 0 0" }}>
                          Найбільший чистий дохід · телефони частково приховані
                        </p>
                      </div>
                      <div style={{ marginLeft:"auto", padding:"3px 10px", borderRadius:6, background:`${t.em}14`, border:`1px solid ${t.em}33`, fontSize:10, fontWeight:700, color:t.em }}>
                        {customerInsights.vip.length} клієнтів
                      </div>
                    </div>
                    <div style={{ overflowX:"auto", maxHeight:360, overflowY:"auto" }}>
                      <table style={{ width:"100%", fontSize:11, borderCollapse:"collapse" }}>
                        <thead style={{ position:"sticky", top:0, background:t.dark?"rgba(4,6,14,0.98)":t.in }}>
                          <tr>
                            {["#","ПІБ / Ключ","Телефон","Замовлень","Сума ₴","Тип"].map((h,hi)=>(
                              <th key={h} style={{ padding:"8px 14px", textAlign: hi>=3?"right":"left", fontWeight:600, fontSize:9, letterSpacing:"0.07em", textTransform:"uppercase", color:t.dim, borderBottom:`1px solid ${t.border}`, whiteSpace:"nowrap" }}>{h}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {customerInsights.vip.map((c, i) => {
                            const isReturning = c.orders >= 2;
                            const displayPhone = c.phone ? maskPhone(c.phone) : (isPhoneString(c.key) ? maskPhone(c.key) : "—");
                            const displayName  = c.displayName || (isPhoneString(c.key) ? "—" : c.key);
                            return (
                              <tr key={i} style={{ borderBottom:`1px solid ${t.dark?"rgba(255,255,255,0.035)":"rgba(0,0,0,0.04)"}`, background: i===0 ? (t.dark?"rgba(255,215,0,0.04)":"rgba(255,215,0,0.06)") : i%2===0?"transparent":(t.dark?"rgba(255,255,255,0.015)":"rgba(0,0,0,0.015)") }}>
                                <td style={{ padding:"8px 14px", color: i===0?t.em:t.dim, fontWeight:700, fontSize:10 }}>
                                  {i===0 ? "👑" : i+1}
                                </td>
                                <td title={displayName} style={{ padding:"8px 14px", color:t.text, fontWeight:500, maxWidth:180, wordBreak:"break-word", whiteSpace:"normal", lineHeight:1.4 }}>{displayName}</td>
                                <td style={{ padding:"8px 14px", color:t.sub, fontFamily:"monospace", letterSpacing:"0.04em" }}>{displayPhone}</td>
                                <td style={{ padding:"8px 14px", color:t.sub, textAlign:"right", fontVariantNumeric:"tabular-nums" }}>{c.orders}</td>
                                <td style={{ padding:"8px 14px", textAlign:"right", fontWeight:700, color:c.net>=0?t.em:t.red, fontVariantNumeric:"tabular-nums", whiteSpace:"nowrap" }}>{fmt(c.net)}</td>
                                <td style={{ padding:"8px 14px", textAlign:"right" }}>
                                  <span style={{ padding:"2px 8px", borderRadius:5, fontSize:9, fontWeight:700,
                                    background: isReturning ? `${t.em}18` : `${t.blue}14`,
                                    border: `1px solid ${isReturning ? t.em : t.blue}44`,
                                    color: isReturning ? t.em : t.blue,
                                  }}>
                                    {isReturning ? "Постійний" : "Новий"}
                                  </span>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </>
            )}

            {/* data table */}
            <div style={{ ...glassBase, overflow:"hidden" }}>
              <div style={{ padding:"12px 18px", borderBottom:`1px solid ${t.border}`, display:"flex", alignItems:"center", justifyContent:"space-between", gap:12 }}>
                <p style={{ color:t.text, fontSize:13, fontWeight:600, margin:0 }}>
                  Всі дані <span style={{ color:t.dim, fontSize:11, fontWeight:400 }}>{tableRows.length} з {filtered.length}</span>
                </p>
                <div style={{ display:"flex", gap:8, alignItems:"center" }}>
                  <div style={{ position:"relative", display:"flex", alignItems:"center" }}>
                    <Search size={12} style={{ position:"absolute", left:10, color:t.dim, pointerEvents:"none" }}/>
                    <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Пошук…"
                      style={{ background:t.in, border:`1px solid ${t.border}`, borderRadius:8, padding:"5px 28px 5px 28px", color:t.text, fontSize:12, outline:"none", width:170 }}/>
                    {search && <button onClick={()=>setSearch("")} style={{ position:"absolute", right:8, background:"none", border:"none", cursor:"pointer", color:t.dim, display:"flex" }}><X size={10}/></button>}
                  </div>
                  <button onClick={()=>{setBrandFilter("All");setMonthFilter("All");setYearFilter("All");setSearch("");}} style={{ padding:"5px 10px", borderRadius:7, background:t.in, border:`1px solid ${t.border}`, color:t.sub, fontSize:11, cursor:"pointer", display:"flex", alignItems:"center", gap:5 }}>
                    <RefreshCw size={10}/> Скинути
                  </button>
                </div>
              </div>
              <div style={{ overflowX:"auto", maxHeight:380, overflowY:"auto" }}>
                <table style={{ width:"100%", fontSize:11, borderCollapse:"collapse" }}>
                  <thead style={{ position:"sticky", top:0, zIndex:10 }}>
                    <tr style={{ background:t.dark?"rgba(4,6,14,0.98)":t.in }}>
                      {fileData.columns.map(col=>{
                        const isFin=[fileData.cols.revenue,fileData.cols.delivery,fileData.cols.commission,fileData.cols.debt].includes(col);
                        return <th key={col} style={{ padding:"8px 13px", textAlign:"left", fontWeight:600, whiteSpace:"nowrap", color:isFin?t.blue:t.dim, borderBottom:`1px solid ${t.border}`, letterSpacing:"0.04em", fontSize:9, textTransform:"uppercase" }}>{col}</th>;
                      })}
                    </tr>
                  </thead>
                  <tbody>
                    {tableRows.map((row,i)=>{
                      const ref=isRefusal(row,fileData.cols);
                      return (
                        <tr key={i} style={{ background:ref?(t.dark?"rgba(244,63,94,0.04)":"rgba(220,38,38,0.03)"):i%2===0?"transparent":(t.dark?"rgba(255,255,255,0.015)":"rgba(0,0,0,0.015)"), borderBottom:`1px solid ${t.dark?"rgba(255,255,255,0.035)":"rgba(0,0,0,0.04)"}` }}>
                          {fileData.columns.map(col=>{
                            const v=row[col];
                            const isFin=[fileData.cols.revenue,fileData.cols.delivery,fileData.cols.commission,fileData.cols.debt].includes(col);
                            const isRef2=col===fileData.cols.status&&ref;
                            const display=isFin&&typeof v==="number"?fmt(v as number):v==null?"":String(v);
                            return <td key={col} style={{ padding:"7px 13px", whiteSpace:"nowrap", color:isRef2?t.red:isFin?t.text:t.sub, fontWeight:isFin?500:400 }}>{display===""?<span style={{ color:t.dim, fontStyle:"italic" }}>—</span>:display}</td>;
                          })}
                        </tr>
                      );
                    })}
                    {tableRows.length===0 && <tr><td colSpan={fileData.columns.length} style={{ textAlign:"center", padding:36, color:t.dim }}>Рядків не знайдено</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>

          {/* ── Last Updated footer ── */}
          {uploadedAt && (
            <div style={{ display:"flex", alignItems:"center", justifyContent:"flex-end", gap:6, marginTop:18, paddingTop:12, borderTop:`1px solid ${t.border}` }}>
              <span style={{ fontSize:11, color:t.dim }}>🕐</span>
              <span style={{ fontSize:11, color:t.dim }}>
                Дані оновлено:{" "}
                <strong style={{ color:t.sub, fontWeight:600 }}>
                  {uploadedAt.toLocaleDateString("uk-UA", { day:"2-digit", month:"long", year:"numeric" })},{" "}
                  {uploadedAt.toLocaleTimeString("uk-UA", { hour:"2-digit", minute:"2-digit" })}
                </strong>
              </span>
            </div>
          )}

          </div>
        )}
        </div>{/* /main content */}
      </div>{/* /flex row */}

      <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display:"none" }} onChange={handleChange}/>

      {/* ── Rejection-reason tooltip portal ── */}
      {rejTooltip && createPortal((() => {
        const adv = getReasonAdvice(rejTooltip.reason);
        const r   = rejTooltip.rect;
        const TW  = 280; // tooltip width
        // Position to the right; flip left if not enough room
        const flipLeft = r.right + TW + 20 > window.innerWidth;
        const leftPx   = flipLeft ? r.left - TW - 12 : r.right + 12;
        const topPx    = r.top + r.height / 2;
        return (
          <div
            className="orbit-tip-pop"
            onMouseEnter={()=>{/* keep open when cursor moves onto tooltip */}}
            style={{
              position:"fixed",
              left:leftPx,
              top:topPx,
              transform:"translateY(-50%)",
              width:TW,
              zIndex:999999,
              pointerEvents:"none",
              borderRadius:12,
              background:"#1C1917",
              border:"1.5px solid #0052FF",
              boxShadow:"0 8px 32px rgba(0,0,0,0.38), 0 0 0 1px rgba(0,82,255,0.12)",
              padding:"13px 15px 12px",
              fontFamily:"'Inter',-apple-system,sans-serif",
            }}
          >
            {/* category badge + emoji */}
            <div style={{ display:"flex", alignItems:"center", gap:6, marginBottom:7 }}>
              <span style={{ fontSize:14, lineHeight:1 }}>{adv.emoji}</span>
              <span style={{ fontSize:9, fontWeight:800, letterSpacing:"0.07em", color:"#0052FF", background:"rgba(0,82,255,0.18)", padding:"2px 7px", borderRadius:4 }}>{adv.cat}</span>
            </div>
            {/* reason label */}
            <p style={{ margin:"0 0 6px", fontSize:10, color:"rgba(255,255,255,0.5)", fontStyle:"italic", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{rejTooltip.reason}</p>
            {/* advice text */}
            <p style={{ margin:0, fontSize:11, color:"#FFFFFF", lineHeight:1.55, fontWeight:500 }}>{adv.text}</p>
            {/* pointer arrow */}
            <div style={{
              position:"absolute",
              top:"50%",
              [flipLeft ? "right" : "left"]: -7,
              transform:"translateY(-50%)",
              width:0, height:0,
              borderTop:"6px solid transparent",
              borderBottom:"6px solid transparent",
              [flipLeft ? "borderLeft" : "borderRight"]:"7px solid #0052FF",
            }}/>
          </div>
        );
      })(), document.body)}
    </div>
  );
}
