import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { motion } from "framer-motion";
import {
  AlertCircle,
  BarChart3,
  CalendarDays,
  CheckCircle2,
  ClipboardList,
  Database,
  Download,
  FileUp,
  RefreshCw,
  Search,
  Settings,
  ShieldAlert,
  Split,
  Tag,
  Trash2,
  Upload,
  Undo2,
  Users,
} from "lucide-react";

// UI (lightweight shadcn-compatible API)
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Badge } from "@/components/ui/badge";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Textarea } from "@/components/ui/textarea";
import { Separator } from "@/components/ui/separator";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import {
  BarChart,
  Bar,
  Cell,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  CartesianGrid,
  Legend,
} from "recharts";

/**
 * Supplier Risk Ops Dashboard
 *
 * End-to-end:
 * - Import XLSX -> normalize -> build Supplier Master (1 supplier code = 1 canonical name)
 * - Dashboard KPIs, funnel, category & country analytics, click-to-drilldown
 * - Supplier Workbench: filters, quick actions, status tracking
 * - Bulk updates: risk/non-risk, contract status, canonical name
 * - Duplicate detection (same code, different names) + recommended canonical
 * - Governance: Settings layer (supplier overrides + category/family rules)
 * - Activity Log of every action (import, edits, bulk updates, decisions)
 * - Notes / action calendar (who / date / what done / what next)
 * - Local persistence (localStorage) + Export/Import JSON snapshot
 */

// -----------------------------
// Types
// -----------------------------

type RiskLevel = "High" | "Non-risk" | "Unknown";
type ContractStatus = "Not sent" | "Sent" | "Signed" | "Review" | "Not compliant" | "N/A";
type BarFillStyle = "solid" | "diagonal" | "dots";
type BarPaletteKey = "country" | "category" | "riskLevel" | "contractStatus" | "funnelStage";
type BarChartSettings = {
  fillStyle: BarFillStyle;
  palettes: Record<BarPaletteKey, Record<string, string>>;
  metricColors: {
    highRisk: string;
    totalSuppliers: string;
    spend: string;
  };
  rulesNote?: string;
};

type EmailAutomationSettings = {
  initialFollowUpDays: number;
  followUpDays: number;
};

type FactRow = {
  supplierName?: string;
  supplierCode?: string;
  country?: string;
  entity?: string;
  year?: string | number;
  category?: string; // Europe Catman category
  family?: string; // FAMILY NAME
  subFamily?: string; // SUB-FAMILY NAME
  spend?: number;
  po?: number;
  riskFlag?: string; // Yes/No/#REF!/blank
  status?: string; // Agreement signed/sent/Asessed/etc
  completion?: string; // Completion status
};

type SupplierOverride = {
  canonicalName?: string;
  risk?: RiskLevel;
  contractStatus?: ContractStatus;
  evaluated?: boolean;
  note?: string;
};

type AppSettings = {
  /** Guardrail: treat a category as "dominant" if its spend share >= this threshold */
  dominanceShareThreshold: number; // 0..1
  /** Guardrail: flag supplier as multi-category when 2nd category share >= this threshold */
  multiCategorySecondShareThreshold: number; // 0..1
  /** Guardrail: "in-scope" spend threshold used for Unknown-in-scope checks */
  scopeSpendThreshold: number;
  /** Data-quality: flag suppliers with missing subFamily when abs spend exceeds this */
  missingSubfamilySpendThreshold: number;
  /** UI palette + styling for bar charts */
  barChartSettings: BarChartSettings;
  /** Email follow-up automation rules */
  emailAutomation: EmailAutomationSettings;
  /** Bulk contact defaults for outreach */
  bulkContactName: string;
  bulkContactEmail: string;
  statusPalette: {
    risk: Record<RiskLevel, string>;
    contract: Record<ContractStatus, string>;
    supplierContacts?: Record<string, string>;
};
  supplierContacts?: Record<string, string>;
};


type UndoBatch = {
  id: string;
  ts: string;
  actor: string;
  action: "risk" | "contractStatus" | "canonicalName" | "categoryFastAction" | "evaluated" | "bulkMark";
  summary: string;
  items: { code: string; prev?: SupplierOverride }[];
};

type CategoryRule = {
  key: string; // category or family name
  scope: "Category" | "Family";
  defaultRisk: RiskLevel; // policy layer
  comment?: string;
};

type WorkSheet = any;

type SupplierMaster = {
  code: string;
  canonicalName: string;
  nameCandidates: { name: string; count: number }[];
  countries: string[];
  categories: string[];
  families: string[];
  subFamilies: string[];
  categoryShares: { category: string; spend: number; share: number }[];
  dominantCategory?: string;
  dominantCategoryShare?: number;
  secondCategory?: string;
  secondCategoryShare?: number;
  multiCategory: boolean;
  totalSpend: number;
  totalPO: number;
  risk: RiskLevel;
  evaluated: boolean;
  contractNeeded: boolean;
  contractStatus: ContractStatus;
  lastTouch?: string; // from activity log
};

type IssueSeverity = "Critical" | "High" | "Medium" | "Low";
type IssueType =
  | "CODE_MANY_NAMES"
  | "NAME_MANY_CODES"
  | "MULTI_CATEGORY_EXPOSURE"
  | "RISK_UNKNOWN_IN_SCOPE"
  | "MISSING_SUBFAMILY_HIGH_SPEND"
  | "POLICY_SUPPRESSED_BY_GUARDRAIL";

type IssueInstance = {
  id: string;
  type: IssueType;
  severity: IssueSeverity;
  title: string;
  code?: string;
  supplierName?: string;
  spendAffected?: number;
  details?: any;
};

type LogEntry = {
  id: string;
  ts: string; // ISO
  actor: string;
  type:
    | "IMPORT"
    | "BULK_UPDATE"
    | "SUPPLIER_OVERRIDE"
    | "CANONICAL_DECISION"
    | "STATUS_UPDATE"
    | "CATEGORY_RULE"
    | "SETTINGS"
    | "TASK"
    | "RESET";
  summary: string;
  details?: any;
};

type Task = {
  id: string;
  date: string;
  createdAt: string;
  owner: string;
  title: string;
  note?: string;
  supplierCode?: string;
  supplierName?: string;
  contactName?: string;
  contactEmail?: string;
  dueDate?: string;
  kind?: "general" | "email" | "contractFollowup";
  followUpStage?: "initial" | "followup" | "urgent";
  status: "todo" | "done";
};

type EvaluationChoice = "Evaluated" | "Not evaluated" | "Clear";
type BulkRiskChoice = RiskLevel | "No change";
type BulkEvalChoice = EvaluationChoice | "No change";

// -----------------------------
// Local persistence
// -----------------------------

const LS_KEY = "supplier-risk-ops-dashboard:v1";

const RISK_LEVEL_OPTIONS: RiskLevel[] = ["High", "Non-risk", "Unknown"];
const CONTRACT_STATUS_OPTIONS: ContractStatus[] = ["Not sent", "Sent", "Signed", "Review", "Not compliant", "N/A"];
const DEFAULT_RISK_PALETTE: Record<RiskLevel, string> = {
  High: "#fca5a5",
  "Non-risk": "#86efac",
  Unknown: "#e5e7eb",
};
const DEFAULT_CONTRACT_PALETTE: Record<ContractStatus, string> = {
  "Not sent": "#fcd34d",
  Sent: "#fde68a",
  Signed: "#86efac",
  Review: "#93c5fd",
  "Not compliant": "#f87171",
  "N/A": "#e5e7eb",
};

const DEFAULT_SETTINGS: AppSettings = {
  dominanceShareThreshold: 0.8,
  multiCategorySecondShareThreshold: 0.12,
  scopeSpendThreshold: 1_000_000,
  missingSubfamilySpendThreshold: 250_000,
  barChartSettings: {
    fillStyle: "solid",
    palettes: {
      country: {},
      category: {},
      riskLevel: { ...DEFAULT_RISK_PALETTE },
      contractStatus: { ...DEFAULT_CONTRACT_PALETTE },
      funnelStage: {},
    },
    metricColors: {
      highRisk: "#fca5a5",
      totalSuppliers: "#93c5fd",
      spend: "#60a5fa",
    },
    rulesNote: "",
  },
  emailAutomation: {
    initialFollowUpDays: 14,
    followUpDays: 7,
  },
  bulkContactName: "",
  bulkContactEmail: "",
  statusPalette: {
    risk: DEFAULT_RISK_PALETTE,
    contract: DEFAULT_CONTRACT_PALETTE,
  },
  supplierContacts: {}, // в†ђ Р”РћР‘РђР’РРўР¬ Р­РўРЈ РЎРўР РћРљРЈ
};

const RISK_LEVELS: RiskLevel[] = ["High", "Non-risk", "Unknown"];
const CONTRACT_STATUSES: ContractStatus[] = ["Not sent", "Sent", "Signed", "Review", "Not compliant", "N/A"];

const DEMO_FACT_ROWS: FactRow[] = [
  {
    supplierName: "Alpine Logistics GmbH",
    supplierCode: "AT-1001",
    country: "Austria",
    entity: "Vienna Ops",
    year: 2024,
    category: "Logistics - Ground",
    family: "Transport",
    subFamily: "Freight",
    spend: 2_450_000,
    po: 120,
    riskFlag: "Yes",
    status: "Sent",
    completion: "Done",
  },
  {
    supplierName: "Alpine Logistics GmbH",
    supplierCode: "AT-1001",
    country: "Austria",
    entity: "Vienna Ops",
    year: 2024,
    category: "Logistics - Air",
    family: "Transport",
    subFamily: "Air Freight",
    spend: 620_000,
    po: 18,
    riskFlag: "Yes",
    status: "Sent",
    completion: "Done",
  },
  {
    supplierName: "Helvetia Metals AG",
    supplierCode: "CH-2007",
    country: "Switzerland",
    entity: "Basel Plant",
    year: 2024,
    category: "Metals & Mining",
    family: "Raw Materials",
    subFamily: "Aluminum",
    spend: 3_800_000,
    po: 45,
    riskFlag: "No",
    status: "Signed",
    completion: "Done",
  },
  {
    supplierName: "Nordic Robotics SA",
    supplierCode: "CH-2015",
    country: "Switzerland",
    entity: "Zurich Lab",
    year: 2024,
    category: "Automation",
    family: "Capital Equipment",
    subFamily: "",
    spend: 1_250_000,
    po: 9,
    riskFlag: "#REF!",
    status: "Not sent",
    completion: "Pending",
  },
  {
    supplierName: "Danube Packaging",
    supplierCode: "AT-1010",
    country: "Austria",
    entity: "Linz Plant",
    year: 2024,
    category: "Packaging",
    family: "Manufacturing",
    subFamily: "Paper",
    spend: 980_000,
    po: 80,
    riskFlag: "No",
    status: "Review",
    completion: "In progress",
  },
  {
    supplierName: "Danube Packaging",
    supplierCode: "AT-1010",
    country: "Austria",
    entity: "Linz Plant",
    year: 2024,
    category: "Logistics - Ground",
    family: "Manufacturing",
    subFamily: "Paper",
    spend: 240_000,
    po: 12,
    riskFlag: "No",
    status: "Review",
    completion: "In progress",
  },
  {
    supplierName: "Baltic Energy Co",
    supplierCode: "DE-3003",
    country: "Germany",
    entity: "Munich Hub",
    year: 2024,
    category: "Utilities",
    family: "Energy",
    subFamily: "Electricity",
    spend: 4_200_000,
    po: 14,
    riskFlag: "Yes",
    status: "Not compliant",
    completion: "Done",
  },
  {
    supplierName: "Alpine Logistik GMBH",
    supplierCode: "AT-1001",
    country: "Austria",
    entity: "Graz Ops",
    year: 2024,
    category: "Logistics - Ground",
    family: "Transport",
    subFamily: "Freight",
    spend: 310_000,
    po: 8,
    riskFlag: "Yes",
    status: "Sent",
    completion: "Done",
  },
  {
    supplierName: "Swiss IT Cloud",
    supplierCode: "CH-2050",
    country: "Switzerland",
    entity: "Geneva HQ",
    year: 2024,
    category: "IT Services",
    family: "Technology",
    subFamily: "Cloud",
    spend: 1_600_000,
    po: 28,
    riskFlag: "No",
    status: "Signed",
    completion: "Done",
  },
  {
    supplierName: "Vienna Office Supplies",
    supplierCode: "AT-1099",
    country: "Austria",
    entity: "Vienna HQ",
    year: 2024,
    category: "Office Supplies",
    family: "Indirect",
    subFamily: "Stationery",
    spend: 420_000,
    po: 120,
    riskFlag: "#REF!",
    status: "Not sent",
    completion: "Pending",
  },
];

const DEMO_RULES: CategoryRule[] = [
  { key: "Logistics - Ground", scope: "Category", defaultRisk: "High", comment: "Demo policy" },
  { key: "Automation", scope: "Category", defaultRisk: "High", comment: "Demo policy" },
  { key: "Raw Materials", scope: "Family", defaultRisk: "High", comment: "Demo policy" },
];

type PersistedDB = {
  factRows: FactRow[];
  overrides: Record<string, SupplierOverride>; // by supplier code
  categoryRules: CategoryRule[];
  settings: AppSettings;
  tasks: Task[];
  log: LogEntry[];
  lastUndo?: UndoBatch | null;
};

function nowIso() {
  return new Date().toISOString();
}

function uuid() {
  return Math.random().toString(16).slice(2) + "-" + Math.random().toString(16).slice(2);
}

function safeStr(v: any) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

function normCode(v: any) {
  // preserve leading zeros by treating as string
  const s = safeStr(v);
  return s;
}

function toNum(v: any): number {
  if (v === null || v === undefined || v === "") return 0;
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function uniq<T>(arr: T[]) {
  return Array.from(new Set(arr)).filter((x) => x !== undefined && x !== null) as T[];
}

function pickCanonicalByFrequency(candidates: { name: string; count: number }[]) {
  if (!candidates.length) return "";
  const sorted = [...candidates].sort((a, b) => b.count - a.count);
  return sorted[0].name;
}

function computeCategorySharesFromRows(rows: FactRow[]) {
  const by: Record<string, number> = {};
  let total = 0;
  for (const r of rows) {
    const cat = safeStr(r.category) || "(no category)";
    const v = Math.abs(r.spend ?? 0);
    if (!v) continue;
    by[cat] = (by[cat] || 0) + v;
    total += v;
  }
  const shares = Object.entries(by)
    .map(([category, spend]) => ({ category, spend, share: total ? spend / total : 0 }))
    .sort((a, b) => b.spend - a.spend);
  const dominant = shares[0];
  const second = shares[1];
  return {
    categoryShares: shares,
    dominantCategory: dominant?.category,
    dominantCategoryShare: dominant?.share,
    secondCategory: second?.category,
    secondCategoryShare: second?.share,
  };
}


function normalizeContractStatus(raw: string): ContractStatus {
  const s = safeStr(raw).toLowerCase();
  if (!s) return "Not sent";
  if (s.includes("not compliant")) return "Not compliant";
  if (s.includes("signed")) return "Signed";
  if (s.includes("sent")) return "Sent";
  if (s.includes("review")) return "Review";
  // common typos in tracker
  if (s.includes("assess") || s.includes("asess") || s.includes("asessed") || s.includes("assessed")) return "Not sent";
  return "Not sent";
}

function normYesNo(raw: string) {
  const s = safeStr(raw).trim().toLowerCase();
  if (!s) return "";
  if (s.startsWith("y")) return "YES";
  if (s.startsWith("n")) return "NO";
  return s.toUpperCase();
}

function riskFromFlags(flags: (string | null | undefined)[]): RiskLevel {
  const normalized = flags
    .map((x) => safeStr(x))
    .filter(Boolean)
    .map((x) => x.toUpperCase());

  // Any explicit YES -> High
  if (normalized.some((x) => x === "YES")) return "High";

  // #REF! / blanks -> Unknown
  if (normalized.some((x) => x.includes("#REF")) || normalized.length === 0) return "Unknown";

  // All NO -> Non-risk
  if (normalized.every((x) => x === "NO")) return "Non-risk";

  return "Unknown";
}

function evaluatedFromFlags(flags: (string | null | undefined)[]) {
  const normalized = flags
    .map((x) => safeStr(x))
    .filter(Boolean)
    .map((x) => x.toUpperCase());
  if (normalized.length === 0) return false;
  if (normalized.some((x) => x.includes("#REF"))) return false;
  // Evaluated means: only YES/NO present
  return normalized.every((x) => x === "YES" || x === "NO");
}

type LegacySettings = Partial<AppSettings> & {
  countryPalette?: Record<string, string>;
  countryBarFillStyle?: BarFillStyle;
};

function addDaysToDate(date: string, days: number) {
  const d = new Date(`${date}T00:00:00`);
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}

function normalizeSettings(settings?: LegacySettings): AppSettings {
  const merged = { ...DEFAULT_SETTINGS, ...(settings ?? {}) };
  const legacyPalette = settings?.countryPalette;
  const legacyFillStyle = settings?.countryBarFillStyle;
  const barChartSettings: BarChartSettings = {
    ...DEFAULT_SETTINGS.barChartSettings,
    ...merged.barChartSettings,
    palettes: {
      ...DEFAULT_SETTINGS.barChartSettings.palettes,
      ...(merged.barChartSettings?.palettes ?? {}),
    },
    metricColors: {
      ...DEFAULT_SETTINGS.barChartSettings.metricColors,
      ...(merged.barChartSettings?.metricColors ?? {}),
    },
  };

  if (legacyPalette) {
    barChartSettings.palettes.country = {
      ...barChartSettings.palettes.country,
      ...legacyPalette,
    };
  }

  if (legacyFillStyle) {
    barChartSettings.fillStyle = legacyFillStyle;
  }

  return {
    ...merged,
    barChartSettings,
  };
}

function loadDB(): PersistedDB {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) {
      return { factRows: [], overrides: {}, categoryRules: [], settings: DEFAULT_SETTINGS, tasks: [], log: [], lastUndo: null };
    }
    const parsed = JSON.parse(raw);
    return {
      factRows: Array.isArray(parsed.factRows) ? parsed.factRows : [],
      overrides: parsed.overrides ?? {},
      categoryRules: Array.isArray(parsed.categoryRules) ? parsed.categoryRules : [],
      settings: normalizeSettings(parsed.settings ?? {}),
      tasks: Array.isArray(parsed.tasks) ? parsed.tasks : [],
      log: Array.isArray(parsed.log) ? parsed.log : [],
      lastUndo: parsed.lastUndo ?? null,
    };
  } catch {
    return { factRows: [], overrides: {}, categoryRules: [], settings: DEFAULT_SETTINGS, tasks: [], log: [], lastUndo: null };
  }
}

function saveDB(db: PersistedDB) {
  localStorage.setItem(LS_KEY, JSON.stringify(db));
}

// -----------------------------
// XLSX import helpers (robust to "header starts at row X")
// -----------------------------

type SheetMatrix = any[][];

function sheetToMatrix(ws: WorkSheet): SheetMatrix {
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as any[][];
}

function findHeaderRowIndex(matrix: SheetMatrix, requiredHeaders: string[], maxScan = 60): number {
  const req = requiredHeaders.map((h) => h.toLowerCase());
  for (let i = 0; i < Math.min(matrix.length, maxScan); i++) {
    const row = matrix[i] ?? [];
    const rowLower = row.map((c) => safeStr(c).toLowerCase());
    const hits = req.filter((h) => rowLower.includes(h)).length;
    if (hits >= Math.max(2, Math.ceil(req.length * 0.5))) {
      return i;
    }
  }
  return 0;
}

function buildHeaderMap(headerRow: any[]): Record<string, number> {
  const map: Record<string, number> = {};
  headerRow.forEach((cell, idx) => {
    const key = safeStr(cell);
    if (!key) return;
    map[key.toLowerCase()] = idx;
  });
  return map;
}

const HEADER_ALIASES: Record<keyof FactRow, string[]> = {
  // In your tracker the "Grouping Supplier" column is the cleanest name to use for canonical selection.
  supplierName: ["grouping supplier", "grouping_supplier", "supplier_name", "supplier name", "name", "supplier"],
  supplierCode: ["supplier_code_erp", "supplier_code", "supplier code", "supplier nr", "suppliercode"],
  country: ["country"],
  entity: ["entity"],
  year: ["year"],
  category: ["europe catman category", "europe catman category risk list"],
  family: ["family name", "family name.1", "family", "family name (raw)"],
  subFamily: ["sub-family name", "sub family name", "sub-family", "sub_family name", "sub-family name.1"],
  spend: ["spend", "sum spent", "sum_spend"],
  po: ["# po", "po", "po count", "po_number"],
  // Tracker variants: Risk / RIsk / Risk Evolution
  riskFlag: ["risk", "risk by given data", "riskflag", "risk (final)", "risk evolution", "risk status", "risk.1", "rIsk".toLowerCase()],
  status: ["status", "agreement status", "contract status"],
  completion: ["completion status", "completion", "evaluation completion"],
};

function getCellByAliases(row: any[], headerMap: Record<string, number>, aliases: string[]) {
  for (const a of aliases) {
    const idx = headerMap[a.toLowerCase()];
    if (idx !== undefined) return row[idx];
  }
  return null;
}

function parseFactSheet(ws: WorkSheet, requiredForHeaderDetection: string[]): FactRow[] {
  const matrix = sheetToMatrix(ws);
  const headerIdx = findHeaderRowIndex(matrix, requiredForHeaderDetection);
  const headerRow = matrix[headerIdx] ?? [];
  const headerMap = buildHeaderMap(headerRow);
  const rows = matrix.slice(headerIdx + 1);

  const out: FactRow[] = [];
  for (const r of rows) {
    if (!r || r.every((c) => safeStr(c) === "")) continue;

    const supplierCode = normCode(getCellByAliases(r, headerMap, HEADER_ALIASES.supplierCode));
    const supplierName = safeStr(getCellByAliases(r, headerMap, HEADER_ALIASES.supplierName));
    if (!supplierCode && !supplierName) continue;

    out.push({
      supplierName,
      supplierCode,
      country: safeStr(getCellByAliases(r, headerMap, HEADER_ALIASES.country)),
      entity: safeStr(getCellByAliases(r, headerMap, HEADER_ALIASES.entity)),
      year: getCellByAliases(r, headerMap, HEADER_ALIASES.year) ?? "",
      category: safeStr(getCellByAliases(r, headerMap, HEADER_ALIASES.category)),
      family: safeStr(getCellByAliases(r, headerMap, HEADER_ALIASES.family)),
      subFamily: safeStr(getCellByAliases(r, headerMap, HEADER_ALIASES.subFamily)),
      spend: toNum(getCellByAliases(r, headerMap, HEADER_ALIASES.spend)),
      po: toNum(getCellByAliases(r, headerMap, HEADER_ALIASES.po)),
      riskFlag: safeStr(getCellByAliases(r, headerMap, HEADER_ALIASES.riskFlag)),
      status: safeStr(getCellByAliases(r, headerMap, HEADER_ALIASES.status)),
      completion: safeStr(getCellByAliases(r, headerMap, HEADER_ALIASES.completion)),
    });
  }

  return out;
}

// Specialized parser: Risk Evaluation Overview has TWO risk columns ("Risk" and "RIsk").
// We explicitly prefer the final one ("RIsk"), because the "Risk" column often contains "Evaluating".
function parseRiskEvalOverview(ws: WorkSheet): FactRow[] {
  const matrix = sheetToMatrix(ws);
  const headerIdx = findHeaderRowIndex(matrix, ["Supplier Name", "Supplier Code", "RIsk", "Europe Catman category", "FAMILY NAME", "Spend", "# PO"]);
  const headerRow = matrix[headerIdx] ?? [];
  const headerMap = buildHeaderMap(headerRow);
  const rows = matrix.slice(headerIdx + 1);

  const getCell = (row: any[], header: string) => {
    const idx = headerMap[header.toLowerCase()];
    return idx === undefined ? null : row[idx];
  };

  const out: FactRow[] = [];
  for (const r of rows) {
    if (!r || r.every((c) => safeStr(c) === "")) continue;

    const supplierCode = normCode(
      getCell(r, "Supplier Code") ??
        getCell(r, "Supplier nr") ??
        getCell(r, "supplier_code_erp") ??
        getCell(r, "supplier_code")
    );
    const supplierName = safeStr(getCell(r, "Supplier Name") ?? getCell(r, "supplier_name") ?? getCell(r, "Name"));
    if (!supplierCode && !supplierName) continue;

    const finalRisk = safeStr(getCell(r, "RIsk")) || safeStr(getCell(r, "Risk"));

    out.push({
      supplierName,
      supplierCode,
      category: safeStr(getCell(r, "Europe Catman category")),
      family: safeStr(getCell(r, "FAMILY NAME")),
      subFamily: safeStr(getCell(r, "SUB-FAMILY NAME")),
      country: safeStr(getCell(r, "Country")),
      entity: safeStr(getCell(r, "Entity")),
      spend: toNum(getCell(r, "Spend")),
      po: toNum(getCell(r, "# PO")),
      riskFlag: finalRisk,
      completion: safeStr(getCell(r, "Completion status")),
      status: safeStr(getCell(r, "Status")),
    });
  }

  return out;
}


// -----------------------------
// Import helpers for non-snapshot JSON + Settings rules
// -----------------------------

function mergeCategoryRules(existing: CategoryRule[], incoming: CategoryRule[]) {
  if (!incoming.length) return existing;
  const seen = new Set(existing.map((r) => `${r.scope}:${safeStr(r.key).toLowerCase()}`));
  const merged = [...existing];
  for (const r of incoming) {
    const k = `${r.scope}:${safeStr(r.key).toLowerCase()}`;
    if (seen.has(k)) continue;
    seen.add(k);
    merged.push(r);
  }
  return merged;
}

function parseFamilyRiskRulesFromSettings(ws: WorkSheet): CategoryRule[] {
  // Settings contains a "Table 2" with columns: FAMILY NAME.1 + Risk (Yes/No/#REF!)
  const matrix = sheetToMatrix(ws);
  const headerIdx = findHeaderRowIndex(matrix, ["family name", "risk"]);
  const headerRow = matrix[headerIdx] ?? [];
  const headerMap = buildHeaderMap(headerRow);

  const famIdx =
    headerMap["family name.1"] ??
    headerMap["family name"] ??
    headerMap["family"] ??
    null;

  const riskIdx = headerMap["risk"] ?? null;

  if (famIdx === null || riskIdx === null) return [];

  const buckets: Record<string, Set<RiskLevel>> = {};
  const rows = matrix.slice(headerIdx + 1);
  for (const r of rows) {
    if (!r) continue;
    const fam = safeStr(r[famIdx]);
    const rawRisk = safeStr(r[riskIdx]);
    if (!fam || !rawRisk) continue;

    let rl: RiskLevel = "Unknown";
    const yn = normYesNo(rawRisk);
    if (yn === "Yes") rl = "High";
    else if (yn === "No") rl = "Non-risk";
    else if (rawRisk.toUpperCase().includes("#REF")) rl = "Unknown";
    else continue; // ignore unrecognized values

    buckets[fam] = buckets[fam] || new Set<RiskLevel>();
    buckets[fam].add(rl);
  }

  const out: CategoryRule[] = [];
  for (const [fam, set] of Object.entries(buckets)) {
    const levels = Array.from(set);
    const defaultRisk: RiskLevel =
      levels.includes("High") ? "High" : levels.includes("Non-risk") ? "Non-risk" : "Unknown";
    if (defaultRisk === "Unknown") continue;
    out.push({ key: fam, scope: "Family", defaultRisk, comment: "Imported from Settings (Table 2)" });
  }
  return out;
}

function normalizeJsonRowsToFactRows(rows: any[]): FactRow[] {
  const pick = (obj: any, aliases: string[]) => {
    for (const a of aliases) {
      if (obj?.[a] !== undefined) return obj[a];
      // try case-insensitive key match
      const hit = Object.keys(obj || {}).find((k) => k.toLowerCase() === a.toLowerCase());
      if (hit) return obj[hit];
    }
    return null;
  };

  const out: FactRow[] = [];
  for (const obj of rows) {
    if (!obj || typeof obj !== "object") continue;

    const supplierCode = normCode(
      pick(obj, ["supplier_code_erp", "supplier_code", "Supplier Code", "Supplier nr", "Supplier Nr"])
    );
    const supplierName = safeStr(
      pick(obj, ["Grouping Supplier", "grouping supplier", "supplier_name", "Supplier Name", "Name"])
    );

    if (!supplierCode && !supplierName) continue;

    out.push({
      supplierName,
      supplierCode,
      country: safeStr(pick(obj, ["Country"])),
      entity: safeStr(pick(obj, ["Entity"])),
      year: pick(obj, ["Year"]) ?? "",
      category: safeStr(pick(obj, ["Europe Catman category", "Europe Catman Category"])),
      family: safeStr(pick(obj, ["FAMILY NAME", "Family Name", "FAMILY"])),
      subFamily: safeStr(pick(obj, ["SUB-FAMILY NAME", "Sub-Family Name", "Sub Family Name"])),
      spend: toNum(pick(obj, ["Spend", "SUM_SPEND", "Sum Spend"])),
      po: toNum(pick(obj, ["# PO", "PO", "PO count"])),
      riskFlag: safeStr(pick(obj, ["RIsk", "Risk", "Risk Evolution"])),
      status: safeStr(pick(obj, ["Status"])),
      completion: safeStr(pick(obj, ["Completion status", "Completion Status"])),
    });
  }
  return out;
}

function pickRowsFromUnknownJson(parsed: any): any[] {
  if (Array.isArray(parsed)) return parsed;
  if (!parsed || typeof parsed !== "object") return [];
  // common export pattern: { "Р›РёСЃС‚1": [...], ... }
  if (Array.isArray((parsed as any)["Р›РёСЃС‚1"])) return (parsed as any)["Р›РёСЃС‚1"];
  // otherwise: pick the first array-of-objects property
  for (const k of Object.keys(parsed)) {
    const v = (parsed as any)[k];
    if (Array.isArray(v) && v.length && typeof v[0] === "object") return v;
  }
  return [];
}

// -----------------------------
// Derivations
// -----------------------------

function buildSupplierMaster(
  factRows: FactRow[],
  overrides: Record<string, SupplierOverride>,
  categoryRules: CategoryRule[],
  activityLog: LogEntry[],
  settings: AppSettings
): SupplierMaster[] {
  const byCode: Record<string, FactRow[]> = {};
  for (const r of factRows) {
    const code = normCode(r.supplierCode);
    if (!code) continue;
    byCode[code] = byCode[code] || [];
    byCode[code].push(r);
  }

  // fast lookup for policy layer
  const ruleIndex = new Map<string, CategoryRule>();
  for (const rule of categoryRules) {
    const k = `${rule.scope}:${safeStr(rule.key).toLowerCase()}`;
    ruleIndex.set(k, rule);
  }

  const lastTouchByCode = new Map<string, string>();
  for (const e of activityLog) {
    const code = safeStr(e.details?.supplierCode);
    if (code) lastTouchByCode.set(code, e.ts);
  }

  const result: SupplierMaster[] = [];

  for (const code of Object.keys(byCode)) {
    const rows = byCode[code];

    // candidates by frequency
    const counts: Record<string, number> = {};
    for (const r of rows) {
      const nm = safeStr(r.supplierName);
      if (!nm) continue;
      counts[nm] = (counts[nm] || 0) + 1;
    }
    const candidates = Object.entries(counts)
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count);

    const autoCanonical = pickCanonicalByFrequency(candidates) || "(missing name)";

    const ov = overrides[code];

    const flags = rows.map((r) => r.riskFlag);
    const dataRisk = riskFromFlags(flags);

    const shareInfo = computeCategorySharesFromRows(rows);

    // policy hints (Category/Family rules imported from Settings or added manually)
    const policySignals: { scope: "Category" | "Family"; risk: RiskLevel; key: string }[] = [];
    for (const r of rows) {
      const cat = safeStr(r.category);
      const fam = safeStr(r.family);
      if (cat) {
        const rule = ruleIndex.get(`Category:${cat.toLowerCase()}`);
        if (rule) policySignals.push({ scope: "Category", risk: rule.defaultRisk, key: cat });
      }
      if (fam) {
        const rule = ruleIndex.get(`Family:${fam.toLowerCase()}`);
        if (rule) policySignals.push({ scope: "Family", risk: rule.defaultRisk, key: fam });
      }
    }

    const pickPolicy = (scope: "Category" | "Family"): RiskLevel => {
      const sig = policySignals.filter((s) => s.scope === scope).map((s) => s.risk);
      return sig.includes("High") ? "High" : sig.includes("Non-risk") ? "Non-risk" : "Unknown";
    };

    let policyFromCategory = pickPolicy("Category");
    const policyFromFamily = pickPolicy("Family");

    // Guardrail: category-based policy is suppressed when supplier is materially multi-category.
    const secondShare = shareInfo.secondCategoryShare ?? 0;
    const multiCategory = secondShare >= settings.multiCategorySecondShareThreshold;
    if (multiCategory) policyFromCategory = "Unknown";

    const policyRisk: RiskLevel =
      [policyFromFamily, policyFromCategory].includes("High")
        ? "High"
        : [policyFromFamily, policyFromCategory].includes("Non-risk")
          ? "Non-risk"
          : "Unknown";

    // Hierarchy of truth for Risk:
    // 1) Supplier overrides (human decision)
    // 2) Data flags (Raw / Risk Evaluation Overview)
    // 3) Policy rules (Settings list) only when data is Unknown
    const risk: RiskLevel = ov?.risk ?? (dataRisk !== "Unknown" ? dataRisk : policyRisk);

    // "Evaluated" means we have a decision that is not Unknown
    const evaluated =
      ov?.evaluated !== undefined
        ? ov.evaluated
        : evaluatedFromFlags(flags) ||
          (ov?.risk !== undefined && ov?.risk !== "Unknown") ||
          (dataRisk === "Unknown" && policyRisk !== "Unknown");

    // Contract status: override wins, else infer from status strings
    const statusStrings = rows.map((r) => r.status).filter(Boolean) as string[];
    const normalized = statusStrings.map((s) => normalizeContractStatus(s));

    const inferredContract: ContractStatus =
      normalized.includes("Not compliant")
        ? "Not compliant"
        : normalized.includes("Signed")
          ? "Signed"
          : normalized.includes("Sent")
            ? "Sent"
            : normalized.includes("Review")
              ? "Review"
              : "Not sent";

    const contractNeeded = risk === "High";
    const contractStatus: ContractStatus = contractNeeded ? (ov?.contractStatus ?? inferredContract) : "N/A";

    const countries = uniq(rows.map((r) => safeStr(r.country)).filter(Boolean));
    const categories = uniq(rows.map((r) => safeStr(r.category)).filter(Boolean));
    const families = uniq(rows.map((r) => safeStr(r.family)).filter(Boolean));
    const subFamilies = uniq(rows.map((r) => safeStr(r.subFamily)).filter(Boolean));

    const totalSpend = rows.reduce((acc, r) => acc + (r.spend ?? 0), 0);
    const totalPO = rows.reduce((acc, r) => acc + (r.po ?? 0), 0);

    const canonicalName = ov?.canonicalName ?? autoCanonical;

    result.push({
      code,
      canonicalName,
      nameCandidates: candidates,
      countries,
      categories,
      families,
      subFamilies,
      categoryShares: shareInfo.categoryShares,
      dominantCategory: shareInfo.dominantCategory,
      dominantCategoryShare: shareInfo.dominantCategoryShare,
      secondCategory: shareInfo.secondCategory,
      secondCategoryShare: shareInfo.secondCategoryShare,
      multiCategory,
      totalSpend,
      totalPO,
      risk,
      evaluated,
      contractNeeded,
      contractStatus,
      lastTouch: lastTouchByCode.get(code),
    });
  }

  return result.sort((a, b) => b.totalSpend - a.totalSpend);
}

function buildDuplicates(suppliers: SupplierMaster[]) {
  return suppliers
    .filter((s) => s.nameCandidates.length > 1)
    .map((s) => {
      const recommended = pickCanonicalByFrequency(s.nameCandidates) || s.canonicalName;
      return {
        code: s.code,
        candidates: s.nameCandidates,
        recommended,
        current: s.canonicalName,
      };
    });
}

function normNameKey(name: string) {
  return safeStr(name).toLowerCase().replace(/\s+/g, " ").trim();
}

function buildNameManyCodes(factRows: FactRow[]) {
  const by: Record<string, { name: string; codes: Set<string>; spend: number }> = {};
  for (const r of factRows) {
    const nm = safeStr(r.supplierName);
    const code = normCode(r.supplierCode);
    if (!nm || !code) continue;
    const key = normNameKey(nm);
    if (!key || key.length < 3) continue;
    if (!by[key]) by[key] = { name: nm, codes: new Set(), spend: 0 };
    by[key].codes.add(code);
    by[key].spend += Math.abs(r.spend ?? 0);
  }
  return Object.values(by)
    .filter((x) => x.codes.size > 1)
    .map((x) => ({ name: x.name, codes: Array.from(x.codes), codeCount: x.codes.size, spend: x.spend }))
    .sort((a, b) => b.codeCount - a.codeCount || b.spend - a.spend);
}

function buildMissingSubfamilySpendByCode(factRows: FactRow[]) {
  const by: Record<string, { code: string; spend: number; rows: number }> = {};
  for (const r of factRows) {
    const code = normCode(r.supplierCode);
    if (!code) continue;
    const sf = safeStr(r.subFamily);
    if (sf) continue;
    const v = Math.abs(r.spend ?? 0);
    if (!v) continue;
    by[code] = by[code] || { code, spend: 0, rows: 0 };
    by[code].spend += v;
    by[code].rows += 1;
  }
  return Object.values(by).sort((a, b) => b.spend - a.spend);
}

function buildIssues(
  factRows: FactRow[],
  suppliers: SupplierMaster[],
  duplicates: ReturnType<typeof buildDuplicates>,
  settings: AppSettings,
  categoryRules: CategoryRule[]
): IssueInstance[] {
  const issues: IssueInstance[] = [];

  // 1) Code -> many names (duplicates)
  for (const d of duplicates) {
    const s = suppliers.find((x) => x.code === d.code);
    const spend = s?.totalSpend ?? 0;
    const sev: IssueSeverity = Math.abs(spend) >= settings.scopeSpendThreshold || d.candidates.length >= 3 ? "High" : "Medium";
    issues.push({
      id: `dup:${d.code}`,
      type: "CODE_MANY_NAMES",
      severity: sev,
      title: `Code has multiple names (${d.candidates.length})`,
      code: d.code,
      supplierName: s?.canonicalName,
      spendAffected: spend,
      details: { candidates: d.candidates, recommended: d.recommended, current: d.current },
    });
  }

  // 2) Name -> many codes
  const nameMany = buildNameManyCodes(factRows);
  for (const n of nameMany.slice(0, 500)) {
    const sev: IssueSeverity = n.codeCount >= 5 ? "High" : "Medium";
    issues.push({
      id: `name-many:${normNameKey(n.name)}`,
      type: "NAME_MANY_CODES",
      severity: sev,
      title: `Name appears under multiple codes (${n.codeCount})`,
      supplierName: n.name,
      spendAffected: n.spend,
      details: { codes: n.codes.slice(0, 25), codeCount: n.codeCount },
    });
  }

  // 3) Multi-category exposure
  for (const s of suppliers) {
    const secondShare = s.secondCategoryShare ?? 0;
    if (secondShare < settings.multiCategorySecondShareThreshold) continue;
    const sev: IssueSeverity = Math.abs(s.totalSpend) >= settings.scopeSpendThreshold ? "High" : "Medium";
    issues.push({
      id: `multicat:${s.code}`,
      type: "MULTI_CATEGORY_EXPOSURE",
      severity: sev,
      title: `Multi-category exposure (2nd share ${(secondShare * 100).toFixed(0)}%)`,
      code: s.code,
      supplierName: s.canonicalName,
      spendAffected: s.totalSpend,
      details: {
        dominantCategory: s.dominantCategory,
        dominantShare: s.dominantCategoryShare,
        secondCategory: s.secondCategory,
        secondShare,
        top3: s.categoryShares.slice(0, 3),
      },
    });
  }

  // 4) Risk unknown but in scope
  for (const s of suppliers) {
    if (s.risk !== "Unknown") continue;
    if (Math.abs(s.totalSpend) < settings.scopeSpendThreshold) continue;
    issues.push({
      id: `unknown-scope:${s.code}`,
      type: "RISK_UNKNOWN_IN_SCOPE",
      severity: "High",
      title: `Unknown risk but in-scope spend`,
      code: s.code,
      supplierName: s.canonicalName,
      spendAffected: s.totalSpend,
      details: { threshold: settings.scopeSpendThreshold },
    });
  }

  // 5) Missing sub-family on meaningful spend
  const missingSF = buildMissingSubfamilySpendByCode(factRows);
  for (const m of missingSF) {
    if (m.spend < settings.missingSubfamilySpendThreshold) continue;
    issues.push({
      id: `missing-sf:${m.code}`,
      type: "MISSING_SUBFAMILY_HIGH_SPEND",
      severity: "Medium",
      title: `Missing sub-family on high spend`,
      code: m.code,
      spendAffected: m.spend,
      details: { rows: m.rows, threshold: settings.missingSubfamilySpendThreshold },
    });
  }

  // 6) Policy suppressed by guardrail (multi-category + category rule exists)
  const ruleIndex = new Map<string, CategoryRule>();
  for (const r of categoryRules) ruleIndex.set(`${r.scope}:${safeStr(r.key).toLowerCase()}`, r);
  for (const s of suppliers) {
    if (!s.multiCategory) continue;
    const matched: { key: string; risk: RiskLevel }[] = [];
    for (const cat of s.categories) {
      const rule = ruleIndex.get(`Category:${cat.toLowerCase()}`);
      if (rule && rule.defaultRisk !== "Unknown") matched.push({ key: cat, risk: rule.defaultRisk });
    }
    if (!matched.length) continue;
    issues.push({
      id: `suppressed:${s.code}`,
      type: "POLICY_SUPPRESSED_BY_GUARDRAIL",
      severity: "Low",
      title: `Category policy suppressed (multi-category)`,
      code: s.code,
      supplierName: s.canonicalName,
      spendAffected: s.totalSpend,
      details: { matchedRules: matched, secondShare: s.secondCategoryShare, threshold: settings.multiCategorySecondShareThreshold },
    });
  }

  const sevWeight: Record<IssueSeverity, number> = { Critical: 4, High: 3, Medium: 2, Low: 1 };
  return issues.sort((a, b) =>
    (sevWeight[b.severity] - sevWeight[a.severity]) || Math.abs((b.spendAffected ?? 0)) - Math.abs((a.spendAffected ?? 0))
  );
}

function buildCategoryMetrics(factRows: FactRow[], suppliers: SupplierMaster[]) {
  const riskByCode = new Map<string, RiskLevel>();
  for (const s of suppliers) riskByCode.set(s.code, s.risk);

  const by: Record<
    string,
    { supplierSet: Set<string>; highSet: Set<string>; spend: number; highSpend: number; po: number; highPO: number }
  > = {};

  for (const r of factRows) {
    const cat = safeStr(r.category) || "(no category)";
    const code = normCode(r.supplierCode);
    if (!by[cat]) by[cat] = { supplierSet: new Set(), highSet: new Set(), spend: 0, highSpend: 0, po: 0, highPO: 0 };

    const entry = by[cat];
    if (code) {
      entry.supplierSet.add(code);
      if (riskByCode.get(code) === "High") entry.highSet.add(code);
    }

    const spend = r.spend ?? 0;
    const po = r.po ?? 0;
    entry.spend += spend;
    entry.po += po;

    if (code && riskByCode.get(code) === "High") {
      entry.highSpend += spend;
      entry.highPO += po;
    }
  }

  return Object.entries(by)
    .map(([category, v]) => ({
      category,
      suppliers: v.supplierSet.size,
      highRiskSuppliers: v.highSet.size,
      riskShare: v.supplierSet.size ? v.highSet.size / v.supplierSet.size : 0,
      totalSpend: v.spend,
      highRiskSpend: v.highSpend,
      totalPO: v.po,
      highRiskPO: v.highPO,
    }))
    .sort((a, b) => b.highRiskSuppliers - a.highRiskSuppliers);
}

function buildCountryMetrics(factRows: FactRow[], suppliers: SupplierMaster[]) {
  const riskByCode = new Map<string, RiskLevel>();
  for (const s of suppliers) riskByCode.set(s.code, s.risk);

  const by: Record<
    string,
    { supplierSet: Set<string>; highSet: Set<string>; spend: number; highSpend: number; po: number; highPO: number }
  > = {};

  for (const r of factRows) {
    const country = safeStr(r.country) || "(unknown)";
    const code = normCode(r.supplierCode);
    if (!by[country]) by[country] = { supplierSet: new Set(), highSet: new Set(), spend: 0, highSpend: 0, po: 0, highPO: 0 };

    const entry = by[country];
    if (code) {
      entry.supplierSet.add(code);
      if (riskByCode.get(code) === "High") entry.highSet.add(code);
    }

    const spend = r.spend ?? 0;
    const po = r.po ?? 0;
    entry.spend += spend;
    entry.po += po;

    if (code && riskByCode.get(code) === "High") {
      entry.highSpend += spend;
      entry.highPO += po;
    }
  }

  return Object.entries(by)
    .map(([country, v]) => ({
      country,
      suppliers: v.supplierSet.size,
      highRiskSuppliers: v.highSet.size,
      totalSpend: v.spend,
      highRiskSpend: v.highSpend,
      totalPO: v.po,
      highRiskPO: v.highPO,
      riskShare: v.supplierSet.size ? v.highSet.size / v.supplierSet.size : 0,
    }))
    .sort((a, b) => b.highRiskSuppliers - a.highRiskSuppliers);
}

function fmtMoney(n: number) {
  const abs = Math.abs(n);
  const sign = n < 0 ? "-" : "";
  if (abs >= 1e9) return `${sign}${(abs / 1e9).toFixed(2)}B`;
  if (abs >= 1e6) return `${sign}${(abs / 1e6).toFixed(2)}M`;
  if (abs >= 1e3) return `${sign}${(abs / 1e3).toFixed(1)}K`;
  return `${sign}${abs.toFixed(0)}`;
}

function clamp01(v: number) {
  if (Number.isNaN(v) || !Number.isFinite(v)) return 0;
  return Math.min(1, Math.max(0, v));
}

function fmtDate(d: string) {
  try {
    const dt = new Date(d);
    return dt.toLocaleString();
  } catch {
    return d;
  }
}

function todayIso() {
  return new Date().toISOString().slice(0, 10);
}

function addDays(dateStr: string, days: number) {
  const d = new Date(`${dateStr}T00:00:00`);
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}

function shortLabel(label: string, max = 22) {
  if (label.length <= max) return label;
  return `${label.slice(0, max - 1)}вЂ¦`;
}

function slugifyKey(value: string) {
  return safeStr(value).toLowerCase().replace(/\s+/g, "-").replace(/[^a-z0-9-_]/g, "");
}

// -----------------------------
// UI components
// -----------------------------

function getReadableTextColor(hex: string) {
  const sanitized = hex.replace("#", "");
  if (sanitized.length !== 6) return "#0f172a";
  const r = parseInt(sanitized.slice(0, 2), 16);
  const g = parseInt(sanitized.slice(2, 4), 16);
  const b = parseInt(sanitized.slice(4, 6), 16);
  const luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
  return luminance > 0.65 ? "#0f172a" : "#f8fafc";
}

const RiskBadge = React.memo(function RiskBadge({ risk, tone }: { risk: RiskLevel; tone?: string }) {
  if (tone) {
    return (
      <Badge
        variant="outline"
        className="border"
        style={{ backgroundColor: tone, borderColor: tone, color: getReadableTextColor(tone) }}
      >
        {risk === "High" ? "High risk" : risk}
      </Badge>
    );
  }
  if (risk === "High") return <Badge variant="destructive">High risk</Badge>;
  if (risk === "Non-risk") return <Badge variant="secondary">Non-risk</Badge>;
  return <Badge variant="outline">Unknown</Badge>;
});

const ContractBadge = React.memo(function ContractBadge({ status, tone }: { status: ContractStatus; tone?: string }) {
  if (tone) {
    return (
      <Badge
        variant="outline"
        className="border"
        style={{ backgroundColor: tone, borderColor: tone, color: getReadableTextColor(tone) }}
      >
        {status}
      </Badge>
    );
  }
  if (status === "Signed") return <Badge variant="secondary">Signed</Badge>;
  if (status === "Sent") return <Badge variant="outline">Sent</Badge>;
  if (status === "Not compliant") return <Badge variant="destructive">Not compliant</Badge>;
  if (status === "Review") return <Badge variant="outline">Review</Badge>;
  if (status === "N/A") return <Badge variant="secondary">N/A</Badge>;
  return <Badge variant="outline">Not sent</Badge>;
});

const SeverityBadge = React.memo(function SeverityBadge({ sev }: { sev: IssueSeverity }) {
  if (sev === "Critical") return <Badge variant="destructive">Critical</Badge>;
  if (sev === "High") return <Badge variant="destructive">High</Badge>;
  if (sev === "Medium") return <Badge variant="outline">Medium</Badge>;
  return <Badge variant="secondary">Low</Badge>;
});

const EmptyState = React.memo(function EmptyState({ title, subtitle }: { title: string; subtitle?: string }) {
  return (
    <Card>
      <CardContent className="py-10 text-center">
        <div className="mx-auto mb-3 w-fit rounded-full bg-muted p-3">
          <AlertCircle className="h-5 w-5" />
        </div>
        <div className="text-lg font-semibold">{title}</div>
        {subtitle ? <div className="mt-1 text-sm text-muted-foreground">{subtitle}</div> : null}
      </CardContent>
    </Card>
  );
});

type SupplierDialogProps = {
  supplier: SupplierMaster;
  overrides: Record<string, SupplierOverride>;
  actor: string;
  bulkContactName: string;
  bulkContactEmail: string;
  supplierContactEmail?: string;
  onSetSupplierOverride: (code: string, patch: Partial<SupplierOverride>) => void;
  onSetCanonical: (code: string, name: string) => void;
  onAddTask: (task: {
    supplierCode: string;
    supplierName: string;
    title: string;
    note?: string;
    date: string;
    owner: string;
    dueDate: string;
    contactName?: string;
    contactEmail?: string;
  }) => void;
  onScheduleEmailFollowUp: (options: {
    supplierCode: string;
    supplierName: string;
    contactName?: string;
    contactEmail?: string;
    stage: "initial" | "followup";
  }) => void;
};

const SupplierDialog = React.memo(function SupplierDialog({
  supplier,
  overrides,
  actor,
  bulkContactName,
  bulkContactEmail,
  supplierContactEmail,
  onSetSupplierOverride,
  onSetCanonical,
  onAddTask,
  onScheduleEmailFollowUp,
}: SupplierDialogProps) {
  const [open, setOpen] = useState(false);
  const [taskDate, setTaskDate] = useState(() => todayIso());
  const [taskOwner, setTaskOwner] = useState(actor);
  const [taskContactName, setTaskContactName] = useState(bulkContactName);
  const [taskContactEmail, setTaskContactEmail] = useState(supplierContactEmail || bulkContactEmail);
  const [taskTitle, setTaskTitle] = useState("");
  const [taskNote, setTaskNote] = useState("");
  const [taskDueDate, setTaskDueDate] = useState(() => todayIso());

  useEffect(() => {
    if (!open) return;
    setTaskDate(todayIso());
    setTaskOwner(actor);
    setTaskContactName(bulkContactName);
    setTaskContactEmail(supplierContactEmail || bulkContactEmail);
    setTaskTitle("");
    setTaskNote("");
    setTaskDueDate(todayIso());
  }, [open, actor, bulkContactName, bulkContactEmail, supplierContactEmail]);

  const override = overrides[supplier.code];

  return (
    <Dialog open={open} onOpenChange={setOpen}>
      <DialogTrigger asChild>
        <Button variant="secondary" size="sm" className="gap-2">
          <Settings className="h-4 w-4" /> Open
        </Button>
      </DialogTrigger>
      <DialogContent className="max-w-3xl">
        <DialogHeader>
          <DialogTitle>Supplier card</DialogTitle>
        </DialogHeader>
        <div className="grid gap-3 md:grid-cols-2">
          <Card>
            <CardHeader className="pb-2">
              <CardTitle className="text-base">Identity</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="text-sm text-muted-foreground">Supplier code (golden key)</div>
              <div className="text-lg font-semibold">{supplier.code}</div>

              <Separator className="my-3" />

              <div className="text-sm font-medium">Canonical name</div>
              <Input
                value={override?.canonicalName ?? supplier.canonicalName}
                onChange={(e) => onSetSupplierOverride(supplier.code, { canonicalName: e.target.value })}
              />

              <div className="mt-3 text-sm font-medium">Name candidates</div>
              <div className="mt-2 space-y-2">
                {supplier.nameCandidates.slice(0, 8).map((candidate) => (
                  <div key={candidate.name} className="flex items-center justify-between gap-2 rounded-xl border p-2">
                    <div>
                      <div className="text-sm font-medium">{candidate.name}</div>
                      <div className="text-xs text-muted-foreground">Seen {candidate.count}×</div>
                    </div>
                    <Button variant="outline" size="sm" onClick={() => onSetCanonical(supplier.code, candidate.name)}>
                      Use
                    </Button>
                  </div>
                ))}
              </div>
            </CardContent>
          </Card>

          <Card>
            <CardHeader className="pb-2">
              <CardTitle className="text-base">Risk & Contract</CardTitle>
            </CardHeader>
            <CardContent className="space-y-3">
              <div className="flex items-center justify-between">
                <div>
                  <div className="text-sm font-medium">Risk</div>
                  <div className="text-xs text-muted-foreground">Overrides beat data & rules.</div>
                </div>
                <Select value={override?.risk ?? supplier.risk} onValueChange={(v: any) => onSetSupplierOverride(supplier.code, { risk: v })}>
                  <SelectTrigger className="w-[160px]">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="High">High</SelectItem>
                    <SelectItem value="Non-risk">Non-risk</SelectItem>
                    <SelectItem value="Unknown">Unknown</SelectItem>
                  </SelectContent>
                </Select>
              </div>

              <div className="flex items-center justify-between">
                <div>
                  <div className="text-sm font-medium">Evaluated</div>
                  <div className="text-xs text-muted-foreground">Override evaluation status.</div>
                </div>
                <Select
                  value={
                    override?.evaluated === undefined
                      ? "auto"
                      : override?.evaluated
                        ? "true"
                        : "false"
                  }
                  onValueChange={(v: any) =>
                    onSetSupplierOverride(supplier.code, {
                      evaluated: v === "auto" ? undefined : v === "true",
                    })
                  }
                >
                  <SelectTrigger className="w-[160px]">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="auto">Auto</SelectItem>
                    <SelectItem value="true">Evaluated</SelectItem>
                    <SelectItem value="false">Not evaluated</SelectItem>
                  </SelectContent>
                </Select>
              </div>

              <div className="flex items-center justify-between">
                <div>
                  <div className="text-sm font-medium">Contract status</div>
                  <div className="text-xs text-muted-foreground">Only relevant for High risk.</div>
                </div>
                <Select
                  value={override?.contractStatus ?? supplier.contractStatus}
                  onValueChange={(v: any) => onSetSupplierOverride(supplier.code, { contractStatus: v })}
                >
                  <SelectTrigger className="w-[180px]">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="Not sent">Not sent</SelectItem>
                    <SelectItem value="Sent">Sent</SelectItem>
                    <SelectItem value="Signed">Signed</SelectItem>
                    <SelectItem value="Review">Review</SelectItem>
                    <SelectItem value="Not compliant">Not compliant</SelectItem>
                    <SelectItem value="N/A">N/A</SelectItem>
                  </SelectContent>
                </Select>
              </div>

              <Separator />

              <div className="grid gap-2 md:grid-cols-2">
                <div className="rounded-2xl border p-3">
                  <div className="text-xs text-muted-foreground">Total spend</div>
                  <div className="text-lg font-semibold">{fmtMoney(supplier.totalSpend)}</div>
                </div>
                <div className="rounded-2xl border p-3">
                  <div className="text-xs text-muted-foreground">Total POs</div>
                  <div className="text-lg font-semibold">{Math.round(supplier.totalPO)}</div>
                </div>
              </div>

              <div>
                <div className="text-sm font-medium">Notes (override)</div>
                <Textarea
                  value={override?.note ?? ""}
                  onChange={(e) => onSetSupplierOverride(supplier.code, { note: e.target.value })}
                  placeholder="Decision rationale, next steps, exceptions…"
                />
              </div>
            </CardContent>
          </Card>
        </div>

        <Separator className="my-2" />

        <div className="grid gap-3 md:grid-cols-2">
          <Card>
            <CardHeader className="pb-2">
              <CardTitle className="text-base">Category footprint</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="flex flex-wrap gap-1">
                {supplier.categories.map((cat) => (
                  <Badge key={cat} variant="outline">
                    {cat}
                  </Badge>
                ))}
              </div>
            </CardContent>
          </Card>

          <Card>
            <CardHeader className="pb-2">
              <CardTitle className="text-base">Add action to calendar</CardTitle>
            </CardHeader>
            <CardContent>
              <div className="grid gap-2 md:grid-cols-2">
                <div>
                  <div className="text-xs text-muted-foreground">Resolve by</div>
                  <Input value={taskDate} onChange={(e) => setTaskDate(e.target.value)} type="date" />
                </div>
                <div>
                  <div className="text-xs text-muted-foreground">Owner</div>
                  <Input value={taskOwner} onChange={(e) => setTaskOwner(e.target.value)} />
                </div>
              </div>
              <div className="mt-2 grid gap-2 md:grid-cols-2">
                <div>
                  <div className="text-xs text-muted-foreground">Contact name</div>
                  <Input value={taskContactName} onChange={(e) => setTaskContactName(e.target.value)} placeholder={bulkContactName || "Procurement contact"} />
                </div>
                <div>
                  <div className="text-xs text-muted-foreground">Contact email</div>
                  <Input value={taskContactEmail} onChange={(e) => setTaskContactEmail(e.target.value)} placeholder={bulkContactEmail || "name@company.com"} />
                </div>
              </div>
              <div className="mt-2">
                <div className="text-xs text-muted-foreground">Title</div>
                <Input value={taskTitle} onChange={(e) => setTaskTitle(e.target.value)} placeholder="Chase signature / send reminder / review docs…" />
              </div>
              <div className="mt-2">
                <div className="text-xs text-muted-foreground">Resolve by</div>
                <Input value={taskDueDate} onChange={(e) => setTaskDueDate(e.target.value)} type="date" />
              </div>
              <div className="mt-2">
                <div className="text-xs text-muted-foreground">Note</div>
                <Textarea value={taskNote} onChange={(e) => setTaskNote(e.target.value)} />
              </div>
              <div className="mt-2 flex flex-wrap justify-end gap-2">
                <Button
                  variant="outline"
                  className="gap-2"
                  onClick={() =>
                    onScheduleEmailFollowUp({
                      supplierCode: supplier.code,
                      supplierName: supplier.canonicalName,
                      contactName: taskContactName || bulkContactName,
                      contactEmail: taskContactEmail || supplierContactEmail || bulkContactEmail,
                      stage: "initial",
                    })
                  }
                >
                  <Upload className="h-4 w-4" /> Email sent (14d)
                </Button>
                <Button
                  className="gap-2"
                  onClick={() =>
                    onAddTask({
                      supplierCode: supplier.code,
                      supplierName: supplier.canonicalName,
                      title: taskTitle,
                      note: taskNote,
                      date: taskDate,
                      owner: taskOwner,
                      dueDate: taskDueDate,
                      contactName: taskContactName || bulkContactName,
                      contactEmail: taskContactEmail || supplierContactEmail || bulkContactEmail,
                    })
                  }
                >
                  <CalendarDays className="h-4 w-4" /> Add task
                </Button>
              </div>
            </CardContent>
          </Card>
        </div>
      </DialogContent>
    </Dialog>
  );
});

function useDebounce<T>(value: T, delayMs: number) {
  const [debounced, setDebounced] = useState(value);

  useEffect(() => {
    const handle = window.setTimeout(() => {
      setDebounced(value);
    }, delayMs);
    return () => window.clearTimeout(handle);
  }, [value, delayMs]);

  return debounced;
}

// -----------------------------
// Main App
// -----------------------------
// вњ… Р’РќР• РєРѕРјРїРѕРЅРµРЅС‚Р° (РЅРµ РёСЃРїРѕР»СЊР·СѓСЋС‚ state)
function getRiskColor(risk: RiskLevel, fallback: string): string {
  return fallback;
}

function getContractColor(status: ContractStatus, fallback: string): string {
  return fallback;
}

export default function SupplierRiskOpsDashboard() {
  const [db, setDB] = useState<PersistedDB>(() => loadDB());
  const [actor, setActor] = useState("Hero");
  const [activeTab, setActiveTab] = useState("overview");
  const [globalCountry, setGlobalCountry] = useState("Overall");

  // filters
  const [q, setQ] = useState("");
  const [drillCategory, setDrillCategory] = useState<string | null>(null);
  const [drillCountry, setDrillCountry] = useState<string | null>(null);
  const [drillFunnelStage, setDrillFunnelStage] = useState<string | null>(null);
  const [riskFilter, setRiskFilter] = useState<RiskLevel | "All">("All");
  const [evalFilter, setEvalFilter] = useState<"All" | "Evaluated" | "Not evaluated">("All");
  const [contractFilter, setContractFilter] = useState<ContractStatus | "All">("All");
  const [categoryFilter, setCategoryFilter] = useState<string | "All">("All");
  const [fastActionSort, setFastActionSort] = useState<{ key: "spend" | "name" | "contract"; dir: "asc" | "desc" }>({
    key: "spend",
    dir: "desc",
  });
  const [categoryTableSort, setCategoryTableSort] = useState<{ key: "category" | "suppliers" | "highRisk" | "riskShare" | "spend"; dir: "asc" | "desc" }>({
    key: "suppliers",
    dir: "desc",
  });

  // quality lab filters
  const [issueTypeFilter, setIssueTypeFilter] = useState<IssueType | "All">("All");
  const [issueSeverityFilter, setIssueSeverityFilter] = useState<IssueSeverity | "All">("All");

  const [topCategorySortBy, setTopCategorySortBy] = useState<"highRiskSuppliers" | "suppliers">("highRiskSuppliers");
  const [topSpendSortBy, setTopSpendSortBy] = useState<"highRiskSpend" | "totalSpend">("highRiskSpend");

  // bulk
  const [bulkCodes, setBulkCodes] = useState("");
  const [bulkRisk, setBulkRisk] = useState<RiskLevel>("High");
  const [bulkContract, setBulkContract] = useState<ContractStatus>("Sent");
  const [bulkCanonicalName, setBulkCanonicalName] = useState("");
  const [bulkEvaluated, setBulkEvaluated] = useState<EvaluationChoice>("Evaluated");
  const [bulkNameQuery, setBulkNameQuery] = useState("");
  const [bulkNameSelected, setBulkNameSelected] = useState<Set<string>>(new Set());
  const [selectedDuplicateCode, setSelectedDuplicateCode] = useState<string | null>(null);
  const [duplicateCandidateDetail, setDuplicateCandidateDetail] = useState<{
    code: string;
    name: string;
    totalPO: number;
    totalSpend: number;
    lastSeen?: string;
  } | null>(null);
  const [manualCanonicalName, setManualCanonicalName] = useState("");

  // category fast action (eligible-only) - uses spend dominance guardrail
  const [fastCategory, setFastCategory] = useState<string>("");
  const [fastCategoryRisk, setFastCategoryRisk] = useState<RiskLevel>("High");

  // rule editor
  const [newRuleScope, setNewRuleScope] = useState<CategoryRule["scope"]>("Category");
  const [newRuleKey, setNewRuleKey] = useState("");
  const [newRuleRisk, setNewRuleRisk] = useState<RiskLevel>("High");
  const [newRuleComment, setNewRuleComment] = useState("");

  // tasks
  const [taskDate, setTaskDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [taskOwner, setTaskOwner] = useState(actor);
  const [taskTitle, setTaskTitle] = useState("");
  const [taskNote, setTaskNote] = useState("");
  const [taskSupplierCode, setTaskSupplierCode] = useState("");
  const [taskSupplierName, setTaskSupplierName] = useState("");
  const [taskContactName, setTaskContactName] = useState("");
  const [taskContactEmail, setTaskContactEmail] = useState("");
  const [taskDueDate, setTaskDueDate] = useState(() => new Date().toISOString().slice(0, 10));

  const [selectedSupplierCodes, setSelectedSupplierCodes] = useState<Set<string>>(new Set());
  const [selectedSuppliers, setSelectedSuppliers] = useState<Set<string>>(new Set());
  const [bulkMarkRisk, setBulkMarkRisk] = useState<BulkRiskChoice>("No change");
  const [bulkMarkEvaluated, setBulkMarkEvaluated] = useState<BulkEvalChoice>("No change");
  const [fastQueueSort, setFastQueueSort] = useState<"SpendDesc" | "SpendAsc" | "NameAsc" | "NameDesc">("SpendDesc");
  const [categoryTopSort, setCategoryTopSort] = useState<"HighRiskDesc" | "TotalDesc">("HighRiskDesc");
  const [categorySpendSort, setCategorySpendSort] = useState<"HighRiskDesc" | "TotalDesc">("HighRiskDesc");
  const [drillSortState, setDrillSortState] = useState<{ key: string; dir: "asc" | "desc" }>({ key: "spendInSlice", dir: "desc" });
  const [categoryTableSortState, setCategoryTableSortState] = useState<{ key: string; dir: "asc" | "desc" }>({ key: "highRiskSuppliers", dir: "desc" });
  const [barInsight, setBarInsight] = useState<{ title: string; label: string; details: string[] } | null>(null);
  const [barAnalysis, setBarAnalysis] = useState<{ title: string; details: Record<string, any> } | null>(null);
  const [bulkContactEntries, setBulkContactEntries] = useState("");

 const fileInputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    saveDB(db);
  }, [db]);

  // вњ… Р’РќРЈРўР Р РєРѕРјРїРѕРЅРµРЅС‚Р° (РїРѕСЃР»Рµ РІСЃРµС… useState/useEffect, РїРµСЂРµРґ useMemo)
  function upsertSupplierContact(code: string, email: string) {
    const trimmedCode = normCode(code);
    const trimmedEmail = safeStr(email);
    if (!trimmedCode || !trimmedEmail) return;
    
    setDB((prev) => ({
      ...prev,
      settings: {
        ...prev.settings,
        supplierContacts: {
          ...(prev.settings.supplierContacts ?? {}),
          [trimmedCode]: trimmedEmail,
        },
      },
    }));
    
    pushLog({
      type: "SETTINGS",
      summary: `Saved contact for ${trimmedCode}: ${trimmedEmail}`,
      details: { code: trimmedCode, email: trimmedEmail },
    });
  }

  const availableCountries = useMemo(() => {
    const list = uniq(db.factRows.map((r) => safeStr(r.country)).filter(Boolean)).sort();
    return ["Overall", ...list];
  }, [db.factRows]);
  const scopedFactRows = useMemo(() => {
    if (globalCountry === "Overall") return db.factRows;
    return db.factRows.filter((r) => safeStr(r.country) === globalCountry);
  }, [db.factRows, globalCountry]);
  const suppliers = useMemo(
    () => buildSupplierMaster(scopedFactRows, db.overrides, db.categoryRules, db.log, db.settings),
    [scopedFactRows, db.overrides, db.categoryRules, db.log, db.settings]
  );
  const allSuppliers = useMemo(
    () => buildSupplierMaster(db.factRows, db.overrides, db.categoryRules, db.log, db.settings),
    [db.factRows, db.overrides, db.categoryRules, db.log, db.settings]
  );

  const duplicates = useMemo(() => buildDuplicates(suppliers), [suppliers]);
  const selectedDuplicate = useMemo(
    () => (selectedDuplicateCode ? duplicates.find((d) => d.code === selectedDuplicateCode) ?? null : null),
    [duplicates, selectedDuplicateCode]
  );
  const issues = useMemo(
    () => buildIssues(scopedFactRows, suppliers, duplicates, db.settings, db.categoryRules),
    [scopedFactRows, suppliers, duplicates, db.settings, db.categoryRules]
  );

  const issueBuckets = useMemo(() => {
    const by: Record<IssueType, number> = {
      CODE_MANY_NAMES: 0,
      NAME_MANY_CODES: 0,
      MULTI_CATEGORY_EXPOSURE: 0,
      RISK_UNKNOWN_IN_SCOPE: 0,
      MISSING_SUBFAMILY_HIGH_SPEND: 0,
      POLICY_SUPPRESSED_BY_GUARDRAIL: 0,
    };
    for (const i of issues) by[i.type] += 1;
    return by;
  }, [issues]);

  const issuesFiltered = useMemo(() => {
    return issues
      .filter((i) => (issueTypeFilter === "All" ? true : i.type === issueTypeFilter))
      .filter((i) => (issueSeverityFilter === "All" ? true : i.severity === issueSeverityFilter))
      .slice(0, 1000);
  }, [issues, issueTypeFilter, issueSeverityFilter]);
  const categoryMetrics = useMemo(() => buildCategoryMetrics(scopedFactRows, suppliers), [scopedFactRows, suppliers]);
  const countryMetrics = useMemo(() => buildCountryMetrics(scopedFactRows, suppliers), [scopedFactRows, suppliers]);
  const barChartSettings = db.settings.barChartSettings;
  const metricColors = barChartSettings.metricColors;
  const categories = useMemo(() => ["All", ...uniq(categoryMetrics.map((c) => c.category)).sort()], [categoryMetrics]);
  const riskyCategoryChart = useMemo(
    () =>
      [...categoryMetrics]
        .filter((c) => c.suppliers >= 15)
        .sort((a, b) => b.riskShare - a.riskShare)
      .slice(0, 12),
    [categoryMetrics]
  );
  const candidateStats = useMemo(() => {
    const map = new Map<
      string,
      {
        rows: number;
        totalSpend: number;
        totalPO: number;
        minYear?: number;
        maxYear?: number;
      }
    >();
    for (const row of db.factRows) {
      const code = normCode(row.supplierCode);
      const name = safeStr(row.supplierName);
      if (!code || !name) continue;
      const key = `${code}||${name}`;
      const entry = map.get(key) ?? { rows: 0, totalSpend: 0, totalPO: 0 };
      entry.rows += 1;
      entry.totalSpend += Math.abs(row.spend ?? 0);
      entry.totalPO += row.po ?? 0;
      const year = Number(row.year);
      if (Number.isFinite(year)) {
        entry.minYear = entry.minYear === undefined ? year : Math.min(entry.minYear, year);
        entry.maxYear = entry.maxYear === undefined ? year : Math.max(entry.maxYear, year);
      }
      map.set(key, entry);
    }
    return map;
  }, [db.factRows]);
  const emailAutomation = db.settings.emailAutomation;

  const redIgnoranceAlertTasks = useMemo(() => {
    const today = new Date().toISOString().slice(0, 10);
    return db.tasks.filter((t) => {
      if (t.kind !== "email") return false;
      const due = t.dueDate ?? t.date;
      if (!due) return false;
      if (t.followUpStage === "urgent") return true;
      return t.followUpStage === "followup" && due < today && t.status !== "done";
    });
  }, [db.tasks]);

  const categoryMetricsSorted = useMemo(() => {
    const dir = categoryTableSort.dir === "asc" ? 1 : -1;
    return [...categoryMetrics].sort((a, b) => {
      switch (categoryTableSort.key) {
        case "category":
          return dir * a.category.localeCompare(b.category);
        case "suppliers":
          return dir * (a.suppliers - b.suppliers);
        case "highRisk":
          return dir * (a.highRiskSuppliers - b.highRiskSuppliers);
        case "riskShare":
          return dir * (a.riskShare - b.riskShare);
        case "spend":
          return dir * (a.totalSpend - b.totalSpend);
        default:
          return 0;
      }
    });
  }, [categoryMetrics, categoryTableSort]);

  const overviewFactRows = useMemo(() => {
    if (globalCountry === "Overall") return db.factRows;
    return db.factRows.filter((r) => safeStr(r.country) === globalCountry);
  }, [db.factRows, globalCountry]);

  const overviewSuppliers = useMemo(() => {
    if (globalCountry === "Overall") return suppliers;
    return suppliers.filter((s) => s.countries.includes(globalCountry));
  }, [globalCountry, suppliers]);

  const overviewSupplierByCode = useMemo(() => {
    const m = new Map<string, SupplierMaster>();
    for (const s of overviewSuppliers) m.set(s.code, s);
    return m;
  }, [overviewSuppliers]);

  const overviewCategoryMetrics = useMemo(
    () => buildCategoryMetrics(overviewFactRows, overviewSuppliers),
    [overviewFactRows, overviewSuppliers]
  );
  const overviewCountryMetrics = useMemo(
    () => buildCountryMetrics(overviewFactRows, overviewSuppliers),
    [overviewFactRows, overviewSuppliers]
  );

  useEffect(() => {
    if (!fastCategory && categories.length > 1) setFastCategory(categories[1]);
  }, [fastCategory, categories]);

  useEffect(() => {
    setDrillCategory(null);
    setDrillCountry(null);
    setDrillFunnelStage(null);
    setBarInsight(null);
  }, [globalCountry]);

  useEffect(() => {
    setBulkNameSelected(new Set());
  }, [bulkNameQuery]);

  useEffect(() => {
    if (selectedDuplicate) {
      setManualCanonicalName(selectedDuplicate.current);
    }
  }, [selectedDuplicate]);

  useEffect(() => {
    if (!availableCountries.includes(globalCountry)) {
      setGlobalCountry("Overall");
    }
  }, [availableCountries, globalCountry]);

  const kpis = useMemo(() => {
    const total = suppliers.length;
    const high = suppliers.filter((s) => s.risk === "High").length;
    const signed = suppliers.filter((s) => s.contractStatus === "Signed").length;
    const sent = suppliers.filter((s) => s.contractStatus === "Sent").length;
    const notEval = suppliers.filter((s) => !s.evaluated).length;
    const completion = high ? signed / high : 0;
    const dqDup = duplicates.length;
    const unknownInScope = issueBuckets.RISK_UNKNOWN_IN_SCOPE;
    const multiCat = issueBuckets.MULTI_CATEGORY_EXPOSURE;
    const nameMany = issueBuckets.NAME_MANY_CODES;
    const qualityTotal = issues.length;
    return { total, high, signed, sent, notEval, completion, dqDup, unknownInScope, multiCat, nameMany, qualityTotal };
  }, [suppliers, duplicates, issueBuckets, issues]);

  const overviewIssues = useMemo(() => {
    if (globalCountry === "Overall") return issues;
    const codes = new Set(overviewSuppliers.map((s) => s.code));
    return issues.filter((i) => (i.code ? codes.has(i.code) : true));
  }, [issues, globalCountry, overviewSuppliers]);

  const overviewIssueBuckets = useMemo(() => {
    const by: Record<IssueType, number> = {
      CODE_MANY_NAMES: 0,
      NAME_MANY_CODES: 0,
      MULTI_CATEGORY_EXPOSURE: 0,
      RISK_UNKNOWN_IN_SCOPE: 0,
      MISSING_SUBFAMILY_HIGH_SPEND: 0,
      POLICY_SUPPRESSED_BY_GUARDRAIL: 0,
    };
    for (const i of overviewIssues) by[i.type] += 1;
    return by;
  }, [overviewIssues]);

  const overviewKpis = useMemo(() => {
    const total = overviewSuppliers.length;
    const high = overviewSuppliers.filter((s) => s.risk === "High").length;
    const signed = overviewSuppliers.filter((s) => s.contractStatus === "Signed").length;
    const sentOnly = overviewSuppliers.filter((s) => s.contractStatus === "Sent").length;
    const sent = sentOnly + signed;
    const notEval = overviewSuppliers.filter((s) => !s.evaluated).length;
    const notEvalHighRisk = overviewSuppliers.filter((s) => !s.evaluated && s.risk === "High").length;
    const completion = high ? signed / high : 0;
    const dqDup = overviewIssues.filter((i) => i.type === "CODE_MANY_NAMES").length;
    const unknownInScope = overviewIssueBuckets.RISK_UNKNOWN_IN_SCOPE;
    const multiCat = overviewIssueBuckets.MULTI_CATEGORY_EXPOSURE;
    const nameMany = overviewIssueBuckets.NAME_MANY_CODES;
    const qualityTotal = overviewIssues.length;
    return { total, high, signed, sent, notEval, notEvalHighRisk, completion, dqDup, unknownInScope, multiCat, nameMany, qualityTotal };
  }, [overviewSuppliers, overviewIssues, overviewIssueBuckets]);

  const funnelData = useMemo(() => {
    const high = overviewSuppliers.filter((s) => s.risk === "High").length;
    const signedHigh = overviewSuppliers.filter((s) => s.risk === "High" && s.contractStatus === "Signed").length;
    const sent = overviewSuppliers.filter((s) => s.contractStatus === "Sent").length;
    const pending = Math.max(high - signedHigh, 0);
    return [
      { stage: "High risk", value: high },
      { stage: "Sent", value: sent },
      { stage: "Signed", value: signedHigh },
      { stage: "Pending", value: pending },
    ];
  }, [overviewSuppliers]);
  const funnelStages = useMemo(() => funnelData.map((stage) => stage.stage), [funnelData]);

  const topRiskyByCount = useMemo(() => {
    return [...categoryMetrics]
      .filter((c) => c.suppliers >= 10)
      .sort((a, b) =>
        categoryTopSort === "HighRiskDesc"
          ? b.highRiskSuppliers - a.highRiskSuppliers
          : b.suppliers - a.suppliers
      )
      .slice(0, 10);
  }, [categoryMetrics, categoryTopSort]);

  const topRiskyBySpend = useMemo(() => {
    return [...categoryMetrics]
      .filter((c) => Math.abs(c.totalSpend) > 0)
      .sort((a, b) =>
        categorySpendSort === "HighRiskDesc"
          ? b.highRiskSpend - a.highRiskSpend
          : b.totalSpend - a.totalSpend
      )
      .slice(0, 10);
  }, [categoryMetrics, categorySpendSort]);

  const topCountriesByCount = useMemo(() => {
    return [...overviewCountryMetrics].sort((a, b) => b.suppliers - a.suppliers);
  }, [overviewCountryMetrics]);

  const topCountriesBySpend = useMemo(() => {
    return [...countryMetrics].sort((a, b) => b.totalSpend - a.totalSpend).slice(0, 10);
  }, [countryMetrics]);
  const topCountriesByRiskShare = useMemo(() => {
    return [...countryMetrics]
      .filter((c) => c.suppliers >= 10)
      .sort((a, b) => b.riskShare - a.riskShare)
      .slice(0, 10);
  }, [countryMetrics]);

  const riskMix = useMemo(() => {
    const buckets: Record<RiskLevel, { suppliers: number; spend: number }> = {
      High: { suppliers: 0, spend: 0 },
      "Non-risk": { suppliers: 0, spend: 0 },
      Unknown: { suppliers: 0, spend: 0 },
    };
    for (const s of overviewSuppliers) {
      buckets[s.risk].suppliers += 1;
      buckets[s.risk].spend += s.totalSpend;
    }
    return (Object.keys(buckets) as RiskLevel[]).map((r) => ({
      risk: r,
      suppliers: buckets[r].suppliers,
      spend: buckets[r].spend,
    }));
  }, [overviewSuppliers]);
  const supplierByCode = useMemo(() => {
    const m = new Map<string, SupplierMaster>();
    for (const s of suppliers) m.set(s.code, s);
    return m;
  }, [suppliers]);

  const duplicateCandidateStats = useMemo(() => {
    const map = new Map<string, { totalPO: number; totalSpend: number; lastSeen?: number }>();
    for (const row of db.factRows) {
      const code = normCode(row.supplierCode);
      const name = safeStr(row.supplierName);
      if (!code || !name) continue;
      const key = getCandidateKey(code, name);
      const entry = map.get(key) ?? { totalPO: 0, totalSpend: 0 };
      entry.totalPO += Number(row.po ?? 0);
      entry.totalSpend += Number(row.spend ?? 0);
      const yearValue = Number(row.year);
      if (!Number.isNaN(yearValue)) {
        entry.lastSeen = entry.lastSeen ? Math.max(entry.lastSeen, yearValue) : yearValue;
      }
      map.set(key, entry);
    }
    return map;
  }, [db.factRows]);

  useEffect(() => {
    const code = normCode(taskSupplierCode);
    if (!code) return;
    if (!taskSupplierName) {
      const name = supplierByCode.get(code)?.canonicalName;
      if (name) setTaskSupplierName(name);
    }
    if (!taskContactEmail) {
      const contact = db.settings.supplierContacts?.[code];
      if (contact) setTaskContactEmail(contact);
    }
  }, [taskSupplierCode, taskSupplierName, taskContactEmail, supplierByCode, db.settings.supplierContacts]);
  const bulkNameMatchesList = useMemo(() => {
    const needle = bulkNameQuery.trim().toLowerCase();
    if (!needle) return [];
    return suppliers
      .filter((s) => s.canonicalName.toLowerCase().includes(needle) || s.code.toLowerCase().includes(needle))
      .slice(0, 200);
  }, [bulkNameQuery, suppliers]);

  // bulk previews (plan-before-apply)
  const bulkParsedCodes = useMemo(() => parseCodes(bulkCodes), [bulkCodes]);
  const bulkRiskPreview = useMemo(() => (bulkParsedCodes.length ? planBulkRisk(bulkParsedCodes, bulkRisk) : null), [bulkParsedCodes, bulkRisk, supplierByCode, db.overrides]);
  const bulkContractPreview = useMemo(
    () => (bulkParsedCodes.length ? planBulkContract(bulkParsedCodes, bulkContract) : null),
    [bulkParsedCodes, bulkContract, supplierByCode, db.overrides]
  );
  const bulkCanonicalPreview = useMemo(() => {
    const nm = safeStr(bulkCanonicalName);
    return bulkParsedCodes.length && nm ? planBulkCanonicalName(bulkParsedCodes, nm) : null;
  }, [bulkParsedCodes, bulkCanonicalName, supplierByCode, db.overrides]);
  const bulkEvaluatedPreviewState = useMemo(
    () => (bulkParsedCodes.length ? planBulkEvaluated(bulkParsedCodes, bulkEvaluated) : null),
    [bulkParsedCodes, bulkEvaluated, supplierByCode, db.overrides]
  );
  const bulkPreviewRows = useMemo(() => {
    return bulkParsedCodes.slice(0, 200).map((code) => {
      const s = supplierByCode.get(code);
      return {
        code,
        name: s?.canonicalName ?? "(not found)",
        risk: s?.risk ?? "Unknown",
        contract: s?.contractStatus ?? "Not sent",
      };
    });
  }, [bulkParsedCodes, supplierByCode]);

  const categoryFastPreview = useMemo(() => {
    const cat = safeStr(fastCategory);
    if (!cat) return { total: 0, eligible: 0, ambiguous: 0, eligibleCodes: [] as string[] };

    const inCat = suppliers.filter((s) => s.categories.includes(cat));
    const eligible = inCat.filter((s) => {
      const domOk = (s.dominantCategory || "") === cat && (s.dominantCategoryShare ?? 0) >= db.settings.dominanceShareThreshold;
      const multiOk = (s.secondCategoryShare ?? 0) < db.settings.multiCategorySecondShareThreshold;
      return domOk && multiOk;
    });
    return {
      total: inCat.length,
      eligible: eligible.length,
      ambiguous: Math.max(inCat.length - eligible.length, 0),
      eligibleCodes: eligible.map((s) => s.code),
    };
  }, [fastCategory, suppliers, db.settings.dominanceShareThreshold, db.settings.multiCategorySecondShareThreshold]);

  const sortedCategoryMetrics = useMemo(() => {
    const sorted = [...categoryMetrics];
    const dir = categoryTableSortState.dir === "asc" ? 1 : -1;
    sorted.sort((a, b) => {
      const key = categoryTableSortState.key;
      if (key === "category") return a.category.localeCompare(b.category) * dir;
      return (Number((a as any)[key] ?? 0) - Number((b as any)[key] ?? 0)) * dir;
    });
    return sorted;
  }, [categoryMetrics, categoryTableSortState]);

  const drillSuppliers = useMemo(() => {
    if (!drillCategory) return [] as any[];
    const by: Record<string, { spend: number; po: number }> = {};
    for (const r of scopedFactRows) {
      if (safeStr(r.category) !== drillCategory) continue;
      const code = normCode(r.supplierCode);
      if (!code) continue;
      by[code] = by[code] || { spend: 0, po: 0 };
      by[code].spend += r.spend ?? 0;
      by[code].po += r.po ?? 0;
    }

    return Object.entries(by)
      .map(([code, v]) => {
        const s = overviewSupplierByCode.get(code);
        return {
          code,
          canonicalName: s?.canonicalName ?? "(missing name)",
          risk: s?.risk ?? "Unknown",
          contractStatus: s?.contractStatus ?? "N/A",
          spendInSlice: v.spend,
          poInSlice: v.po,
          totalSpend: s?.totalSpend ?? v.spend,
          totalPO: s?.totalPO ?? v.po,
          categories: s?.categories ?? [drillCategory],
        };
      })
      .sort((a: any, b: any) => b.spendInSlice - a.spendInSlice)
      .slice(0, 500);
  }, [drillCategory, scopedFactRows, supplierByCode]);

  const drillSuppliersByCountry = useMemo(() => {
    if (!drillCountry) return [] as any[];
    const by: Record<string, { spend: number; po: number }> = {};
    for (const r of scopedFactRows) {
      const c = safeStr(r.country) || "(unknown)";
      if (c !== drillCountry) continue;
      const code = normCode(r.supplierCode);
      if (!code) continue;
      by[code] = by[code] || { spend: 0, po: 0 };
      by[code].spend += r.spend ?? 0;
      by[code].po += r.po ?? 0;
    }

    return Object.entries(by)
      .map(([code, v]) => {
        const s = overviewSupplierByCode.get(code);
        return {
          code,
          canonicalName: s?.canonicalName ?? "(missing name)",
          risk: s?.risk ?? "Unknown",
          contractStatus: s?.contractStatus ?? "N/A",
          spendInSlice: v.spend,
          poInSlice: v.po,
          totalSpend: s?.totalSpend ?? v.spend,
          totalPO: s?.totalPO ?? v.po,
          categories: s?.categories ?? [],
        };
      })
      .sort((a: any, b: any) => b.spendInSlice - a.spendInSlice)
      .slice(0, 500);
  }, [drillCountry, scopedFactRows, supplierByCode]);

  const drillSuppliersByFunnel = useMemo(() => {
    if (!drillFunnelStage) return [] as any[];
    let filtered = overviewSuppliers;
    if (drillFunnelStage === "High risk") {
      filtered = overviewSuppliers.filter((s) => s.risk === "High");
    } else if (drillFunnelStage === "Sent") {
      filtered = overviewSuppliers.filter((s) => s.contractStatus === "Sent");
    } else if (drillFunnelStage === "Signed") {
      filtered = overviewSuppliers.filter((s) => s.contractStatus === "Signed");
    } else if (drillFunnelStage === "Pending") {
      filtered = overviewSuppliers.filter((s) => s.risk === "High" && s.contractStatus !== "Signed");
    }

    return filtered.map((s) => ({
      code: s.code,
      canonicalName: s.canonicalName,
      risk: s.risk,
      contractStatus: s.contractStatus,
      spendInSlice: s.totalSpend,
      poInSlice: s.totalPO,
      totalSpend: s.totalSpend,
      totalPO: s.totalPO,
      categories: s.categories,
    }));
  }, [drillFunnelStage, overviewSuppliers]);

  const debouncedQuery = useDebounce(q, 300);

  const supplierFiltered = useMemo(() => {
    const needle = debouncedQuery.trim().toLowerCase();
    return suppliers
      .filter((s) => {
        if (riskFilter !== "All" && s.risk !== riskFilter) return false;
        if (evalFilter === "Evaluated" && !s.evaluated) return false;
        if (evalFilter === "Not evaluated" && s.evaluated) return false;
        if (contractFilter !== "All" && s.contractStatus !== contractFilter) return false;
        if (categoryFilter !== "All" && !s.categories.includes(categoryFilter)) return false;
        if (!needle) return true;
        return (
          s.code.toLowerCase().includes(needle) ||
          s.canonicalName.toLowerCase().includes(needle) ||
          s.categories.some((c) => c.toLowerCase().includes(needle)) ||
          s.families.some((f) => f.toLowerCase().includes(needle))
        );
      })
      .slice(0, 1500); // guardrail for UI
  }, [suppliers, debouncedQuery, riskFilter, evalFilter, contractFilter, categoryFilter]);

  const worklist = useMemo(() => {
    // Fast Action list: High risk and not signed
    const base = suppliers.filter((s) => s.risk === "High" && s.contractStatus !== "Signed");
    const sorted = [...base].sort((a, b) => {
      const dir = fastActionSort.dir === "asc" ? 1 : -1;
      if (fastActionSort.key === "name") return dir * a.canonicalName.localeCompare(b.canonicalName);
      if (fastActionSort.key === "contract") return dir * a.contractStatus.localeCompare(b.contractStatus);
      return dir * (a.totalSpend - b.totalSpend);
    });
    return sorted.slice(0, 200);
  }, [suppliers, fastActionSort]);



  const pushBarAnalysis = (title: string, details: Record<string, any>) => {
    setBarAnalysis({ title, details });
  };

  const setCategorySortKey = (key: typeof categoryTableSort.key) => {
    setCategoryTableSort((prev) => ({
      key,
      dir: prev.key === key ? (prev.dir === "asc" ? "desc" : "asc") : "desc",
    }));
  };

  const worklistSorted = useMemo(() => {
    const sorted = [...worklist];
    const dir = fastQueueSort.includes("Asc") ? 1 : -1;
    if (fastQueueSort.startsWith("Spend")) {
      sorted.sort((a, b) => (a.totalSpend - b.totalSpend) * dir);
    } else {
      sorted.sort((a, b) => a.canonicalName.localeCompare(b.canonicalName) * dir);
    }
    return sorted;
  }, [worklist, fastQueueSort]);

  const drillRows = useMemo(() => {
    const rows = (drillFunnelStage ? drillSuppliersByFunnel : drillCategory ? drillSuppliers : drillSuppliersByCountry) as any[];
    const sorted = [...rows];
    const dir = drillSortState.dir === "asc" ? 1 : -1;
    const orderRisk: Record<RiskLevel, number> = { High: 3, "Non-risk": 2, Unknown: 1 };
    sorted.sort((a, b) => {
      const key = drillSortState.key;
      if (key === "risk") return (orderRisk[a.risk as RiskLevel] - orderRisk[b.risk as RiskLevel]) * dir;
      if (key === "contractStatus") return String(a.contractStatus).localeCompare(String(b.contractStatus)) * dir;
      if (key === "canonicalName") return String(a.canonicalName).localeCompare(String(b.canonicalName)) * dir;
      if (key === "categories") return String(a.categories?.[0] ?? "").localeCompare(String(b.categories?.[0] ?? "")) * dir;
      return (Number(a[key] ?? 0) - Number(b[key] ?? 0)) * dir;
    });
    return sorted;
  }, [drillCategory, drillFunnelStage, drillSuppliers, drillSuppliersByCountry, drillSuppliersByFunnel, drillSortState]);

  function jumpToSupplier(code: string) {
    setQ(code);
    setActiveTab("suppliers");
  }

  const pushLog = useCallback((entry: Omit<LogEntry, "id" | "ts" | "actor">) => {
    const e: LogEntry = { id: uuid(), ts: nowIso(), actor, ...entry };
    setDB((prev) => ({ ...prev, log: [e, ...prev.log] }));
  }, [actor]);
  async function onImportXlsx(file: File) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });

    const sheetNames = wb.SheetNames;

    const findSheet = (predicate: (n: string) => boolean) => sheetNames.find(predicate);

    const rawName = findSheet((n) => n.toLowerCase() === "raw" || n.toLowerCase().includes("raw"));
    const metaName = findSheet((n) => n.toLowerCase() === "metadatabase" || n.toLowerCase().includes("metadata"));
    const reoName = findSheet((n) => n.toLowerCase().includes("risk evaluation overview"));
    const settingsName = findSheet((n) => n.toLowerCase() === "settings" || n.toLowerCase().startsWith("settings"));

    const tryParse = (name: string | undefined, required: string[]) => {
      if (!name) return null;
      const ws = wb.Sheets[name];
      if (!ws) return null;
      return parseFactSheet(ws, required);
    };

    // Raw is the operational truth table (already contains Grouping Supplier, Risk, Status, Spend, PO)
    const raw = tryParse(rawName, ["Grouping Supplier", "supplier_code_erp", "Europe Catman category", "FAMILY NAME", "Spend", "# PO", "Risk", "Status"]);
    const meta = tryParse(metaName, ["supplier_name", "supplier_code", "Europe Catman category", "FAMILY NAME", "Spend", "# PO"]);
    const reo = reoName && wb.Sheets[reoName] ? parseRiskEvalOverview(wb.Sheets[reoName]) : null;

    // Policy rules from Settings (Family risk list)
    const settingsRules = settingsName && wb.Sheets[settingsName] ? parseFamilyRiskRulesFromSettings(wb.Sheets[settingsName]) : [];

    const base = (raw && raw.length ? raw : meta && meta.length ? meta : reo && reo.length ? reo : []) as FactRow[];

    // Overlay risk/completion from Risk Evaluation Overview by Supplier Code (more stable than row-level matching)
    let merged = base;
    if (reo && reo.length && base.length) {
      const bucket: Record<string, Set<string>> = {};
      const completionBucket: Record<string, Set<string>> = {};

      for (const r of reo) {
        const code = normCode(r.supplierCode);
        if (!code) continue;
        bucket[code] = bucket[code] || new Set<string>();
        const rf = safeStr(r.riskFlag);
        if (rf) bucket[code].add(rf);

        const comp = safeStr(r.completion);
        if (comp) {
          completionBucket[code] = completionBucket[code] || new Set<string>();
          completionBucket[code].add(comp);
        }
      }

      const pickRisk = (set: Set<string>) => {
        const vals = Array.from(set).map((v) => safeStr(v).toUpperCase());
        if (vals.some((v) => v === "YES")) return "Yes";
        if (vals.some((v) => v === "NO")) return "No";
        if (vals.some((v) => v.includes("#REF"))) return "#REF!";
        // fallback (keep original casing from set)
        return Array.from(set)[0] ?? "";
      };

      const pickCompletion = (set: Set<string>) => Array.from(set)[0] ?? "";

      merged = base.map((r) => {
        const code = normCode(r.supplierCode);
        const riskSet = code ? bucket[code] : undefined;
        const compSet = code ? completionBucket[code] : undefined;

        return {
          ...r,
          supplierCode: code,
          riskFlag: riskSet && riskSet.size ? pickRisk(riskSet) : r.riskFlag,
          completion: compSet && compSet.size ? pickCompletion(compSet) : r.completion,
        };
      });
    } else {
      merged = base.map((r) => ({ ...r, supplierCode: normCode(r.supplierCode) }));
    }

    setDB((prev) => ({
      ...prev,
      factRows: merged,
      categoryRules: mergeCategoryRules(prev.categoryRules, settingsRules),
    }));

    pushLog({
      type: "IMPORT",
      summary: `Imported XLSX: ${file.name} (${merged.length} rows)`,
      details: {
        fileName: file.name,
        sheets: sheetNames,
        rows: merged.length,
        baseSheet: rawName && raw?.length ? rawName : metaName && meta?.length ? metaName : reoName ?? "unknown",
        importedFamilyRules: settingsRules.length,
      },
    });

    setActiveTab("overview");
  }


  function exportDB() {
    const payload = JSON.stringify(db, null, 2);
    const blob = new Blob([payload], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `supplier-risk-ops-db-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
    pushLog({ type: "RESET", summary: "Exported DB snapshot", details: { size: payload.length } });
  }
  async function importDBJson(file: File) {
    const txt = await file.text();
    const parsed = JSON.parse(txt);

    const isSnapshot =
      parsed &&
      typeof parsed === "object" &&
      !Array.isArray(parsed) &&
      ("factRows" in parsed || "overrides" in parsed || "categoryRules" in parsed || "tasks" in parsed || "log" in parsed);

    if (isSnapshot) {
      const next: PersistedDB = {
        factRows: Array.isArray((parsed as any).factRows) ? (parsed as any).factRows : [],
        overrides: (parsed as any).overrides ?? {},
        categoryRules: Array.isArray((parsed as any).categoryRules) ? (parsed as any).categoryRules : [],
        settings: normalizeSettings(((parsed as any).settings ?? {}) as LegacySettings),
        tasks: Array.isArray((parsed as any).tasks) ? (parsed as any).tasks : [],
        log: Array.isArray((parsed as any).log) ? (parsed as any).log : [],
        lastUndo: (parsed as any).lastUndo ?? null,
      };
      setDB(next);
      pushLog({ type: "IMPORT", summary: `Imported DB snapshot: ${file.name}`, details: { fileName: file.name } });
      setActiveTab("overview");
      return;
    }

    // Otherwise treat it as "Raw-like" data exported to JSON (array of objects)
    const rows = pickRowsFromUnknownJson(parsed);
    const factRows = normalizeJsonRowsToFactRows(rows);

    setDB((prev) => ({ ...prev, factRows }));
    pushLog({
      type: "IMPORT",
      summary: `Imported JSON rows: ${file.name} (${factRows.length} rows)`,
      details: { fileName: file.name, rows: factRows.length },
    });
    setActiveTab("overview");
  }


  function resetAll() {
    setDB({ factRows: [], overrides: {}, categoryRules: [], settings: DEFAULT_SETTINGS, tasks: [], log: [], lastUndo: null });
    pushLog({ type: "RESET", summary: "Reset dashboard database", details: {} });
    setActiveTab("overview");
  }

  function loadDemoData() {
    const entry: LogEntry = {
      id: uuid(),
      ts: nowIso(),
      actor,
      type: "IMPORT",
      summary: `Loaded demo dataset (${DEMO_FACT_ROWS.length} rows)`,
      details: { rows: DEMO_FACT_ROWS.length, countries: uniq(DEMO_FACT_ROWS.map((r) => r.country)).filter(Boolean) },
    };
    setDB({
      factRows: DEMO_FACT_ROWS,
      overrides: {},
      categoryRules: DEMO_RULES,
      settings: DEFAULT_SETTINGS,
      tasks: [],
      log: [entry],
      lastUndo: null,
    });
    setActiveTab("overview");
  }

  function parseCodes(text: string) {
    return uniq(
      text
        .split(/\s|,|;|\n|\t/)
        .map((x) => normCode(x))
        .filter(Boolean)
    );
  }

  function parseContactEntries(text: string) {
    return text
      .split(/\n+/)
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => {
        const [code, email] = line.split(/\s+/);
        return { code: code ?? "", email: email ?? "" };
      })
      .filter((entry) => entry.code && entry.email);
  }

  function getCandidateKey(code: string, name: string) {
    return `${normCode(code)}::${safeStr(name).toLowerCase()}`;
  }

  function handleToggleSupplierSelection(code: string) {
    setSelectedSuppliers((prev) => {
      const next = new Set(prev);
      if (next.has(code)) next.delete(code);
      else next.add(code);
      return next;
    });
  }

  function toggleSelectAllSuppliers() {
    setSelectedSuppliers((prev) => {
      const next = new Set(prev);
      const allVisible = supplierFiltered.map((s) => s.code);
      const hasAll = allVisible.every((code) => next.has(code));
      if (hasAll) {
        allVisible.forEach((code) => next.delete(code));
      } else {
        allVisible.forEach((code) => next.add(code));
      }
      return next;
    });
  }

  function handleClearSelectedSuppliers() {
    setSelectedSuppliers(new Set());
  }

  function handleToggleBulkNameSelection(code: string) {
    setBulkNameSelected((prev) => {
      const next = new Set(prev);
      if (next.has(code)) next.delete(code);
      else next.add(code);
      return next;
    });
  }

  function handleApplyBulkNameSelection() {
    const codes = uniq([...parseCodes(bulkCodes), ...Array.from(bulkNameSelected)]);
    setBulkCodes(codes.join("\n"));
    setBulkNameSelected(new Set());
  }

  function cleanOverride(ov?: SupplierOverride): SupplierOverride | undefined {
    if (!ov) return undefined;
    const out: SupplierOverride = {};
    const cn = safeStr(ov.canonicalName);
    if (cn) out.canonicalName = cn;
    if (ov.risk && ov.risk !== "Unknown") out.risk = ov.risk;
    if (ov.contractStatus && ov.contractStatus !== "N/A") out.contractStatus = ov.contractStatus;
    if (ov.evaluated !== undefined) out.evaluated = ov.evaluated;
    const note = safeStr(ov.note);
    if (note) out.note = note;
    return Object.keys(out).length ? out : undefined;
  }

  function planBulkRisk(codes: string[], target: RiskLevel) {
    const found: string[] = [];
    const notFound: string[] = [];
    const changes: { code: string; prev?: SupplierOverride; next?: SupplierOverride }[] = [];

    for (const c of codes) {
      const s = supplierByCode.get(c);
      if (!s) {
        notFound.push(c);
        continue;
      }
      found.push(c);
      const prev = db.overrides[c];
      const effectiveRisk = s.risk;

      if (target === "Unknown") {
        if (!prev?.risk) continue; // nothing to clear
        const patched = cleanOverride({ ...prev, risk: undefined });
        changes.push({ code: c, prev: prev ? { ...prev } : undefined, next: patched });
        continue;
      }

      if (prev?.risk === target) continue;
      // Avoid writing redundant overrides when the effective risk already equals the target.
      if (!prev?.risk && effectiveRisk === target) continue;

      const patched = cleanOverride({ ...(prev || {}), risk: target });
      changes.push({ code: c, prev: prev ? { ...prev } : undefined, next: patched });
    }

    return { found, notFound, changes };
  }

  function planBulkContract(codes: string[], target: ContractStatus) {
    const found: string[] = [];
    const notFound: string[] = [];
    const changes: { code: string; prev?: SupplierOverride; next?: SupplierOverride }[] = [];

    for (const c of codes) {
      const s = supplierByCode.get(c);
      if (!s) {
        notFound.push(c);
        continue;
      }
      found.push(c);
      const prev = db.overrides[c];
      const effective = s.contractStatus;

      if (target === "N/A") {
        if (!prev?.contractStatus) continue;
        const patched = cleanOverride({ ...prev, contractStatus: undefined });
        changes.push({ code: c, prev: prev ? { ...prev } : undefined, next: patched });
        continue;
      }

      if (prev?.contractStatus === target) continue;
      if (!prev?.contractStatus && effective === target) continue;

      const patched = cleanOverride({ ...(prev || {}), contractStatus: target });
      changes.push({ code: c, prev: prev ? { ...prev } : undefined, next: patched });
    }

    return { found, notFound, changes };
  }

  function planBulkCanonicalName(codes: string[], name: string) {
    const found: string[] = [];
    const notFound: string[] = [];
    const changes: { code: string; prev?: SupplierOverride; next?: SupplierOverride }[] = [];

    for (const c of codes) {
      const s = supplierByCode.get(c);
      if (!s) {
        notFound.push(c);
        continue;
      }
      found.push(c);
      const prev = db.overrides[c];
      const effective = s.canonicalName;

      if (prev?.canonicalName === name) continue;
      if (!safeStr(prev?.canonicalName) && effective === name) continue;

      const patched = cleanOverride({ ...(prev || {}), canonicalName: name });
      changes.push({ code: c, prev: prev ? { ...prev } : undefined, next: patched });
    }

    return { found, notFound, changes };
  }

  function planBulkEvaluated(codes: string[], target: EvaluationChoice) {
    const found: string[] = [];
    const notFound: string[] = [];
    const changes: { code: string; prev?: SupplierOverride; next?: SupplierOverride }[] = [];

    for (const c of codes) {
      const s = supplierByCode.get(c);
      if (!s) {
        notFound.push(c);
        continue;
      }
      found.push(c);
      const prev = db.overrides[c];
      const effective = s.evaluated;

      if (target === "Clear") {
        if (prev?.evaluated === undefined) continue;
        const patched = cleanOverride({ ...prev, evaluated: undefined });
        changes.push({ code: c, prev: prev ? { ...prev } : undefined, next: patched });
        continue;
      }

      const nextValue = target === "Evaluated";
      if (prev?.evaluated === nextValue) continue;
      if (prev?.evaluated === undefined && effective === nextValue) continue;

      const patched = cleanOverride({ ...(prev || {}), evaluated: nextValue });
      changes.push({ code: c, prev: prev ? { ...prev } : undefined, next: patched });
    }

    return { found, notFound, changes };
  }

  function planBulkMark(codes: string[], riskChoice: BulkRiskChoice, evalChoice: BulkEvalChoice) {
    const found: string[] = [];
    const notFound: string[] = [];
    const changes: { code: string; prev?: SupplierOverride; next?: SupplierOverride }[] = [];

    for (const c of codes) {
      const s = supplierByCode.get(c);
      if (!s) {
        notFound.push(c);
        continue;
      }
      found.push(c);
      const prev = db.overrides[c];
      let next: SupplierOverride | undefined = prev ? { ...prev } : {};

      if (riskChoice !== "No change") {
        if (riskChoice === "Unknown") {
          next = { ...(next || {}), risk: undefined };
        } else {
          next = { ...(next || {}), risk: riskChoice };
        }
      }

      if (evalChoice !== "No change") {
        if (evalChoice === "Clear") {
          next = { ...(next || {}), evaluated: undefined };
        } else {
          next = { ...(next || {}), evaluated: evalChoice === "Evaluated" };
        }
      }

      const cleaned = cleanOverride(next);
      const prevClean = cleanOverride(prev);
      if (JSON.stringify(prevClean ?? {}) === JSON.stringify(cleaned ?? {})) continue;
      changes.push({ code: c, prev: prev ? { ...prev } : undefined, next: cleaned });
    }

    return { found, notFound, changes };
  }

  function applyPlan(action: UndoBatch["action"], summary: string, plan: { changes: { code: string; prev?: SupplierOverride; next?: SupplierOverride }[] }) {
    if (!plan.changes.length) return false;

    const undo: UndoBatch = {
      id: uuid(),
      ts: nowIso(),
      actor: actor || "Unknown",
      action,
      summary,
      items: plan.changes.map((c) => ({ code: c.code, prev: c.prev ? { ...c.prev } : undefined })),
    };

    setDB((prev) => {
      const nextOverrides: Record<string, SupplierOverride> = { ...prev.overrides };
      for (const ch of plan.changes) {
        if (!ch.next) {
          delete nextOverrides[ch.code];
        } else {
          nextOverrides[ch.code] = ch.next;
        }
      }
      return { ...prev, overrides: nextOverrides, lastUndo: undo };
    });

    return true;
  }

  function undoLast() {
    const u = db.lastUndo;
    if (!u) return;

    setDB((prev) => {
      const nextOverrides: Record<string, SupplierOverride> = { ...prev.overrides };
      for (const it of u.items) {
        if (!it.prev) delete nextOverrides[it.code];
        else nextOverrides[it.code] = { ...it.prev };
      }
      return { ...prev, overrides: nextOverrides, lastUndo: null };
    });

    pushLog({
      type: "BULK_UPDATE",
      summary: `Undo: ${u.summary}`,
      details: { undoId: u.id, action: u.action, count: u.items.length, sample: u.items.slice(0, 20).map((x) => x.code) },
    });
  }

  function bulkApplyRisk() {
    const codes = parseCodes(bulkCodes);
    if (!codes.length) return;

    const plan = planBulkRisk(codes, bulkRisk);
    const didApply = applyPlan(
      "risk",
      `Bulk risk ${bulkRisk === "Unknown" ? "cleared" : "set"}: ${bulkRisk} for ${plan.changes.length} suppliers`,
      plan
    );

    pushLog({
      type: "BULK_UPDATE",
      summary: didApply
        ? `Bulk risk ${bulkRisk === "Unknown" ? "cleared" : "set"}: ${bulkRisk} for ${plan.changes.length} suppliers`
        : `Bulk risk: no changes (already compliant)`,
      details: {
        action: "risk",
        risk: bulkRisk,
        changed: plan.changes.length,
        found: plan.found.length,
        notFound: plan.notFound.length,
        notFoundSample: plan.notFound.slice(0, 25),
        sample: plan.changes.slice(0, 20).map((x) => x.code),
      },
    });
  }

  function bulkApplyContractStatus() {
    const codes = parseCodes(bulkCodes);
    if (!codes.length) return;

    const plan = planBulkContract(codes, bulkContract);
    const didApply = applyPlan(
      "contractStatus",
      `Bulk contract ${bulkContract === "N/A" ? "cleared" : "set"}: ${bulkContract} for ${plan.changes.length} suppliers`,
      plan
    );

    pushLog({
      type: "BULK_UPDATE",
      summary: didApply
        ? `Bulk contract ${bulkContract === "N/A" ? "cleared" : "set"}: ${bulkContract} for ${plan.changes.length} suppliers`
        : `Bulk contract: no changes (already compliant)`,
      details: {
        action: "contractStatus",
        status: bulkContract,
        changed: plan.changes.length,
        found: plan.found.length,
        notFound: plan.notFound.length,
        notFoundSample: plan.notFound.slice(0, 25),
        sample: plan.changes.slice(0, 20).map((x) => x.code),
      },
    });
  }

  function bulkApplyCanonicalName() {
    const codes = parseCodes(bulkCodes);
    const name = safeStr(bulkCanonicalName);
    if (!codes.length || !name) return;

    const plan = planBulkCanonicalName(codes, name);
    const didApply = applyPlan("canonicalName", `Bulk canonical name set for ${plan.changes.length} suppliers`, plan);

    pushLog({
      type: "BULK_UPDATE",
      summary: didApply ? `Bulk canonical name set for ${plan.changes.length} suppliers` : `Bulk canonical name: no changes`,
      details: {
        action: "canonicalName",
        name,
        changed: plan.changes.length,
        found: plan.found.length,
        notFound: plan.notFound.length,
        notFoundSample: plan.notFound.slice(0, 25),
        sample: plan.changes.slice(0, 20).map((x) => x.code),
      },
    });
  }

  function bulkApplyEvaluated() {
    const codes = parseCodes(bulkCodes);
    if (!codes.length) return;

    const plan = planBulkEvaluated(codes, bulkEvaluated);
    const didApply = applyPlan(
      "evaluated",
      `Bulk evaluated ${bulkEvaluated === "Clear" ? "cleared" : "set"} for ${plan.changes.length} suppliers`,
      plan
    );

    pushLog({
      type: "BULK_UPDATE",
      summary: didApply
        ? `Bulk evaluated ${bulkEvaluated === "Clear" ? "cleared" : "set"} for ${plan.changes.length} suppliers`
        : `Bulk evaluated: no changes`,
      details: {
        action: "evaluated",
        evaluated: bulkEvaluated,
        changed: plan.changes.length,
        found: plan.found.length,
        notFound: plan.notFound.length,
        notFoundSample: plan.notFound.slice(0, 25),
        sample: plan.changes.slice(0, 20).map((x) => x.code),
      },
    });
  }

  function bulkApplyMarkedSuppliers() {
    const codes = Array.from(selectedSuppliers);
    if (!codes.length) return;
    if (bulkMarkRisk === "No change" && bulkMarkEvaluated === "No change") return;

    const plan = planBulkMark(codes, bulkMarkRisk, bulkMarkEvaluated);
    const didApply = applyPlan(
      "bulkMark",
      `Bulk mark applied to ${plan.changes.length} suppliers`,
      plan
    );

    pushLog({
      type: "BULK_UPDATE",
      summary: didApply ? `Bulk mark applied to ${plan.changes.length} suppliers` : `Bulk mark: no changes`,
      details: {
        action: "bulkMark",
        risk: bulkMarkRisk,
        evaluated: bulkMarkEvaluated,
        changed: plan.changes.length,
        found: plan.found.length,
        notFound: plan.notFound.length,
        notFoundSample: plan.notFound.slice(0, 25),
        sample: plan.changes.slice(0, 20).map((x) => x.code),
      },
    });
  }

  function applyCategoryFastAction() {
    const cat = safeStr(fastCategory);
    if (!cat) return;

    const eligible = suppliers
      .filter((s) => s.categories.includes(cat))
      .filter((s) => {
        const domOk = (s.dominantCategory || "") === cat && (s.dominantCategoryShare ?? 0) >= db.settings.dominanceShareThreshold;
        const multiOk = (s.secondCategoryShare ?? 0) < db.settings.multiCategorySecondShareThreshold;
        return domOk && multiOk;
      })
      .map((s) => s.code);

    if (!eligible.length) {
      pushLog({ type: "BULK_UPDATE", summary: `Category FastAction: no eligible suppliers for ${cat}`, details: { category: cat } });
      return;
    }

    const plan = planBulkRisk(eligible, fastCategoryRisk);
    const didApply = applyPlan(
      "categoryFastAction",
      `Category FastAction: ${cat} -> ${fastCategoryRisk} for ${plan.changes.length} suppliers (eligible-only)`,
      plan
    );

    pushLog({
      type: "BULK_UPDATE",
      summary: didApply
        ? `Category FastAction: ${cat} -> ${fastCategoryRisk} for ${plan.changes.length} suppliers (eligible-only)`
        : `Category FastAction: no changes (already compliant)`,
      details: {
        action: "categoryFastAction",
        category: cat,
        risk: fastCategoryRisk,
        eligible: eligible.length,
        changed: plan.changes.length,
        sample: plan.changes.slice(0, 20).map((x) => x.code),
        thresholds: {
          dominanceShareThreshold: db.settings.dominanceShareThreshold,
          multiCategorySecondShareThreshold: db.settings.multiCategorySecondShareThreshold,
        },
      },
    });
  }

  const setCanonical = useCallback((code: string, name: string) => {
    setDB((prev) => ({
      ...prev,
      overrides: {
        ...prev.overrides,
        [code]: { ...(prev.overrides[code] || {}), canonicalName: name },
      },
    }));

    pushLog({
      type: "CANONICAL_DECISION",
      summary: `Canonical name fixed for ${code}`,
      details: { supplierCode: code, canonicalName: name },
    });
  }, [pushLog]);

  const setSupplierOverride = useCallback((code: string, patch: Partial<SupplierOverride>) => {
    setDB((prev) => ({
      ...prev,
      overrides: {
        ...prev.overrides,
        [code]: { ...(prev.overrides[code] || {}), ...patch },
      },
    }));

    pushLog({
      type: "SUPPLIER_OVERRIDE",
      summary: `Supplier override updated: ${code}`,
      details: { supplierCode: code, patch },
    });
  }, [pushLog]);

  type AppSettingsPatch = Omit<Partial<AppSettings>, "barChartSettings"> & {
    barChartSettings?: Partial<BarChartSettings>;
  };

  function setAppSettings(patch: AppSettingsPatch) {
    setDB((prev) => {
      const nextSettings: AppSettings = {
        ...prev.settings,
        ...patch,
        barChartSettings: {
          ...prev.settings.barChartSettings,
          ...(patch.barChartSettings ?? {}),
          palettes: {
            ...prev.settings.barChartSettings.palettes,
            ...(patch.barChartSettings?.palettes ?? {}),
          },
          metricColors: {
            ...prev.settings.barChartSettings.metricColors,
            ...(patch.barChartSettings?.metricColors ?? {}),
          },
        },
      };
      return { ...prev, settings: nextSettings };
    });
    pushLog({ type: "SETTINGS", summary: "Updated app settings", details: patch });
  }

  function updateBarPalette(kind: BarPaletteKey, keyRaw: string, color: string) {
    const key = safeStr(keyRaw);
    if (!key) return;
    const next = { ...(db.settings.barChartSettings.palettes[kind] ?? {}) };
    if (!color) delete next[key];
    else next[key] = color;
    setAppSettings({
      barChartSettings: {
        palettes: {
          ...db.settings.barChartSettings.palettes,
          [kind]: next,
        },
      },
    });
  }

  function clearBarPalette(kind: BarPaletteKey, keyRaw: string) {
    const key = safeStr(keyRaw);
    if (!key) return;
    const next = { ...(db.settings.barChartSettings.palettes[kind] ?? {}) };
    delete next[key];
    setAppSettings({
      barChartSettings: {
        palettes: {
          ...db.settings.barChartSettings.palettes,
          [kind]: next,
        },
      },
    });
  }

  function getBarColor(kind: BarPaletteKey, key: string, fallback: string) {
    return db.settings.barChartSettings.palettes[kind]?.[key] || fallback;
  }

  function getMetricColor(key: keyof BarChartSettings["metricColors"], fallback: string) {
    return db.settings.barChartSettings.metricColors?.[key] || fallback;
  }

  function getBarFill(kind: BarPaletteKey, key: string, fallback: string) {
    const color = getBarColor(kind, key, fallback);
    if (db.settings.barChartSettings.fillStyle === "solid") return color;
    return `url(#${kind}-${slugifyKey(key)}-${db.settings.barChartSettings.fillStyle})`;
  }

  function renderBarDefs(kind: BarPaletteKey, keys: string[], fallback: string) {
    const style = db.settings.barChartSettings.fillStyle;
    if (style === "solid") return null;
    const unique = uniq(keys);
    return (
      <defs>
        {unique.map((key) => {
          const color = getBarColor(kind, key, fallback);
          const id = `${kind}-${slugifyKey(key)}-${style}`;
          if (style === "dots") {
            return (
              <pattern key={id} id={id} width="6" height="6" patternUnits="userSpaceOnUse">
                <rect width="6" height="6" fill="transparent" />
                <circle cx="2" cy="2" r="1.3" fill={color} />
              </pattern>
            );
          }
          return (
            <pattern key={id} id={id} width="8" height="8" patternUnits="userSpaceOnUse" patternTransform="rotate(45)">
              <rect width="8" height="8" fill="transparent" />
              <line x1="0" y1="0" x2="0" y2="8" stroke={color} strokeWidth="2" />
            </pattern>
          );
        })}
      </defs>
    );
  }

  function addRule() {
    const key = safeStr(newRuleKey);
    if (!key) return;
    const rule: CategoryRule = {
      key,
      scope: newRuleScope,
      defaultRisk: newRuleRisk,
      comment: safeStr(newRuleComment) || undefined,
    };
    setDB((prev) => ({ ...prev, categoryRules: [rule, ...prev.categoryRules] }));
    pushLog({ type: "CATEGORY_RULE", summary: `Added ${newRuleScope} rule: ${key} -> ${newRuleRisk}`, details: rule });
    setNewRuleKey("");
    setNewRuleComment("");
  }

  function deleteRule(idx: number) {
    setDB((prev) => {
      const copy = [...prev.categoryRules];
      const removed = copy.splice(idx, 1)[0];
      pushLog({ type: "CATEGORY_RULE", summary: `Deleted rule: ${removed.scope} ${removed.key}`, details: removed });
      return { ...prev, categoryRules: copy };
    });
  }

  const addTaskEntry = useCallback((partial: Partial<Task> & { title: string }) => {
    const t: Task = {
      id: uuid(),
      date: partial.date ?? taskDate,
      owner: partial.owner ?? (safeStr(taskOwner) || actor),
      title: partial.title,
      note: safeStr(partial.note) || undefined,
      supplierCode: safeStr(partial.supplierCode) || undefined,
      supplierName: safeStr(partial.supplierName) || undefined,
      contactName: safeStr(partial.contactName) || undefined,
      contactEmail: safeStr(partial.contactEmail) || undefined,
      createdAt: partial.createdAt ?? nowIso(),
      dueDate: partial.dueDate ?? taskDueDate,
      kind: partial.kind ?? "general",
      followUpStage: partial.followUpStage,
      status: partial.status ?? "todo",
    };
    setDB((prev) => ({ ...prev, tasks: [t, ...prev.tasks] }));
    pushLog({ type: "TASK", summary: `Task added (${t.date}): ${t.title}`, details: t });
  }, [actor, taskDate, taskDueDate, taskOwner, taskContactName, taskContactEmail, db.settings.bulkContactEmail, db.settings.bulkContactName, pushLog]);

  const addTask = useCallback(() => {
    const title = safeStr(taskTitle);
    if (!title) return;
    const supplierName = safeStr(taskSupplierName);
    const resolvedTitle =
      supplierName && !title.toLowerCase().includes(supplierName.toLowerCase()) ? `${title} — ${supplierName}` : title;
    addTaskEntry({
      title: resolvedTitle,
      date: taskDate,
      note: taskNote,
      supplierCode: taskSupplierCode,
      supplierName: supplierName || taskSupplierName,
      contactName: taskContactName || db.settings.bulkContactName,
      contactEmail: taskContactEmail || db.settings.bulkContactEmail,
      dueDate: taskDueDate,
      kind: "general",
    });
    setTaskTitle("");
    setTaskNote("");
    setTaskSupplierCode("");
    setTaskSupplierName("");
  }, [
    addTaskEntry,
    taskTitle,
    taskSupplierName,
    taskDate,
    taskNote,
    taskSupplierCode,
    taskContactName,
    taskContactEmail,
    taskDueDate,
    db.settings.bulkContactEmail,
    db.settings.bulkContactName,
  ]);

  const addSupplierTask = useCallback(
    (task: {
      supplierCode: string;
      supplierName: string;
      title: string;
      note?: string;
      date: string;
      owner: string;
      dueDate: string;
      contactName?: string;
      contactEmail?: string;
    }) => {
      const title = safeStr(task.title);
      if (!title) return;
      const supplierName = safeStr(task.supplierName);
      const resolvedTitle =
        supplierName && !title.toLowerCase().includes(supplierName.toLowerCase()) ? `${title} — ${supplierName}` : title;
      addTaskEntry({
        title: resolvedTitle,
        date: task.date,
        note: task.note,
        supplierCode: task.supplierCode,
        supplierName: supplierName || task.supplierName,
        contactName: task.contactName || db.settings.bulkContactName,
        contactEmail: task.contactEmail || db.settings.bulkContactEmail,
        dueDate: task.dueDate,
        kind: "general",
        owner: task.owner,
      });
    },
    [addTaskEntry, db.settings.bulkContactEmail, db.settings.bulkContactName]
  );

  const updateTask = useCallback((id: string, patch: Partial<Task>) => {
    setDB((prev) => {
      const tasks = prev.tasks.map((t) => (t.id === id ? { ...t, ...patch } : t));
      const updated = tasks.find((t) => t.id === id);
      pushLog({ type: "TASK", summary: `Task updated: ${updated?.title}`, details: updated });
      return { ...prev, tasks };
    });
  }, [pushLog]);

  const scheduleEmailFollowUp = useCallback((options: {
    supplierCode?: string;
    supplierName?: string;
    contactName?: string;
    contactEmail?: string;
    stage: "initial" | "followup";
  }) => {
    const baseDate = new Date().toISOString().slice(0, 10);
    const followUpDays = emailAutomation.followUpDays ?? 7;
    const initialDays = emailAutomation.initialFollowUpDays ?? 14;
    const dueDate = options.stage === "initial" ? addDaysToDate(baseDate, initialDays) : addDaysToDate(baseDate, followUpDays);
    const title =
      options.stage === "initial"
        ? `Follow-up: signature in ${initialDays} days`
        : `Escalation check: signature in ${followUpDays} days`;
    addTaskEntry({
      title,
      date: dueDate,
      dueDate,
      kind: "email",
      followUpStage: options.stage,
      supplierCode: options.supplierCode,
      supplierName: options.supplierName,
      contactName: options.contactName || db.settings.bulkContactName,
      contactEmail: options.contactEmail || db.settings.bulkContactEmail,
      note:
        options.stage === "initial"
          ? `Email sent вЂ” reminder scheduled for ${initialDays} days.`
          : `Follow-up email sent вЂ” reminder scheduled for ${followUpDays} days.`,
    });
  }, [addTaskEntry, emailAutomation.followUpDays, emailAutomation.initialFollowUpDays, db.settings.bulkContactEmail, db.settings.bulkContactName]);

  const toggleTask = useCallback((id: string) => {
    setDB((prev) => {
      const tasks = prev.tasks.map((t) =>
        t.id === id
          ? { ...t, status: (t.status === "todo" ? "done" : "todo") as Task["status"] }
          : t,
      );
      const updated = tasks.find((t) => t.id === id);
      pushLog({ type: "TASK", summary: `Task updated: ${updated?.title}`, details: updated });
      return { ...prev, tasks };
    });
  }, [pushLog]);

  const deleteTask = useCallback((id: string) => {
    setDB((prev) => {
      const removed = prev.tasks.find((t) => t.id === id);
      const tasks = prev.tasks.filter((t) => t.id !== id);
      pushLog({ type: "TASK", summary: `Task deleted: ${removed?.title ?? id}`, details: removed });
      return { ...prev, tasks };
    });
  }, [pushLog]);

  useEffect(() => {
    const today = todayIso();
    const tasks = db.tasks;
    const contactMap = db.settings.supplierContacts ?? {};
    const byKey = new Map<string, Task>();
    for (const t of tasks) {
      if (t.kind === "contractFollowup" && t.supplierCode && t.followUpStage) {
        byKey.set(`${t.supplierCode}::${t.followUpStage}`, t);
      }
    }
    const signedCodes = new Set(allSuppliers.filter((s) => s.contractStatus === "Signed").map((s) => s.code));
    const toAdd: Task[] = [];
    const toUpdate: Task[] = [];

    const makeTask = (
      supplier: SupplierMaster,
      followUpStage: Task["followUpStage"],
      dueDate: string,
      title: string,
      note?: string
    ): Task => ({
      id: uuid(),
      date: dueDate,
      dueDate,
      createdAt: nowIso(),
      owner: actor,
      title,
      note,
      supplierCode: supplier.code,
      supplierName: supplier.canonicalName,
      contactEmail: contactMap[supplier.code],
      kind: "contractFollowup",
      followUpStage,
      status: "todo",
    });

    for (const supplier of allSuppliers) {
      if (supplier.contractStatus === "Sent") {
        const initialKey = `${supplier.code}::initial`;
        if (!byKey.has(initialKey)) {
          const dueDate = addDays(today, 14);
          toAdd.push(
            makeTask(
              supplier,
              "initial",
              dueDate,
              `Contract signature expected: ${supplier.canonicalName}`,
              "Auto: 14-day follow-up after first email."
            )
          );
        }
      }
    }

    for (const t of tasks) {
      if (t.kind === "contractFollowup" && t.supplierCode && signedCodes.has(t.supplierCode) && t.status !== "done") {
        toUpdate.push({ ...t, status: "done" });
      }
    }

    for (const supplier of allSuppliers) {
      if (supplier.contractStatus === "Signed") continue;
      const initial = byKey.get(`${supplier.code}::initial`);
      if (initial && (initial.dueDate ?? initial.date) < today && !byKey.has(`${supplier.code}::followup`)) {
        const dueDate = addDays(today, 7);
        toAdd.push(
          makeTask(
            supplier,
            "followup",
            dueDate,
            `Send reminder email: ${supplier.canonicalName}`,
            "Auto: signature overdue. Ignore Risk until resolved. Reminder required (7 days)."
          )
        );
      }
      const followup = byKey.get(`${supplier.code}::followup`);
      if (followup && (followup.dueDate ?? followup.date) < today && !byKey.has(`${supplier.code}::urgent`)) {
        const dueDate = addDays(today, 1);
        toAdd.push(
          makeTask(
            supplier,
            "urgent",
            dueDate,
            `Urgent call required: ${supplier.canonicalName}`,
            "Auto: second reminder overdue. РўСЂРµР±СѓРµС‚СЃСЏ СЃСЂРѕС‡РЅС‹Р№ Р·РІРѕРЅРѕРє."
          )
        );
      }
    }

    if (!toAdd.length && !toUpdate.length) return;
    setDB((prev) => {
      const updated = prev.tasks.map((t) => toUpdate.find((u) => u.id === t.id) ?? t);
      return { ...prev, tasks: [...toAdd, ...updated] };
    });
  }, [allSuppliers, db.settings.supplierContacts, db.tasks, actor]);

  const tasksByDate = useMemo(() => {
    const by: Record<string, Task[]> = {};
    for (const t of db.tasks) {
      by[t.date] = by[t.date] || [];
      by[t.date].push(t);
    }
    return by;
  }, [db.tasks]);
  const redIgnoranceAlerts = useMemo(
    () => db.tasks.filter((t) => t.kind === "contractFollowup" && t.followUpStage === "urgent" && t.status !== "done"),
    [db.tasks]
  );

  const calendarDays = useMemo(() => {
    // basic month grid for current taskDate month
    const d = new Date(taskDate + "T00:00:00");
    const year = d.getFullYear();
    const month = d.getMonth();
    const first = new Date(year, month, 1);
    const last = new Date(year, month + 1, 0);
    const startWeekday = first.getDay(); // 0 Sun
    const daysInMonth = last.getDate();

    const cells: { date: string; day: number; inMonth: boolean }[] = [];
    // pad previous
    for (let i = 0; i < startWeekday; i++) {
      const padDate = new Date(year, month, -(startWeekday - 1 - i));
      cells.push({ date: padDate.toISOString().slice(0, 10), day: padDate.getDate(), inMonth: false });
    }
    for (let day = 1; day <= daysInMonth; day++) {
      const dt = new Date(year, month, day);
      cells.push({ date: dt.toISOString().slice(0, 10), day, inMonth: true });
    }
    // pad next to complete weeks
    while (cells.length % 7 !== 0) {
      const lastCellDate = new Date(cells[cells.length - 1].date + "T00:00:00");
      const next = new Date(lastCellDate);
      next.setDate(lastCellDate.getDate() + 1);
      cells.push({ date: next.toISOString().slice(0, 10), day: next.getDate(), inMonth: false });
    }
    return cells;
  }, [taskDate]);

  const pctOfTotal = (value: number) => (overviewKpis.total ? Math.round((value / overviewKpis.total) * 100) : 0);
  const pctOfHighRisk = (value: number) => (overviewKpis.high ? Math.round((value / overviewKpis.high) * 100) : 0);
  const getFunnelColor = (stage: string) => {
    if (stage === "High risk") return getBarColor("riskLevel", "High", getMetricColor("highRisk", DEFAULT_RISK_PALETTE.High));
    if (stage === "Sent") return getBarColor("contractStatus", "Sent", DEFAULT_CONTRACT_PALETTE.Sent);
    if (stage === "Signed") return getBarColor("contractStatus", "Signed", DEFAULT_CONTRACT_PALETTE.Signed);
    if (stage === "Pending") return getBarColor("funnelStage", "Pending", getMetricColor("totalSuppliers", "#93c5fd"));
    return getBarColor("funnelStage", stage, getMetricColor("totalSuppliers", "#93c5fd"));
  };

  // -----------------------------
  // Render
  // -----------------------------

  return (
    <div className="min-h-screen bg-background">
      <div className="border-b bg-background/80 backdrop-blur">
        <div className="mx-auto flex max-w-7xl items-center justify-between gap-3 p-4">
          <div className="flex items-center gap-3">
            <div className="flex h-10 w-10 items-center justify-center rounded-2xl bg-muted">
              <ShieldAlert className="h-5 w-5" />
            </div>
            <div>
              <div className="text-lg font-semibold">Supplier Risk Ops Dashboard</div>
              <div className="text-xs text-muted-foreground">Single source of truth вЂў Bulk actions вЂў Audit log вЂў Calendar</div>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <div className="hidden items-center gap-2 md:flex">
              <span className="text-xs text-muted-foreground">Actor</span>
              <Input value={actor} onChange={(e) => setActor(e.target.value)} className="h-9 w-[160px]" />
            </div>
            <div className="hidden items-center gap-2 md:flex">
              <span className="text-xs text-muted-foreground">Country</span>
              <Select value={globalCountry} onValueChange={(v: any) => setGlobalCountry(v)}>
                <SelectTrigger className="h-9 w-[180px]">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  {availableCountries.map((country) => (
                    <SelectItem key={country} value={country}>
                      {country}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.json"
              className="hidden"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (!f) return;
                if (f.name.toLowerCase().endsWith(".json")) {
                  importDBJson(f);
                } else {
                  onImportXlsx(f);
                }
                e.currentTarget.value = "";
              }}
            />

            <Button variant="secondary" className="gap-2" onClick={() => fileInputRef.current?.click()}>
              <Upload className="h-4 w-4" /> Import XLSX/JSON
            </Button>
            <Button variant="outline" className="gap-2" onClick={exportDB}>
              <Download className="h-4 w-4" /> Export DB
            </Button>

            <Dialog>
              <DialogTrigger asChild>
                <Button variant="destructive" className="gap-2">
                  <Trash2 className="h-4 w-4" /> Reset
                </Button>
              </DialogTrigger>
              <DialogContent>
                <DialogHeader>
                  <DialogTitle>Reset local database?</DialogTitle>
                </DialogHeader>
                <div className="text-sm text-muted-foreground">
                  This will wipe imported data, overrides, rules, tasks, and log from localStorage.
                </div>
                <div className="flex justify-end gap-2">
                  <Button variant="outline" type="button">
                    Cancel
                  </Button>
                  <Button
                    variant="destructive"
                    onClick={() => {
                      resetAll();
                    }}
                  >
                    Confirm reset
                  </Button>
                </div>
              </DialogContent>
            </Dialog>
          </div>
        </div>
      </div>

      <div className="mx-auto max-w-7xl p-4">
        <Tabs value={activeTab} onValueChange={setActiveTab}>
          <TabsList className="grid w-full grid-cols-2 gap-2 md:grid-cols-7">
            <TabsTrigger value="overview" className="gap-2">
              <BarChart3 className="h-4 w-4" /> Overview
            </TabsTrigger>
            <TabsTrigger value="suppliers" className="gap-2">
              <Users className="h-4 w-4" /> Suppliers
            </TabsTrigger>
            <TabsTrigger value="categories" className="gap-2">
              <Database className="h-4 w-4" /> Categories
            </TabsTrigger>
            <TabsTrigger value="duplicates" className="gap-2">
              <Split className="h-4 w-4" /> Duplicates
            </TabsTrigger>
            <TabsTrigger value="bulk" className="gap-2">
              <FileUp className="h-4 w-4" /> Bulk
            </TabsTrigger>
            <TabsTrigger value="log" className="gap-2">
              <ClipboardList className="h-4 w-4" /> Log
            </TabsTrigger>
            <TabsTrigger value="calendar" className="gap-2">
              <CalendarDays className="h-4 w-4" /> Calendar
            </TabsTrigger>
          </TabsList>

          <TabsContent value="overview" className="mt-4 space-y-4">
            {activeTab === "overview" ? (
              db.factRows.length === 0 ? (
                <EmptyState title="No data loaded" subtitle="Import your XLSX tracker (recommended) or a DB snapshot (JSON)." />
              ) : (
                <>
                <div className="flex flex-wrap items-center justify-end gap-2">
                  <Dialog>
                    <DialogTrigger asChild>
                      <Button variant="outline" size="sm" className="h-9 w-9 p-0" aria-label="Bar chart settings">
                        <Settings className="h-4 w-4" />
                      </Button>
                    </DialogTrigger>
                    <DialogContent className="ml-auto mr-0 h-full max-h-none w-full max-w-xs rounded-none border-l">
                      <DialogHeader>
                        <DialogTitle>Bar chart settings</DialogTitle>
                      </DialogHeader>
                      <div className="text-sm text-muted-foreground">
                        Quick styling tweaks for the overview charts.
                      </div>
                      <div className="mt-4 space-y-4">
                        <div className="rounded-2xl border p-3">
                          <div className="text-xs text-muted-foreground">Bar fill style</div>
                          <div className="mt-2 flex items-center gap-2">
                            {[
                              {
                                value: "solid" as BarFillStyle,
                                label: "Solid",
                                style: { backgroundColor: "#93c5fd" },
                              },
                              {
                                value: "diagonal" as BarFillStyle,
                                label: "Diagonal",
                                style: {
                                  backgroundImage:
                                    "repeating-linear-gradient(45deg, #93c5fd 0, #93c5fd 2px, transparent 2px, transparent 6px)",
                                },
                              },
                              {
                                value: "dots" as BarFillStyle,
                                label: "Dots",
                                style: {
                                  backgroundImage: "radial-gradient(#93c5fd 1.4px, transparent 1.4px)",
                                  backgroundSize: "6px 6px",
                                },
                              },
                            ].map((opt) => (
                              <button
                                key={opt.value}
                                type="button"
                                className={`flex items-center gap-2 rounded-full border px-2 py-1 text-xs ${
                                  barChartSettings.fillStyle === opt.value ? "border-primary text-primary" : "text-muted-foreground"
                                }`}
                                onClick={() => setAppSettings({ barChartSettings: { fillStyle: opt.value } })}
                              >
                                <span className="h-4 w-4 rounded-sm border" style={opt.style} />
                                {opt.label}
                              </button>
                            ))}
                          </div>
                        </div>

                        <div className="rounded-2xl border p-3">
                          <div className="text-xs text-muted-foreground">Funnel palette</div>
                          <div className="mt-2 flex flex-wrap gap-2">
                            {funnelStages.map((stage) => (
                              <div key={stage} className="flex items-center gap-2 rounded-full border px-3 py-1 text-xs">
                                <span className="font-medium">{stage}</span>
                                <input
                                  type="color"
                                  value={getBarColor("funnelStage", stage, getFunnelColor(stage))}
                                  onChange={(e) => updateBarPalette("funnelStage", stage, e.target.value)}
                                />
                              </div>
                            ))}
                          </div>
                        </div>

                        <div className="rounded-2xl border p-3">
                          <div className="text-xs text-muted-foreground">Overview colors</div>
                          <div className="mt-2 flex flex-wrap gap-2">
                            {[
                              { key: "highRisk", label: "High risk" },
                              { key: "totalSuppliers", label: "Total suppliers" },
                              { key: "spend", label: "Spend" },
                            ].map((opt) => (
                              <div key={opt.key} className="flex items-center gap-2 rounded-full border px-3 py-1 text-xs">
                                <span className="font-medium">{opt.label}</span>
                                <input
                                  type="color"
                                  value={metricColors[opt.key as keyof typeof metricColors]}
                                  onChange={(e) =>
                                    setAppSettings({
                                      barChartSettings: {
                                        metricColors: {
                                          ...metricColors,
                                          [opt.key]: e.target.value,
                                        },
                                      },
                                    })
                                  }
                                />
                              </div>
                            ))}
                          </div>
                        </div>

                        <div className="rounded-2xl border p-3">
                          <div className="text-xs text-muted-foreground">Risk level palette</div>
                          <div className="mt-2 space-y-2">
                            {RISK_LEVEL_OPTIONS.map((risk) => (
                              <div key={risk} className="flex items-center justify-between gap-2 rounded-full border px-3 py-2 text-xs">
                                <div className="font-medium">{risk}</div>
                                <div className="flex items-center gap-2">
                                  <input
                                    type="color"
                                    value={getBarColor("riskLevel", risk, DEFAULT_RISK_PALETTE[risk])}
                                    onChange={(e) => updateBarPalette("riskLevel", risk, e.target.value)}
                                  />
                                  <Input
                                    value={getBarColor("riskLevel", risk, DEFAULT_RISK_PALETTE[risk])}
                                    onChange={(e) => updateBarPalette("riskLevel", risk, e.target.value)}
                                    className="h-8 w-[110px]"
                                  />
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>
                    </DialogContent>
                  </Dialog>
                </div>

                <div className="grid gap-3 md:grid-cols-5">
                  <Card>
                    <CardHeader className="pb-2">
                      <CardTitle className="flex items-center gap-2 text-base">
                        <Database className="h-4 w-4" /> Supplier base
                      </CardTitle>
                    </CardHeader>
                    <CardContent>
                      <div className="text-3xl font-semibold">{overviewKpis.total}</div>
                      <div className="mt-1 text-sm text-muted-foreground">Unique supplier codes</div>
                      <div className="mt-2 text-xs text-muted-foreground">
                        Not evaluated: {overviewKpis.notEvalHighRisk} вЂў{" "}
                        {overviewKpis.high ? Math.round((overviewKpis.notEvalHighRisk / overviewKpis.high) * 100) : 0}% of High risk
                      </div>
                    </CardContent>
                  </Card>
                  <Card>
                    <CardHeader className="pb-2">
                      <CardTitle className="flex items-center gap-2 text-base">
                        <ShieldAlert className="h-4 w-4" /> High risk suppliers
                      </CardTitle>
                    </CardHeader>
                    <CardContent>
                      <div className="text-3xl font-semibold">{overviewKpis.high}</div>
                      <div className="mt-1 text-sm text-muted-foreground">Marked as High risk</div>
                      <div className="mt-2 text-xs text-muted-foreground">{pctOfTotal(overviewKpis.high)}% of total suppliers</div>
                    </CardContent>
                  </Card>
                  <Card>
                    <CardHeader className="pb-2">
                      <CardTitle className="flex items-center gap-2 text-base">
                        <FileUp className="h-4 w-4" /> Agreement sent
                      </CardTitle>
                    </CardHeader>
                    <CardContent>
                      <div className="text-3xl font-semibold">{overviewKpis.sent}</div>
                      <div className="mt-1 text-sm text-muted-foreground">Sent or signed (High risk)</div>
                      <div className="mt-2 text-xs text-muted-foreground">{pctOfHighRisk(overviewKpis.sent)}% of High risk</div>
                    </CardContent>
                  </Card>
                  <Card>
                    <CardHeader className="pb-2">
                      <CardTitle className="flex items-center gap-2 text-base">
                        <CheckCircle2 className="h-4 w-4" /> Agreement signed
                      </CardTitle>
                    </CardHeader>
                    <CardContent>
                      <div className="text-3xl font-semibold">{overviewKpis.signed}</div>
                      <div className="mt-1 text-sm text-muted-foreground">Signed by High risk suppliers</div>
                      <div className="mt-2 text-xs text-muted-foreground">{pctOfHighRisk(overviewKpis.signed)}% of High risk</div>
                    </CardContent>
                  </Card>
                  <Card>
                    <CardHeader className="pb-2">
                      <CardTitle className="flex items-center gap-2 text-base">
                        <Split className="h-4 w-4" /> Quality alerts
                      </CardTitle>
                    </CardHeader>
                    <CardContent>
                      <div className="text-3xl font-semibold">{overviewKpis.qualityTotal}</div>
                      <div className="mt-1 text-sm text-muted-foreground">
                        Dup {overviewKpis.dqDup} вЂў Multi-cat {overviewKpis.multiCat} вЂў Unknown-in-scope {overviewKpis.unknownInScope}
                      </div>
                    </CardContent>
                  </Card>
                </div>

                <div className="grid gap-3 md:grid-cols-2">
                  <Card>
                    <CardHeader>
                      <CardTitle className="text-base">Funnel</CardTitle>
                    </CardHeader>
                    <CardContent className="h-[280px]">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={funnelData}>
                          {renderBarDefs("funnelStage", funnelStages, getMetricColor("totalSuppliers", "#93c5fd"))}
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis dataKey="stage" />
                          <YAxis />
                          <Tooltip />
                          <Bar
                            dataKey="value"
                            onClick={(d: any) =>
                              (() => {
                                const stage = d?.payload?.stage ?? "Unknown";
                                setDrillCategory(null);
                                setDrillCountry(null);
                                setDrillFunnelStage(stage);
                                setBarInsight({
                                  title: "Funnel stage",
                                  label: stage,
                                  details: [
                                    `Value: ${d?.payload?.value ?? 0}`,
                                    `Scope: ${globalCountry}`,
                                    "Source: Supplier Master derived from imported fact rows.",
                                  ],
                                });
                              })()
                            }
                          >
                            {funnelData.map((entry) => (
                              <Cell
                                key={`funnel-${entry.stage}`}
                                fill={getBarFill("funnelStage", entry.stage, getFunnelColor(entry.stage))}
                              />
                            ))}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </CardContent>
                  </Card>

                  <Card>
                    <CardHeader>
                      <CardTitle className="flex items-center justify-between text-base">
                        <span>Fast Action Queue (Top 200)</span>
                        <Select value={fastQueueSort} onValueChange={(v: any) => setFastQueueSort(v)}>
                          <SelectTrigger className="w-[190px]">
                            <SelectValue />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="SpendDesc">Spend (High в†’ Low)</SelectItem>
                            <SelectItem value="SpendAsc">Spend (Low в†’ High)</SelectItem>
                            <SelectItem value="NameAsc">Name (A в†’ Z)</SelectItem>
                            <SelectItem value="NameDesc">Name (Z в†’ A)</SelectItem>
                          </SelectContent>
                        </Select>
                      </CardTitle>
                    </CardHeader>
                    <CardContent>
                      <div className="text-sm text-muted-foreground">High risk and not signed вЂ” sorted by your selection.</div>
                      <div className="mt-3 max-h-[240px] overflow-auto rounded-xl border">
                        <Table>
                          <TableHeader>
                            <TableRow>
                              <TableHead>Supplier</TableHead>
                              <TableHead>Risk</TableHead>
                              <TableHead>Contract</TableHead>
                              <TableHead className="text-right">Spend</TableHead>
                            </TableRow>
                          </TableHeader>
                          <TableBody>
                            {worklistSorted.slice(0, 12).map((s) => (
                              <TableRow key={s.code}>
                                <TableCell>
                                  <button
                                    type="button"
                                    className="text-left font-medium text-primary hover:underline"
                                    onClick={() => jumpToSupplier(s.code)}
                                  >
                                    {s.canonicalName}
                                  </button>
                                  <div className="text-xs text-muted-foreground">{s.code}</div>
                                </TableCell>
                                <TableCell>
                                  <RiskBadge risk={s.risk} tone={getRiskColor(s.risk, DEFAULT_SETTINGS.statusPalette.risk[s.risk])} />
                                </TableCell>
                                <TableCell>
                                  <ContractBadge status={s.contractStatus} tone={getContractColor(s.contractStatus, DEFAULT_SETTINGS.statusPalette.contract[s.contractStatus])} />
                                </TableCell>
                                <TableCell className="text-right font-medium">{fmtMoney(s.totalSpend)}</TableCell>
                              </TableRow>
                            ))}
                          </TableBody>
                        </Table>
                      </div>
                      <div className="mt-3 flex items-center justify-between">
                        <div className="text-xs text-muted-foreground">Completion (Signed / High risk)</div>
                        <div className="text-sm font-semibold">{(overviewKpis.completion * 100).toFixed(0)}%</div>
                      </div>
                    </CardContent>
                  </Card>
                </div>

                <div className="grid gap-3 md:grid-cols-2">
                  <Card>
                    <CardHeader>
                      <CardTitle className="flex items-center justify-between text-base">
                        <span>Top risky categories (by suppliers)</span>
                        <Select value={categoryTopSort} onValueChange={(v: any) => setCategoryTopSort(v)}>
                          <SelectTrigger className="w-[190px]">
                            <SelectValue />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="HighRiskDesc">High risk (desc)</SelectItem>
                            <SelectItem value="TotalDesc">Total suppliers (desc)</SelectItem>
                          </SelectContent>
                        </Select>
                      </CardTitle>
                    </CardHeader>
                    <CardContent className="h-[320px]">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart
                          data={topRiskyByCount}
                          layout="vertical"
                          margin={{ left: 120 }}
                        >
                          {renderBarDefs("category", topRiskyByCount.map((c) => c.category), "#93c5fd")}
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis type="number" />
                          <YAxis dataKey="category" type="category" width={160} interval={0} tickFormatter={(v: any) => shortLabel(String(v))} />
                          <Tooltip />
                          <Legend />
                          <Bar
                            dataKey="highRiskSuppliers"
                            name="High risk suppliers"
                            onClick={(d: any) => {
                              const category = d?.payload?.category ?? null;
                              setDrillCategory(category);
                              setDrillFunnelStage(null);
                              setBarInsight({
                                title: "Category risk (by suppliers)",
                                label: category ?? "Unknown",
                                details: [
                                  `High risk suppliers: ${d?.payload?.highRiskSuppliers ?? 0}`,
                                  `Total suppliers: ${d?.payload?.suppliers ?? 0}`,
                                  `Scope: ${globalCountry}`,
                                ],
                              });
                            }}
                          >
                            {topRiskyByCount.map((entry) => (
                              <Cell
                                key={`risk-${entry.category}`}
                                fill={getBarFill("category", entry.category, getMetricColor("highRisk", "#93c5fd"))}
                              />
                            ))}
                          </Bar>
                          <Bar
                            dataKey="suppliers"
                            name="Total suppliers"
                            onClick={(d: any) => {
                              const category = d?.payload?.category ?? null;
                              setDrillCategory(category);
                              setDrillFunnelStage(null);
                              setBarInsight({
                                title: "Category volume (by suppliers)",
                                label: category ?? "Unknown",
                                details: [
                                  `Total suppliers: ${d?.payload?.suppliers ?? 0}`,
                                  `High risk suppliers: ${d?.payload?.highRiskSuppliers ?? 0}`,
                                  `Scope: ${globalCountry}`,
                                ],
                              });
                            }}
                          >
                            {topRiskyByCount.map((entry) => (
                              <Cell
                                key={`total-${entry.category}`}
                                fill={getBarFill("category", entry.category, getMetricColor("totalSuppliers", "#93c5fd"))}
                                fillOpacity={0.35}
                              />
                            ))}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </CardContent>
                  </Card>

                  <Card>
                    <CardHeader>
                      <CardTitle className="flex items-center justify-between text-base">
                        <span>Top risky categories (by spend)</span>
                        <Select value={categorySpendSort} onValueChange={(v: any) => setCategorySpendSort(v)}>
                          <SelectTrigger className="w-[190px]">
                            <SelectValue />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="HighRiskDesc">High risk spend (desc)</SelectItem>
                            <SelectItem value="TotalDesc">Total spend (desc)</SelectItem>
                          </SelectContent>
                        </Select>
                      </CardTitle>
                    </CardHeader>
                    <CardContent className="h-[320px]">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart
                          data={topRiskyBySpend}
                          layout="vertical"
                          margin={{ left: 60 }}
                        >
                          {renderBarDefs("category", topRiskyBySpend.map((c) => c.category), "#93c5fd")}
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis type="number" tickFormatter={(v: any) => fmtMoney(Number(v))} />
                          <YAxis dataKey="category" type="category" width={160} interval={0} tickFormatter={(v: any) => shortLabel(String(v))} />
                          <Tooltip formatter={(v: any) => fmtMoney(Number(v))} />
                          <Legend />
                          <Bar
                            dataKey="highRiskSpend"
                            name="High risk spend"
                            onClick={(d: any) => {
                              const category = d?.payload?.category ?? null;
                              setDrillCategory(category);
                              setDrillFunnelStage(null);
                              setBarInsight({
                                title: "Category risk (by spend)",
                                label: category ?? "Unknown",
                                details: [
                                  `High risk spend: ${fmtMoney(Number(d?.payload?.highRiskSpend ?? 0))}`,
                                  `Total spend: ${fmtMoney(Number(d?.payload?.totalSpend ?? 0))}`,
                                  `Scope: ${globalCountry}`,
                                ],
                              });
                            }}
                          >
                            {topRiskyBySpend.map((entry) => (
                              <Cell
                                key={`risk-${entry.category}`}
                                fill={getBarFill("category", entry.category, getMetricColor("highRisk", "#93c5fd"))}
                              />
                            ))}
                          </Bar>
                          <Bar
                            dataKey="totalSpend"
                            name="Total spend"
                            onClick={(d: any) => {
                              const category = d?.payload?.category ?? null;
                              setDrillCategory(category);
                              setDrillFunnelStage(null);
                              setBarInsight({
                                title: "Category volume (by spend)",
                                label: category ?? "Unknown",
                                details: [
                                  `Total spend: ${fmtMoney(Number(d?.payload?.totalSpend ?? 0))}`,
                                  `High risk spend: ${fmtMoney(Number(d?.payload?.highRiskSpend ?? 0))}`,
                                  `Scope: ${globalCountry}`,
                                ],
                              });
                            }}
                          >
                            {topRiskyBySpend.map((entry) => (
                              <Cell
                                key={`total-${entry.category}`}
                                fill={getBarFill("category", entry.category, getMetricColor("spend", "#93c5fd"))}
                                fillOpacity={0.35}
                              />
                            ))}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </CardContent>
                  </Card>
                </div>

                <div className="grid gap-3 md:grid-cols-3">
                  <Card className="md:col-span-2">
                    <CardHeader>
                      <CardTitle className="text-base">Country split (suppliers)</CardTitle>
                    </CardHeader>
                    <CardContent className="h-[320px]">
                      <div className="mb-2 text-xs text-muted-foreground">Click a country bar to drill down to suppliers.</div>
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart
                          data={topCountriesByCount}
                          margin={{ left: 10, right: 10 }}
                        >
                          {renderBarDefs("country", topCountriesByCount.map((c) => c.country), "#93c5fd")}
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis dataKey="country" />
                          <YAxis />
                          <Tooltip />
                          <Legend />
                          <Bar
                            dataKey="highRiskSuppliers"
                            name="High risk"
                            onClick={(d: any) => {
                              const country = d?.payload?.country ?? null;
                              setDrillCountry(country);
                              setDrillFunnelStage(null);
                              setBarInsight({
                                title: "Country risk (by suppliers)",
                                label: country ?? "Unknown",
                                details: [
                                  `High risk suppliers: ${d?.payload?.highRiskSuppliers ?? 0}`,
                                  `Total suppliers: ${d?.payload?.suppliers ?? 0}`,
                                  `Scope: ${globalCountry}`,
                                ],
                              });
                            }}
                          >
                            {topCountriesByCount.map((entry) => (
                              <Cell key={`risk-${entry.country}`} fill={getBarFill("country", entry.country, getMetricColor("highRisk", "#fca5a5"))} />
                            ))}
                          </Bar>
                          <Bar
                            dataKey="suppliers"
                            name="Total"
                            onClick={(d: any) => {
                              const country = d?.payload?.country ?? null;
                              setDrillCountry(country);
                              setDrillFunnelStage(null);
                              setBarInsight({
                                title: "Country volume (by suppliers)",
                                label: country ?? "Unknown",
                                details: [
                                  `Total suppliers: ${d?.payload?.suppliers ?? 0}`,
                                  `High risk suppliers: ${d?.payload?.highRiskSuppliers ?? 0}`,
                                  `Scope: ${globalCountry}`,
                                ],
                              });
                            }}
                          >
                            {topCountriesByCount.map((entry) => (
                              <Cell
                                key={`total-${entry.country}`}
                                fill={getBarFill("country", entry.country, getMetricColor("totalSuppliers", "#93c5fd"))}
                                fillOpacity={0.4}
                              />
                            ))}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </CardContent>
                  </Card>

                  <Card>
                    <CardHeader>
                      <CardTitle className="text-base">Risk mix</CardTitle>
                    </CardHeader>
                    <CardContent className="h-[320px]">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={riskMix}>
                          {renderBarDefs("riskLevel", riskMix.map((r) => r.risk), "#93c5fd")}
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis dataKey="risk" />
                          <YAxis />
                          <Tooltip formatter={(v: any, n: any) => (n === "spend" ? fmtMoney(Number(v)) : v)} />
                          <Legend />
                          <Bar
                            dataKey="suppliers"
                            name="Suppliers"
                            onClick={(d: any) =>
                              setBarInsight({
                                title: "Risk mix (suppliers)",
                                label: d?.payload?.risk ?? "Unknown",
                                details: [
                                  `Suppliers: ${d?.payload?.suppliers ?? 0}`,
                                  `Spend: ${fmtMoney(Number(d?.payload?.spend ?? 0))}`,
                                  `Scope: ${globalCountry}`,
                                ],
                              })
                            }
                          >
                            {riskMix.map((entry) => (
                              <Cell
                                key={`risk-suppliers-${entry.risk}`}
                                fill={getBarFill("riskLevel", entry.risk, DEFAULT_RISK_PALETTE[entry.risk as RiskLevel] ?? "#e5e7eb")}
                              />
                            ))}
                          </Bar>
                          <Bar
                            dataKey="spend"
                            name="Spend"
                            onClick={(d: any) =>
                              setBarInsight({
                                title: "Risk mix (spend)",
                                label: d?.payload?.risk ?? "Unknown",
                                details: [
                                  `Spend: ${fmtMoney(Number(d?.payload?.spend ?? 0))}`,
                                  `Suppliers: ${d?.payload?.suppliers ?? 0}`,
                                  `Scope: ${globalCountry}`,
                                ],
                              })
                            }
                          >
                            {riskMix.map((entry) => (
                              <Cell
                                key={`risk-spend-${entry.risk}`}
                                fill={getBarFill("riskLevel", entry.risk, DEFAULT_RISK_PALETTE[entry.risk as RiskLevel] ?? "#93c5fd")}
                                fillOpacity={0.5}
                              />
                            ))}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </CardContent>
                  </Card>
                </div>

                <div className="grid gap-3 md:grid-cols-2">
                  <Card>
                    <CardHeader>
                      <CardTitle className="text-base">Country split (spend)</CardTitle>
                    </CardHeader>
                    <CardContent className="h-[320px]">
                      <div className="mb-2 text-xs text-muted-foreground">Top countries by spend. Click to drill down.</div>
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart
                          data={topCountriesBySpend}
                          margin={{ left: 10, right: 10 }}
                        >
                          {renderBarDefs("country", topCountriesBySpend.map((c) => c.country), "#93c5fd")}
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis dataKey="country" />
                          <YAxis tickFormatter={(v: any) => fmtMoney(Number(v))} />
                          <Tooltip formatter={(v: any) => fmtMoney(Number(v))} />
                          <Legend />
                          <Bar
                            dataKey="highRiskSpend"
                            name="High risk spend"
                            onClick={(d: any) => {
                              const country = d?.payload?.country ?? null;
                              setDrillCountry(country);
                              setDrillFunnelStage(null);
                              setBarInsight({
                                title: "Country risk (by spend)",
                                label: country ?? "Unknown",
                                details: [
                                  `High risk spend: ${fmtMoney(Number(d?.payload?.highRiskSpend ?? 0))}`,
                                  `Total spend: ${fmtMoney(Number(d?.payload?.totalSpend ?? 0))}`,
                                  `Scope: ${globalCountry}`,
                                ],
                              });
                            }}
                          >
                            {topCountriesBySpend.map((entry) => (
                              <Cell key={`risk-${entry.country}`} fill={getBarFill("country", entry.country, getMetricColor("highRisk", "#fca5a5"))} />
                            ))}
                          </Bar>
                          <Bar
                            dataKey="totalSpend"
                            name="Total spend"
                            onClick={(d: any) => {
                              const country = d?.payload?.country ?? null;
                              setDrillCountry(country);
                              setDrillFunnelStage(null);
                              setBarInsight({
                                title: "Country volume (by spend)",
                                label: country ?? "Unknown",
                                details: [
                                  `Total spend: ${fmtMoney(Number(d?.payload?.totalSpend ?? 0))}`,
                                  `High risk spend: ${fmtMoney(Number(d?.payload?.highRiskSpend ?? 0))}`,
                                  `Scope: ${globalCountry}`,
                                ],
                              });
                            }}
                          >
                            {topCountriesBySpend.map((entry) => (
                              <Cell
                                key={`total-${entry.country}`}
                                fill={getBarFill("country", entry.country, getMetricColor("spend", "#93c5fd"))}
                                fillOpacity={0.4}
                              />
                            ))}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </CardContent>
                  </Card>

                  <Card>
                    <CardHeader>
                      <CardTitle className="text-base">Country risk share</CardTitle>
                    </CardHeader>
                    <CardContent className="h-[320px]">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart
                          data={topCountriesByRiskShare}
                          margin={{ left: 10, right: 10 }}
                        >
                          {renderBarDefs("country", topCountriesByRiskShare.map((c) => c.country), "#fca5a5")}
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis dataKey="country" />
                          <YAxis tickFormatter={(v: any) => `${Math.round(Number(v) * 100)}%`} />
                          <Tooltip formatter={(v: any) => `${Math.round(Number(v) * 100)}%`} />
                          <Bar
                            dataKey="riskShare"
                            name="High risk share"
                            onClick={(d: any) => {
                              const country = d?.payload?.country ?? null;
                              setDrillCountry(country);
                              setDrillFunnelStage(null);
                              setBarInsight({
                                title: "Country risk share",
                                label: country ?? "Unknown",
                                details: [
                                  `High risk share: ${Math.round(Number(d?.payload?.riskShare ?? 0) * 100)}%`,
                                  `Suppliers: ${d?.payload?.suppliers ?? 0}`,
                                  `Scope: ${globalCountry}`,
                                ],
                              });
                            }}
                          >
                            {topCountriesByRiskShare.map((entry) => (
                              <Cell key={`share-${entry.country}`} fill={getBarFill("country", entry.country, getMetricColor("highRisk", "#fca5a5"))} />
                            ))}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </CardContent>
                  </Card>
                </div>

                {barInsight ? (
                  <Card>
                    <CardHeader>
                      <CardTitle className="flex items-center justify-between text-base">
                        <span>Bar analysis: {barInsight.label}</span>
                        <Button variant="outline" size="sm" onClick={() => setBarInsight(null)}>
                          Clear
                        </Button>
                      </CardTitle>
                    </CardHeader>
                    <CardContent>
                      <div className="text-sm text-muted-foreground">{barInsight.title}</div>
                      <ul className="mt-2 list-disc space-y-1 pl-5 text-sm">
                        {barInsight.details.map((d, idx) => (
                          <li key={`${d}-${idx}`}>{d}</li>
                        ))}
                      </ul>
                    </CardContent>
                  </Card>
                ) : null}

                {(drillCategory || drillCountry || drillFunnelStage) ? (
                  <Card>
                    <CardHeader>
                      <CardTitle className="flex items-center justify-between text-base">
                        <span>
                          Drill-down:{" "}
                          {drillFunnelStage
                            ? `Funnel — ${drillFunnelStage}`
                            : drillCategory
                              ? `Category — ${drillCategory}`
                              : `Country — ${drillCountry}`}
                        </span>
                        <Button
                          variant="outline"
                          size="sm"
                          onClick={() => {
                            setDrillCategory(null);
                            setDrillCountry(null);
                            setDrillFunnelStage(null);
                          }}
                        >
                          Clear
                        </Button>
                      </CardTitle>
                    </CardHeader>
                    <CardContent>
                      <div className="text-xs text-muted-foreground">Clicked chart в†’ suppliers table (Top 500 by spend).</div>
                      <div className="mt-3 max-h-[420px] overflow-auto rounded-2xl border">
                        <Table>
                          <TableHeader>
                            <TableRow>
                              <TableHead>
                                  <button
                                    type="button"
                                    className="flex items-center gap-1"
                                    onClick={() =>
                                    setDrillSortState((prev) => ({
                                      key: "canonicalName",
                                      dir: prev.key === "canonicalName" && prev.dir === "asc" ? "desc" : "asc",
                                    }))
                                  }
                                >
                                  Supplier
                                </button>
                              </TableHead>
                              <TableHead>
                                  <button
                                    type="button"
                                    className="flex items-center gap-1"
                                    onClick={() =>
                                    setDrillSortState((prev) => ({
                                      key: "risk",
                                      dir: prev.key === "risk" && prev.dir === "asc" ? "desc" : "asc",
                                    }))
                                  }
                                >
                                  Risk
                                </button>
                              </TableHead>
                              <TableHead>
                                  <button
                                    type="button"
                                    className="flex items-center gap-1"
                                    onClick={() =>
                                    setDrillSortState((prev) => ({
                                      key: "contractStatus",
                                      dir: prev.key === "contractStatus" && prev.dir === "asc" ? "desc" : "asc",
                                    }))
                                  }
                                >
                                  Contract
                                </button>
                              </TableHead>
                              <TableHead className="text-right">
                                  <button
                                    type="button"
                                    className="ml-auto flex items-center gap-1"
                                    onClick={() =>
                                    setDrillSortState((prev) => ({
                                      key: "poInSlice",
                                      dir: prev.key === "poInSlice" && prev.dir === "asc" ? "desc" : "asc",
                                    }))
                                  }
                                >
                                  PO (slice)
                                </button>
                              </TableHead>
                              <TableHead className="text-right">
                                  <button
                                    type="button"
                                    className="ml-auto flex items-center gap-1"
                                    onClick={() =>
                                    setDrillSortState((prev) => ({
                                      key: "spendInSlice",
                                      dir: prev.key === "spendInSlice" && prev.dir === "asc" ? "desc" : "asc",
                                    }))
                                  }
                                >
                                  Spend (slice)
                                </button>
                              </TableHead>
                              <TableHead className="text-right">
                                  <button
                                    type="button"
                                    className="ml-auto flex items-center gap-1"
                                    onClick={() =>
                                    setDrillSortState((prev) => ({
                                      key: "totalSpend",
                                      dir: prev.key === "totalSpend" && prev.dir === "asc" ? "desc" : "asc",
                                    }))
                                  }
                                >
                                  Overall spend
                                </button>
                              </TableHead>
                              <TableHead>
                                  <button
                                    type="button"
                                    className="flex items-center gap-1"
                                    onClick={() =>
                                    setDrillSortState((prev) => ({
                                      key: "categories",
                                      dir: prev.key === "categories" && prev.dir === "asc" ? "desc" : "asc",
                                    }))
                                  }
                                >
                                  Category
                                </button>
                              </TableHead>
                            </TableRow>
                          </TableHeader>
                          <TableBody>
                            {drillRows.map((s) => (
                              <TableRow key={s.code}>
                                <TableCell>
                                  <div className="font-medium">{s.canonicalName}</div>
                                  <div className="text-xs text-muted-foreground">{s.code}</div>
                                </TableCell>
                                <TableCell>
                                  <RiskBadge risk={s.risk} tone={getRiskColor(s.risk, DEFAULT_SETTINGS.statusPalette.risk[s.risk])} />
                                </TableCell>
                                <TableCell>
                                  <ContractBadge status={s.contractStatus} tone={getContractColor(s.contractStatus, DEFAULT_SETTINGS.statusPalette.contract[s.contractStatus])} />
                                </TableCell>
                                <TableCell className="text-right">{Math.round(s.poInSlice)}</TableCell>
                                <TableCell className="text-right font-medium">{fmtMoney(s.spendInSlice)}</TableCell>
                                <TableCell className="text-right text-muted-foreground">{fmtMoney(s.totalSpend)}</TableCell>
                                <TableCell className="min-w-[280px]">
                                  {drillCategory ? (
                                    <Badge variant="outline">{drillCategory}</Badge>
                                  ) : (
                                    <div className="flex flex-wrap gap-1">
                                      {s.categories.slice(0, 2).map((c: string) => (
                                        <Badge key={c} variant="outline">{c}</Badge>
                                      ))}
                                      {s.categories.length > 2 ? <Badge variant="outline">+{s.categories.length - 2}</Badge> : null}
                                    </div>
                                  )}
                                </TableCell>
                              </TableRow>
                            ))}
                          </TableBody>
                        </Table>
                      </div>
                    </CardContent>
                  </Card>
                ) : null}

                {barAnalysis ? (
                  <Card>
                    <CardHeader>
                      <CardTitle className="flex items-center justify-between text-base">
                        <span>Bar analysis вЂ” {barAnalysis.title}</span>
                        <Button variant="outline" size="sm" onClick={() => setBarAnalysis(null)}>
                          Clear
                        </Button>
                      </CardTitle>
                    </CardHeader>
                    <CardContent>
                      <div className="grid gap-2 md:grid-cols-2">
                        {Object.entries(barAnalysis.details).map(([key, value]) => (
                          <div key={key} className="rounded-2xl border p-3">
                            <div className="text-xs text-muted-foreground">{key}</div>
                            <div className="text-sm font-medium">
                              {typeof value === "number"
                                ? key.toLowerCase().includes("spend")
                                  ? fmtMoney(value)
                                  : key.toLowerCase().includes("share")
                                    ? `${Math.round(value * 100)}%`
                                    : value
                                : String(value)}
                            </div>
                          </div>
                        ))}
                      </div>
                    </CardContent>
                  </Card>
                ) : null}
                </>
              )
            ) : null}
          </TabsContent>

          <TabsContent value="suppliers" className="mt-4 space-y-4">
            {activeTab === "suppliers" ? (
              db.factRows.length === 0 ? (
                <EmptyState title="Import data first" subtitle="Suppliers tab is powered by the Supplier Master derived from your XLSX." />
              ) : (
                <>
                <Card>
                  <CardHeader>
                    <CardTitle className="flex items-center justify-between text-base">
                      <span className="flex items-center gap-2">
                        <Users className="h-4 w-4" /> Supplier Workbench
                      </span>
                      <div className="text-xs text-muted-foreground">Showing up to 1500 rows (performance guardrail)</div>
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="grid gap-2 md:grid-cols-5">
                      <div className="md:col-span-2">
                        <div className="relative">
                          <Search className="absolute left-2 top-2.5 h-4 w-4 text-muted-foreground" />
                          <Input
                            value={q}
                            onChange={(e) => setQ(e.target.value)}
                            placeholder="Search supplier code, name, category, familyвЂ¦"
                            className="pl-8"
                          />
                        </div>
                      </div>

                      <Select value={riskFilter} onValueChange={(v: any) => setRiskFilter(v)}>
                        <SelectTrigger>
                          <SelectValue placeholder="Risk" />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="All">All risks</SelectItem>
                          <SelectItem value="High">High</SelectItem>
                          <SelectItem value="Non-risk">Non-risk</SelectItem>
                          <SelectItem value="Unknown">Unknown</SelectItem>
                        </SelectContent>
                      </Select>

                      <Select value={evalFilter} onValueChange={(v: any) => setEvalFilter(v)}>
                        <SelectTrigger>
                          <SelectValue placeholder="Evaluation" />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="All">All</SelectItem>
                          <SelectItem value="Evaluated">Evaluated</SelectItem>
                          <SelectItem value="Not evaluated">Not evaluated</SelectItem>
                        </SelectContent>
                      </Select>

                      <Select value={categoryFilter} onValueChange={(v: any) => setCategoryFilter(v)}>
                        <SelectTrigger>
                          <SelectValue placeholder="Category" />
                        </SelectTrigger>
                        <SelectContent>
                          {categories.map((c) => (
                            <SelectItem key={c} value={c}>
                              {c}
                            </SelectItem>
                          ))}
                        </SelectContent>
                      </Select>

                      <Select value={contractFilter} onValueChange={(v: any) => setContractFilter(v)}>
                        <SelectTrigger>
                          <SelectValue placeholder="Contract" />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="All">All</SelectItem>
                          <SelectItem value="Not sent">Not sent</SelectItem>
                          <SelectItem value="Sent">Sent</SelectItem>
                          <SelectItem value="Signed">Signed</SelectItem>
                          <SelectItem value="Review">Review</SelectItem>
                          <SelectItem value="Not compliant">Not compliant</SelectItem>
                          <SelectItem value="N/A">N/A</SelectItem>
                        </SelectContent>
                      </Select>
                    </div>

                    <Separator className="my-3" />

                    <div className="flex flex-wrap items-center justify-between gap-3 rounded-2xl border border-red-200 bg-red-50 p-3">
                      <div>
                        <div className="text-sm font-medium">Bulk mark selection</div>
                        <div className="text-xs text-muted-foreground">
                          Selected {selectedSuppliers.size} suppliers. Choose parameters and apply.
                        </div>
                      </div>
                      <div className="flex flex-wrap items-center gap-2">
                        <Select value={bulkMarkRisk} onValueChange={(v: any) => setBulkMarkRisk(v)}>
                          <SelectTrigger className="w-[150px]">
                            <SelectValue placeholder="Risk" />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="No change">Risk: no change</SelectItem>
                            <SelectItem value="High">High</SelectItem>
                            <SelectItem value="Non-risk">Non-risk</SelectItem>
                            <SelectItem value="Unknown">Unknown</SelectItem>
                          </SelectContent>
                        </Select>
                        <Select value={bulkMarkEvaluated} onValueChange={(v: any) => setBulkMarkEvaluated(v)}>
                          <SelectTrigger className="w-[170px]">
                            <SelectValue placeholder="Evaluated" />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="No change">Evaluated: no change</SelectItem>
                            <SelectItem value="Evaluated">Evaluated: Yes</SelectItem>
                            <SelectItem value="Not evaluated">Evaluated: No</SelectItem>
                            <SelectItem value="Clear">Evaluated: Clear override</SelectItem>
                          </SelectContent>
                        </Select>
                        <Button
                          variant="destructive"
                          className="gap-2 bg-red-600 text-white hover:bg-red-700"
                          onClick={bulkApplyMarkedSuppliers}
                        >
                          Set As
                        </Button>
                        <Button variant="outline" onClick={handleClearSelectedSuppliers}>
                          Clear
                        </Button>
                      </div>
                    </div>

                    <div className="max-h-[560px] overflow-auto rounded-2xl border">
                      <Table>
                        <TableHeader>
                          <TableRow>
                            <TableHead className="w-[40px]">
                              <input
                                type="checkbox"
                                checked={supplierFiltered.length > 0 && supplierFiltered.every((s) => selectedSuppliers.has(s.code))}
                                onChange={toggleSelectAllSuppliers}
                              />
                            </TableHead>
                            <TableHead>Supplier</TableHead>
                            <TableHead>Risk</TableHead>
                            <TableHead>Evaluated</TableHead>
                            <TableHead>Contract</TableHead>
                            <TableHead>Categories</TableHead>
                            <TableHead className="text-right">Spend</TableHead>
                            <TableHead className="text-right">PO</TableHead>
                            <TableHead>Actions</TableHead>
                          </TableRow>
                        </TableHeader>
                        <TableBody>
                          {supplierFiltered.map((s) => (
                            <TableRow key={s.code}>
                              <TableCell>
                                <input
                                  type="checkbox"
                                  checked={selectedSuppliers.has(s.code)}
                                  onChange={() => handleToggleSupplierSelection(s.code)}
                                />
                              </TableCell>
                              <TableCell className="min-w-[280px]">
                                <div className="font-medium">{s.canonicalName}</div>
                                <div className="text-xs text-muted-foreground">{s.code}</div>
                                {s.lastTouch ? (
                                  <div className="mt-1 text-[11px] text-muted-foreground">Last touch: {fmtDate(s.lastTouch)}</div>
                                ) : null}
                              </TableCell>
                              <TableCell>
                                <RiskBadge risk={s.risk} tone={getRiskColor(s.risk, DEFAULT_SETTINGS.statusPalette.risk[s.risk])} />
                              </TableCell>
                              <TableCell>
                                {s.evaluated ? <Badge variant="secondary">Yes</Badge> : <Badge variant="outline">No</Badge>}
                              </TableCell>
                              <TableCell>
                                <ContractBadge status={s.contractStatus} tone={getContractColor(s.contractStatus, DEFAULT_SETTINGS.statusPalette.contract[s.contractStatus])} />
                              </TableCell>
                              <TableCell className="min-w-[240px]">
                                <div className="flex flex-wrap gap-1">
                                  {s.categories.slice(0, 3).map((c) => (
                                    <Badge key={c} variant="outline" className="gap-1">
                                      <Tag className="h-3 w-3" /> {c}
                                    </Badge>
                                  ))}
                                  {s.categories.length > 3 ? <Badge variant="outline">+{s.categories.length - 3}</Badge> : null}
                                </div>
                              </TableCell>
                              <TableCell className="text-right font-medium">{fmtMoney(s.totalSpend)}</TableCell>
                              <TableCell className="text-right">{Math.round(s.totalPO)}</TableCell>
                              <TableCell className="min-w-[240px]">
                                <div className="flex flex-wrap gap-2">
                                  <SupplierDialog
                                    supplier={s}
                                    overrides={db.overrides}
                                    actor={actor}
                                    bulkContactName={db.settings.bulkContactName}
                                    bulkContactEmail={db.settings.bulkContactEmail}
                                    supplierContactEmail={db.settings.supplierContacts?.[s.code]}
                                    onSetSupplierOverride={setSupplierOverride}
                                    onSetCanonical={setCanonical}
                                    onAddTask={addSupplierTask}
                                    onScheduleEmailFollowUp={scheduleEmailFollowUp}
                                  />

                                  <Button variant="outline" size="sm" className="gap-2" onClick={() => setSupplierOverride(s.code, { contractStatus: "Sent" })}>
                                    <FileUp className="h-4 w-4" /> Mark Sent
                                  </Button>

                                  <Button variant="outline" size="sm" className="gap-2" onClick={() => setSupplierOverride(s.code, { contractStatus: "Signed" })}>
                                    <CheckCircle2 className="h-4 w-4" /> Mark Signed
                                  </Button>
                                </div>
                              </TableCell>
                            </TableRow>
                          ))}
                        </TableBody>
                      </Table>
                    </div>
                  </CardContent>
                </Card>

                <Card>
                  <CardHeader>
                    <CardTitle className="text-base">Worklist (High risk not signed)</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="max-h-[340px] overflow-auto rounded-2xl border">
                      <Table>
                        <TableHeader>
                          <TableRow>
                            <TableHead>Supplier</TableHead>
                            <TableHead>Contract</TableHead>
                            <TableHead className="text-right">Spend</TableHead>
                            <TableHead>Quick action</TableHead>
                          </TableRow>
                        </TableHeader>
                        <TableBody>
                          {worklist.slice(0, 60).map((s) => (
                            <TableRow key={s.code}>
                              <TableCell>
                                <div className="font-medium">{s.canonicalName}</div>
                                <div className="text-xs text-muted-foreground">{s.code}</div>
                              </TableCell>
                              <TableCell>
                                <ContractBadge status={s.contractStatus} tone={getContractColor(s.contractStatus, DEFAULT_SETTINGS.statusPalette.contract[s.contractStatus])} />
                              </TableCell>
                              <TableCell className="text-right font-medium">{fmtMoney(s.totalSpend)}</TableCell>
                              <TableCell>
                                <div className="flex gap-2">
                                  <Button variant="outline" size="sm" onClick={() => setSupplierOverride(s.code, { contractStatus: "Sent" })}>
                                    Sent
                                  </Button>
                                  <Button variant="secondary" size="sm" onClick={() => setSupplierOverride(s.code, { contractStatus: "Signed" })}>
                                    Signed
                                  </Button>
                                  <Button
                                    variant="outline"
                                    size="sm"
                                    onClick={() => {
                                      setTaskSupplierCode(s.code);
                                      setTaskSupplierName(s.canonicalName);
                                      setTaskContactEmail(db.settings.supplierContacts?.[s.code] ?? "");
                                      setTaskTitle(`Chase signature: ${s.canonicalName}`);
                                      setActiveTab("calendar");
                                    }}
                                  >
                                    Add task
                                  </Button>
                                </div>
                              </TableCell>
                            </TableRow>
                          ))}
                        </TableBody>
                      </Table>
                    </div>
                  </CardContent>
                </Card>
                </>
              )
            ) : null}
          </TabsContent>

          <TabsContent value="categories" className="mt-4 space-y-4">
            {activeTab === "categories" ? (
              db.factRows.length === 0 ? (
                <EmptyState title="Import data first" subtitle="Category analytics needs supplier master." />
              ) : (
                <>
                <div className="grid gap-3 md:grid-cols-2">
                  <Card>
                    <CardHeader>
                      <CardTitle className="text-base">Risky categories (share)</CardTitle>
                    </CardHeader>
                    <CardContent className="h-[360px]">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart
                          data={riskyCategoryChart}
                          layout="vertical"
                          margin={{ left: 60 }}
                        >
                          {renderBarDefs("category", riskyCategoryChart.map((c) => c.category), "#93c5fd")}
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis type="number" tickFormatter={(v: any) => `${Math.round(Number(v) * 100)}%`} />
                          <YAxis type="category" dataKey="category" width={260} interval={0} />
                          <Tooltip formatter={(v: any) => `${Math.round(Number(v) * 100)}%`} />
                          <Bar dataKey="riskShare" name="Risk share">
                            {riskyCategoryChart.map((entry) => (
                              <Cell
                                key={`category-risk-${entry.category}`}
                                fill={getBarFill("category", entry.category, "#93c5fd")}
                              />
                            ))}
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </CardContent>
                  </Card>

                  <Card>
                    <CardHeader>
                      <CardTitle className="text-base">All categories (table)</CardTitle>
                    </CardHeader>
                    <CardContent>
                      <div className="max-h-[360px] overflow-auto rounded-2xl border">
                        <Table>
                          <TableHeader>
                            <TableRow>
                              <TableHead>
                                <button
                                  type="button"
                                  className="flex items-center gap-1"
                                  onClick={() =>
                                    setCategoryTableSortState((prev) => ({
                                      key: "category",
                                      dir: prev.key === "category" && prev.dir === "asc" ? "desc" : "asc",
                                    }))
                                  }
                                >
                                  Category
                                </button>
                              </TableHead>
                              <TableHead className="text-right">
                                <button
                                  type="button"
                                  className="ml-auto flex items-center gap-1"
                                  onClick={() =>
                                    setCategoryTableSortState((prev) => ({
                                      key: "suppliers",
                                      dir: prev.key === "suppliers" && prev.dir === "asc" ? "desc" : "asc",
                                    }))
                                  }
                                >
                                  Suppliers
                                </button>
                              </TableHead>
                              <TableHead className="text-right">
                                <button
                                  type="button"
                                  className="ml-auto flex items-center gap-1"
                                  onClick={() =>
                                    setCategoryTableSortState((prev) => ({
                                      key: "highRiskSuppliers",
                                      dir: prev.key === "highRiskSuppliers" && prev.dir === "asc" ? "desc" : "asc",
                                    }))
                                  }
                                >
                                  High risk
                                </button>
                              </TableHead>
                              <TableHead className="text-right">
                                <button
                                  type="button"
                                  className="ml-auto flex items-center gap-1"
                                  onClick={() =>
                                    setCategoryTableSortState((prev) => ({
                                      key: "riskShare",
                                      dir: prev.key === "riskShare" && prev.dir === "asc" ? "desc" : "asc",
                                    }))
                                  }
                                >
                                  Risk share
                                </button>
                              </TableHead>
                              <TableHead className="text-right">
                                <button
                                  type="button"
                                  className="ml-auto flex items-center gap-1"
                                  onClick={() =>
                                    setCategoryTableSortState((prev) => ({
                                      key: "totalSpend",
                                      dir: prev.key === "totalSpend" && prev.dir === "asc" ? "desc" : "asc",
                                    }))
                                  }
                                >
                                  Spend
                                </button>
                              </TableHead>
                            </TableRow>
                          </TableHeader>
                          <TableBody>
                            {sortedCategoryMetrics.map((c) => (
                              <TableRow key={c.category}>
                                <TableCell className="min-w-[240px]">{c.category}</TableCell>
                                <TableCell className="text-right">{c.suppliers}</TableCell>
                                <TableCell className="text-right">{c.highRiskSuppliers}</TableCell>
                                <TableCell className="text-right">{Math.round(c.riskShare * 100)}%</TableCell>
                                <TableCell className="text-right font-medium">{fmtMoney(c.totalSpend)}</TableCell>
                              </TableRow>
                            ))}
                          </TableBody>
                        </Table>
                      </div>
                    </CardContent>
                  </Card>
                </div>

                <Card>
                  <CardHeader>
                    <CardTitle className="flex items-center gap-2 text-base">
                      <ShieldAlert className="h-4 w-4" /> Guardrails & scope thresholds
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="text-sm text-muted-foreground">
                      These thresholds power multi-category protection and "unknown in-scope" checks. They do not change raw data вЂ” only decision rules.
                    </div>

                    <div className="mt-4 grid gap-3 md:grid-cols-4">
                      <div className="rounded-2xl border p-3">
                        <div className="text-xs text-muted-foreground">Dominant category share</div>
                        <Input
                          type="number"
                          step="0.01"
                          value={db.settings.dominanceShareThreshold}
                          onChange={(e) => setAppSettings({ dominanceShareThreshold: clamp01(Number(e.target.value)) })}
                          className="mt-2"
                        />
                        <div className="mt-1 text-xs text-muted-foreground">e.g. 0.80 means 80%+</div>
                      </div>

                      <div className="rounded-2xl border p-3">
                        <div className="text-xs text-muted-foreground">Multi-category 2nd share</div>
                        <Input
                          type="number"
                          step="0.01"
                          value={db.settings.multiCategorySecondShareThreshold}
                          onChange={(e) => setAppSettings({ multiCategorySecondShareThreshold: clamp01(Number(e.target.value)) })}
                          className="mt-2"
                        />
                        <div className="mt-1 text-xs text-muted-foreground">Triggers ambiguity alert</div>
                      </div>

                      <div className="rounded-2xl border p-3">
                        <div className="text-xs text-muted-foreground">In-scope spend threshold</div>
                        <Input
                          type="number"
                          step="1000"
                          value={db.settings.scopeSpendThreshold}
                          onChange={(e) => setAppSettings({ scopeSpendThreshold: Math.max(0, Number(e.target.value) || 0) })}
                          className="mt-2"
                        />
                        <div className="mt-1 text-xs text-muted-foreground">Flags "Unknown in-scope"</div>
                      </div>

                      <div className="rounded-2xl border p-3">
                        <div className="text-xs text-muted-foreground">Missing sub-family threshold</div>
                        <Input
                          type="number"
                          step="1000"
                          value={db.settings.missingSubfamilySpendThreshold}
                          onChange={(e) => setAppSettings({ missingSubfamilySpendThreshold: Math.max(0, Number(e.target.value) || 0) })}
                          className="mt-2"
                        />
                        <div className="mt-1 text-xs text-muted-foreground">DQ alert for taxonomy gaps</div>
                      </div>
                    </div>
                  </CardContent>
                </Card>

                <Card>
                  <CardHeader>
                    <CardTitle className="flex items-center gap-2 text-base">
                      <Settings className="h-4 w-4" /> Policy layer (Category/Family rules)
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="grid gap-2 md:grid-cols-4">
                      <Select value={newRuleScope} onValueChange={(v: any) => setNewRuleScope(v)}>
                        <SelectTrigger>
                          <SelectValue />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="Category">Category</SelectItem>
                          <SelectItem value="Family">Family</SelectItem>
                        </SelectContent>
                      </Select>

                      <Input value={newRuleKey} onChange={(e) => setNewRuleKey(e.target.value)} placeholder="Exact name (e.g., Transport - Catman)" />

                      <Select value={newRuleRisk} onValueChange={(v: any) => setNewRuleRisk(v)}>
                        <SelectTrigger>
                          <SelectValue />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value="High">High</SelectItem>
                          <SelectItem value="Non-risk">Non-risk</SelectItem>
                          <SelectItem value="Unknown">Unknown</SelectItem>
                        </SelectContent>
                      </Select>

                      <Button className="gap-2" onClick={addRule}>
                        <PlusIcon /> Add rule
                      </Button>
                    </div>

                    <div className="mt-2">
                      <Textarea value={newRuleComment} onChange={(e) => setNewRuleComment(e.target.value)} placeholder="Comment / rationale (optional)" />
                    </div>

                    <Separator className="my-4" />

                    <div className="max-h-[360px] overflow-auto rounded-2xl border">
                      <Table>
                        <TableHeader>
                          <TableRow>
                            <TableHead>Scope</TableHead>
                            <TableHead>Key</TableHead>
                            <TableHead>Default risk</TableHead>
                            <TableHead>Comment</TableHead>
                            <TableHead className="text-right">Action</TableHead>
                          </TableRow>
                        </TableHeader>
                        <TableBody>
                          {db.categoryRules.map((r, idx) => (
                            <TableRow key={`${r.scope}:${r.key}:${idx}`}>
                              <TableCell>
                                <Badge variant="outline">{r.scope}</Badge>
                              </TableCell>
                              <TableCell className="min-w-[260px]">{r.key}</TableCell>
                              <TableCell>
                                <RiskBadge risk={r.defaultRisk} tone={getRiskColor(r.defaultRisk, DEFAULT_SETTINGS.statusPalette.risk[r.defaultRisk])} />
                              </TableCell>
                              <TableCell className="text-sm text-muted-foreground">{r.comment ?? ""}</TableCell>
                              <TableCell className="text-right">
                                <Button variant="destructive" size="sm" onClick={() => deleteRule(idx)}>
                                  Delete
                                </Button>
                              </TableCell>
                            </TableRow>
                          ))}
                        </TableBody>
                      </Table>
                    </div>
                  </CardContent>
                </Card>
                </>
              )
            ) : null}
          </TabsContent>

          <TabsContent value="duplicates" className="mt-4 space-y-4">
            {activeTab === "duplicates" ? (
              db.factRows.length === 0 ? (
                <EmptyState title="Import data first" subtitle="Duplicate detection relies on supplier master." />
              ) : (
                <>
                <Card>
                  <CardHeader>
                    <CardTitle className="flex items-center gap-2 text-base">
                      <Split className="h-4 w-4" /> Quality command center
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="grid gap-3 md:grid-cols-4">
                      <Card>
                        <CardContent className="p-4">
                          <div className="text-xs text-muted-foreground">Total issues</div>
                          <div className="text-2xl font-semibold">{kpis.qualityTotal}</div>
                          <div className="mt-1 text-xs text-muted-foreground">Sorted by severity then spend.</div>
                        </CardContent>
                      </Card>
                      <Card>
                        <CardContent className="p-4">
                          <div className="text-xs text-muted-foreground">Unknown in-scope</div>
                          <div className="text-2xl font-semibold">{kpis.unknownInScope}</div>
                          <div className="mt-1 text-xs text-muted-foreground">Spend в‰Ґ {fmtMoney(db.settings.scopeSpendThreshold)}</div>
                        </CardContent>
                      </Card>
                      <Card>
                        <CardContent className="p-4">
                          <div className="text-xs text-muted-foreground">Multi-category</div>
                          <div className="text-2xl font-semibold">{kpis.multiCat}</div>
                          <div className="mt-1 text-xs text-muted-foreground">2nd share в‰Ґ {(db.settings.multiCategorySecondShareThreshold * 100).toFixed(0)}%</div>
                        </CardContent>
                      </Card>
                      <Card>
                        <CardContent className="p-4">
                          <div className="text-xs text-muted-foreground">Name в†’ codes</div>
                          <div className="text-2xl font-semibold">{kpis.nameMany}</div>
                          <div className="mt-1 text-xs text-muted-foreground">Potential ID hygiene issue</div>
                        </CardContent>
                      </Card>
                    </div>

                    <Separator className="my-4" />

                    <div className="flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
                      <div>
                        <div className="text-sm font-medium">Issue inbox</div>
                        <div className="text-xs text-muted-foreground">
                          Click a row to inspect and resolve duplicates. Use filters to focus your workstream.
                        </div>
                      </div>
                      <div className="flex flex-wrap gap-2">
                        <Select value={issueTypeFilter} onValueChange={(v: any) => setIssueTypeFilter(v)}>
                          <SelectTrigger className="w-[220px]">
                            <SelectValue placeholder="Issue type" />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="All">All types</SelectItem>
                            <SelectItem value="RISK_UNKNOWN_IN_SCOPE">Unknown in-scope</SelectItem>
                            <SelectItem value="MULTI_CATEGORY_EXPOSURE">Multi-category exposure</SelectItem>
                            <SelectItem value="CODE_MANY_NAMES">Code в†’ many names</SelectItem>
                            <SelectItem value="NAME_MANY_CODES">Name в†’ many codes</SelectItem>
                            <SelectItem value="MISSING_SUBFAMILY_HIGH_SPEND">Missing sub-family</SelectItem>
                            <SelectItem value="POLICY_SUPPRESSED_BY_GUARDRAIL">Policy suppressed</SelectItem>
                          </SelectContent>
                        </Select>
                        <Select value={issueSeverityFilter} onValueChange={(v: any) => setIssueSeverityFilter(v)}>
                          <SelectTrigger className="w-[180px]">
                            <SelectValue placeholder="Severity" />
                          </SelectTrigger>
                          <SelectContent>
                            <SelectItem value="All">All severity</SelectItem>
                            <SelectItem value="High">High</SelectItem>
                            <SelectItem value="Medium">Medium</SelectItem>
                            <SelectItem value="Low">Low</SelectItem>
                          </SelectContent>
                        </Select>
                      </div>
                    </div>

                    <div className="mt-3 max-h-[520px] overflow-auto rounded-2xl border">
                      <Table>
                        <TableHeader>
                          <TableRow>
                            <TableHead>Severity</TableHead>
                            <TableHead>Type</TableHead>
                            <TableHead>Supplier</TableHead>
                            <TableHead>Title</TableHead>
                            <TableHead className="text-right">Spend</TableHead>
                          </TableRow>
                        </TableHeader>
                        <TableBody>
                          {issuesFiltered.length === 0 ? (
                            <TableRow>
                              <TableCell colSpan={5} className="py-10 text-center text-sm text-muted-foreground">
                                No issues match this filter.
                              </TableCell>
                            </TableRow>
                          ) : (
                            issuesFiltered.slice(0, 300).map((i) => (
                              <TableRow
                                key={i.id}
                                className={`cursor-pointer ${i.code && i.code === selectedDuplicateCode ? "bg-slate-50" : ""}`}
                                onClick={() => {
                                  if (i.type === "CODE_MANY_NAMES") {
                                    setSelectedDuplicateCode(i.code ?? null);
                                  } else {
                                    setSelectedDuplicateCode(null);
                                  }
                                }}
                              >
                                <TableCell>
                                  <SeverityBadge sev={i.severity} />
                                </TableCell>
                                <TableCell>
                                  <Badge variant="outline">{i.type}</Badge>
                                </TableCell>
                                <TableCell className="min-w-[220px]">
                                  <div className="font-medium">{i.supplierName ?? "(name missing)"}</div>
                                  {i.code ? <div className="text-xs text-muted-foreground">{i.code}</div> : null}
                                </TableCell>
                                <TableCell className="min-w-[280px]">{i.title}</TableCell>
                                <TableCell className="text-right font-medium">{fmtMoney(i.spendAffected ?? 0)}</TableCell>
                              </TableRow>
                            ))
                          )}
                        </TableBody>
                      </Table>
                    </div>
                  </CardContent>
                </Card>

                <Card>
                  <CardHeader>
                    <CardTitle className="flex items-center gap-2 text-base">
                      <Split className="h-4 w-4" /> Canonical name resolver
                    </CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="text-sm text-muted-foreground">
                      Duplicates here mean: <span className="font-medium">same Supplier Code</span> with <span className="font-medium">multiple names</span>.
                      Choose a canonical once, and the dashboard normalizes all rollups.
                    </div>

                    <Separator className="my-3" />

                    {duplicates.length === 0 ? (
                      <EmptyState title="No codeв†’name duplicates" subtitle="Your Supplier Dictionary looks clean." />
                    ) : (
                      <div className="space-y-4">
                        {selectedDuplicate ? (
                          <Card>
                            <CardContent className="space-y-3 p-4">
                              <div className="flex items-center justify-between">
                                <div>
                                  <div className="text-sm text-muted-foreground">Selected supplier code</div>
                                  <div className="text-lg font-semibold">{selectedDuplicate.code}</div>
                                </div>
                                <Badge variant="outline">Candidates: {selectedDuplicate.candidates.length}</Badge>
                              </div>
                              <div className="grid gap-3 md:grid-cols-2">
                                <div className="rounded-2xl border border-red-200 bg-red-50 p-3">
                                  <div className="text-xs font-semibold text-red-600">Problem</div>
                                  <div className="mt-1 text-sm">
                                    Code has multiple names ({selectedDuplicate.candidates.length}).
                                  </div>
                                  <div className="mt-2 text-xs text-muted-foreground">
                                    Current canonical: <span className="font-medium">{selectedDuplicate.current}</span>
                                  </div>
                                </div>
                                <div className="rounded-2xl border border-green-200 bg-green-50 p-3">
                                  <div className="text-xs font-semibold text-green-600">Suggestion</div>
                                  <div className="mt-1 text-sm">
                                    Recommended canonical: <span className="font-medium">{selectedDuplicate.recommended}</span>
                                  </div>
                                  <Button
                                    variant="secondary"
                                    className="mt-2"
                                    onClick={() => setCanonical(selectedDuplicate.code, selectedDuplicate.recommended)}
                                  >
                                    Apply recommendation
                                  </Button>
                                </div>
                              </div>
                              <div className="rounded-2xl border p-3">
                                <div className="text-xs text-muted-foreground">Manual canonical name</div>
                                <div className="mt-2 flex flex-wrap items-center gap-2">
                                  <Input
                                    value={manualCanonicalName}
                                    onChange={(e) => setManualCanonicalName(e.target.value)}
                                    placeholder="Type corrected canonical name"
                                    className="w-full md:flex-1"
                                  />
                                  <Button
                                    variant="outline"
                                    onClick={() => setCanonical(selectedDuplicate.code, manualCanonicalName.trim())}
                                    disabled={!manualCanonicalName.trim()}
                                  >
                                    Save
                                  </Button>
                                </div>
                              </div>
                              <div className="grid gap-2 md:grid-cols-2">
                                {selectedDuplicate.candidates.map((c) => {
                                  const stats = duplicateCandidateStats.get(getCandidateKey(selectedDuplicate.code, c.name));
                                  return (
                                    <div
                                      key={c.name}
                                      className={`flex items-center justify-between gap-2 rounded-2xl border p-3 ${
                                        c.name === selectedDuplicate.recommended
                                          ? "border-green-200 bg-green-50"
                                          : "border-red-200 bg-red-50"
                                      }`}
                                    >
                                      <div>
                                        <div className="font-medium">{c.name}</div>
                                        <div className="text-xs text-muted-foreground">Seen {c.count}Г—</div>
                                      </div>
                                      <div className="flex items-center gap-2">
                                        <Dialog>
                                          <DialogTrigger asChild>
                                            <Button variant="outline" size="sm">
                                              Details
                                            </Button>
                                          </DialogTrigger>
                                          <DialogContent className="max-w-lg">
                                            <DialogHeader>
                                              <DialogTitle>Candidate details</DialogTitle>
                                            </DialogHeader>
                                            <div className="space-y-3 text-sm">
                                              <div className="rounded-2xl border p-3">
                                                <div className="text-xs text-muted-foreground">Supplier name</div>
                                                <div className="text-lg font-semibold">{c.name}</div>
                                                <div className="mt-1 text-xs text-muted-foreground">Supplier code: {selectedDuplicate.code}</div>
                                              </div>
                                              <div className="grid gap-2 md:grid-cols-3">
                                                <div className="rounded-2xl border p-3">
                                                  <div className="text-xs text-muted-foreground">PO count</div>
                                                  <div className="text-lg font-semibold">{Math.round(stats?.totalPO ?? 0)}</div>
                                                </div>
                                                <div className="rounded-2xl border p-3">
                                                  <div className="text-xs text-muted-foreground">SUM</div>
                                                  <div className="text-lg font-semibold">{fmtMoney(stats?.totalSpend ?? 0)}</div>
                                                </div>
                                                <div className="rounded-2xl border p-3">
                                                  <div className="text-xs text-muted-foreground">Date when</div>
                                                  <div className="text-lg font-semibold">{stats?.lastSeen ? String(stats.lastSeen) : "N/A"}</div>
                                                </div>
                                              </div>
                                            </div>
                                          </DialogContent>
                                        </Dialog>
                                        <Button variant="outline" size="sm" onClick={() => setCanonical(selectedDuplicate.code, c.name)}>
                                          Select
                                        </Button>
                                      </div>
                                    </div>
                                  );
                                })}
                              </div>
                            </CardContent>
                          </Card>
                        ) : (
                          <div className="rounded-2xl border border-slate-200 bg-slate-50 p-4 text-sm text-muted-foreground">
                            Select a "Code в†’ many names" issue to resolve it here.
                          </div>
                        )}
                      </div>
                    )}
                  </CardContent>
                </Card>

                </>
              )
            ) : null}
          </TabsContent>

          <TabsContent value="bulk" className="mt-4 space-y-4">
            {activeTab === "bulk" ? (
              <Card>
                <CardHeader>
                  <CardTitle className="flex items-center gap-2 text-base">
                    <FileUp className="h-4 w-4" /> Bulk Operations
                  </CardTitle>
                </CardHeader>
                <CardContent className="space-y-4">
                  <div className="grid gap-3 md:grid-cols-2">
                    <div className="space-y-3">
                      {db.lastUndo ? (
                        <Card>
                          <CardContent className="p-4">
                            <div className="flex items-start justify-between gap-3">
                              <div>
                                <div className="text-sm font-medium">Undo last bulk action</div>
                                <div className="text-xs text-muted-foreground">{db.lastUndo.summary}</div>
                                <div className="mt-1 text-xs text-muted-foreground">Changed: {db.lastUndo.items.length}</div>
                              </div>
                              <Button variant="secondary" onClick={undoLast} className="gap-2">
                                <Undo2 className="h-4 w-4" /> Undo
                              </Button>
                            </div>
                          </CardContent>
                        </Card>
                      ) : null}

                    <Card>
                      <CardContent className="p-4">
                        <div className="text-sm font-medium">Supplier code contact</div>
                        <div className="text-xs text-muted-foreground">
                          Save contacts in bulk (one supplier code and email per line).
                        </div>
                        <Textarea
                          value={bulkContactEntries}
                          onChange={(e) => setBulkContactEntries(e.target.value)}
                          placeholder="SUPPLIER_CODE contact@email.com"
                          className="mt-3 min-h-[120px]"
                        />
                        <div className="mt-2 text-xs text-muted-foreground">Format: SUPPLIER_CODE email@company.com (one per line).</div>
                        <div className="mt-3 flex justify-end">
                          <Button
                            variant="secondary"
                            onClick={() => {
                              parseContactEntries(bulkContactEntries).forEach((entry) =>
                                upsertSupplierContact(entry.code, entry.email)
                              );
                              setBulkContactEntries("");
                            }}
                          >
                            Save contacts
                          </Button>
                        </div>
                        {Object.keys(db.settings.supplierContacts ?? {}).length ? (
                          <div className="mt-3 space-y-1 text-xs text-muted-foreground">
                            {Object.entries(db.settings.supplierContacts ?? {}).slice(0, 5).map(([code, email]) => (
                              <div key={code} className="flex items-center justify-between">
                                <span className="font-medium">{code}</span>
                                <span>{email}</span>
                              </div>
                            ))}
                          </div>
                        ) : (
                          <div className="mt-3 text-xs text-muted-foreground">No contacts saved yet.</div>
                        )}
                      </CardContent>
                    </Card>
                    <Card>
                      <CardContent className="p-4">
                        <div className="flex items-center justify-between">
                          <div>
                            <div className="text-sm font-medium">Set risk</div>
                            <div className="text-xs text-muted-foreground">Supplier override (wins over data & rules)</div>
                          </div>
                          <Select value={bulkRisk} onValueChange={(v: any) => setBulkRisk(v)}>
                            <SelectTrigger className="w-[160px]">
                              <SelectValue />
                            </SelectTrigger>
                            <SelectContent>
                              <SelectItem value="High">High</SelectItem>
                              <SelectItem value="Non-risk">Non-risk</SelectItem>
                              <SelectItem value="Unknown">Unknown</SelectItem>
                            </SelectContent>
                          </Select>
                        </div>
                        {bulkRiskPreview ? (
                          <div className="mt-2 text-xs text-muted-foreground">
                            Found {bulkRiskPreview.found.length} вЂў Changes {bulkRiskPreview.changes.length} вЂў Not found {bulkRiskPreview.notFound.length}
                          </div>
                        ) : (
                          <div className="mt-2 text-xs text-muted-foreground">Paste codes to generate a change preview.</div>
                        )}
                        <div className="mt-3 flex justify-end">
                          <Button className="gap-2" onClick={bulkApplyRisk}>
                            <RefreshCw className="h-4 w-4" /> Apply
                          </Button>
                        </div>
                      </CardContent>
                    </Card>

                    <Card>
                      <CardContent className="p-4">
                        <div className="flex items-center justify-between">
                          <div>
                            <div className="text-sm font-medium">Set contract status</div>
                            <div className="text-xs text-muted-foreground">Contract ops tracking</div>
                          </div>
                          <Select value={bulkContract} onValueChange={(v: any) => setBulkContract(v)}>
                            <SelectTrigger className="w-[180px]">
                              <SelectValue />
                            </SelectTrigger>
                            <SelectContent>
                              <SelectItem value="Not sent">Not sent</SelectItem>
                              <SelectItem value="Sent">Sent</SelectItem>
                              <SelectItem value="Signed">Signed</SelectItem>
                              <SelectItem value="Review">Review</SelectItem>
                              <SelectItem value="Not compliant">Not compliant</SelectItem>
                              <SelectItem value="N/A">N/A</SelectItem>
                            </SelectContent>
                          </Select>
                        </div>
                        {bulkContractPreview ? (
                          <div className="mt-2 text-xs text-muted-foreground">
                            Found {bulkContractPreview.found.length} вЂў Changes {bulkContractPreview.changes.length} вЂў Not found {bulkContractPreview.notFound.length}
                          </div>
                        ) : (
                          <div className="mt-2 text-xs text-muted-foreground">Paste codes to generate a change preview.</div>
                        )}
                        <div className="mt-3 flex justify-end">
                          <Button className="gap-2" variant="secondary" onClick={bulkApplyContractStatus}>
                            <RefreshCw className="h-4 w-4" /> Apply
                          </Button>
                        </div>
                      </CardContent>
                    </Card>

                    <Card>
                      <CardContent className="p-4">
                        <div className="flex items-center justify-between">
                          <div>
                            <div className="text-sm font-medium">Set evaluated status</div>
                            <div className="text-xs text-muted-foreground">Override evaluation state.</div>
                          </div>
                          <Select value={bulkEvaluated} onValueChange={(v: any) => setBulkEvaluated(v)}>
                            <SelectTrigger className="w-[190px]">
                              <SelectValue />
                            </SelectTrigger>
                            <SelectContent>
                              <SelectItem value="Evaluated">Evaluated</SelectItem>
                              <SelectItem value="Not evaluated">Not evaluated</SelectItem>
                              <SelectItem value="Clear">Clear override</SelectItem>
                            </SelectContent>
                          </Select>
                        </div>
                        {bulkEvaluatedPreviewState ? (
                          <div className="mt-2 text-xs text-muted-foreground">
                            Found {bulkEvaluatedPreviewState.found.length} вЂў Changes {bulkEvaluatedPreviewState.changes.length} вЂў Not found {bulkEvaluatedPreviewState.notFound.length}
                          </div>
                        ) : (
                          <div className="mt-2 text-xs text-muted-foreground">Paste codes to generate a change preview.</div>
                        )}
                        <div className="mt-3 flex justify-end">
                          <Button className="gap-2" variant="secondary" onClick={bulkApplyEvaluated}>
                            <RefreshCw className="h-4 w-4" /> Apply
                          </Button>
                        </div>
                      </CardContent>
                    </Card>

                    <Card>
                      <CardContent className="p-4">
                        <div>
                          <div className="text-sm font-medium">Set canonical name</div>
                          <div className="text-xs text-muted-foreground">Use this to unify name across the database.</div>
                        </div>
                        <Input value={bulkCanonicalName} onChange={(e) => setBulkCanonicalName(e.target.value)} placeholder="Canonical supplier name" className="mt-2" />
                        {bulkCanonicalPreview ? (
                          <div className="mt-2 text-xs text-muted-foreground">
                            Found {bulkCanonicalPreview.found.length} вЂў Changes {bulkCanonicalPreview.changes.length} вЂў Not found {bulkCanonicalPreview.notFound.length}
                          </div>
                        ) : (
                          <div className="mt-2 text-xs text-muted-foreground">Add a canonical name + codes to preview changes.</div>
                        )}
                        <div className="mt-3 flex justify-end">
                          <Button className="gap-2" variant="outline" onClick={bulkApplyCanonicalName}>
                            <RefreshCw className="h-4 w-4" /> Apply
                          </Button>
                        </div>
                      </CardContent>
                    </Card>

                    <Card>
                      <CardContent className="p-4">
                        <div className="flex items-start justify-between gap-3">
                          <div>
                            <div className="text-sm font-medium">Category FastAction</div>
                            <div className="text-xs text-muted-foreground">
                              Applies risk override to eligible suppliers in a category (dominant + not multi-category).
                            </div>
                          </div>
                        </div>
                        <div className="mt-3 grid gap-2 md:grid-cols-2">
                          <Select value={fastCategory} onValueChange={(v: any) => setFastCategory(v)}>
                            <SelectTrigger>
                              <SelectValue placeholder="Category" />
                            </SelectTrigger>
                            <SelectContent>
                              {categories.filter((c) => c !== "All").map((c) => (
                                <SelectItem key={c} value={c}>
                                  {c}
                                </SelectItem>
                              ))}
                            </SelectContent>
                          </Select>
                          <Select value={fastCategoryRisk} onValueChange={(v: any) => setFastCategoryRisk(v)}>
                            <SelectTrigger>
                              <SelectValue />
                            </SelectTrigger>
                            <SelectContent>
                              <SelectItem value="High">High</SelectItem>
                              <SelectItem value="Non-risk">Non-risk</SelectItem>
                              <SelectItem value="Unknown">Unknown</SelectItem>
                            </SelectContent>
                          </Select>
                        </div>
                        <div className="mt-2 text-xs text-muted-foreground">
                          In category {categoryFastPreview.total} вЂў Eligible {categoryFastPreview.eligible} вЂў Ambiguous {categoryFastPreview.ambiguous}
                        </div>
                        <div className="mt-3 flex justify-end gap-2">
                          <Button
                            variant="outline"
                            onClick={() => {
                              setBulkCodes(categoryFastPreview.eligibleCodes.join("\n"));
                            }}
                          >
                            Copy eligible to Bulk
                          </Button>
                          <Button className="gap-2" onClick={applyCategoryFastAction}>
                            <RefreshCw className="h-4 w-4" /> Apply
                          </Button>
                        </div>
                      </CardContent>
                    </Card>
                  </div>
                  <div className="space-y-3">
                    <Card>
                      <CardContent className="p-4">
                        <div className="text-sm font-medium">Bulk target list</div>
                        <div className="text-xs text-muted-foreground">
                          Paste supplier codes below to preview and apply bulk rules.
                        </div>
                        <Textarea
                          value={bulkCodes}
                          onChange={(e) => setBulkCodes(e.target.value)}
                          placeholder="Paste supplier codes (space/comma/new line separated)"
                          className="mt-3 min-h-[140px]"
                        />
                        <div className="mt-2 flex items-center justify-between text-xs text-muted-foreground">
                          <span>Loaded: {bulkParsedCodes.length} codes</span>
                          <Button variant="outline" size="sm" onClick={() => setBulkCodes("")}>
                            Clear
                          </Button>
                        </div>
                        <div className="mt-3 rounded-xl border">
                          {bulkPreviewRows.length ? (
                            <Table>
                              <TableHeader>
                                <TableRow>
                                  <TableHead>Code</TableHead>
                                  <TableHead>Supplier</TableHead>
                                  <TableHead>Risk</TableHead>
                                  <TableHead>Contract</TableHead>
                                </TableRow>
                              </TableHeader>
                              <TableBody>
                                {bulkPreviewRows.map((row) => (
                                  <TableRow key={row.code}>
                                    <TableCell className="font-medium">{row.code}</TableCell>
                                    <TableCell>{row.name}</TableCell>
                                    <TableCell>{row.risk}</TableCell>
                                    <TableCell>{row.contract}</TableCell>
                                  </TableRow>
                                ))}
                              </TableBody>
                            </Table>
                          ) : (
                            <div className="p-4 text-xs text-muted-foreground">No codes yet. Paste above to preview names.</div>
                          )}
                        </div>
                      </CardContent>
                    </Card>

                    <Card>
                      <CardContent className="p-4">
                        <div className="text-sm font-medium">Find suppliers by name</div>
                        <div className="text-xs text-muted-foreground">
                          Type a name fragment to find supplier codes and add them to the bulk list.
                        </div>
                        <Input
                          value={bulkNameQuery}
                          onChange={(e) => setBulkNameQuery(e.target.value)}
                          placeholder="Start typing a supplier name"
                          className="mt-3"
                        />
                        <div className="mt-3 max-h-[240px] overflow-auto rounded-xl border">
                          {bulkNameMatchesList.length ? (
                            <Table>
                              <TableHeader>
                                <TableRow>
                                  <TableHead className="w-[40px]" />
                                  <TableHead>Supplier</TableHead>
                                  <TableHead>Code</TableHead>
                                </TableRow>
                              </TableHeader>
                              <TableBody>
                                {bulkNameMatchesList.map((match) => (
                                  <TableRow key={match.code}>
                                    <TableCell>
                                      <input
                                        type="checkbox"
                                        checked={bulkNameSelected.has(match.code)}
                                        onChange={() => handleToggleBulkNameSelection(match.code)}
                                      />
                                    </TableCell>
                                    <TableCell className="font-medium">{match.canonicalName}</TableCell>
                                    <TableCell className="text-xs text-muted-foreground">{match.code}</TableCell>
                                  </TableRow>
                                ))}
                              </TableBody>
                            </Table>
                          ) : (
                            <div className="p-4 text-xs text-muted-foreground">No matches yet.</div>
                          )}
                        </div>
                        <div className="mt-3 flex items-center justify-between text-xs text-muted-foreground">
                          <span>Selected: {bulkNameSelected.size}</span>
                          <Button variant="secondary" size="sm" onClick={handleApplyBulkNameSelection}>
                            Add to bulk list
                          </Button>
                        </div>
                      </CardContent>
                    </Card>
                  </div>
                </div>

                  <Separator />

                  <div className="text-xs text-muted-foreground">
                    Governance tip: use Bulk to mass-set known risk/non-risk suppliers and immediately shrink the evaluation queue.
                  </div>
                </CardContent>
              </Card>
            ) : null}
          </TabsContent>

          <TabsContent value="log" className="mt-4 space-y-4">
            {activeTab === "log" ? (
              <Card>
                <CardHeader>
                  <CardTitle className="flex items-center gap-2 text-base">
                    <ClipboardList className="h-4 w-4" /> Activity Log (audit trail)
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  {db.log.length === 0 ? (
                    <EmptyState title="No activity yet" subtitle="Import a file or perform actions to generate an audit trail." />
                  ) : (
                    <div className="max-h-[640px] overflow-auto rounded-2xl border">
                      <Table>
                        <TableHeader>
                          <TableRow>
                            <TableHead>When</TableHead>
                            <TableHead>Actor</TableHead>
                            <TableHead>Type</TableHead>
                            <TableHead>Summary</TableHead>
                          </TableRow>
                        </TableHeader>
                        <TableBody>
                          {db.log.slice(0, 800).map((e) => (
                            <TableRow key={e.id}>
                              <TableCell className="whitespace-nowrap">{fmtDate(e.ts)}</TableCell>
                              <TableCell>{e.actor}</TableCell>
                              <TableCell>
                                <Badge variant="outline">{e.type}</Badge>
                              </TableCell>
                              <TableCell className="min-w-[520px]">
                                <div className="font-medium">{e.summary}</div>
                                {e.details ? (
                                  <pre className="mt-1 max-w-[780px] overflow-auto rounded-xl bg-muted p-2 text-[11px] text-muted-foreground">
                                    {JSON.stringify(e.details, null, 2)}
                                  </pre>
                                ) : null}
                              </TableCell>
                            </TableRow>
                          ))}
                        </TableBody>
                      </Table>
                    </div>
                  )}
                </CardContent>
              </Card>
            ) : null}
          </TabsContent>

          <TabsContent value="calendar" className="mt-4 space-y-4">
            {activeTab === "calendar" ? (
            <>
            <div className="grid gap-3 md:grid-cols-2">
              <Card>
                <CardHeader>
                  <CardTitle className="flex items-center gap-2 text-base">
                    <CalendarDays className="h-4 w-4" /> Notes & Action Calendar
                  </CardTitle>
                </CardHeader>
                <CardContent className="space-y-3">
                  <div className="grid gap-2 md:grid-cols-2">
                    <div>
                      <div className="text-xs text-muted-foreground">Resolve by</div>
                      <Input value={taskDate} onChange={(e) => setTaskDate(e.target.value)} type="date" />
                    </div>
                    <div>
                      <div className="text-xs text-muted-foreground">Owner</div>
                      <Input value={taskOwner} onChange={(e) => setTaskOwner(e.target.value)} />
                    </div>
                  </div>
                  <div>
                    <div className="text-xs text-muted-foreground">Supplier code (optional)</div>
                    <Input value={taskSupplierCode} onChange={(e) => setTaskSupplierCode(e.target.value)} placeholder="3302858" />
                  </div>
                  <div>
                    <div className="text-xs text-muted-foreground">Supplier name (optional)</div>
                    <Input value={taskSupplierName} onChange={(e) => setTaskSupplierName(e.target.value)} placeholder="Alpine Logistics GmbH" />
                  </div>
                  <div className="grid gap-2 md:grid-cols-2">
                    <div>
                      <div className="text-xs text-muted-foreground">Contact name</div>
                      <Input
                        value={taskContactName}
                        onChange={(e) => setTaskContactName(e.target.value)}
                        placeholder={db.settings.bulkContactName || "Procurement contact"}
                      />
                    </div>
                    <div>
                      <div className="text-xs text-muted-foreground">Contact email</div>
                      <Input
                        value={taskContactEmail}
                        onChange={(e) => setTaskContactEmail(e.target.value)}
                        placeholder={db.settings.bulkContactEmail || "name@company.com"}
                      />
                    </div>
                  </div>
                  <div>
                    <div className="text-xs text-muted-foreground">Title</div>
                    <Input value={taskTitle} onChange={(e) => setTaskTitle(e.target.value)} placeholder="Send contract / chase signature / review docsвЂ¦" />
                  </div>
                  <div>
                    <div className="text-xs text-muted-foreground">Resolve by</div>
                    <Input value={taskDueDate} onChange={(e) => setTaskDueDate(e.target.value)} type="date" />
                  </div>
                  <div>
                    <div className="text-xs text-muted-foreground">Note</div>
                    <Textarea value={taskNote} onChange={(e) => setTaskNote(e.target.value)} placeholder="What was done, what is blocked, next step, by whomвЂ¦" />
                  </div>
                  <div className="flex justify-end">
                    <Button className="gap-2" onClick={() => addTask()}>
                      <PlusIcon /> Add
                    </Button>
                  </div>

                  {redIgnoranceAlerts.length ? (
                    <>
                      <Separator />
                      <div className="rounded-2xl border border-red-200 bg-red-50 p-3">
                        <div className="text-sm font-medium text-red-700">Red Ignorance Alert</div>
                        <div className="text-xs text-red-700/80">
                          Suppliers requiring urgent contact after two unanswered reminders.
                        </div>
                        <div className="mt-2 space-y-1 text-xs">
                          {redIgnoranceAlerts.slice(0, 6).map((t) => (
                            <div key={t.id} className="flex items-center justify-between">
                              <span className="font-medium">{t.supplierName ?? t.supplierCode ?? "Supplier"}</span>
                              <span>Due {t.dueDate ?? t.date}</span>
                            </div>
                          ))}
                        </div>
                      </div>
                    </>
                  ) : null}

                  <Separator />

                  <div className="text-sm font-medium">Tasks on selected day</div>
                  <div className="space-y-2">
                    {(tasksByDate[taskDate] ?? []).length === 0 ? (
                      <div className="text-sm text-muted-foreground">No tasks for {taskDate}</div>
                    ) : (
                      (tasksByDate[taskDate] ?? []).map((t) => {
                        const today = new Date().toISOString().slice(0, 10);
                        const due = t.dueDate ?? t.date;
                        const overdue = due ? due < today : false;
                        const overdueClass =
                          overdue && t.kind === "email"
                            ? t.followUpStage === "followup"
                              ? "border-red-200 bg-red-50"
                              : "border-yellow-200 bg-yellow-50"
                            : "";
                        return (
                          <div key={t.id} className={`flex items-start justify-between gap-2 rounded-2xl border p-3 ${overdueClass}`}>
                          <div>
                            <div className="flex items-center gap-2">
                              {t.status === "done" ? <CheckCircle2 className="h-4 w-4" /> : <AlertCircle className="h-4 w-4" />}
                              <div className="font-medium">{t.title}</div>
                              {t.kind === "email" ? (
                                <Badge variant={t.followUpStage === "urgent" ? "destructive" : "outline"} className="text-[10px]">
                                  {t.followUpStage === "urgent"
                                    ? "Urgent call"
                                    : t.followUpStage === "followup"
                                      ? "Follow-up"
                                      : "Initial email"}
                                </Badge>
                              ) : null}
                            </div>
                            <div className="mt-1 text-xs text-muted-foreground">
                              Owner: {t.owner}
                              {t.supplierCode ? ` вЂў Supplier: ${t.supplierCode}` : ""}
                              {t.supplierName ? ` вЂў ${t.supplierName}` : ""}
                            </div>
                            <div className="mt-1 text-xs text-muted-foreground">
                              Created: {t.createdAt ? fmtDate(t.createdAt) : "N/A"}
                              {t.dueDate ? ` вЂў Due: ${t.dueDate}` : ""}
                              {t.contactName || t.contactEmail
                                ? ` вЂў Contact: ${t.contactName || ""}${t.contactEmail ? ` (${t.contactEmail})` : ""}`
                                : ""}
                            </div>
                            {t.note ? <div className="mt-2 text-sm">{t.note}</div> : null}
                          </div>
                          <div className="flex gap-2">
                            {t.kind === "email" && t.followUpStage === "initial" && t.status !== "done" ? (
                              <Button
                                variant="outline"
                                size="sm"
                                onClick={() => {
                                  const baseDate = new Date().toISOString().slice(0, 10);
                                  const followUpDays = emailAutomation.followUpDays ?? 7;
                                  updateTask(t.id, {
                                    followUpStage: "followup",
                                    dueDate: addDaysToDate(baseDate, followUpDays),
                                    date: addDaysToDate(baseDate, followUpDays),
                                    note: `${t.note ?? ""}\nFollow-up email sent.`,
                                  });
                                }}
                              >
                                Send follow-up
                              </Button>
                            ) : null}
                            {t.kind === "email" && t.followUpStage === "followup" && t.status !== "done" ? (
                              <Button
                                variant="destructive"
                                size="sm"
                                onClick={() =>
                                  updateTask(t.id, {
                                    followUpStage: "urgent",
                                    note: `${t.note ?? ""}\nEscalated: urgent call required.`,
                                  })
                                }
                              >
                                Escalate call
                              </Button>
                            ) : null}
                            <Button variant="outline" size="sm" onClick={() => toggleTask(t.id)}>
                              {t.status === "done" ? "Undo" : "Done"}
                            </Button>
                            <Button variant="destructive" size="sm" onClick={() => deleteTask(t.id)}>
                              Delete
                            </Button>
                          </div>
                        </div>
                        );
                      })
                    )}
                  </div>
                </CardContent>
              </Card>

              <Card>
                <CardHeader>
                  <CardTitle className="text-base">Month view</CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="grid grid-cols-7 gap-2 text-xs text-muted-foreground">
                    {["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].map((d) => (
                      <div key={d} className="text-center">{d}</div>
                    ))}
                  </div>
                  <div className="mt-2 grid grid-cols-7 gap-2">
                    {calendarDays.map((c) => {
                      const has = (tasksByDate[c.date] ?? []).length;
                      const isSelected = c.date === taskDate;
                      return (
                        <button
                          key={c.date}
                          onClick={() => setTaskDate(c.date)}
                          className={`rounded-2xl border p-2 text-left transition ${isSelected ? "bg-muted" : "hover:bg-muted/50"}`}
                        >
                          <div className={`text-sm font-medium ${c.inMonth ? "" : "text-muted-foreground"}`}>{c.day}</div>
                          {has ? <div className="mt-1 text-[11px] text-muted-foreground">{has} tasks</div> : <div className="mt-1 text-[11px] text-muted-foreground">&nbsp;</div>}
                        </button>
                      );
                    })}
                  </div>
                </CardContent>
              </Card>
            </div>

            <Card>
              <CardHeader>
                <CardTitle className="text-base text-red-600">Red Ignorance Alert</CardTitle>
              </CardHeader>
              <CardContent>
                {redIgnoranceAlertTasks.length === 0 ? (
                  <div className="text-sm text-muted-foreground">No escalations pending.</div>
                ) : (
                  <div className="space-y-2">
                    {redIgnoranceAlertTasks.map((t) => (
                      <div key={t.id} className="rounded-2xl border border-red-200 bg-red-50 p-3">
                        <div className="flex items-center justify-between gap-2">
                          <div>
                            <div className="font-medium">{t.supplierName ?? t.supplierCode ?? "Unknown supplier"}</div>
                            <div className="text-xs text-muted-foreground">
                              {t.contactName || t.contactEmail
                                ? `Contact: ${t.contactName || ""}${t.contactEmail ? ` (${t.contactEmail})` : ""}`
                                : "No contact provided"}
                            </div>
                          </div>
                          <Badge variant="destructive">Urgent call</Badge>
                        </div>
                      <div className="mt-1 text-xs text-muted-foreground">
                        Due: {t.dueDate ?? t.date} вЂў Status: Ignore Risk
                      </div>
                    </div>
                  ))}
                  </div>
                )}
              </CardContent>
            </Card>
            </>
            ) : null}
          </TabsContent>
        </Tabs>

        <Separator className="my-6" />

        <div className="grid gap-3 md:grid-cols-3">
          <Card className="md:col-span-2">
            <CardHeader>
              <CardTitle className="text-base">Operational guidance baked into the model</CardTitle>
            </CardHeader>
            <CardContent className="space-y-2 text-sm text-muted-foreground">
              <div>
                <span className="font-medium text-foreground">Golden key:</span> Supplier Code. Names are attributes, not identifiers.
              </div>
              <div>
                <span className="font-medium text-foreground">Hierarchy of truth:</span> Supplier overrides в†’ Category/Family rules в†’ Data flags.
              </div>
              <div>
                <span className="font-medium text-foreground">Fast action:</span> Worklist is High risk & not signed, sorted by spend.
              </div>
              <div>
                <span className="font-medium text-foreground">Auditability:</span> Every change writes to Activity Log.
              </div>
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle className="text-base">System health</CardTitle>
            </CardHeader>
            <CardContent className="space-y-2">
              <div className="flex items-center justify-between">
                <div className="text-sm text-muted-foreground">Rows loaded</div>
                <div className="font-semibold">{db.factRows.length}</div>
              </div>
              <div className="flex items-center justify-between">
                <div className="text-sm text-muted-foreground">Overrides</div>
                <div className="font-semibold">{Object.keys(db.overrides).length}</div>
              </div>
              <div className="flex items-center justify-between">
                <div className="text-sm text-muted-foreground">Rules</div>
                <div className="font-semibold">{db.categoryRules.length}</div>
              </div>
              <div className="flex items-center justify-between">
                <div className="text-sm text-muted-foreground">Tasks</div>
                <div className="font-semibold">{db.tasks.length}</div>
              </div>
              <div className="flex items-center justify-between">
                <div className="text-sm text-muted-foreground">Log entries</div>
                <div className="font-semibold">{db.log.length}</div>
              </div>
            </CardContent>
          </Card>
        </div>
      </div>
    </div>
  );
}

function PlusIcon() {
  return (
    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path d="M12 5V19" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
      <path d="M5 12H19" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
    </svg>
  );
}
