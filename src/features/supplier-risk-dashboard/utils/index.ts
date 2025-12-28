import * as XLSX from "xlsx";
import type {
  AppSettings,
  BarChartSettings,
  BarFillStyle,
  CategoryRule,
  ContractStatus,
  FactRow,
  IssueInstance,
  IssueSeverity,
  IssueType,
  LegacySettings,
  LogEntry,
  PersistedDB,
  RiskLevel,
  SheetMatrix,
  SupplierMaster,
  SupplierOverride,
  WorkSheet,
} from "../types";

export const LS_KEY = "supplier-risk-ops-dashboard:v1";

export const RISK_LEVEL_OPTIONS: RiskLevel[] = ["High", "Non-risk", "Unknown"];
export const CONTRACT_STATUS_OPTIONS: ContractStatus[] = ["Not sent", "Sent", "Signed", "Review", "Not compliant", "N/A"];

export const DEFAULT_RISK_PALETTE: Record<RiskLevel, string> = {
  High: "#fca5a5",
  "Non-risk": "#86efac",
  Unknown: "#e5e7eb",
};

export const DEFAULT_CONTRACT_PALETTE: Record<ContractStatus, string> = {
  "Not sent": "#fcd34d",
  Sent: "#fde68a",
  Signed: "#86efac",
  Review: "#93c5fd",
  "Not compliant": "#f87171",
  "N/A": "#e5e7eb",
};

export const DEFAULT_SETTINGS: AppSettings = {
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
  supplierContacts: {},
};

export const DEMO_FACT_ROWS: FactRow[] = [
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

export const DEMO_RULES: CategoryRule[] = [
  { key: "Logistics - Ground", scope: "Category", defaultRisk: "High", comment: "Demo policy" },
  { key: "Automation", scope: "Category", defaultRisk: "High", comment: "Demo policy" },
  { key: "Raw Materials", scope: "Family", defaultRisk: "High", comment: "Demo policy" },
];

export function nowIso() {
  return new Date().toISOString();
}

export function uuid() {
  return Math.random().toString(16).slice(2) + "-" + Math.random().toString(16).slice(2);
}

export function safeStr(v: any) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

export function normCode(v: any) {
  return safeStr(v);
}

export function toNum(v: any): number {
  if (v === null || v === undefined || v === "") return 0;
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

export function uniq<T>(arr: T[]) {
  return Array.from(new Set(arr)).filter((x) => x !== undefined && x !== null) as T[];
}

export function pickCanonicalByFrequency(candidates: { name: string; count: number }[]) {
  if (!candidates.length) return "";
  const sorted = [...candidates].sort((a, b) => b.count - a.count);
  return sorted[0].name;
}

export function computeCategorySharesFromRows(rows: FactRow[]) {
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

export function normalizeContractStatus(raw: string): ContractStatus {
  const s = safeStr(raw).toLowerCase();
  if (!s) return "Not sent";
  if (s.includes("not compliant")) return "Not compliant";
  if (s.includes("signed")) return "Signed";
  if (s.includes("sent")) return "Sent";
  if (s.includes("review")) return "Review";
  if (s.includes("assess") || s.includes("asess") || s.includes("asessed") || s.includes("assessed")) return "Not sent";
  return "Not sent";
}

export function normYesNo(raw: string) {
  const s = safeStr(raw).trim().toLowerCase();
  if (!s) return "";
  if (s.startsWith("y")) return "YES";
  if (s.startsWith("n")) return "NO";
  return s.toUpperCase();
}

export function riskFromFlags(flags: (string | null | undefined)[]): RiskLevel {
  const normalized = flags
    .map((x) => safeStr(x))
    .filter(Boolean)
    .map((x) => x.toUpperCase());

  if (normalized.some((x) => x === "YES")) return "High";
  if (normalized.some((x) => x.includes("#REF")) || normalized.length === 0) return "Unknown";
  if (normalized.every((x) => x === "NO")) return "Non-risk";

  return "Unknown";
}

export function evaluatedFromFlags(flags: (string | null | undefined)[]) {
  const normalized = flags
    .map((x) => safeStr(x))
    .filter(Boolean)
    .map((x) => x.toUpperCase());
  if (normalized.length === 0) return false;
  if (normalized.some((x) => x.includes("#REF"))) return false;
  return normalized.every((x) => x === "YES" || x === "NO");
}

export function addDaysToDate(date: string, days: number) {
  const d = new Date(`${date}T00:00:00`);
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}

export function normalizeSettings(settings?: LegacySettings): AppSettings {
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

export function loadDB(): PersistedDB {
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

export function saveDB(db: PersistedDB) {
  localStorage.setItem(LS_KEY, JSON.stringify(db));
}

export function sheetToMatrix(ws: WorkSheet): SheetMatrix {
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null }) as any[][];
}

export function findHeaderRowIndex(matrix: SheetMatrix, requiredHeaders: string[], maxScan = 60): number {
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

export function buildHeaderMap(headerRow: any[]): Record<string, number> {
  const map: Record<string, number> = {};
  headerRow.forEach((cell, idx) => {
    const key = safeStr(cell);
    if (!key) return;
    map[key.toLowerCase()] = idx;
  });
  return map;
}

export const HEADER_ALIASES: Record<keyof FactRow, string[]> = {
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
  riskFlag: ["risk", "risk by given data", "riskflag", "risk (final)", "risk evolution", "risk status", "risk.1", "risk"],
  status: ["status", "agreement status", "contract status"],
  completion: ["completion status", "completion", "evaluation completion"],
};

export function getCellByAliases(row: any[], headerMap: Record<string, number>, aliases: string[]) {
  for (const a of aliases) {
    const idx = headerMap[a.toLowerCase()];
    if (idx !== undefined) return row[idx];
  }
  return null;
}

export function parseFactSheet(ws: WorkSheet, requiredForHeaderDetection: string[]): FactRow[] {
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

export function parseRiskEvalOverview(ws: WorkSheet): FactRow[] {
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

export function mergeCategoryRules(existing: CategoryRule[], incoming: CategoryRule[]) {
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

export function parseFamilyRiskRulesFromSettings(ws: WorkSheet): CategoryRule[] {
  const matrix = sheetToMatrix(ws);
  const headerIdx = findHeaderRowIndex(matrix, ["family name", "risk"]);
  const headerRow = matrix[headerIdx] ?? [];
  const headerMap = buildHeaderMap(headerRow);

  const famIdx = headerMap["family name.1"] ?? headerMap["family name"] ?? headerMap["family"] ?? null;
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
    else continue;

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

export function normalizeJsonRowsToFactRows(rows: any[]): FactRow[] {
  const pick = (obj: any, aliases: string[]) => {
    for (const a of aliases) {
      if (obj?.[a] !== undefined) return obj[a];
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

export function pickRowsFromUnknownJson(parsed: any): any[] {
  if (Array.isArray(parsed)) return parsed;
  if (!parsed || typeof parsed !== "object") return [];
  if (Array.isArray((parsed as any)["Р›РёСЃС‚1"])) return (parsed as any)["Р›РёСЃС‚1"];
  for (const k of Object.keys(parsed)) {
    const v = (parsed as any)[k];
    if (Array.isArray(v) && v.length && typeof v[0] === "object") return v;
  }
  return [];
}

export function buildSupplierMaster(
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

    const secondShare = shareInfo.secondCategoryShare ?? 0;
    const multiCategory = secondShare >= settings.multiCategorySecondShareThreshold;
    if (multiCategory) policyFromCategory = "Unknown";

    const policyRisk: RiskLevel =
      [policyFromFamily, policyFromCategory].includes("High")
        ? "High"
        : [policyFromFamily, policyFromCategory].includes("Non-risk")
          ? "Non-risk"
          : "Unknown";

    const risk: RiskLevel = ov?.risk ?? (dataRisk !== "Unknown" ? dataRisk : policyRisk);

    const evaluated =
      ov?.evaluated !== undefined
        ? ov.evaluated
        : evaluatedFromFlags(flags) ||
          (ov?.risk !== undefined && ov?.risk !== "Unknown") ||
          (dataRisk === "Unknown" && policyRisk !== "Unknown");

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

export function buildDuplicates(suppliers: SupplierMaster[]) {
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

export function normNameKey(name: string) {
  return safeStr(name).toLowerCase().replace(/\s+/g, " ").trim();
}

export function buildNameManyCodes(factRows: FactRow[]) {
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

export function buildMissingSubfamilySpendByCode(factRows: FactRow[]) {
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

export function buildIssues(
  factRows: FactRow[],
  suppliers: SupplierMaster[],
  duplicates: ReturnType<typeof buildDuplicates>,
  settings: AppSettings,
  categoryRules: CategoryRule[]
): IssueInstance[] {
  const issues: IssueInstance[] = [];

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
    sevWeight[b.severity] - sevWeight[a.severity] || Math.abs((b.spendAffected ?? 0)) - Math.abs((a.spendAffected ?? 0))
  );
}

export function buildCategoryMetrics(factRows: FactRow[], suppliers: SupplierMaster[]) {
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

export function buildCountryMetrics(factRows: FactRow[], suppliers: SupplierMaster[]) {
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

export function fmtMoney(n: number) {
  const abs = Math.abs(n);
  const sign = n < 0 ? "-" : "";
  if (abs >= 1e9) return `${sign}${(abs / 1e9).toFixed(2)}B`;
  if (abs >= 1e6) return `${sign}${(abs / 1e6).toFixed(2)}M`;
  if (abs >= 1e3) return `${sign}${(abs / 1e3).toFixed(1)}K`;
  return `${sign}${abs.toFixed(0)}`;
}

export function clamp01(v: number) {
  if (Number.isNaN(v) || !Number.isFinite(v)) return 0;
  return Math.min(1, Math.max(0, v));
}

export function fmtDate(d: string) {
  try {
    const dt = new Date(d);
    return dt.toLocaleString();
  } catch {
    return d;
  }
}

export function todayIso() {
  return new Date().toISOString().slice(0, 10);
}

export function addDays(dateStr: string, days: number) {
  const d = new Date(`${dateStr}T00:00:00`);
  d.setDate(d.getDate() + days);
  return d.toISOString().slice(0, 10);
}

export function shortLabel(label: string, max = 22) {
  if (label.length <= max) return label;
  return `${label.slice(0, max - 1)}…`;
}

export function slugifyKey(value: string) {
  return safeStr(value).toLowerCase().replace(/\s+/g, "-").replace(/[^a-z0-9-_]/g, "");
}

export function getReadableTextColor(hex: string) {
  const sanitized = hex.replace("#", "");
  if (sanitized.length !== 6) return "#0f172a";
  const r = parseInt(sanitized.slice(0, 2), 16);
  const g = parseInt(sanitized.slice(2, 4), 16);
  const b = parseInt(sanitized.slice(4, 6), 16);
  const luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
  return luminance > 0.65 ? "#0f172a" : "#f8fafc";
}

export function getRiskColor(_: RiskLevel, fallback: string): string {
  return fallback;
}

export function getContractColor(_: ContractStatus, fallback: string): string {
  return fallback;
}
