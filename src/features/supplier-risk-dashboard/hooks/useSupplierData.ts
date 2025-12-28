import { useMemo } from "react";
import type {
  ContractStatus,
  IssueSeverity,
  IssueType,
  PersistedDB,
  RiskLevel,
  SupplierMaster,
} from "../types";
import {
  buildCategoryMetrics,
  buildCountryMetrics,
  buildDuplicates,
  buildIssues,
  buildSupplierMaster,
  normCode,
  safeStr,
  uniq,
} from "../utils";

export type SupplierDataFilters = {
  issueTypeFilter: IssueType | "All";
  issueSeverityFilter: IssueSeverity | "All";
  categoryTableSort: { key: "category" | "suppliers" | "highRisk" | "riskShare" | "spend"; dir: "asc" | "desc" };
  q: string;
  riskFilter: RiskLevel | "All";
  evalFilter: "All" | "Evaluated" | "Not evaluated";
  contractFilter: ContractStatus | "All";
  categoryFilter: string | "All";
  fastActionSort: { key: "spend" | "name" | "contract"; dir: "asc" | "desc" };
  fastCategory: string;
  bulkNameQuery: string;
  drillCategory: string | null;
  drillCountry: string | null;
  categoryTableSortState: { key: string; dir: "asc" | "desc" };
  drillSortState: { key: string; dir: "asc" | "desc" };
  fastQueueSort: "SpendDesc" | "SpendAsc" | "NameAsc" | "NameDesc";
};

type SupplierDataInput = {
  db: PersistedDB;
  globalCountry: string;
  selectedDuplicateCode: string | null;
  filters: SupplierDataFilters;
};

export function useSupplierData({ db, globalCountry, selectedDuplicateCode, filters }: SupplierDataInput) {
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
      .filter((i) => (filters.issueTypeFilter === "All" ? true : i.type === filters.issueTypeFilter))
      .filter((i) => (filters.issueSeverityFilter === "All" ? true : i.severity === filters.issueSeverityFilter))
      .slice(0, 1000);
  }, [issues, filters.issueTypeFilter, filters.issueSeverityFilter]);

  const categoryMetrics = useMemo(() => buildCategoryMetrics(scopedFactRows, suppliers), [scopedFactRows, suppliers]);
  const countryMetrics = useMemo(() => buildCountryMetrics(scopedFactRows, suppliers), [scopedFactRows, suppliers]);

  const categories = useMemo(() => ["All", ...uniq(categoryMetrics.map((c) => c.category)).sort()], [categoryMetrics]);
  const paletteCountries = useMemo(() => availableCountries.filter((c) => c !== "Overall"), [availableCountries]);
  const paletteCategories = useMemo(() => categories.filter((c) => c !== "All"), [categories]);

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
    const dir = filters.categoryTableSort.dir === "asc" ? 1 : -1;
    return [...categoryMetrics].sort((a, b) => {
      switch (filters.categoryTableSort.key) {
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
  }, [categoryMetrics, filters.categoryTableSort]);

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

  const supplierByCode = useMemo(() => {
    const m = new Map<string, SupplierMaster>();
    for (const s of suppliers) m.set(s.code, s);
    return m;
  }, [suppliers]);

  const bulkNameMatchesList = useMemo(() => {
    const needle = filters.bulkNameQuery.trim().toLowerCase();
    if (!needle) return [];
    return suppliers
      .filter((s) => s.canonicalName.toLowerCase().includes(needle) || s.code.toLowerCase().includes(needle))
      .slice(0, 200);
  }, [filters.bulkNameQuery, suppliers]);

  const categoryFastPreview = useMemo(() => {
    const cat = safeStr(filters.fastCategory);
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
  }, [filters.fastCategory, suppliers, db.settings.dominanceShareThreshold, db.settings.multiCategorySecondShareThreshold]);

  const sortedCategoryMetrics = useMemo(() => {
    const sorted = [...categoryMetrics];
    const dir = filters.categoryTableSortState.dir === "asc" ? 1 : -1;
    sorted.sort((a, b) => {
      const key = filters.categoryTableSortState.key;
      if (key === "category") return a.category.localeCompare(b.category) * dir;
      return (Number((a as any)[key] ?? 0) - Number((b as any)[key] ?? 0)) * dir;
    });
    return sorted;
  }, [categoryMetrics, filters.categoryTableSortState]);

  const drillSuppliers = useMemo(() => {
    if (!filters.drillCategory) return [] as any[];
    const by: Record<string, { spend: number; po: number }> = {};
    for (const r of scopedFactRows) {
      if (safeStr(r.category) !== filters.drillCategory) continue;
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
          categories: s?.categories ?? [filters.drillCategory],
        };
      })
      .sort((a: any, b: any) => b.spendInSlice - a.spendInSlice)
      .slice(0, 500);
  }, [filters.drillCategory, scopedFactRows, overviewSupplierByCode]);

  const drillSuppliersByCountry = useMemo(() => {
    if (!filters.drillCountry) return [] as any[];
    const by: Record<string, { spend: number; po: number }> = {};
    for (const r of scopedFactRows) {
      const c = safeStr(r.country) || "(unknown)";
      if (c !== filters.drillCountry) continue;
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
  }, [filters.drillCountry, scopedFactRows, overviewSupplierByCode]);

  const supplierFiltered = useMemo(() => {
    const needle = filters.q.trim().toLowerCase();
    return suppliers
      .filter((s) => {
        if (filters.riskFilter !== "All" && s.risk !== filters.riskFilter) return false;
        if (filters.evalFilter === "Evaluated" && !s.evaluated) return false;
        if (filters.evalFilter === "Not evaluated" && s.evaluated) return false;
        if (filters.contractFilter !== "All" && s.contractStatus !== filters.contractFilter) return false;
        if (filters.categoryFilter !== "All" && !s.categories.includes(filters.categoryFilter)) return false;
        if (!needle) return true;
        return (
          s.code.toLowerCase().includes(needle) ||
          s.canonicalName.toLowerCase().includes(needle) ||
          s.categories.some((c) => c.toLowerCase().includes(needle)) ||
          s.families.some((f) => f.toLowerCase().includes(needle))
        );
      })
      .slice(0, 1500);
  }, [
    suppliers,
    filters.q,
    filters.riskFilter,
    filters.evalFilter,
    filters.contractFilter,
    filters.categoryFilter,
  ]);

  const worklist = useMemo(() => {
    const base = suppliers.filter((s) => s.risk === "High" && s.contractStatus !== "Signed");
    const sorted = [...base].sort((a, b) => {
      const dir = filters.fastActionSort.dir === "asc" ? 1 : -1;
      if (filters.fastActionSort.key === "name") return dir * a.canonicalName.localeCompare(b.canonicalName);
      if (filters.fastActionSort.key === "contract") return dir * a.contractStatus.localeCompare(b.contractStatus);
      return dir * (a.totalSpend - b.totalSpend);
    });
    return sorted.slice(0, 200);
  }, [suppliers, filters.fastActionSort]);

  const worklistSorted = useMemo(() => {
    const sorted = [...worklist];
    const dir = filters.fastQueueSort.includes("Asc") ? 1 : -1;
    if (filters.fastQueueSort.startsWith("Spend")) {
      sorted.sort((a, b) => (a.totalSpend - b.totalSpend) * dir);
    } else {
      sorted.sort((a, b) => a.canonicalName.localeCompare(b.canonicalName) * dir);
    }
    return sorted;
  }, [worklist, filters.fastQueueSort]);

  const drillRows = useMemo(() => {
    const rows = (filters.drillCategory ? drillSuppliers : drillSuppliersByCountry) as any[];
    const sorted = [...rows];
    const dir = filters.drillSortState.dir === "asc" ? 1 : -1;
    const orderRisk: Record<RiskLevel, number> = { High: 3, "Non-risk": 2, Unknown: 1 };
    sorted.sort((a, b) => {
      const key = filters.drillSortState.key;
      if (key === "risk") return (orderRisk[a.risk as RiskLevel] - orderRisk[b.risk as RiskLevel]) * dir;
      if (key === "contractStatus") return String(a.contractStatus).localeCompare(String(b.contractStatus)) * dir;
      if (key === "canonicalName") return String(a.canonicalName).localeCompare(String(b.canonicalName)) * dir;
      if (key === "categories") return String(a.categories?.[0] ?? "").localeCompare(String(b.categories?.[0] ?? "")) * dir;
      return (Number(a[key] ?? 0) - Number(b[key] ?? 0)) * dir;
    });
    return sorted;
  }, [filters.drillCategory, drillSuppliers, drillSuppliersByCountry, filters.drillSortState]);

  return {
    availableCountries,
    scopedFactRows,
    suppliers,
    allSuppliers,
    duplicates,
    selectedDuplicate,
    issues,
    issueBuckets,
    issuesFiltered,
    categoryMetrics,
    countryMetrics,
    categories,
    paletteCountries,
    paletteCategories,
    riskyCategoryChart,
    candidateStats,
    redIgnoranceAlertTasks,
    categoryMetricsSorted,
    overviewFactRows,
    overviewSuppliers,
    overviewSupplierByCode,
    overviewCategoryMetrics,
    overviewCountryMetrics,
    supplierByCode,
    bulkNameMatchesList,
    categoryFastPreview,
    sortedCategoryMetrics,
    drillSuppliers,
    drillSuppliersByCountry,
    supplierFiltered,
    worklist,
    worklistSorted,
    drillRows,
  };
}
