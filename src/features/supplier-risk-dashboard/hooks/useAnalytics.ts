import { useMemo } from "react";
import type { IssueInstance, IssueType, RiskLevel, SupplierMaster } from "../types";

export type AnalyticsInput = {
  suppliers: SupplierMaster[];
  duplicates: Array<{ code: string }>;
  issues: IssueInstance[];
  issueBuckets: Record<IssueType, number>;
  overviewSuppliers: SupplierMaster[];
  overviewCategoryMetrics: Array<{ category: string; suppliers: number; highRiskSuppliers: number; riskShare: number; totalSpend: number; highRiskSpend: number }>;
  overviewCountryMetrics: Array<{ country: string; suppliers: number; highRiskSuppliers: number; totalSpend: number; highRiskSpend: number; riskShare: number }>;
  categoryMetrics: Array<{ category: string; suppliers: number; highRiskSuppliers: number; riskShare: number; totalSpend: number; highRiskSpend: number }>;
  countryMetrics: Array<{ country: string; suppliers: number; highRiskSuppliers: number; totalSpend: number; highRiskSpend: number; riskShare: number }>;
  globalCountry: string;
  categoryTopSort: "HighRiskDesc" | "TotalDesc";
  categorySpendSort: "HighRiskDesc" | "TotalDesc";
};

export function useAnalytics({
  suppliers,
  duplicates,
  issues,
  issueBuckets,
  overviewSuppliers,
  overviewCategoryMetrics,
  overviewCountryMetrics,
  categoryMetrics,
  countryMetrics,
  globalCountry,
  categoryTopSort,
  categorySpendSort,
}: AnalyticsInput) {
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
    const sent = overviewSuppliers.filter((s) => s.contractStatus === "Sent").length;
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

  return {
    kpis,
    overviewIssues,
    overviewIssueBuckets,
    overviewKpis,
    funnelData,
    funnelStages,
    topRiskyByCount,
    topRiskyBySpend,
    topCountriesByCount,
    topCountriesBySpend,
    topCountriesByRiskShare,
    riskMix,
  };
}
