import { useMemo } from "react";
import type {
  BulkEvalChoice,
  BulkRiskChoice,
  ContractStatus,
  EvaluationChoice,
  RiskLevel,
  SupplierOverride,
  SupplierMaster,
} from "../types";
import { safeStr, uniq } from "../utils";

export type BulkOpsInput = {
  bulkCodes: string;
  bulkRisk: RiskLevel;
  bulkContract: ContractStatus;
  bulkCanonicalName: string;
  bulkEvaluated: EvaluationChoice;
  supplierByCode: Map<string, SupplierMaster>;
  overrides: Record<string, SupplierOverride>;
};

export function useBulkOps({
  bulkCodes,
  bulkRisk,
  bulkContract,
  bulkCanonicalName,
  bulkEvaluated,
  supplierByCode,
  overrides,
}: BulkOpsInput) {
  const parseCodes = (text: string) => {
    return uniq(
      text
        .split(/[\n,;\t]+/)
        .map((s) => safeStr(s))
        .filter(Boolean)
    );
  };

  const cleanOverride = (ov?: SupplierOverride): SupplierOverride | undefined => {
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
  };

  const planBulkRisk = (codes: string[], target: RiskLevel) => {
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
      const prev = overrides[c];
      const effectiveRisk = s.risk;

      if (target === "Unknown") {
        if (!prev?.risk) continue;
        const patched = cleanOverride({ ...prev, risk: undefined });
        changes.push({ code: c, prev: prev ? { ...prev } : undefined, next: patched });
        continue;
      }

      if (prev?.risk === target) continue;
      if (!prev?.risk && effectiveRisk === target) continue;

      const patched = cleanOverride({ ...(prev || {}), risk: target });
      changes.push({ code: c, prev: prev ? { ...prev } : undefined, next: patched });
    }

    return { found, notFound, changes };
  };

  const planBulkContract = (codes: string[], target: ContractStatus) => {
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
      const prev = overrides[c];
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
  };

  const planBulkCanonicalName = (codes: string[], name: string) => {
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
      const prev = overrides[c];
      const effective = s.canonicalName;

      if (prev?.canonicalName === name) continue;
      if (!safeStr(prev?.canonicalName) && effective === name) continue;

      const patched = cleanOverride({ ...(prev || {}), canonicalName: name });
      changes.push({ code: c, prev: prev ? { ...prev } : undefined, next: patched });
    }

    return { found, notFound, changes };
  };

  const planBulkEvaluated = (codes: string[], target: EvaluationChoice) => {
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
      const prev = overrides[c];
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
  };

  const planBulkMark = (codes: string[], riskChoice: BulkRiskChoice, evalChoice: BulkEvalChoice) => {
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
      const prev = overrides[c];
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

      const patched = cleanOverride(next);
      if (!patched) {
        if (!prev) continue;
        changes.push({ code: c, prev: prev ? { ...prev } : undefined, next: undefined });
        continue;
      }

      if (JSON.stringify(patched) === JSON.stringify(prev ?? {})) continue;

      changes.push({ code: c, prev: prev ? { ...prev } : undefined, next: patched });
    }

    return { found, notFound, changes };
  };

  const bulkParsedCodes = useMemo(() => parseCodes(bulkCodes), [bulkCodes]);
  const bulkRiskPreview = useMemo(
    () => (bulkParsedCodes.length ? planBulkRisk(bulkParsedCodes, bulkRisk) : null),
    [bulkParsedCodes, bulkRisk, supplierByCode, overrides]
  );
  const bulkContractPreview = useMemo(
    () => (bulkParsedCodes.length ? planBulkContract(bulkParsedCodes, bulkContract) : null),
    [bulkParsedCodes, bulkContract, supplierByCode, overrides]
  );
  const bulkCanonicalPreview = useMemo(() => {
    const nm = safeStr(bulkCanonicalName);
    return bulkParsedCodes.length && nm ? planBulkCanonicalName(bulkParsedCodes, nm) : null;
  }, [bulkParsedCodes, bulkCanonicalName, supplierByCode, overrides]);
  const bulkEvaluatedPreviewState = useMemo(
    () => (bulkParsedCodes.length ? planBulkEvaluated(bulkParsedCodes, bulkEvaluated) : null),
    [bulkParsedCodes, bulkEvaluated, supplierByCode, overrides]
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

  return {
    bulkParsedCodes,
    bulkRiskPreview,
    bulkContractPreview,
    bulkCanonicalPreview,
    bulkEvaluatedPreviewState,
    bulkPreviewRows,
    parseCodes,
    planBulkRisk,
    planBulkContract,
    planBulkCanonicalName,
    planBulkEvaluated,
    planBulkMark,
  };
}
