export type RiskLevel = "High" | "Non-risk" | "Unknown";
export type ContractStatus = "Not sent" | "Sent" | "Signed" | "Review" | "Not compliant" | "N/A";
export type BarFillStyle = "solid" | "diagonal" | "dots";
export type BarPaletteKey = "country" | "category" | "riskLevel" | "contractStatus" | "funnelStage";

export type BarChartSettings = {
  fillStyle: BarFillStyle;
  palettes: Record<BarPaletteKey, Record<string, string>>;
  rulesNote?: string;
};

export type EmailAutomationSettings = {
  initialFollowUpDays: number;
  followUpDays: number;
};

export type FactRow = {
  supplierName?: string;
  supplierCode?: string;
  country?: string;
  entity?: string;
  year?: string | number;
  category?: string;
  family?: string;
  subFamily?: string;
  spend?: number;
  po?: number;
  riskFlag?: string;
  status?: string;
  completion?: string;
};

export type SupplierOverride = {
  canonicalName?: string;
  risk?: RiskLevel;
  contractStatus?: ContractStatus;
  evaluated?: boolean;
  note?: string;
};

export type AppSettings = {
  dominanceShareThreshold: number;
  multiCategorySecondShareThreshold: number;
  scopeSpendThreshold: number;
  missingSubfamilySpendThreshold: number;
  barChartSettings: BarChartSettings;
  emailAutomation: EmailAutomationSettings;
  bulkContactName: string;
  bulkContactEmail: string;
  statusPalette: {
    risk: Record<RiskLevel, string>;
    contract: Record<ContractStatus, string>;
    supplierContacts?: Record<string, string>;
  };
  supplierContacts?: Record<string, string>;
};

export type UndoBatch = {
  id: string;
  ts: string;
  actor: string;
  action: "risk" | "contractStatus" | "canonicalName" | "categoryFastAction" | "evaluated" | "bulkMark";
  summary: string;
  items: { code: string; prev?: SupplierOverride }[];
};

export type CategoryRule = {
  key: string;
  scope: "Category" | "Family";
  defaultRisk: RiskLevel;
  comment?: string;
};

export type WorkSheet = any;
export type SheetMatrix = any[][];

export type SupplierMaster = {
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
  lastTouch?: string;
};

export type IssueSeverity = "Critical" | "High" | "Medium" | "Low";
export type IssueType =
  | "CODE_MANY_NAMES"
  | "NAME_MANY_CODES"
  | "MULTI_CATEGORY_EXPOSURE"
  | "RISK_UNKNOWN_IN_SCOPE"
  | "MISSING_SUBFAMILY_HIGH_SPEND"
  | "POLICY_SUPPRESSED_BY_GUARDRAIL";

export type IssueInstance = {
  id: string;
  type: IssueType;
  severity: IssueSeverity;
  title: string;
  code?: string;
  supplierName?: string;
  spendAffected?: number;
  details?: any;
};

export type LogEntry = {
  id: string;
  ts: string;
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

export type Task = {
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

export type EvaluationChoice = "Evaluated" | "Not evaluated" | "Clear";
export type BulkRiskChoice = RiskLevel | "No change";
export type BulkEvalChoice = EvaluationChoice | "No change";

export type PersistedDB = {
  factRows: FactRow[];
  overrides: Record<string, SupplierOverride>;
  categoryRules: CategoryRule[];
  settings: AppSettings;
  tasks: Task[];
  log: LogEntry[];
  lastUndo?: UndoBatch | null;
};

export type LegacySettings = Partial<AppSettings> & {
  countryPalette?: Record<string, string>;
  countryBarFillStyle?: BarFillStyle;
};
