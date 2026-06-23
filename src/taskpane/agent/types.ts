import { WorkbookContextSnapshot } from "../context/types";

export type AgentMacroName =
  | "place_output"
  | "write_formula"
  | "apply_format"
  | "create_chart"
  | "join_tables"
  | "table_view"
  | "summarize_actions"
  | "validate_data";

export type PlanConfirmationChoice = { id: string; label: string };

export interface PlanConfirmation {
  id: string;
  question: string;
  choices: PlanConfirmationChoice[];
  required: boolean;
}

export interface PlanArtifact {
  id: string;
  type: "table" | "chart";
  sheet: string;
  anchor: string;
  fromStep: string;
  sheetName?: string;
  blockRef?: string;
  addressA1?: string;
  headers?: string[];
  headerAliases?: Record<string, string[]>;
  rowCount?: number;
  colCount?: number;
  counts?: Record<string, number>;
}

export interface PlanStep {
  id: string;
  macro: AgentMacroName;
  params: Record<string, any>;
}

export interface ChartColumnRef {
  colIndex?: number;
  headerName?: string;
  header?: string;
}

export type JoinMatchStrategy = "case_insensitive_trim" | "numeric" | "exact";
export type JoinMultipleMatchesBehavior = "explode_rows" | "first";

export interface JoinTablesKey {
  left: string;
  right: string;
  strategy?: JoinMatchStrategy;
}

export interface JoinSelectConfig {
  mode: "all" | "list" | "all_except_keys";
  columns?: string[];
}

export interface JoinConflictConfig {
  onDuplicateRightColumns?: "suffix" | "overwrite_left" | "skip";
  rightSuffix?: string;
  onMultipleMatches?: JoinMultipleMatchesBehavior;
}

export interface JoinOutputConfig {
  mode?: "right" | "below" | "newSheet";
  anchor?: { blockRef?: string; sheet?: string; cell?: string };
  sheetName?: string;
  tableName?: string;
}

export interface JoinTablesParamsNormalized {
  left: { blockRef: string; artifactRef?: string };
  right: { blockRef: string; artifactRef?: string };
  keys: JoinTablesKey[];
  joinType: "left" | "inner" | "anti_left" | "full";
  allowKeyFallback?: boolean;
  keepRightKeyColumns?: boolean;
  selectionPolicy?: "defaultAll" | "explicit";
  select?: { left?: JoinSelectConfig; right?: JoinSelectConfig };
  conflict?: JoinConflictConfig;
  output?: JoinOutputConfig;
  match?: { defaultStrategy?: JoinMatchStrategy };
}

export interface ValidateDataInternalState {
  decisions?: {
    fixMissing?: boolean;
    fixDuplicates?: boolean;
    fixBadType?: boolean;
  };
  executions?: number;
  phase?: "confirmations";
}

export interface ValidateDataParams {
  source?: { blockRef?: string; artifactRef?: string };
  detect?: { missing?: boolean; duplicates?: boolean; badType?: boolean };
  missingColumns?: string[];
  duplicateKeyColumns?: string[];
  typeRules?: { col: string; type: "date" | "number" | "text" }[];
  options?: { maxIssues?: number };
  __internal?: ValidateDataInternalState;
}

export interface PlanContextTrace {
  rawPlan?: any;
  canonicalPlan?: any;
  sanitizedPlan?: any;
  repairedPlan?: any;
  validatedPlan?: any;
}

export interface AgentPlan {
  version: "1.0";
  goal: string;
  steps: PlanStep[];
  confirmations?: PlanConfirmation[];
  artifacts?: PlanArtifact[];
}

export type AgentLogLevel = "info" | "warn" | "error";

export interface AgentLogEntry {
  ts: string;
  level: AgentLogLevel;
  message: string;
  macro?: AgentMacroName;
  stepId?: string;
  data?: any;
}

export interface ConfirmationRequest {
  id: string;
  question: string;
  choices: PlanConfirmationChoice[];
}

export interface ArtifactRecord {
  id?: string;
  type: "table" | "chart" | "range";
  kind?: "table" | "chart" | "range";
  sheet: string;
  anchor: string;
  fromStep?: string;
  tableName?: string;
  rows?: number;
  cols?: number;
  rowCount?: number;
  colCount?: number;
  headers?: string[];
  headerAliases?: Record<string, string[]>;
  joinType?: string;
  matchedRows?: number;
  outputRows?: number;
  details?: any;
  blockRef?: string;
  sheetName?: string;
  address?: string;
  addressA1?: string;
  counts?: Record<string, number>;
}

export type ExecutionStatus = "ok" | "need_user_confirmation" | "error" | "skipped";

export interface ExecutionResult {
  logs: AgentLogEntry[];
  artifacts: ArtifactRecord[];
  confirmationsRequested?: ConfirmationRequest[];
  status: ExecutionStatus;
  errors?: string[];
  ok?: boolean;
  warnings?: string[];
}

export interface ExecutionOptions {
  confirmationDecisions?: Record<string, string>;
  attempt?: number;
  autoAnswerMode?: "demoEval" | "interactive" | "none";
  initialArtifacts?: ArtifactRecord[];
}

export interface MacroContext {
  excelCtx: Excel.RequestContext;
  plan: AgentPlan;
  context: WorkbookContextSnapshot;
  step: PlanStep;
  decisions: Record<string, string>;
  artifacts: ArtifactRecord[];
  lastAddedHeader?: string | null;
  headerRegistry?: { artifactRef?: string; blockRef?: string; datasetRef?: string; headerName: string; sourceStepId: string }[];
  datasetRef?: string;
  log(entry: Omit<AgentLogEntry, "ts">): void;
}

export interface MacroResult {
  artifacts?: ArtifactRecord[];
  requiresConfirmation?: ConfirmationRequest;
  status?: ExecutionStatus;
}

export type CreateChartSource = { blockRef?: string; tableName?: string; sheetName?: string; artifactRef?: string };
export type CreateChartDest = {
  mode: "right" | "below" | "newSheet";
  anchor?: { blockRef?: string; artifactRef?: string };
  sheet?: string;
  sheetName?: string;
  sheetNameHint?: string;
  titleHint?: string;
};
export type CreateChartParams = {
  source: CreateChartSource;
  mapping: { xCol: ChartColumnRef; yCols: ChartColumnRef[] };
  chartType?: "columnClustered" | "line" | "barClustered";
  dest: CreateChartDest;
};
