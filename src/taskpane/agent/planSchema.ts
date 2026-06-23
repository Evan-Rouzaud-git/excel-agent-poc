import Ajv, { JSONSchemaType } from "ajv";
import { AgentMacroName, AgentPlan, PlanArtifact, PlanConfirmation, PlanStep } from "./types";

const macroEnum: AgentMacroName[] = ["place_output", "write_formula", "apply_format", "create_chart", "join_tables", "table_view", "summarize_actions", "validate_data"];

const confirmationChoiceSchema = {
  type: "object",
  properties: {
    id: { type: "string", minLength: 1 },
    label: { type: "string", minLength: 1 },
  },
  required: ["id", "label"],
  additionalProperties: false,
} as const;

const confirmationSchema = {
  type: "object",
  properties: {
    id: { type: "string", minLength: 1 },
    question: { type: "string", minLength: 1 },
    choices: { type: "array", items: confirmationChoiceSchema, minItems: 1 },
    required: { type: "boolean" },
  },
  required: ["id", "question", "choices", "required"],
  additionalProperties: false,
} as const satisfies JSONSchemaType<PlanConfirmation>;

const headerAliasesSchema = {
  type: "object",
  nullable: true,
  additionalProperties: {
    type: "array",
    items: { type: "string", minLength: 1 },
  },
  required: [],
} as const;

const artifactSchema = {
  type: "object",
  properties: {
    id: { type: "string", minLength: 1 },
    type: { type: "string", enum: ["table", "chart"] },
    sheet: { type: "string", minLength: 1 },
    anchor: { type: "string", minLength: 1 },
    fromStep: { type: "string", minLength: 1 },
    sheetName: { type: "string", minLength: 1, nullable: true },
    blockRef: { type: "string", minLength: 1, nullable: true },
    addressA1: { type: "string", minLength: 1, nullable: true },
    headers: { type: "array", nullable: true, items: { type: "string" } },
    rowCount: { type: "number", nullable: true },
    colCount: { type: "number", nullable: true },
    counts: {
      type: "object",
      nullable: true,
      properties: {},
      required: [],
      additionalProperties: { type: "number" },
    },
    headerAliases: headerAliasesSchema,
  },
  required: ["id", "type", "sheet", "anchor", "fromStep"],
  additionalProperties: false,
} as const satisfies JSONSchemaType<PlanArtifact>;

const joinKeySchema = {
  type: "object",
  properties: {
    left: { type: "string", minLength: 1 },
    right: { type: "string", minLength: 1 },
    strategy: { type: "string", enum: ["case_insensitive_trim", "numeric", "exact"], nullable: true },
  },
  required: ["left", "right"],
  additionalProperties: false,
} as const;

const anchorSchema = {
  type: "object",
  properties: {
    blockRef: { type: "string", minLength: 1, nullable: true },
    artifactRef: { type: "string", minLength: 1, nullable: true },
    sheet: { type: "string", minLength: 1, nullable: true },
    cell: { type: "string", minLength: 1, nullable: true },
  },
  additionalProperties: true,
} as const;

const sourceRefSchema = {
  type: "object",
  properties: {
    blockRef: { type: "string", minLength: 1, nullable: true },
    artifactRef: { type: "string", minLength: 1, nullable: true },
  },
  required: [],
  additionalProperties: false,
} as const;

const tableViewParamsSchema = {
  type: "object",
  properties: {
    source: { ...sourceRefSchema },
    select: { type: "array", nullable: true, items: { type: "string", minLength: 1 } },
    rename: { type: "object", nullable: true, additionalProperties: { type: "string" } },
    filter: {
      type: "array",
      nullable: true,
      items: {
        type: "object",
        properties: {
          col: { type: "string", minLength: 1 },
          op: { type: "string", enum: ["gt", "gte", "lt", "lte", "eq", "neq", "contains", "not_contains", "in", "isEmpty", "notEmpty", "between"] },
          value: { type: ["string", "number", "boolean", "object", "array", "null"], nullable: true },
          type: { type: "string", enum: ["number", "text", "date"], nullable: true },
        },
        required: ["col", "op"],
        additionalProperties: false,
      },
    },
    sort: {
      type: "object",
      nullable: true,
      properties: {
        col: { type: "string", minLength: 1 },
        dir: { type: "string", enum: ["asc", "desc"] },
      },
      required: ["col", "dir"],
      additionalProperties: false,
    },
    dest: {
      type: "object",
      properties: {
        mode: { type: "string", enum: ["right", "below", "newSheet", "inPlace"] },
        anchor: { ...anchorSchema, nullable: true },
        sheetName: { type: "string", minLength: 1, nullable: true },
        sheetNameHint: { type: "string", minLength: 1, nullable: true },
        sheet: { type: "string", minLength: 1, nullable: true },
        cell: { type: "string", minLength: 1, nullable: true },
      },
      required: ["mode"],
      additionalProperties: true,
    },
    outputTableName: { type: "string", minLength: 1, nullable: true },
    options: {
      type: "object",
      nullable: true,
      properties: {
        styleAsTable: { type: "boolean", nullable: true },
        freezeHeader: { type: "boolean", nullable: true },
      },
      additionalProperties: false,
    },
  },
  required: ["source", "dest"],
  additionalProperties: false,
  allOf: [
    {
      if: { type: "object", properties: { dest: { type: "object", properties: { mode: { const: "inPlace" } } } } },
      then: { properties: { select: { type: "array", maxItems: 0 } } },
    },
    {
      if: { type: "object", properties: { dest: { type: "object", properties: { mode: { enum: ["newSheet", "right", "below"] } } } } },
      then: { required: ["select"], properties: { select: { type: "array", minItems: 1 } } },
    },
  ],
} as const;

const joinParamsSchema = {
  type: "object",
  properties: {
    left: {
      type: "object",
      properties: { blockRef: { type: "string", minLength: 1 } },
      required: ["blockRef"],
      additionalProperties: true,
    },
    right: {
      type: "object",
      properties: { blockRef: { type: "string", minLength: 1 } },
      required: ["blockRef"],
      additionalProperties: true,
    },
    keys: { type: "array", items: joinKeySchema, minItems: 1 },
    joinType: { type: "string", enum: ["left", "inner", "anti_left", "full"], nullable: true },
    allowKeyFallback: { type: "boolean", nullable: true },
    keepRightKeyColumns: { type: "boolean", nullable: true },
    selectionPolicy: { type: "string", enum: ["defaultAll", "explicit"], nullable: true },
    select: { type: "object", nullable: true, additionalProperties: true },
    conflict: { type: "object", nullable: true, additionalProperties: true },
    output: {
      type: "object",
      nullable: true,
      properties: {
        mode: { type: "string", enum: ["right", "below", "newSheet"], nullable: true },
        sheetName: { type: "string", minLength: 1, nullable: true },
        anchor: { ...anchorSchema, nullable: true },
        tableName: { type: "string", minLength: 1, nullable: true },
      },
      additionalProperties: true,
    },
    match: { type: "object", nullable: true, additionalProperties: true },
  },
  required: ["left", "right", "keys"],
  additionalProperties: true,
} as const;

const createChartParamsSchema = {
  type: "object",
  properties: {
    source: {
      type: "object",
      properties: {
        blockRef: { type: "string", minLength: 1, nullable: true },
        tableName: { type: "string", minLength: 1, nullable: true },
        sheetName: { type: "string", minLength: 1, nullable: true },
        artifactRef: { type: "string", minLength: 1, nullable: true },
      },
      required: [],
      additionalProperties: false,
    },
    mapping: {
      type: "object",
      properties: {
        xCol: {
          type: "object",
          properties: {
            colIndex: { type: "number", nullable: true },
            headerName: { type: "string", nullable: true },
            header: { type: "string", nullable: true },
          },
          required: [],
          additionalProperties: false,
        },
        yCols: {
          type: "array",
          items: {
            type: "object",
            properties: {
              colIndex: { type: "number", nullable: true },
              headerName: { type: "string", nullable: true },
              header: { type: "string", nullable: true },
            },
            required: [],
            additionalProperties: false,
          },
        },
      },
      required: ["xCol", "yCols"],
      additionalProperties: false,
    },
    chartType: { type: "string", enum: ["columnClustered", "line", "barClustered"], nullable: true },
    titleHint: { type: "string", nullable: true },
    dest: {
      type: "object",
      properties: {
        mode: { type: "string", enum: ["right", "below", "newSheet"] },
        anchor: {
          type: "object",
          properties: { blockRef: { type: "string", nullable: true }, artifactRef: { type: "string", minLength: 1, nullable: true } },
          required: [],
          additionalProperties: false,
          nullable: true,
        },
        sheet: { type: "string", nullable: true },
        sheetName: { type: "string", nullable: true },
        sheetNameHint: { type: "string", nullable: true },
        titleHint: { type: "string", nullable: true },
        newSheetNameHint: { type: "string", nullable: true },
      },
      required: ["mode"],
      additionalProperties: false,
    },
  },
  required: ["source", "mapping", "dest"],
  additionalProperties: false,
} as const;

const validateDataSourceSchema = {
  type: "object",
  properties: {
    blockRef: { type: "string", minLength: 1 },
    artifactRef: { type: "string", minLength: 1 },
  },
  additionalProperties: false,
  oneOf: [
    { required: ["blockRef"], not: { required: ["artifactRef"] } },
    { required: ["artifactRef"], not: { required: ["blockRef"] } },
  ],
} as const;

const validateDetectSchema = {
  type: "object",
  properties: {
    missing: { type: "boolean", nullable: true },
    duplicates: { type: "boolean", nullable: true },
    badType: { type: "boolean", nullable: true },
  },
  additionalProperties: false,
} as const;

const validateTypeRuleSchema = {
  type: "object",
  properties: {
    col: { type: "string", minLength: 1 },
    type: { type: "string", enum: ["date", "number", "text"] },
  },
  required: ["col", "type"],
  additionalProperties: false,
} as const;

const validateOptionsSchema = {
  type: "object",
  properties: {
    maxIssues: { type: "integer", minimum: 1, nullable: true },
  },
  additionalProperties: false,
} as const;

const validateInternalDecisionsSchema = {
  type: "object",
  properties: {
    fixMissing: { type: "boolean", nullable: true },
    fixDuplicates: { type: "boolean", nullable: true },
    fixBadType: { type: "boolean", nullable: true },
  },
  required: [],
  additionalProperties: false,
} as const;

const validateInternalSchema = {
  type: "object",
  properties: {
    decisions: { ...validateInternalDecisionsSchema, nullable: true },
    executions: { type: "integer", minimum: 0, nullable: true },
    phase: { type: "string", enum: ["confirmations"], nullable: true },
  },
  additionalProperties: false,
} as const;

const validateDataParamsSchema = {
  type: "object",
  properties: {
    source: validateDataSourceSchema,
    detect: validateDetectSchema,
    missingColumns: { type: "array", nullable: true, items: { type: "string", minLength: 1 } },
    duplicateKeyColumns: { type: "array", nullable: true, items: { type: "string", minLength: 1 } },
    typeRules: { type: "array", nullable: true, items: validateTypeRuleSchema },
    options: validateOptionsSchema,
    __internal: validateInternalSchema,
  },
  required: ["source"],
  additionalProperties: false,
} as const;

const stepSchema: any = {
  type: "object",
  properties: {
    id: { type: "string", minLength: 1 },
    macro: { type: "string", enum: macroEnum },
    params: { type: "object", default: {}, additionalProperties: true },
  },
  required: ["id", "macro", "params"],
  additionalProperties: false,
  allOf: [
    {
      if: { properties: { macro: { const: "join_tables" } } },
      then: { properties: { params: joinParamsSchema } },
    },
    {
      if: { properties: { macro: { const: "create_chart" } } },
      then: { properties: { params: createChartParamsSchema } },
    },
    {
      if: { properties: { macro: { const: "table_view" } } },
      then: { properties: { params: tableViewParamsSchema } },
    },
    {
      if: { properties: { macro: { const: "validate_data" } } },
      then: { properties: { params: validateDataParamsSchema } },
    },
  ],
};

const planSchema: JSONSchemaType<AgentPlan> = {
  type: "object",
  properties: {
    version: { type: "string", const: "1.0" },
    goal: { type: "string", minLength: 1 },
    steps: { type: "array", items: stepSchema, minItems: 1 },
    confirmations: { type: "array", items: confirmationSchema, nullable: true },
    artifacts: { type: "array", items: artifactSchema, nullable: true },
  },
  required: ["version", "goal", "steps"],
  additionalProperties: false,
};

const ajv = new Ajv({ allErrors: true, allowUnionTypes: true });
const compiledValidator = ajv.compile(planSchema);

export function validatePlan(plan: any): { valid: true } | { valid: false; errors: string[] } {
  const ok = compiledValidator(plan);
  if (ok) return { valid: true };
  const errors = (compiledValidator.errors || []).map((err) => ajv.errorsText([err], { dataVar: "plan" }));
  return { valid: false, errors };
}

export { planSchema };
