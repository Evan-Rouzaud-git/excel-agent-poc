import { WorkbookContextSnapshot } from "../../context/types";
import { sanitizeTableViewParams } from "../tableViewUtils";
import { normalizeHeader } from "../normalizeHeader";
const normalizeHeaderText = normalizeHeader;
import { AgentPlan, AgentMacroName } from "../types";
import { findBlock } from "../utils";
import { collectExtraHeaders } from "./extraHeaders";
import { VALIDATE_DATA_QUESTIONS, VALIDATE_DATA_CONFIRM_CHOICES } from "../validateDataFlow";

const planKeys = new Set(["version", "goal", "steps", "confirmations"]);

const macroParamsAllow: Record<AgentMacroName, Record<string, true>> = {
  place_output: {
    mode: true,
    anchor: true,
    avoidOverwrite: true,
    minBlankArea: true,
    newSheetNameHint: true,
  },
  write_formula: {
    target: true,
    formula: true,
    fillDown: true,
    numberFormat: true,
    ifOverwrite: true,
  },
  apply_format: {
    target: true,
    options: true,
  },
  create_chart: {
    source: true,
    mapping: true,
    chartType: true,
    dest: true,
    titleHint: true,
  },
  join_tables: {
    left: true,
    right: true,
    keys: true,
    joinType: true,
    keepRightKeyColumns: true,
    allowKeyFallback: true,
    selectionPolicy: true,
    select: true,
    conflict: true,
    output: true,
    match: true,
  },
  table_view: {
    source: true,
    select: true,
    rename: true,
    filter: true,
    sort: true,
    dest: true,
    outputTableName: true,
    options: true,
  },
  validate_data: {
    source: true,
    detect: true,
    missingColumns: true,
    duplicateKeyColumns: true,
    typeRules: true,
    options: true,
  },
  summarize_actions: {
    params: true,
  },
};

const targetKeys = new Set(["blockRef", "artifactRef", "writeMode", "headerName", "sheet", "rangeA1"]);
const anchorKeys = new Set(["blockRef", "artifactRef", "sheet", "cell"]);
const mappingKeys = new Set(["xCol", "yCols"]);
const colKeys = new Set(["colIndex", "headerName", "header"]);
const optionsKeys = new Set(["header", "bandedRows", "numberFormats", "columnWidth", "freezeHeaderRow", "preset"]);
const headerKeys = new Set(["bold", "background", "fontColor"]);
const minBlankKeys = new Set(["rows", "cols"]);
const destKeys = new Set(["mode", "anchor", "sheet", "sheetName", "sheetNameHint", "titleHint", "newSheetNameHint"]);
const confirmationKeys = new Set(["id", "question", "choices", "required"]);
const confirmationChoiceKeys = new Set(["id", "label"]);
const artifactKeys = new Set(["id", "type", "sheet", "anchor", "fromStep"]);
const validateDetectKeys = new Set(["missing", "duplicates", "badType"]);
const validateOptionsKeys = new Set(["maxIssues"]);

function stripObj(obj: any, allowed: Set<string>) {
  if (!obj || typeof obj !== "object") return obj;
  const out: any = Array.isArray(obj) ? [] : {};
  Object.keys(obj).forEach((k) => {
    if (allowed.has(k)) out[k] = obj[k];
  });
  return out;
}

const TABLE_VIEW_ALLOW_TOKENS = ["$lastAddedColumn"];

type JoinColumnAlias = {
  left: string;
  right: string;
  leftNorm: string;
  rightNorm: string;
};

type JoinColumnAliasMap = Record<string, JoinColumnAlias[]>;

function buildJoinColumnAliasMap(steps: any[]): JoinColumnAliasMap {
  const map: JoinColumnAliasMap = {};
  if (!Array.isArray(steps)) return map;
  steps.forEach((step) => {
    if (!step || step.macro !== "join_tables") return;
    const artifactRef = step.id;
    if (!artifactRef) return;
    const keys = Array.isArray(step.params?.keys) ? step.params.keys : [];
    keys.forEach((key: any) => {
      const left = typeof key?.left === "string" ? key.left.trim() : "";
      const right = typeof key?.right === "string" ? key.right.trim() : "";
      if (!left || !right) return;
      const leftNorm = normalizeHeaderText(left);
      const rightNorm = normalizeHeaderText(right);
      if (!leftNorm || !rightNorm) return;
      const list = map[artifactRef] || [];
      list.push({ left, right, leftNorm, rightNorm });
      map[artifactRef] = list;
    });
  });
  return map;
}

function buildArtifactHeaderAliasMap(aliasMap: JoinColumnAliasMap): Record<string, Record<string, string[]>> {
  const out: Record<string, Record<string, string[]>> = {};
  Object.entries(aliasMap).forEach(([artifactRef, aliases]) => {
    const group: Record<string, Set<string>> = {};
    aliases.forEach((entry) => {
      const labels = [entry.left, entry.right].filter((label) => typeof label === "string" && label.trim()).map((label) => label.trim());
      if (!labels.length) return;
      const norms = [
        { norm: entry.leftNorm, label: entry.left },
        { norm: entry.rightNorm, label: entry.right },
      ];
      norms.forEach(({ norm }) => {
        if (!norm) return;
        const set = group[norm] || new Set<string>();
        labels.forEach((label) => set.add(label));
        group[norm] = set;
      });
    });
    const record = Object.fromEntries(
      Object.entries(group).map(([norm, set]) => [norm, Array.from(set)])
    );
    if (Object.keys(record).length) out[artifactRef] = record;
  });
  return out;
}

function applyJoinAliasesToTableView(
  params: any,
  artifactRef: string | undefined,
  aliasMap: JoinColumnAliasMap,
  artifactHeaders: string[] | undefined,
  allowTokens: string[] = []
) {
  if (!artifactRef) return;
  const aliasPairs = aliasMap[artifactRef];
  if (!aliasPairs?.length) return;
  const headers = (artifactHeaders || []).filter((h) => typeof h === "string");
  if (!headers.length) return;
  const headerLookup = new Map<string, string>();
  headers.forEach((header) => {
    const normalized = normalizeHeaderText(header);
    if (normalized && !headerLookup.has(normalized)) headerLookup.set(normalized, header);
  });
  if (!headerLookup.size) return;

  const renameProvided = params.rename && typeof params.rename === "object" && !Array.isArray(params.rename);
  const renameMap: Record<string, string> = renameProvided ? { ...params.rename } : {};
  const renameKeys = new Set(Object.keys(renameMap));
  const tokenSet = new Set(
    (allowTokens || [])
      .map((token) => (typeof token === "string" ? token.trim() : ""))
      .filter(Boolean)
  );

  const addRename = (actual: string, requested: string) => {
    if (!actual || !requested || actual === requested) return;
    if (renameKeys.has(actual)) return;
    renameMap[actual] = requested;
    renameKeys.add(actual);
  };

  const remapValue = (raw: string) => {
    const trimmed = raw?.trim();
    if (!trimmed) return raw;
    if (tokenSet.has(trimmed)) return raw;
    const normalized = normalizeHeaderText(trimmed);
    if (headerLookup.has(normalized)) return headerLookup.get(normalized) as string;
    for (const alias of aliasPairs) {
      const targetNorm = alias.leftNorm === normalized ? alias.rightNorm : alias.rightNorm === normalized ? alias.leftNorm : "";
      if (!targetNorm) continue;
      const target = headerLookup.get(targetNorm);
      if (!target) continue;
      addRename(target, trimmed);
      return target;
    }
    return trimmed;
  };

  if (Array.isArray(params.select)) {
    params.select = params.select.map((col: any) => (typeof col === "string" ? remapValue(col) : col));
  }
  if (Array.isArray(params.filter)) {
    params.filter = params.filter.map((filterObj: any) => {
      if (!filterObj || typeof filterObj !== "object") return filterObj;
      const rawCol = (filterObj.col || filterObj.column || filterObj.field || "" + "").toString();
      if (!rawCol) return filterObj;
      const remapped = remapValue(rawCol);
      if (remapped && remapped !== rawCol) {
        filterObj.col = remapped;
        filterObj.column = remapped;
        filterObj.field = remapped;
      }
      return filterObj;
    });
  }
  if (params.sort && typeof params.sort === "object") {
    const rawSort = (params.sort.col || params.sort.column || params.sort.field || "" + "").toString();
    if (rawSort) {
      const remapped = remapValue(rawSort);
      params.sort.col = remapped;
      params.sort.column = remapped;
      params.sort.field = remapped;
    }
  }

  if (renameProvided || Object.keys(renameMap).length > 0) {
    params.rename = renameMap;
  } else {
    delete params.rename;
  }
}

type JoinKeyCandidate = {
  left: string;
  right: string;
};

function isJoinKeyPlaceholder(value: any): boolean {
  if (!value || typeof value !== "string") return false;
  const trimmed = value.trim();
  if (!trimmed) return false;
  const lower = trimmed.toLowerCase();
  if (trimmed.includes("?") || lower.includes("join_key") || trimmed === "???" || (trimmed.startsWith("<") && trimmed.endsWith(">"))) {
    return true;
  }
  return false;
}

function describeJoinSide(side: any): string {
  if (!side || typeof side !== "object") return "table";
  return (side.tableName || side.blockRef || side.artifactRef || "table").toString();
}

function findTableHeaders(context: WorkbookContextSnapshot | undefined, tableName?: string, sheetName?: string): string[] {
  if (!context || !tableName) return [];
  const target = tableName.trim().toLowerCase();
  if (!target) return [];
  const matches: any[] = [];
  (context.sheets || []).forEach((sheet) => {
    const sheetMatches = !sheetName || sheet.name === sheetName;
    if (!sheetMatches) return;
    (sheet.tables || []).forEach((table) => {
      if (!table || typeof table.name !== "string") return;
      if (table.name.trim().toLowerCase() === target) {
        matches.push(table);
      }
    });
  });
  if (matches.length) {
    return matches[0].headers || [];
  }
  // fallback: search across all sheets if sheetName restricted and no match yet
  const fallback = (context.sheets || [])
    .flatMap((sheet) => (sheet.tables || []).filter((table) => table && typeof table.name === "string"))
    .find((table) => (table.name || "").trim().toLowerCase() === target);
  return fallback?.headers || [];
}

function gatherJoinSideHeaders(
  side: any,
  context?: WorkbookContextSnapshot,
  extraHeaders?: Record<string, string[]>
): string[] {
  const seen = new Set<string>();
  const addValue = (val?: string) => {
    if (!val || typeof val !== "string") return;
    const trimmed = val.trim();
    if (trimmed) seen.add(trimmed);
  };
  const addList = (list?: string[]) => {
    if (!Array.isArray(list)) return;
    list.forEach((item) => addValue(item));
  };

  const fallbackName = (side?.tableName || side?.blockRef || side?.artifactRef || "").toString();
  const addExtra = (key?: string) => {
    if (!key) return;
    addList(extraHeaders?.[key]);
  };

  addExtra(side?.artifactRef);
  addExtra(side?.blockRef);

  if (context) {
    if (typeof side?.blockRef === "string") {
      const block = findBlock(side.blockRef, context).block;
      addList(block?.headers);
    }
    if (typeof side?.tableName === "string") {
      addList(findTableHeaders(context, side.tableName, side.sheetName));
    }
  }

  if (!seen.size) {
    addValue(fallbackName);
    if (context) {
      const aggregate = (context.sheets || []).flatMap((sheet) => (sheet.blocks || []).flatMap((block) => block.headers || []));
      addList(aggregate);
    }
  }

  return Array.from(seen);
}

type JoinKeyChoice = {
  left: string;
  right: string;
};

const HEURISTICS: Array<{
  left: (header: string) => boolean;
  right: (header: string) => boolean;
}> = [
  { left: () => true, right: () => true },
  { left: (h) => normalizeHeaderText(h).includes("id"), right: (h) => normalizeHeaderText(h).includes("id") },
  { left: (h) => normalizeHeaderText(h).includes("code"), right: (h) => normalizeHeaderText(h).includes("code") },
  {
    left: (h) => {
      const norm = normalizeHeaderText(h);
      return norm.includes("projet") || norm.includes("project");
    },
    right: (h) => {
      const norm = normalizeHeaderText(h);
      return norm.includes("projet") || norm.includes("project");
    },
  },
  {
    left: (h) => {
      const norm = normalizeHeaderText(h);
      return norm.includes("nom") || norm.includes("name");
    },
    right: (h) => {
      const norm = normalizeHeaderText(h);
      return norm.includes("nom") || norm.includes("name");
    },
  },
];

function buildJoinKeyChoices(leftHeaders: string[], rightHeaders: string[], prompt?: string, limit = 6): JoinKeyChoice[] {
  const normalizedPrompt = normalizeHeaderText(prompt);
  const choices: JoinKeyChoice[] = [];
  const seen = new Set<string>();
  const addChoice = (left: string, right: string) => {
    if (!left || !right) return;
    const key = `${left}|${right}`;
    if (seen.has(key)) return;
    seen.add(key);
    if (choices.length < limit) {
      choices.push({ left, right });
    }
  };
  const promptMatches = (headers: string[]) => {
    if (!normalizedPrompt) return [];
    return headers.filter((header) => {
      const norm = normalizeHeaderText(header);
      return norm && normalizedPrompt.includes(norm);
    });
  };
  const leftPrompt = promptMatches(leftHeaders);
  const rightPrompt = promptMatches(rightHeaders);
  if (leftPrompt.length && rightPrompt.length) {
    leftPrompt.some((left) =>
      rightPrompt.some((right) => {
        if (choices.length >= limit) return true;
        addChoice(left, right);
        return false;
      })
    );
  }
  for (const heuristic of HEURISTICS) {
    if (choices.length >= limit) break;
    const leftMatch = leftHeaders.find((header) => heuristic.left(header));
    const rightMatch = rightHeaders.find((header) => heuristic.right(header));
    if (leftMatch && rightMatch) {
      addChoice(leftMatch, rightMatch);
    }
  }
  if (!choices.length && leftHeaders.length && rightHeaders.length && leftHeaders[0] && rightHeaders[0]) {
    addChoice(leftHeaders[0], rightHeaders[0]);
  }
  return choices;
}

function handleJoinKeyPlaceholders(
  step: any,
  context: WorkbookContextSnapshot | undefined,
  extraHeaders: Record<string, string[]> | undefined,
  warnings: string[],
  joinConfirmations: any[],
  existingConfirmationIds: Set<string>,
  planPrompt?: string
) {
  if (!step || step.macro !== "join_tables") return;
  const params = step.params || {};
  const keys = Array.isArray(params.keys) ? params.keys : [];
  const hasPlaceholder = keys.some((k: any) => isJoinKeyPlaceholder(k?.left) || isJoinKeyPlaceholder(k?.right));

  const leftHeaders = gatherJoinSideHeaders(params.left, context, extraHeaders);
  const rightHeaders = gatherJoinSideHeaders(params.right, context, extraHeaders);
  const normalizedLeft = new Set(leftHeaders.map((header) => normalizeHeaderText(header)).filter(Boolean));
  const normalizedRight = new Set(rightHeaders.map((header) => normalizeHeaderText(header)).filter(Boolean));
  const hasValidKeys =
    keys.length > 0 &&
    keys.every((k: any) => {
      const leftNorm = normalizeHeaderText(k?.left);
      const rightNorm = normalizeHeaderText(k?.right);
      return (
        leftNorm &&
        rightNorm &&
        normalizedLeft.has(leftNorm) &&
        normalizedRight.has(rightNorm)
      );
    });

  const needsConfirmation = hasPlaceholder || !hasValidKeys;
  if (!needsConfirmation) return;

  const choices = buildJoinKeyChoices(leftHeaders, rightHeaders, planPrompt);
  if (!choices.length) return;

  const existingStrategy = keys[0]?.strategy;
  const fallbackChoice = choices[0];
  if (!fallbackChoice?.left || !fallbackChoice?.right) return;
  params.keys = [
    {
      left: fallbackChoice.left,
      right: fallbackChoice.right,
      strategy: existingStrategy || "case_insensitive_trim",
    },
  ];
  warnings.push(`join_keys_placeholder_resolved:${step.id}`);

  const confirmId = `joinKey:${step.id}`;
  if (existingConfirmationIds.has(confirmId)) return;
  existingConfirmationIds.add(confirmId);
  const leftLabel = describeJoinSide(params.left);
  const rightLabel = describeJoinSide(params.right);
  const confirmChoices = choices.map((choice) => ({
    id: `${choice.left}|${choice.right}`,
    label: `${choice.left} ↔ ${choice.right}`,
  }));
  joinConfirmations.push({
    id: confirmId,
    question: `Choisir la clé de jointure pour ${leftLabel} ↔ ${rightLabel}`,
    choices: confirmChoices,
    required: true,
  });
}

function repairConfirmationChoices(rawChoices: any[]): any[] {
  if (!Array.isArray(rawChoices)) return [];
  const seenIds = new Set<string>();
  const repaired: any[] = [];
  rawChoices.forEach((choice, idx) => {
    let candidate: any;
    if (typeof choice === "string") {
      candidate = { id: `c${idx + 1}`, label: choice };
    } else if (choice && typeof choice === "object") {
      candidate = stripObj(choice, confirmationChoiceKeys);
    } else {
      candidate = { id: `c${idx + 1}`, label: `choice${idx + 1}` };
    }
    const id = (candidate.id || "").toString().trim() || `c${idx + 1}`;
    const label = (candidate.label || id).toString().trim() || id;
    if (seenIds.has(id)) return;
    seenIds.add(id);
    repaired.push({ id, label });
  });
  return repaired;
}

export function sanitizePlan(
  plan: any,
  context?: WorkbookContextSnapshot,
  warningsOut?: string[],
  extraHeaders?: Record<string, string[]>,
  userPrompt?: string
) {
  if (!plan || typeof plan !== "object") return plan;
  const warnings = warningsOut || [];
  const planExtraHeaders = extraHeaders || collectExtraHeaders(plan);
  const planPrompt = typeof userPrompt === "string" ? userPrompt : typeof plan?.goal === "string" ? plan?.goal : "";
  const joinAliasMap = buildJoinColumnAliasMap(plan.steps || []);
  const artifactHeaderAliases = buildArtifactHeaderAliasMap(joinAliasMap);
  const existingConfirmationIds = new Set<string>();
  if (Array.isArray(plan.confirmations)) {
    plan.confirmations.forEach((c: any) => {
      if (c && typeof c.id === "string" && c.id.trim()) existingConfirmationIds.add(c.id);
    });
  }
  const joinConfirmations: any[] = [];
  const validateDataConfirmations: any[] = [];
  const clean: any = {};
  Object.keys(plan).forEach((k) => {
    if (planKeys.has(k)) clean[k] = plan[k];
  });
  delete clean.confirmations;
  delete clean.confirmations;
  clean.steps = (clean.steps || []).map((step: any) => {
    const base: any = { id: step?.id, macro: step?.macro, params: step?.params || {} };
    const allow = macroParamsAllow[step?.macro as AgentMacroName] || {};
    // migrate misplaced fields for write_formula into params
    if (step?.macro === "write_formula") {
      base.params = base.params || {};
      if (!base.params.target && step.target) base.params.target = step.target;
      if (!base.params.formula && step.formula) base.params.formula = step.formula;
      if (typeof step.fillDown !== "undefined" && typeof base.params.fillDown === "undefined") base.params.fillDown = step.fillDown;
      if (step.numberFormat && !base.params.numberFormat) base.params.numberFormat = step.numberFormat;
      if (step.ifOverwrite && !base.params.ifOverwrite) base.params.ifOverwrite = step.ifOverwrite;
    }
    const beforeKeys = Object.keys(base.params || {});
    base.params = Object.keys(base.params || {}).reduce((acc: any, key: string) => {
      if (allow[key]) acc[key] = base.params[key];
      return acc;
    }, {});
    const afterKeys = Object.keys(base.params || {});
    const dropped = beforeKeys.filter((k) => !afterKeys.includes(k));
    const shouldLog = (typeof process !== "undefined" && (process as any).env?.SANITIZE_DEBUG === "1") || (typeof window !== "undefined" && (window as any).SANITIZE_DEBUG);
    if (dropped.length) {
      const dropInfo = { step: base.id, macro: base.macro, keys: dropped };
      if (shouldLog) {
        // eslint-disable-next-line no-console
        console.debug?.("sanitize_dropped_keys", dropInfo);
      }
      const strictDrop = typeof process !== "undefined" && ((process as any).env?.SANITIZE_STRICT === "1" || (process as any).env?.NODE_ENV === "development");
      if (strictDrop) {
        warnings.push(`sanitize_dropped_keys ${JSON.stringify(dropInfo)}`);
  }
}

    if (step?.macro === "create_chart" && !base.params.dest) {
      base.params.dest = { mode: "right" };
    }
    if (step?.macro === "table_view") {
      const artifactRef = base.params?.source?.artifactRef;
      const artifactHeaders = artifactRef ? planExtraHeaders?.[artifactRef] : undefined;
      const headerAliases = artifactRef ? artifactHeaderAliases[artifactRef] : undefined;
      applyJoinAliasesToTableView(base.params, artifactRef, joinAliasMap, artifactHeaders, TABLE_VIEW_ALLOW_TOKENS);
      const headers =
        [
          artifactRef && planExtraHeaders?.[artifactRef],
          step?.params?.source?.blockRef && planExtraHeaders?.[step.params.source.blockRef],
          planExtraHeaders?.[step?.id || ""],
        ]
          .filter(Boolean)
          .flat() || [];
      base.params = sanitizeTableViewParams(base.params, context, warnings, {
        allowTokens: ["$lastAddedColumn"],
        extraHeaders: headers as string[],
        headerAliases,
      });
    }
    // nested strips
    if (base.params.anchor) base.params.anchor = stripObj(base.params.anchor, anchorKeys);
    if (base.params.dest) {
      base.params.dest = stripObj(base.params.dest, destKeys);
      if (base.params.dest?.anchor) base.params.dest.anchor = stripObj(base.params.dest.anchor, anchorKeys);
    }
    if (base.params.minBlankArea) base.params.minBlankArea = stripObj(base.params.minBlankArea, minBlankKeys);
    if (base.params.source) base.params.source = stripObj(base.params.source, new Set(["blockRef", "artifactRef", "tableName", "sheetName"]));
    if (base.params.target) base.params.target = stripObj(base.params.target, targetKeys);
    if (base.params.mapping) {
      const m = stripObj(base.params.mapping, mappingKeys);
      if (m.xCol) m.xCol = stripObj(m.xCol, colKeys);
      if (Array.isArray(m.yCols)) m.yCols = m.yCols.map((c: any) => stripObj(c, colKeys));
      base.params.mapping = m;
    }
    if (base.params.options && base.macro !== "validate_data") {
      const o = stripObj(base.params.options, optionsKeys);
      if (o.header) o.header = stripObj(o.header, headerKeys);
      if (Array.isArray(o.numberFormats)) {
        o.numberFormats = o.numberFormats.map((nf: any) => stripObj(nf, new Set(["headerHints", "format"])));
      }
      base.params.options = o;
    }
    if (base.macro === "validate_data") {
      if (base.params.source) {
        base.params.source = stripObj(base.params.source, new Set(["blockRef", "artifactRef"]));
      }
      if (base.params.detect) {
        base.params.detect = stripObj(base.params.detect, validateDetectKeys);
        if (!Object.keys(base.params.detect).length) delete base.params.detect;
      }
      if (base.params.options) {
        base.params.options = stripObj(base.params.options, validateOptionsKeys);
      }
      const sourceBlockRef = base.params.source?.blockRef;
      const blockInfo = sourceBlockRef && context ? findBlock(sourceBlockRef, context).block : undefined;
      if (sourceBlockRef && !blockInfo) {
        warnings.push(`validate_data_block_not_found:${sourceBlockRef}`);
        delete base.params.source.blockRef;
      }
      const headers = blockInfo?.headers || [];
      const headerLookup = new Map<string, string>();
      headers.forEach((hdr: string) => {
        const norm = normalizeHeaderText(hdr);
        if (norm && !headerLookup.has(norm)) headerLookup.set(norm, hdr);
      });
      const filterColumns = (values?: any[]): string[] | undefined => {
        if (!Array.isArray(values)) return undefined;
        const filtered: string[] = [];
        const seen = new Set<string>();
        values.forEach((col) => {
          if (typeof col !== "string") return;
          const norm = normalizeHeaderText(col);
          if (!norm) return;
          const actual = headerLookup.get(norm);
          if (!actual || seen.has(norm)) return;
          seen.add(norm);
          filtered.push(actual);
        });
        return filtered.length ? filtered : undefined;
      };
      const sanitizeColumnList = (
        field: "missingColumns" | "duplicateKeyColumns",
        warning: string,
        warningNoHeaders: string
      ) => {
        const raw = Array.isArray(base.params[field]) ? base.params[field] : [];
        if (!raw.length) {
          delete base.params[field];
          return;
        }
        if (!headers.length) {
          delete base.params[field];
          warnings.push(warningNoHeaders);
          return;
        }
        const filtered = filterColumns(raw);
        if (filtered) {
          base.params[field] = filtered;
          return;
        }
        delete base.params[field];
        warnings.push(warning);
      };
      sanitizeColumnList("missingColumns", "validate_data_missing_columns_cleared", "validate_data_missing_columns_cleared_no_headers");
      sanitizeColumnList("duplicateKeyColumns", "validate_data_duplicate_columns_cleared", "validate_data_duplicate_columns_cleared_no_headers");
      const sanitizeTypeRules = (rules: any[]): Array<{ col: string; type: string }> | undefined => {
        if (!rules || !rules.length) return undefined;
        if (!headers.length) {
          warnings.push("validate_data_type_rules_cleared_no_headers");
          return undefined;
        }
        const result: Array<{ col: string; type: string }> = [];
        const seen = new Set<string>();
        rules.forEach((rule) => {
          if (!rule || typeof rule !== "object") return;
          const rawCol = typeof rule.col === "string" ? rule.col : "";
          const norm = normalizeHeaderText(rawCol);
          if (!norm || seen.has(norm)) return;
          const actual = headerLookup.get(norm);
          if (!actual) return;
          const ruleType = typeof rule.type === "string" ? rule.type : "";
          if (!ruleType) return;
          if (!["date", "number", "text"].includes(ruleType)) return;
          seen.add(norm);
          result.push({ col: actual, type: ruleType });
        });
        if (!result.length) {
          warnings.push("validate_data_type_rules_cleared");
          return undefined;
        }
        return result;
      };
      if (Array.isArray(base.params.typeRules)) {
        const filtered = sanitizeTypeRules(base.params.typeRules);
        if (filtered) {
          base.params.typeRules = filtered;
        } else {
          delete base.params.typeRules;
        }
      }
      if (base.params.options?.maxIssues !== undefined && typeof base.params.options.maxIssues === "number") {
        base.params.options.maxIssues = Math.max(1, Math.min(5000, Math.floor(base.params.options.maxIssues)));
      }
      if (base.params.action) delete base.params.action;
      if (base.params.__internal) delete base.params.__internal;
      const detectFlags = {
        missing: base.params.detect?.missing ?? true,
        duplicates: base.params.detect?.duplicates ?? true,
        badType: base.params.detect?.badType ?? true,
      };
      VALIDATE_DATA_QUESTIONS.forEach((question) => {
        if (!detectFlags[question.detectKey]) return;
        if (existingConfirmationIds.has(question.id)) return;
        existingConfirmationIds.add(question.id);
        const choices = VALIDATE_DATA_CONFIRM_CHOICES.map((choice) => ({ ...choice }));
        validateDataConfirmations.push({
          id: question.id,
          question: question.question,
          choices,
          required: true,
        });
      });
    }
    if (base.macro === "join_tables") {
      handleJoinKeyPlaceholders(
        base,
        context,
        planExtraHeaders,
        warnings,
        joinConfirmations,
        existingConfirmationIds,
        planPrompt
      );
    }
    return base;
  });

  const baseConfirmations = Array.isArray(plan.confirmations)
    ? plan.confirmations
        .map((c: any) => {
          const cc = stripObj(c, confirmationKeys);
          if (Array.isArray(cc.choices)) {
            const repaired = repairConfirmationChoices(cc.choices);
            if (!repaired.length) return null;
            cc.choices = repaired;
          }
          if (!cc.choices || !Array.isArray(cc.choices) || !cc.choices.length) return null;
          return cc;
        })
        .filter(Boolean)
    : [];
  const combinedConfirmations = [...baseConfirmations, ...joinConfirmations, ...validateDataConfirmations];
  const normalizedConfirmations = combinedConfirmations
    .map((c: any) => {
      const cc = stripObj(c, confirmationKeys);
      if (Array.isArray(cc.choices)) {
        const repaired = repairConfirmationChoices(cc.choices);
        if (!repaired.length) return null;
        cc.choices = repaired;
      }
      if (!cc.choices || !Array.isArray(cc.choices) || !cc.choices.length) return null;
      return cc;
    })
    .filter(Boolean);
  if (normalizedConfirmations.length) {
    clean.confirmations = normalizedConfirmations;
  } else {
    delete clean.confirmations;
  }
  if (Array.isArray(clean.artifacts)) {
    clean.artifacts = clean.artifacts.map((a: any) => stripObj(a, artifactKeys));
  }
  delete clean.artifacts; // artifacts must not come from planner
  return clean;
}
