import { AgentPlan } from "../types";
import { sanitizeTableViewParams } from "../tableViewUtils";
import { normalizeHeader } from "../normalizeHeader";

type ContextLike = any;

const chartTypeMap: Record<string, string> = {
  column: "columnClustered",
  columnclustered: "columnClustered",
  bar: "barClustered",
  barclustered: "barClustered",
  line: "line",
};

const allowedChartTypes = ["columnClustered", "barClustered", "line"];

const normalizeHeaderText = normalizeHeader;

function isCodeLike(h: string) {
  const norm = normalizeHeaderText(h);
  return norm.includes("code") || norm.includes("id");
}
function isLabelLike(h: string) {
  const norm = normalizeHeaderText(h);
  return norm.includes("nom") || norm.includes("name") || norm.includes("label") || norm.includes("libell") || norm.includes("region");
}

function listBlocks(context: ContextLike) {
  const blocks: { id: string; headers: string[]; columnTypes?: string[]; sheetName?: string }[] = [];
  (context?.sheets || []).forEach((s: any) => {
    (s.blocks || []).forEach((b: any) => {
      blocks.push({ id: b.id, headers: b.headers || [], columnTypes: b.columnTypes || [], sheetName: s.name });
    });
  });
  return blocks;
}

function hoistWriteFormulaBeforeJoin(steps: any[]): any[] {
  if (!Array.isArray(steps) || steps.length === 0) return steps;
  const formulaByHeader = new Map<string, any>();
  steps.forEach((step) => {
    if (step?.macro === "write_formula") {
      const header = step.params?.target?.headerName;
      const norm = header ? normalizeHeaderText(header) : "";
      if (norm) formulaByHeader.set(norm, step);
    }
  });
  const inserted = new Set<any>();
  const result: any[] = [];
  steps.forEach((step) => {
    if (step?.macro === "join_tables") {
      const keys: any[] = Array.isArray(step.params?.keys) ? step.params.keys : [];
      const needed = keys
        .map((key) => [key?.left, key?.right])
        .flat()
        .map((name) => normalizeHeaderText(name))
        .filter(Boolean);
      needed.forEach((norm) => {
        const formula = formulaByHeader.get(norm);
        if (formula && !inserted.has(formula)) {
          result.push(formula);
          inserted.add(formula);
        }
      });
    }
    if (step?.macro === "write_formula") {
      if (inserted.has(step)) return;
      inserted.add(step);
      result.push(step);
      return;
    }
    result.push(step);
  });
  return result;
}

function fallbackBlockId(context: ContextLike, blocks: { id: string }[]): string | undefined {
  const first = blocks[0];
  return context?.active?.selectionInBlockId || context?.active?.nearestBlockId || (first ? first.id : undefined);
}

function isValidIdx(idx: any, colCount: number) {
  return Number.isInteger(idx) && idx >= 0 && idx < colCount;
}

const normKey = (ref?: string | null) => (ref || "").toString().toLowerCase();

export function normalizePlan(planInput: any, context: ContextLike, userPrompt?: string, warningsOut?: string[]): AgentPlan | any {
  if (!planInput || typeof planInput !== "object") return planInput;
  // strict root keys
  const plan: any = {
    version: planInput.version,
    goal: planInput.goal,
    steps: planInput.steps,
    confirmations: Array.isArray(planInput.confirmations) ? [...planInput.confirmations] : undefined,
  };
  const warnings = warningsOut || [];
  const blocks = listBlocks(context);
  const fbBlock = fallbackBlockId(context, blocks);
  const normalizeBlockRef = (ref: string | undefined) => {
    if (ref && blocks.some((b) => b.id === ref)) return ref;
    return fbBlock;
  };
  const findBlockInfo = (ref?: string) => blocks.find((b) => b.id === ref);
  const findBlockByTableName = (tableName?: string, sheetName?: string) => {
    if (!tableName) return undefined;
    const tn = tableName.toLowerCase();
    const candidates = blocks.filter((b) => {
      const srcName = (b as any)?.source?.tableName;
      const sheetOk = !sheetName || b.sheetName === sheetName || (context?.sheets || []).some((s: any) => s.name === sheetName && s.blocks?.some((bb: any) => bb.id === b.id));
      return srcName && srcName.toLowerCase() === tn && sheetOk;
    });
    return candidates[0];
  };
  const uniqueSheetName = (base: string) => {
    const existing = new Set<string>((context?.sheets || []).map((s: any) => (s.name || "").toLowerCase()));
    let suffix = 0;
    while (suffix < 50) {
      const candidate = suffix === 0 ? base : `${base}_${suffix + 1}`;
      if (!existing.has(candidate.toLowerCase())) return candidate;
      suffix += 1;
    }
    return `${base}_${Date.now()}`;
  };

  const rawSteps = plan.steps || [];
  const mergedSteps: any[] = [];
  for (let i = 0; i < rawSteps.length; i += 1) {
    const current = rawSteps[i];
    const next = rawSteps[i + 1];
    const canMergeTableView =
      current &&
      next &&
      current.macro === "table_view" &&
      next.macro === "table_view" &&
      !(next.params?.source && next.params.source.artifactRef === current.id) &&
      !(current.params?.dest && current.params.dest.anchor?.artifactRef) &&
      !(next.params?.dest && next.params.dest.anchor?.artifactRef);
    if (canMergeTableView) {
      const merged = { ...current, params: { ...(current.params || {}) } };
      // merge select
      const sel1 = merged.params.select;
      const sel2 = next.params?.select;
      merged.params.select = (Array.isArray(sel1) && sel1.length ? sel1 : sel2) || sel1 || sel2;
      // merge rename (second wins)
      const r1 = merged.params.rename && typeof merged.params.rename === "object" ? merged.params.rename : {};
      const r2 = next.params?.rename && typeof next.params.rename === "object" ? next.params.rename : {};
      merged.params.rename = { ...r1, ...r2 };
      // carry filter/sort if missing
      merged.params.filter = merged.params.filter || next.params?.filter;
      merged.params.sort = merged.params.sort || next.params?.sort;
      merged.params.dest = merged.params.dest || next.params?.dest;
      merged.params.outputTableName = merged.params.outputTableName || next.params?.outputTableName;
      mergedSteps.push(merged);
      i += 1; // skip next
    } else {
      mergedSteps.push(current);
    }
  }

  // Dataset lineage & headers registry (datasetRef centric)
  type DatasetMeta = { artifactRef?: string | null; blockRef?: string | null };
  const latestDatasetBySource: Record<string, string> = {};
  const headersByDataset: Record<string, string[]> = {};
  const datasetMeta: Record<string, DatasetMeta> = {};
  const createdHeadersByDataset: Record<string, string[]> = {};
  const createdAliasesByDataset: Record<string, { header: string; norm: string }[]> = {};
  const lastCreatedByDataset: Record<string, string | undefined> = {};

  const sourceKey = (ref?: { blockRef?: string | null; artifactRef?: string | null }) => {
    if (!ref) return "";
    if (ref.artifactRef) return `artifact:${normKey(ref.artifactRef)}`;
    if (ref.blockRef) return `block:${normKey(ref.blockRef)}`;
    return "";
  };

  const registerDataset = (datasetRef: string, headers: string[], keys: string[], meta?: DatasetMeta) => {
    const uniqueHeaders = Array.from(new Set(headers.filter(Boolean)));
    headersByDataset[datasetRef] = uniqueHeaders;
    datasetMeta[datasetRef] = { ...(datasetMeta[datasetRef] || {}), ...(meta || {}) };
    keys
      .filter((k) => k)
      .forEach((k) => {
        latestDatasetBySource[k] = datasetRef;
      });
  };

  // seed with existing blocks
  blocks.forEach((b) => {
    const ds = normKey(b.id);
    registerDataset(ds, b.headers || [], [`block:${ds}`], { blockRef: b.id });
  });

  const headersForDataset = (datasetRef?: string) => (datasetRef ? headersByDataset[datasetRef] || [] : []);

  const ensureDatasetForSource = (refObj: any, metaHint?: DatasetMeta) => {
    const key = sourceKey(refObj) || (refObj?.blockRef ? `block:${normKey(refObj.blockRef)}` : "");
    const fallbackHeaders =
      (refObj?.blockRef && findBlockInfo(refObj.blockRef)?.headers) || (fbBlock && findBlockInfo(fbBlock)?.headers) || [];
    let ds = key ? latestDatasetBySource[key] : undefined;
    if (!ds) {
      ds = normKey(refObj?.artifactRef || refObj?.blockRef || `ds_${Object.keys(headersByDataset).length}`);
      registerDataset(ds, fallbackHeaders, key ? [key] : [], metaHint);
    }
    return { datasetRef: ds, headers: headersForDataset(ds), key };
  };

  const resolveSourceRef = (ref: any): any => {
    if (!ref) return { blockRef: fbBlock };
    if (typeof ref === "string") return resolveSourceRef({ blockRef: ref });
    const obj = { ...ref };
    const dsInfo = ensureDatasetForSource(obj, { artifactRef: obj.artifactRef, blockRef: obj.blockRef || fbBlock });
    const meta = datasetMeta[dsInfo.datasetRef] || {};
    if (!obj.artifactRef && meta.artifactRef) obj.artifactRef = meta.artifactRef;
    if (!obj.blockRef) {
      if (meta.blockRef) obj.blockRef = meta.blockRef;
      else if (!obj.artifactRef) obj.blockRef = fbBlock;
    }
    return obj;
  };

  const normSteps = mergedSteps.map((step: any) => {
    if (!step || typeof step !== "object") return step;
    if (step.macro === "table_view") {
      // translation layer: map requested output headers back to source headers when rename aliases exist
      const pRaw = { ...(step.params || {}) };
      // tolerate legacy/blockRef on params or string source
      const srcRaw = pRaw.source;
      const srcObj =
        typeof srcRaw === "string"
          ? { blockRef: srcRaw }
          : srcRaw && typeof srcRaw === "object"
            ? { ...srcRaw }
            : {};
      if (typeof (pRaw as any).blockRef === "string") {
        srcObj.blockRef = srcObj.blockRef || (pRaw as any).blockRef;
        delete (pRaw as any).blockRef;
      }
      const resolvedSource = resolveSourceRef(srcObj);
      const sourceDataset = ensureDatasetForSource(resolvedSource, { artifactRef: resolvedSource.artifactRef, blockRef: resolvedSource.blockRef });
      const sourceHeaders = sourceDataset.headers;
      const createdHeaders = createdHeadersByDataset[sourceDataset.datasetRef] || [];
      const lastCreated = lastCreatedByDataset[sourceDataset.datasetRef];
      const createdAliases = createdAliasesByDataset[sourceDataset.datasetRef] || [];
      const p = { ...pRaw, source: resolvedSource };
      const headerExists = (h: string) => {
        const norm = normalizeHeaderText(h);
        return (
          sourceHeaders.some((x: string) => normalizeHeaderText(x) === norm) ||
          createdHeaders.some((x) => normalizeHeaderText(x) === norm) ||
          norm === normalizeHeaderText("$lastAddedColumn")
        );
      };
      const maybeRewriteToLastAdded = (h: string) => {
        if (h === "$lastAddedColumn") return h;
        if (headerExists(h)) return h;
        const norm = normalizeHeaderText(h);
        const aliasHit = createdAliases.slice().reverse().find((c) => c.norm === norm);
        if (aliasHit) return "$lastAddedColumn";
        return h;
      };
      if (Array.isArray(p.select)) p.select = p.select.map(maybeRewriteToLastAdded);
      if (p.sort?.col) p.sort.col = maybeRewriteToLastAdded(p.sort.col);
      if (Array.isArray(p.filter)) p.filter = p.filter.map((f: any) => (f && f.col ? { ...f, col: maybeRewriteToLastAdded(f.col) } : f));
      if (createdHeaders.length && Array.isArray(p.select)) {
        const hasCreated =
          p.select.some((h: string) => createdHeaders.some((c) => normalizeHeaderText(c) === normalizeHeaderText(h))) ||
          p.select.includes("$lastAddedColumn");
        if (!hasCreated) p.select = [...p.select, "$lastAddedColumn"];
      }
      const warnStart = warnings.length;
      let params = sanitizeTableViewParams(
        { ...p },
        context,
        warnings,
        { allowTokens: ["$lastAddedColumn"], extraHeaders: createdHeaders, sourceHeaders }
      );
      if (!params.dest) params.dest = { mode: "newSheet" };
      if (
        params.dest?.mode === "inPlace" &&
        ((Array.isArray(params.select) && params.select.length > 0) ||
          (Array.isArray(params.filter) && params.filter.length > 0) ||
          !!params.sort)
      ) {
        params.dest.mode = "newSheet";
      }
      const stepWarnings = warnings.slice(warnStart);
      if (!params.dest) params.dest = { mode: "newSheet" };
      if (!params.dest.mode) params.dest.mode = "newSheet";
      const hasSelect = Array.isArray(params.select) && params.select.length > 0;
      const hasFilter = Array.isArray(params.filter) && params.filter.length > 0;
      const hasSort = !!params.sort;
      const hasRename = params.rename && Object.keys(params.rename || {}).length > 0;
      const ensureSelect = () => {
        if (params.dest?.mode !== "inPlace" && (!params.select || params.select.length === 0)) {
          const fromFilters =
            Array.isArray(params.filter) && params.filter.length
              ? Array.from(new Set(params.filter.map((f: any) => f?.col).filter(Boolean)))
              : [];
        let chosen = fromFilters;
        if (!chosen.length && sourceHeaders.length) {
          const proj = sourceHeaders.find((h: string) => normalizeHeaderText(h) === normalizeHeaderText("Projet"));
          if (proj) {
            const normProj = normalizeHeaderText(proj);
            const rest = sourceHeaders.filter((h: string) => normalizeHeaderText(h) !== normProj);
            chosen = [proj, ...rest];
          } else {
            chosen = [...sourceHeaders];
          }
        }
          if (chosen.length) {
            params.select = chosen;
          }
        }
        if (params.dest?.mode === "inPlace" && Array.isArray(params.select) && params.select.length > 0) {
          // projection + inPlace incompatible -> move to newSheet to keep selection
          params.dest.mode = "newSheet";
        }
      };
      ensureSelect();

      if (createdHeaders.length) {
        const createdNorms = createdHeaders.map(normalizeHeaderText);
        if (params.dest?.mode !== "inPlace") {
          const selArr = Array.isArray(params.select) ? params.select : [];
          const selNorms = selArr.map(normalizeHeaderText);
          if (!selArr.includes("$lastAddedColumn") && !selNorms.some((n) => createdNorms.includes(n))) {
            params.select = [...selArr, "$lastAddedColumn"];
          }
        }
        if (params.sort?.col && createdNorms.includes(normalizeHeaderText(params.sort.col))) {
          params.sort.col = "$lastAddedColumn";
        }
        if (Array.isArray(params.filter)) {
          params.filter = params.filter.map((f: any) =>
            f && f.col && createdNorms.includes(normalizeHeaderText(f.col)) ? { ...f, col: "$lastAddedColumn" } : f
          );
        }
      }
      if (createdHeaders.length && (!params.sort || !params.sort.col)) {
        params.sort = { col: "$lastAddedColumn", dir: "desc" };
      }

      const selectEqualsSource =
        hasSelect &&
        sourceHeaders.length > 0 &&
        params.select.every((s: string) => sourceHeaders.some((h: string) => normalizeHeaderText(h) === normalizeHeaderText(s))) &&
        sourceHeaders.length === params.select.length;
      // keep select even in inPlace to surface explicit error at execution
      if (!params.dest.sheetName && params.dest.sheetNameHint) params.dest.sheetName = params.dest.sheetNameHint;
      if (!params.options) params.options = {};
      if (typeof params.options.styleAsTable === "undefined") params.options.styleAsTable = true;
      if (typeof params.options.freezeHeader === "undefined") params.options.freezeHeader = true;
      if (!params.outputTableName && step.id) params.outputTableName = `View_${step.id}`;
      // keep artifactRef if provided; normalize blockRef only when no artifactRef
      if (params.source && !params.source.artifactRef) params.source.blockRef = normalizeBlockRef(params.source.blockRef);
      if (params.dest?.anchor && typeof params.dest.anchor === "object") {
        const a = params.dest.anchor;
        if (!a.artifactRef) a.blockRef = normalizeBlockRef(a.blockRef);
      }
      if (!params.dest.anchor && params.dest.mode !== "newSheet") {
        params.dest.anchor = { blockRef: params.source?.blockRef || normalizeBlockRef(undefined) };
      }
      if (!params.dest.sheetName && params.dest.sheet) params.dest.sheetName = params.dest.sheet;
      if (params.sort && params.dest?.mode !== "inPlace") {
        const selectNorm = new Set((params.select || []).map((s: string) => normalizeHeaderText(s)));
        const sortNorm = normalizeHeaderText(params.sort.col);
        if (!selectNorm.has(sortNorm)) {
          params.select = [...(params.select || []), params.sort.col];
        }
      }
      // Ambiguity confirmations
      const hdrs = sourceHeaders;
      const dynamicHeaders = [
        ...(params.select || []),
        ...((Array.isArray(params.filter) ? params.filter.map((f: any) => f?.col) : []).filter(Boolean) as string[]),
        ...(params.sort?.col ? [params.sort.col] : []),
      ];
      const headersList = [...hdrs, ...createdHeaders, ...dynamicHeaders, ...(blocks[0]?.headers || [])]
        .filter(Boolean)
        .slice(0, 12);
      const confirmations: any[] = plan.confirmations ? [...plan.confirmations] : [];
      const addConfirmation = (id: string, question: string, headers: string[]) => {
        if (!headers || headers.length === 0) return;
        confirmations.push({
          id,
          question,
          choices: headers.slice(0, 5).map((h, idx) => ({ id: `c${idx}`, label: h })),
          required: true,
        });
      };
      const unknownHeaderWarnings = stepWarnings.filter((w) => w.startsWith("unknown_header:"));
      if (stepWarnings.includes("table_view_select_unresolved")) {
        addConfirmation(`${step.id}:select`, "Choisir les colonnes à conserver", headersList);
      }
      if (unknownHeaderWarnings.length) {
        const requested = unknownHeaderWarnings.map((w) => w.split(":")[1]).join(", ");
        try {
          console.warn?.(`unknown_header requested=${requested} headers=[${headersList.join(", ")}]`);
        } catch {
          // ignore logging failures
        }
        addConfirmation(`${step.id}:unknown_header`, "Colonne introuvable, choisir dans la liste", headersList);
      }
      if (stepWarnings.some((w) => w.startsWith("table_view_filter_col_unresolved"))) {
        addConfirmation(`${step.id}:filter_col`, "Choisir la colonne pour filtrer", headersList);
      }
      if (stepWarnings.includes("filter_op_ambiguous")) {
        confirmations.push({
          id: `${step.id}:filter_op`,
          question: "Choisir l'opérateur du filtre",
          choices: [
            { id: "gt", label: ">" },
            { id: "gte", label: ">=" },
            { id: "lt", label: "<" },
            { id: "lte", label: "<=" },
            { id: "eq", label: "=" },
            { id: "neq", label: "!=" },
          ],
          required: true,
        });
      }
      if (stepWarnings.includes("table_view_sort_col_unresolved")) {
        addConfirmation(`${step.id}:sort_col`, "Choisir la colonne de tri", headersList);
      }
      if (confirmations.length) plan.confirmations = confirmations;

      // lineage update: compute output headers
      const inputHeaders = hdrs;
      const outputHeaders = (() => {
        if (Array.isArray(params.select) && params.select.length) {
          return params.select.map((h: string) => (params.rename && params.rename[h] ? params.rename[h] : h));
        }
        return inputHeaders;
      })();
      const datasetRefOut = params.dest?.mode === "inPlace" ? sourceDataset.datasetRef : normKey(step.id || sourceDataset.datasetRef);
      const keysToMap: string[] = [];
      const srcKey = sourceKey(params.source);
      if (srcKey) keysToMap.push(srcKey);
      if (params.dest?.mode !== "inPlace") keysToMap.push(`artifact:${normKey(step.id || "")}`);
      registerDataset(datasetRefOut, outputHeaders, keysToMap, {
        artifactRef: params.dest?.mode !== "inPlace" ? step.id : datasetMeta[sourceDataset.datasetRef]?.artifactRef || params.source?.artifactRef,
        blockRef: params.dest?.mode === "inPlace" ? params.source?.blockRef || fbBlock : datasetMeta[sourceDataset.datasetRef]?.blockRef,
      });
      return { ...step, params };
    }
    if (step.macro === "validate_data") {
      const params = { ...(step.params || {}) };
      const resolvedSource = resolveSourceRef(params.source);
      params.source = resolvedSource;
      if (!params.detect) {
        params.detect = { missing: true, duplicates: true, badType: true };
      }
      params.options = params.options || {};
      if (typeof params.options.maxIssues === "undefined") params.options.maxIssues = 1000;
      delete params.dest;
      delete params.highlight;
      delete params.checks;
      delete params.action;
      delete params.__internal;
      return { ...step, params };
    }
    if (step.macro === "write_formula" && step.params?.target?.headerName) {
      const targetRaw = step.params.target || {};
      const resolvedTarget = resolveSourceRef(targetRaw);
      const targetDataset = ensureDatasetForSource(resolvedTarget, { artifactRef: resolvedTarget.artifactRef, blockRef: resolvedTarget.blockRef || fbBlock });
      const targetKey = targetDataset.key || sourceKey(resolvedTarget);
      const updatedHeaders = Array.from(new Set([...headersForDataset(targetDataset.datasetRef), step.params.target.headerName]));
      registerDataset(targetDataset.datasetRef, updatedHeaders, targetKey ? [targetKey] : [], {
        artifactRef: resolvedTarget.artifactRef,
        blockRef: resolvedTarget.blockRef || fbBlock,
      });
      createdHeadersByDataset[targetDataset.datasetRef] = [...(createdHeadersByDataset[targetDataset.datasetRef] || []), step.params.target.headerName];
      lastCreatedByDataset[targetDataset.datasetRef] = step.params.target.headerName;
      const norm = normalizeHeaderText(step.params.target.headerName);
      const aliasList = createdAliasesByDataset[targetDataset.datasetRef] || [];
      aliasList.push({ header: step.params.target.headerName, norm });
      createdAliasesByDataset[targetDataset.datasetRef] = aliasList;
      const hadExplicitBlockRef = typeof targetRaw.blockRef === "string" && targetRaw.blockRef.trim().length > 0;
      (resolvedTarget as any).__explicitBlockRefFromUser = hadExplicitBlockRef;
      if (hadExplicitBlockRef) {
        (resolvedTarget as any).__userBlockRefRaw = targetRaw.blockRef;
      }
      step = { ...step, params: { ...step.params, target: resolvedTarget } };
    }
    if (step.macro === "join_tables") {
      const params = { ...(step.params || {}) };
      params.left = params.left || {};
      params.right = params.right || {};
      params.selectionPolicy = params.selectionPolicy || "defaultAll";
      params.keys = Array.isArray(params.keys) ? params.keys.filter((k: any) => k && k.left && k.right).map((k: any) => ({ ...k })) : [];
      const leftFromTable = findBlockByTableName(params.left.tableName, params.left.sheetName);
      const rightFromTable = findBlockByTableName(params.right.tableName, params.right.sheetName);
      const leftRef = normalizeBlockRef(params.left.blockRef || leftFromTable?.id);
      const rightRefCandidate =
        params.right.blockRef && blocks.some((b) => b.id === params.right.blockRef)
          ? params.right.blockRef
          : rightFromTable?.id || blocks.find((b) => b.id !== leftRef)?.id || fbBlock;
      params.left.blockRef = leftRef;
      params.right.blockRef = rightRefCandidate;
      if (!params.left.tableName && leftFromTable?.id) params.left.tableName = (leftFromTable as any).source?.tableName;
      if (!params.right.tableName && rightFromTable?.id) params.right.tableName = (rightFromTable as any).source?.tableName;
      params.joinType = params.joinType || "left";
      params.allowKeyFallback = params.allowKeyFallback === true;

    const leftInfo = findBlockInfo(params.left.blockRef);
    const rightInfo = findBlockInfo(params.right.blockRef);
      const defaultStrategy = "case_insensitive_trim";
      const detectStrategy = (k: any) => {
        const lIdx = (leftInfo?.headers || []).findIndex((h: string) => normalizeHeaderText(h) === normalizeHeaderText(k.left));
        const rIdx = (rightInfo?.headers || []).findIndex((h: string) => normalizeHeaderText(h) === normalizeHeaderText(k.right));
        const lt = lIdx >= 0 ? leftInfo?.columnTypes?.[lIdx] : undefined;
        const rt = rIdx >= 0 ? rightInfo?.columnTypes?.[rIdx] : undefined;
        if (lt === "number" && rt === "number") return "numeric";
        return defaultStrategy;
      };
      params.keys = params.keys.map((k: any) => ({ ...k, strategy: k.strategy || detectStrategy(k) }));
      // heuristic to avoid code vs label mismatch
      const rightHeadersNorm = (rightInfo?.headers || []).map((h: string) => normalizeHeaderText(h));
      const leftHeadersNorm = (leftInfo?.headers || []).map((h: string) => normalizeHeaderText(h));
      const findCodePair = () => {
        const codeIdxLeft = leftHeadersNorm.findIndex((h) => h.includes("code") || h.includes("id"));
        const codeIdxRight = rightHeadersNorm.findIndex((h) => h.includes("code") || h.includes("id"));
        if (codeIdxLeft >= 0 && codeIdxRight >= 0) {
          return { left: leftInfo?.headers?.[codeIdxLeft], right: rightInfo?.headers?.[codeIdxRight] };
        }
        return null;
      };
      params.keys = params.keys.map((k: any) => {
        const lNorm = normalizeHeaderText(k.left);
        const rNorm = normalizeHeaderText(k.right);
        const mismatch = isCodeLike(k.left) && isLabelLike(k.right);
        if (mismatch) {
          const replacement = findCodePair();
          if (replacement) {
            warnings.push("join_keys_rewritten_code_match");
            return { ...k, left: replacement.left, right: replacement.right, strategy: detectStrategy(replacement) };
          }
        }
        return k;
      });

      // auto-swap keys if left/right appear inverted across tables
      params.keys = params.keys.map((k: any) => {
        const leftIdx = leftHeadersNorm.indexOf(normalizeHeaderText(k.left));
        const rightIdx = rightHeadersNorm.indexOf(normalizeHeaderText(k.right));
        const swapLeftIdx = leftHeadersNorm.indexOf(normalizeHeaderText(k.right));
        const swapRightIdx = rightHeadersNorm.indexOf(normalizeHeaderText(k.left));
        if ((leftIdx === -1 || rightIdx === -1) && swapLeftIdx >= 0 && swapRightIdx >= 0) {
          warnings.push("join_keys_auto_swapped");
          return { ...k, left: leftHeaders[swapLeftIdx], right: rightHeaders[swapRightIdx], strategy: detectStrategy({ left: leftHeaders[swapLeftIdx], right: rightHeaders[swapRightIdx] }) };
        }
        return k;
      });

      const leftHeaders = leftInfo?.headers || [];
      const rightHeaders = rightInfo?.headers || [];
      const leftNorms = new Set(leftHeaders.map((h: string) => normalizeHeaderText(h)));
      const rightKeyNorms = new Set(params.keys.map((k: any) => normalizeHeaderText(k.right)));
      const defaultRightCols = rightHeaders.filter((h: string) => {
        const norm = normalizeHeaderText(h);
        if (!params.keepRightKeyColumns && rightKeyNorms.has(norm) && leftNorms.has(norm)) return false;
        if (leftNorms.has(norm)) return false;
        return true;
      });
      const keepRightKeyColumns = params.keepRightKeyColumns === true;
      params.select = params.select || {};
      // legacy select.right as array or select.right string[]
      const legacyRightArray: any = (params as any)["select.right"];
      if (Array.isArray(legacyRightArray)) {
        params.select.right = { mode: "list", columns: legacyRightArray };
        delete (params as any)["select.right"];
      }
      if (Array.isArray(params.select.right)) {
        params.select.right = { mode: "list", columns: params.select.right as any };
      }
      if (params.select.right && !params.select.right.mode) {
        params.select.right = { ...params.select.right, mode: params.select.right.columns ? "list" : "all" };
      }
      if (params.select.left && Array.isArray(params.select.left as any)) {
        params.select.left = { mode: "list", columns: params.select.left as any };
      }
      const explicitSelect =
        params.selectionPolicy === "explicit" ||
        (params.select?.right && Array.isArray((params.select.right as any).columns)) ||
        (params.select?.left && Array.isArray((params.select.left as any).columns));
      const forceAll = !explicitSelect;
      if (forceAll) {
        params.select.left = { mode: "all", columns: leftHeaders };
        params.select.right = { mode: "all", columns: rightHeaders };
      } else {
        if (!params.select.left) params.select.left = { mode: "all", columns: leftHeaders };
        if (!params.select.right || (Array.isArray((params.select.right as any).columns) && (params.select.right as any).columns.length === 0)) {
          params.select.right = { mode: "all", columns: rightHeaders };
        }
      }
      const conflict = params.conflict || {};
      params.conflict = {
        onDuplicateRightColumns: conflict.onDuplicateRightColumns || "suffix",
        rightSuffix: conflict.rightSuffix || "_r",
        onMultipleMatches: conflict.onMultipleMatches || "explode_rows",
      };
      const output = params.output || {};
      if (output.name && !output.sheetName) output.sheetName = output.name;
      if (output.newSheetNameHint && !output.sheetName) output.sheetName = output.newSheetNameHint;
      output.mode = output.mode || "newSheet";
      if (output.mode === "newSheet") {
        output.sheetName = output.sheetName || uniqueSheetName("Join_Result");
        delete output.anchor;
      } else {
        output.anchor = output.anchor || { blockRef: params.left.blockRef || normalizeBlockRef(undefined) };
      }
      output.tableName = output.tableName || "Join_Result";
      params.output = output;
      params.match = params.match || { defaultStrategy };
      const joinHeaders = Array.from(
        new Set([
          ...(Array.isArray((params.select.left as any)?.columns) ? ((params.select.left as any).columns as string[]) : leftHeaders),
          ...(Array.isArray((params.select.right as any)?.columns) ? ((params.select.right as any).columns as string[]) : rightHeaders),
        ])
      );
      const joinDatasetRef = normKey(step.id || output.tableName || `join_${Date.now()}`);
      registerDataset(joinDatasetRef, joinHeaders, [`artifact:${normKey(step.id || "")}`], { artifactRef: step.id });
      return { ...step, params };
    }
    if (step.macro === "create_chart") {
      const params = { ...(step.params || {}) };
      params.source = params.source || {};
      params.mapping = params.mapping || {};
      params.dest = params.dest || { mode: "right" };
      if (!params.dest.sheetName && (params.dest as any).sheetNameHint) params.dest.sheetName = (params.dest as any).sheetNameHint;
      if (params.dest.sheetName && !params.source.sheetName) {
        params.source.sheetName = params.dest.sheetName;
      }
      if (params.dest.sheet && !params.dest.sheetName) params.dest.sheetName = params.dest.sheet;
      if (!params.dest.newSheetNameHint && params.dest.sheetName) params.dest.newSheetNameHint = params.dest.sheetName;
      if (!params.mapping.xCol) params.mapping.xCol = {};
      if (!Array.isArray(params.mapping.yCols)) params.mapping.yCols = [];
      params.mapping.yCols = params.mapping.yCols.map((c: any) => ({ ...c }));
      step = { ...step, params };
    }
    if (step.macro !== "create_chart") {
      if (step.macro === "apply_format") {
        const opts = step.params?.options || {};
        const promptLower = (userPrompt || "").toLowerCase();
        const wantsPresentation = promptLower.includes("presentation") || promptLower.includes("professionnel");
        if (opts.preset === "corporate_blue" || opts.header?.background === "lightGray" || wantsPresentation) {
          step.params = step.params || {};
          step.params.options = {
            ...opts,
            preset: "corporate_blue",
            header: { ...(opts.header || {}), background: "#0e2a80", fontColor: "#ffffff", bold: true },
            bandedRows: opts.bandedRows ?? true,
            columnWidth: opts.columnWidth ?? "auto",
          };
        }
      }
      if (step.params?.target?.blockRef) {
        step = { ...step, params: { ...step.params, target: { ...step.params.target, blockRef: normalizeBlockRef(step.params.target.blockRef) } } };
      }
      return step;
    }

    const params = { ...(step.params || {}) };

    // chartType
    const ctRaw = (params.chartType || "").toString();
    const ctMapped = chartTypeMap[ctRaw.toLowerCase()] || ctRaw;
    params.chartType = allowedChartTypes.includes(ctMapped) ? ctMapped : "columnClustered";

    // source block
    if (!params.source) params.source = {};
    if (!params.source.artifactRef) params.source.blockRef = normalizeBlockRef(params.source.blockRef);

    // mapping
    const blockInfo = blocks.find((b) => b.id === params.source.blockRef);
    const colCount = Math.max(1, blockInfo?.headers?.length || 1);
    const mapping = params.mapping || {};
    const hasHeaderX = !!mapping.xCol?.headerName;
    if (!hasHeaderX) {
      const providedX = mapping.xCol?.colIndex;
      let xIdx = isValidIdx(providedX, colCount) ? (providedX as number) : 0;
      if (!isValidIdx(xIdx, colCount)) {
        warnings.push("mapping_rewritten_out_of_range");
        xIdx = Math.min(Math.max(0, Number(providedX) || 0), colCount - 1);
      }
      mapping.xCol = { colIndex: xIdx };
    }

    const hasHeaderY = Array.isArray(mapping.yCols) && mapping.yCols.some((c: any) => !!c?.headerName);
    if (!hasHeaderY) {
      const providedX = mapping.xCol?.colIndex;
      const xIdx = isValidIdx(providedX, colCount) ? (providedX as number) : 0;
      const providedY = Array.isArray(mapping.yCols) ? mapping.yCols.map((c: any) => c?.colIndex).filter((n: any) => Number.isInteger(n)) : [];
      let yIdxs = providedY.filter((n: number) => isValidIdx(n, colCount)).filter((n: number) => n !== xIdx);
      if (providedY.length !== yIdxs.length) warnings.push("mapping_rewritten_out_of_range");
      if (yIdxs.length === 0) {
        const fallbackY = isValidIdx(1, colCount) && xIdx !== 1 ? 1 : xIdx === 0 && colCount > 1 ? 1 : 0;
        yIdxs = [fallbackY];
        warnings.push("mapping_missing_ycols");
      }
      mapping.yCols = yIdxs.map((n: number) => ({ colIndex: n }));
    }
    params.mapping = mapping;
    // dest
    const dest = params.dest || {};
    const promptLower = (userPrompt || "").toLowerCase();
    const userWantsNewSheet =
      promptLower.includes("nouvelle feuille") || promptLower.includes("new sheet") || promptLower.includes("nouvel onglet") || promptLower.includes("new worksheet");
    if (dest.mode === "newSheet" && !userWantsNewSheet) dest.mode = "right";
    dest.mode = dest.mode || "right";
    const anchor = dest.anchor;
    let anchorBlockRef: string | undefined;
    let anchorArtifactRef: string | undefined;
    if (anchor && typeof anchor === "object") {
      anchorBlockRef = anchor.blockRef;
      anchorArtifactRef = (anchor as any).artifactRef;
    }
    if (typeof anchor === "string") {
      anchorBlockRef = anchor.toLowerCase() === "topleft" ? fbBlock : anchor;
    }
    dest.anchor = anchorArtifactRef ? { artifactRef: anchorArtifactRef, blockRef: anchorBlockRef } : { blockRef: normalizeBlockRef(anchorBlockRef) };
    params.dest = dest;

    return { ...step, params };
  });
  const normalizedSteps = hoistWriteFormulaBeforeJoin(normSteps);

  // Chain artifactRef to subsequent steps when missing
  let lastTableStepId: string | undefined;
  const tableViews = normalizedSteps.filter((s: any) => s?.macro === "table_view");

  const chainedSteps = normalizedSteps.map((s: any) => {
    if (s?.macro === "table_view") {
      lastTableStepId = s.id;
      return s;
    }
    if (lastTableStepId && s?.macro === "write_formula") {
      const tgt = s.params?.target || {};
      if (!tgt.artifactRef && !tgt.blockRef) {
        s = { ...s, params: { ...s.params, target: { ...tgt, artifactRef: lastTableStepId } } };
      }
    }
    if (lastTableStepId && (s?.macro === "apply_format" || s?.macro === "create_chart")) {
      const target = s.params?.target || s.params?.source || {};
      if (!target.artifactRef && !target.blockRef) {
        if (s.macro === "apply_format") s = { ...s, params: { ...s.params, target: { ...target, artifactRef: lastTableStepId } } };
        if (s.macro === "create_chart") s = { ...s, params: { ...s.params, source: { ...target, artifactRef: lastTableStepId } } };
      }
    }
    return s;
  });

  return { ...plan, steps: chainedSteps };
}
