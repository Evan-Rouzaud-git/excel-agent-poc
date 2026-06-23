import { WorkbookContextSnapshot } from "../context/types";
import { macros } from "./macros";
import { validatePlan } from "./planSchema";
import { validatePlanInvariants } from "./planInvariants";
import { canonicalizePlan } from "./canonicalizePlan";
import {
  AgentMacroName,
  AgentPlan,
  AgentLogEntry,
  ExecutionOptions,
  ExecutionResult,
  ExecutionStatus,
  MacroContext,
} from "./types";
import { nowIso, parseA1Address, rowColToA1 } from "./utils";
import { autoAnswerConfirmations } from "./autoConfirm";

function planHash(obj: any) {
  try {
    const str = JSON.stringify(obj);
    let h = 0;
    for (let i = 0; i < str.length; i += 1) h = (Math.imul(31, h) + str.charCodeAt(i)) | 0;
    return `h${(h >>> 0).toString(16)}`;
  } catch {
    return "h0000";
  }
}

export async function executePlan(
  plan: AgentPlan,
  contextSnapshot: WorkbookContextSnapshot,
  excelCtx: Excel.RequestContext,
  options: ExecutionOptions = {}
): Promise<ExecutionResult> {
  const contextState: WorkbookContextSnapshot = JSON.parse(JSON.stringify(contextSnapshot));
  const logs: AgentLogEntry[] = [];
  const artifacts: ExecutionResult["artifacts"] = [...(options.initialArtifacts || [])];
  const confirmationsRequested: ExecutionResult["confirmationsRequested"] = [];
  const decisions = options.confirmationDecisions || {};
  let errors: string[] = [];
  const warnings: string[] = [];
  const headerRegistry: { artifactRef?: string; blockRef?: string; datasetRef?: string; headerName: string; sourceStepId: string }[] = [];
  type DatasetMeta = {
    artifactRef?: string | null;
    blockRef?: string | null;
    headerAliases?: Record<string, string[]>;
  };
  const latestDatasetBySource: Record<string, string> = {};
  const headersByDataset: Record<string, string[]> = {};
  const datasetMeta: Record<string, DatasetMeta> = {};

  const sourceKey = (ref?: { blockRef?: string | null; artifactRef?: string | null }) => {
    if (!ref) return "";
    if (ref.artifactRef) return `artifact:${(ref.artifactRef || "").toLowerCase()}`;
    if (ref.blockRef) return `block:${(ref.blockRef || "").toLowerCase()}`;
    return "";
  };
  const registerDataset = (datasetRef: string, headers: string[], keys: string[], meta?: DatasetMeta) => {
    headersByDataset[datasetRef] = Array.from(new Set(headers.filter(Boolean)));
    datasetMeta[datasetRef] = { ...(datasetMeta[datasetRef] || {}), ...(meta || {}) };
    keys
      .filter(Boolean)
      .forEach((k) => {
        latestDatasetBySource[k] = datasetRef;
      });
  };
  const ensureDatasetForRef = (ref: any, metaHint?: DatasetMeta) => {
    const key = sourceKey(ref);
    const fallbackHeaders: string[] = [];
    let ds: string;
    if (key && latestDatasetBySource[key]) {
      ds = latestDatasetBySource[key];
    } else {
      ds = (ref?.artifactRef || ref?.blockRef || `ds_${Object.keys(headersByDataset).length}`).toString().toLowerCase();
      registerDataset(ds, fallbackHeaders, key ? [key] : [], metaHint);
    }
    return { datasetRef: ds, headers: headersByDataset[ds] || [], key };
  };

  const hasResolvableExplicitBlockRef = (blockRef?: string): boolean => {
    if (!blockRef || !contextState.sheets?.length) return false;
    const parsed = parseA1Address(blockRef);
    if (!parsed) return false;
    const sheetName = parsed.sheet || (blockRef.includes("!") ? blockRef.split("!")[0] : null);
    if (!sheetName) return false;
    return contextState.sheets.some((sheet) => sheet.name === sheetName);
  };

  // seed datasets from context
  (contextSnapshot?.sheets || []).forEach((s) => {
    (s.blocks || []).forEach((b: any) => {
      const ds = (b.id || "").toString().toLowerCase();
      registerDataset(ds, b.headers || [], [`block:${ds}`], { blockRef: b.id });
    });
  });

  const addLog = (entry: Omit<AgentLogEntry, "ts">) => {
    logs.push({
      ts: nowIso(),
      ...entry,
    });
  };

  plan = canonicalizePlan(plan);

  const validation = validatePlan(plan);
  if (!validation.valid) {
    validation.errors?.forEach((err) => addLog({ level: "error", message: err }));
    errors = logs.filter((l) => l.level === "error").map((l) => l.message);
    if (errors.length === 0) errors.push("plan_schema_invalid");
    return { logs, artifacts, status: "error", errors, ok: false, warnings };
  }

  const invariant = validatePlanInvariants(plan);
  const invariantIssues = invariant.valid ? [] : invariant.issues || [];
  const nonBlocking = new Set<string>(["place_output_fallback", "created_column_not_using_token"]);
  const blockingIssues = invariantIssues.filter((i) => !nonBlocking.has(i.code || ""));
  if (invariantIssues.length && blockingIssues.length === 0) {
    invariantIssues.forEach((issue) => addLog({ level: "warn", message: issue.message }));
  } else if (blockingIssues.length) {
    blockingIssues.forEach((issue) => addLog({ level: "error", message: issue.message }));
    errors = blockingIssues.map((i) => i.message);
    return { logs, artifacts, status: "error", errors, ok: false, warnings };
  }

  // auto-answer confirmations if allowed
  const autoMode = (options.autoAnswerMode || "interactive") as "demoEval" | "interactive" | "none";
  const auto = autoAnswerConfirmations(plan, contextSnapshot, autoMode);
  Object.assign(decisions, auto.decisions);
  plan = auto.plan;
  addLog({ level: "info", message: `PLAN_EXEC stage=validated hash=${planHash(plan)}`, stepId: "-", macro: undefined as any });

  let status: ExecutionStatus = "ok";
  let lastJoinArtifact: any = null;
  let lastJoinSources: { left?: string; right?: string } | null = null;
  let lastProducedTable: { sheetName: string; tableName?: string; blockRef?: string; address?: string } | null = null;
  let lastTableArtifact: any = (artifacts || []).slice().reverse().find((a) => a.type === "table") || null;
  let currentTableRef: { blockRef?: string; artifactRef?: string; headers?: string[]; tableName?: string } | null = lastTableArtifact
    ? { blockRef: lastTableArtifact.blockRef, artifactRef: lastTableArtifact.fromStep, headers: (lastTableArtifact as any).headers, tableName: (lastTableArtifact as any).tableName }
    : null;
  let lastAddedHeader: string | null = null;

  for (const step of plan.steps) {
    addLog({ level: "info", message: `Debut etape ${step.id}`, stepId: step.id, macro: step.macro });
    if (step.macro === "summarize_actions") {
      const summaryMessage = `Actions resumeees: logs=${logs.length}, artifacts=${artifacts.length}, stepsDone=${logs.filter((l) => l.stepId).length}`;
      addLog({ level: "info", message: summaryMessage, stepId: step.id, macro: step.macro as AgentMacroName });
      continue;
    }

    // Redirect formatting/writes to last join result if needed
    if (lastJoinArtifact && (step.macro === "apply_format" || step.macro === "write_formula")) {
      const target = (step.params && (step.params.target || step.params.source || {})) as any;
      const blockRef = target?.blockRef;
      const hasArtifactRef = !!target?.artifactRef;
      const sourceRefMatch =
        !blockRef || blockRef === lastJoinSources?.left || blockRef === lastJoinSources?.right || blockRef === contextSnapshot.active?.selectionInBlockId;
      if (!hasArtifactRef && sourceRefMatch) {
        const newTarget = {
          ...(step.params?.target || {}),
          blockRef: lastJoinArtifact.blockRef,
          tableName: lastJoinArtifact.tableName,
          sheetName: lastJoinArtifact.sheet || lastJoinArtifact.sheetName,
        };
        step.params = { ...step.params, target: newTarget };
        addLog({ level: "info", message: "target redirige vers table join", stepId: step.id, macro: step.macro as AgentMacroName });
      }
    }

    if (step.macro === "create_chart") {
      const src = (step.params as any)?.source || {};
      const noSource = !src.blockRef && !src.tableName;
      if (noSource && lastProducedTable) {
        const newSource: any = { ...src };
        if (lastProducedTable.tableName) newSource.tableName = lastProducedTable.tableName;
        if (lastProducedTable.blockRef) newSource.blockRef = lastProducedTable.blockRef;
        newSource.sheetName = lastProducedTable.sheetName;
        step.params = { ...(step.params || {}), source: newSource };
        const dest = (step.params as any).dest || {};
        if (!dest.sheetName) dest.sheetName = lastProducedTable.sheetName;
        step.params.dest = dest;
        addLog({ level: "info", message: `create_chart: defaultedToLastProducedTable ${lastProducedTable.tableName || lastProducedTable.blockRef}`, stepId: step.id, macro: step.macro as AgentMacroName });
      }
    }

    // Runtime fallback to current table when targets/sources are missing or stale
    const retargetToCurrent = (targetObj: any) => {
      if (!currentTableRef || !targetObj) return false;
      if (targetObj.artifactRef || targetObj.blockRef) return false;
      targetObj.blockRef = currentTableRef.blockRef;
      if (currentTableRef.artifactRef) targetObj.artifactRef = currentTableRef.artifactRef;
      addLog({ level: "warn", message: "retargeted_to_current_table", stepId: step.id, macro: step.macro as AgentMacroName });
      return true;
    };

    if (step.macro === "table_view") {
      // Do not retarget or fill select for table_view; fail explicitly if invalid
    }
    if (step.macro === "write_formula" || step.macro === "apply_format") {
      const tgt = (step.params as any)?.target || {};
      const changed = retargetToCurrent(tgt);
      if (changed) (step.params as any).target = tgt;
    }

    if (step.macro === "create_chart") {
      const src = (step.params as any)?.source || {};
      const changed = retargetToCurrent(src);
      if (changed) (step.params as any).source = src;
    }

    let datasetRefForStep: string | undefined;
    if (step.macro === "table_view") {
      const src = (step.params as any)?.source || {};
      const info = ensureDatasetForRef(src, { artifactRef: src.artifactRef, blockRef: src.blockRef });
      datasetRefForStep = info.datasetRef;
      const meta = datasetMeta[info.datasetRef] || {};
      const blockRef = src.artifactRef && meta.blockRef ? meta.blockRef : src.blockRef || meta.blockRef;
      (step.params as any).source = { ...src, artifactRef: src.artifactRef || meta.artifactRef, blockRef };
    }
    if (step.macro === "write_formula") {
      const tgt = (step.params as any)?.target || {};
      const info = ensureDatasetForRef(tgt, { artifactRef: tgt.artifactRef, blockRef: tgt.blockRef });
      datasetRefForStep = info.datasetRef;
      const meta = datasetMeta[info.datasetRef] || {};
      const shouldAdoptArtifact = !tgt.artifactRef && (!tgt.blockRef || (meta.blockRef && meta.blockRef === tgt.blockRef));
      const blockRefOverride = tgt.artifactRef && meta.blockRef && tgt.blockRef && tgt.blockRef !== meta.blockRef ? meta.blockRef : null;
      (step.params as any).target = {
        ...tgt,
        artifactRef: shouldAdoptArtifact ? meta.artifactRef || tgt.artifactRef : tgt.artifactRef,
        blockRef: blockRefOverride || tgt.blockRef || meta.blockRef,
      };
    }
    if (step.macro === "write_formula") {
      const tgt = (step.params as any)?.target || {};
      const userBlockRef = (tgt as any).__userBlockRefRaw;
      if (userBlockRef) {
        tgt.blockRef = userBlockRef;
        (step.params as any).target = { ...tgt };
      }
    }
    if (step.macro === "create_chart") {
      const src = (step.params as any)?.source || {};
      const info = ensureDatasetForRef(src, { artifactRef: src.artifactRef, blockRef: src.blockRef });
      datasetRefForStep = info.datasetRef;
    }

    const macro = macros[step.macro as AgentMacroName];
    if (!macro) {
      addLog({ level: "warn", message: `Macro inconnue: ${step.macro}`, stepId: step.id, macro: step.macro as AgentMacroName });
      continue;
    }

    if (step.macro === "write_formula" && lastTableArtifact) {
      const tgt = (step.params as any).target || {};
      const explicitBlockRefFromUser = Boolean((tgt as any).__explicitBlockRefFromUser);
      const userBlockRef = (tgt as any).__userBlockRefRaw || tgt.blockRef;
      const explicitTargetResolvable = explicitBlockRefFromUser && hasResolvableExplicitBlockRef(userBlockRef);
      if (explicitTargetResolvable) {
        addLog({
          level: "info",
          message: "write_formula retarget skipped because explicit target provided",
          stepId: step.id,
          macro: "write_formula",
        });
      } else {
        const targetArtifactRef = ((step.params || {}).target || {}).artifactRef;
        const artifactMatchesCurrent = targetArtifactRef && currentTableRef && targetArtifactRef === currentTableRef.artifactRef;
        const allowRetarget = !targetArtifactRef || artifactMatchesCurrent;
        if (allowRetarget && (!tgt.blockRef || (currentTableRef && tgt.blockRef !== currentTableRef.blockRef))) {
          (step.params as any).target = {
            ...tgt,
            blockRef: (currentTableRef && currentTableRef.blockRef) || lastTableArtifact.blockRef,
            artifactRef: (currentTableRef && currentTableRef.artifactRef) || lastTableArtifact.fromStep,
          };
          addLog({ level: "warn", message: "write_formula retargeted to last produced table", stepId: step.id, macro: "write_formula" });
        }
      }
    }

    const macroCtx: MacroContext = {
      excelCtx,
      plan,
      context: contextState,
      step,
      decisions,
      artifacts,
      lastAddedHeader,
      headerRegistry,
      datasetRef: datasetRefForStep,
      log: (entry) =>
        addLog({
          stepId: step.id,
          macro: step.macro as AgentMacroName,
          ...entry,
        }),
    };

    try {
      const result = await macro(step.params || {}, macroCtx);
      if (result.artifacts) artifacts.push(...result.artifacts);
      if (result.requiresConfirmation) {
        confirmationsRequested.push(result.requiresConfirmation);
        status = "need_user_confirmation";
        addLog({ level: "warn", message: "Confirmation requise", stepId: step.id, macro: step.macro as AgentMacroName });
        break;
      }
      // auto-apply corporate format to any newly produced table artifacts for table_view outputs
      if (step.macro === "table_view" && result.artifacts) {
        const applyFmt = macros.apply_format;
        const tableArtifacts = result.artifacts.filter((a) => a.type === "table");
        for (const art of tableArtifacts) {
          try {
            await applyFmt(
              {
                target: { tableName: art.tableName, sheetName: art.sheet || art.sheetName, blockRef: art.blockRef, rangeA1: art.address },
                options: { preset: "corporate_blue" },
              },
              {
                excelCtx,
                plan,
                context: contextState,
                step: { ...step, id: `${step.id}:corporate_blue`, macro: "apply_format", params: {} },
                decisions,
                artifacts,
                log: (entry: any) => addLog({ ...entry, stepId: step.id, macro: "apply_format" as AgentMacroName }),
              } as any
            );
            addLog({ level: "info", message: `corporate_blue_format_applied table=${art.tableName || art.blockRef}`, stepId: step.id, macro: step.macro as AgentMacroName });
          } catch (err: any) {
            addLog({ level: "warn", message: `corporate_blue format failed: ${err?.message || err}`, stepId: step.id, macro: step.macro as AgentMacroName });
          }
        }
        const renameMap = (step.params as any)?.rename || {};
        const outputHeaders =
          (tableArtifacts[0] as any)?.headers ||
          (Array.isArray((step.params as any)?.select) && (step.params as any).select.length
            ? (step.params as any).select.map((h: string) => (renameMap[h] ? renameMap[h] : h))
            : headersByDataset[datasetRefForStep || ""] || []);
        const outDatasetRef = (step.params as any)?.dest?.mode === "inPlace" ? datasetRefForStep : (step.id || tableArtifacts[0]?.fromStep || `ds_${Date.now()}`);
        const keys: string[] = [];
        const srcKey = sourceKey((step.params as any)?.source || {});
        if (srcKey) keys.push(srcKey);
        if ((step.params as any)?.dest?.mode !== "inPlace") keys.push(`artifact:${(step.id || "").toLowerCase()}`);
        registerDataset(
          (outDatasetRef || "").toString().toLowerCase(),
          outputHeaders,
          keys,
          {
            artifactRef: (step.params as any)?.dest?.mode !== "inPlace" ? step.id : datasetMeta[datasetRefForStep || ""]?.artifactRef,
            blockRef: (step.params as any)?.dest?.mode === "inPlace" ? (step.params as any)?.source?.blockRef : tableArtifacts[0]?.blockRef,
          }
        );
      }
      if (result.artifacts) {
        const tableArt = result.artifacts.slice().reverse().find((a) => a.type === "table");
        if (tableArt) {
          lastTableArtifact = tableArt;
          currentTableRef = {
            blockRef: tableArt.blockRef,
            artifactRef: tableArt.fromStep || step.id,
            headers: (tableArt as any).headers,
            tableName: (tableArt as any).tableName,
          };
          lastAddedHeader = null;
        }
      }
      if (step.macro === "join_tables" && result.artifacts) {
        const tableArtifact = result.artifacts.find((a) => a.type === "table");
        if (tableArtifact) {
          lastJoinArtifact = tableArtifact;
          lastJoinSources = { left: (step.params as any)?.left?.blockRef, right: (step.params as any)?.right?.blockRef };
          const isNonEmptyString = (v: any): v is string => typeof v === "string" && v.length > 0;
          const joinHeaders = ((tableArtifact as any).headers as string[]) || [];
          registerDataset(
            (step.id || tableArtifact.fromStep || `join_${Date.now()}`).toLowerCase(),
            joinHeaders,
            [`artifact:${(step.id || "").toLowerCase()}`],
            {
              artifactRef: step.id,
              blockRef: tableArtifact.blockRef,
              headerAliases: (tableArtifact as any).headerAliases,
            }
          );
          const resolvedSheetName =
            (tableArtifact.sheet && isNonEmptyString(tableArtifact.sheet) && tableArtifact.sheet) ||
            (tableArtifact.sheetName && isNonEmptyString(tableArtifact.sheetName) && tableArtifact.sheetName) ||
            (contextSnapshot.active?.sheetName && isNonEmptyString(contextSnapshot.active.sheetName) ? contextSnapshot.active.sheetName : "Sheet1");
          lastProducedTable = {
            sheetName: resolvedSheetName,
            tableName: tableArtifact.tableName,
            blockRef: tableArtifact.blockRef || (tableArtifact.address ? `${resolvedSheetName}!${tableArtifact.address}` : undefined),
            address: tableArtifact.address,
          };
          // auto format corporate preset on join output
          try {
            const fmtMacro = macros.apply_format;
            if (fmtMacro) {
              await fmtMacro(
                {
                  target: {
                    tableName: tableArtifact.tableName,
                    sheetName: tableArtifact.sheet || tableArtifact.sheetName,
                    blockRef: tableArtifact.blockRef,
                    rangeA1: tableArtifact.address,
                  },
                  options: { preset: "corporate_blue" },
                },
                {
                  excelCtx,
                  plan,
                  context: contextState,
                  step: { ...step, id: `${step.id}:autoformat`, macro: "apply_format", params: {} },
                  decisions,
                  artifacts,
                  log: (entry: any) => addLog({ ...entry, stepId: step.id, macro: "apply_format" as AgentMacroName }),
                } as any
              );
            }
          } catch (err: any) {
            addLog({ level: "warn", message: `autoformat join_tables failed: ${err?.message || err}`, stepId: step.id, macro: "apply_format" as AgentMacroName });
          }
        }
      }
      if (result.status === "error") {
        if (step.macro === "apply_format") {
          addLog({ level: "warn", message: `apply_format ignored error status`, stepId: step.id, macro: step.macro as AgentMacroName });
          continue;
        }
        status = "error";
        errors.push(`step ${step.id} (${step.macro}) error`);
        break;
      }
      // propagate context changes when new column added (write_formula newColumnRight)
      if (
        step.macro === "write_formula" &&
        (result.status === undefined || result.status === "ok" || result.status === "skipped")
      ) {
        const target = step.params?.target;
        const headerName = (result as any)?.createdHeader || step.params?.headerName || target?.headerName;
        const blockRef = (result as any)?.blockRef || target?.blockRef;
        const writeMode = target?.writeMode || "newColumnRight";
        if (blockRef && writeMode === "newColumnRight" && headerName) {
          const sheet = contextState.sheets.find((s) => s.blocks.some((b) => b.id === blockRef || `${s.name}!${b.address}` === blockRef));
          const block = sheet?.blocks.find((b) => b.id === blockRef || `${sheet?.name}!${b.address}` === blockRef);
          if (sheet && block) {
            if (!block.headers.includes(headerName)) block.headers.push(headerName);
            lastAddedHeader = headerName;
            currentTableRef = { ...(currentTableRef || {}), blockRef, artifactRef: target?.artifactRef || (currentTableRef?.artifactRef ?? step.id), headers: block.headers.slice() };
            const dsRef = (datasetRefForStep || target?.artifactRef || target?.blockRef || step.id || "").toString().toLowerCase();
            registerDataset(dsRef, block.headers.slice(), [sourceKey(target)].filter(Boolean) as string[], { artifactRef: target?.artifactRef, blockRef });
            headerRegistry.push({ artifactRef: target?.artifactRef, blockRef, datasetRef: dsRef, headerName, sourceStepId: step.id });
            const addressWithSheet = block.address.includes("!") ? block.address : `${sheet.name}!${block.address}`;
            const bnd = parseA1Address(addressWithSheet);
            if (bnd) {
              const newEndCol = bnd.endCol + 1;
              const newAddr = `${rowColToA1(bnd.startRow, bnd.startCol)}:${rowColToA1(bnd.endRow, newEndCol)}`;
              block.address = newAddr; // keep id stable
              if (sheet.tables?.[0]) {
                const tbl = sheet.tables[0];
                tbl.address = `${sheet.name}!${newAddr}`;
                tbl.dataBodyAddress = `${sheet.name}!${rowColToA1(bnd.startRow + 1, bnd.startCol)}:${rowColToA1(bnd.endRow, newEndCol)}`;
                tbl.headerAddress = `${sheet.name}!${rowColToA1(bnd.startRow, bnd.startCol)}:${rowColToA1(bnd.startRow, newEndCol)}`;
                tbl.headers = block.headers;
              }
              sheet.valueBounds = {
                firstRow: bnd.startRow,
                firstCol: bnd.startCol,
                lastRow: bnd.endRow,
                lastCol: newEndCol,
                address: `${sheet.name}!${newAddr}`,
              };
            }
          }
        }
        // register dataset even when block not present in context (artifact-based)
        const dsRef = (datasetRefForStep || target?.artifactRef || target?.blockRef || step.id || "").toString().toLowerCase();
        const artMatch =
          artifacts
            .slice()
            .reverse()
            .find(
              (a) =>
                a.type === "table" &&
                ((target?.artifactRef && a.fromStep === target.artifactRef) || (target?.blockRef && a.blockRef === target.blockRef))
            ) || lastTableArtifact;
        if (artMatch && Array.isArray((artMatch as any).headers)) {
          registerDataset(
            dsRef,
            (artMatch as any).headers,
            [sourceKey({ artifactRef: target?.artifactRef, blockRef: target?.blockRef })].filter(Boolean) as string[],
            { artifactRef: target?.artifactRef || artMatch.fromStep, blockRef: (artMatch as any).blockRef }
          );
        }
      }
    } catch (err: any) {
      const msg = err?.message || String(err);
      if (step.macro === "apply_format" && typeof msg === "string" && msg.toLowerCase().includes("keynotfound")) {
        addLog({ level: "warn", message: `apply_format ignore erreur KeyNotFound`, stepId: step.id, macro: step.macro as AgentMacroName });
        continue;
      }
      addLog({ level: "error", message: msg, stepId: step.id, macro: step.macro as AgentMacroName, data: err });
      errors.push(`[${step.id} ${step.macro}] ${msg}`);
      status = "error";
      break;
    }
  }

  errors = errors.concat(logs.filter((l) => l.level === "error").map((l) => l.message));
  if (status !== "ok" && errors.length === 0) errors.push("execution_failed");
  const ok = status === "ok";
  return { logs, artifacts, confirmationsRequested: confirmationsRequested.length ? confirmationsRequested : undefined, status, errors, ok, warnings };
}
