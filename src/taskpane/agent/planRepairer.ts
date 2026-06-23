import { WorkbookContextSnapshot } from "../context/types";
import { AgentPlan } from "./types";
import { RepairOptions, repairWriteFormulaStep } from "./planner/formulaRepairer";

export interface RepairPlanOptions {
  logger?: (level: "info" | "warn" | "error", message: string) => void;
  formulaRepairOpts?: RepairOptions;
}

export async function repairPlanWriteFormulas(
  plan: AgentPlan,
  context: WorkbookContextSnapshot,
  userPrompt: string,
  opts: RepairPlanOptions = {}
) {
  const notes: string[] = [];
  const repairedSteps = [];
  for (const step of plan.steps) {
    if (step.macro === "write_formula") {
      try {
        const repairOpts: RepairOptions = {
          ...(opts.formulaRepairOpts || {}),
          logger: opts.logger
            ? (level, message) => opts.logger?.(level, `${message} (${step.id})`)
            : undefined,
        };
        const res = await repairWriteFormulaStep(step as any, context, userPrompt, repairOpts);
        if (res.repairLog.applied) {
          repairedSteps.push(res.patchedStep as any);
          notes.push(`step ${step.id}: ${res.repairLog.finalFormula || "patched"}`);
        } else {
          repairedSteps.push(step);
          notes.push(`step ${step.id}: réparateur skip (${res.repairLog.reason || "unknown"})`);
        }
      } catch (err: any) {
        repairedSteps.push(step);
        notes.push(`step ${step.id}: réparateur erreur ${err?.message || err}`);
      }
    } else {
      repairedSteps.push(step);
    }
  }
  const repairedPlan: AgentPlan = { ...plan, steps: repairedSteps };
  return { repairedPlan, notes };
}
