import { AgentPlan, PlanStep } from "./types";

export type InvariantIssueCode = "place_output_fallback" | "table_view_select_required";

export interface InvariantIssue {
  code: InvariantIssueCode;
  message: string;
  stepId?: string;
  details?: any;
}

function needsSelect(step: PlanStep): boolean {
  if (!step || step.macro !== "table_view") return false;
  const destMode = step.params?.dest?.mode || "newSheet";
  if (destMode === "inPlace") return false;
  const select = Array.isArray(step.params?.select) ? step.params?.select : [];
  return select.length === 0;
}

export function validatePlanInvariants(plan: AgentPlan): { valid: boolean; issues: InvariantIssue[] } {
  const issues: InvariantIssue[] = [];
  const steps = Array.isArray(plan?.steps) ? plan.steps : [];

  const placeOnly = steps.length > 0 && steps.every((s) => s?.macro === "place_output");
  if (placeOnly) {
    issues.push({
      code: "place_output_fallback",
      message: "Plan ne peut pas se limiter à place_output. Fournir une action concrète (table_view, write_formula...).",
    });
  }

  steps.forEach((step) => {
    if (needsSelect(step)) {
      issues.push({
        code: "table_view_select_required",
        message: `table_view ${step.id} with dest.mode=${step.params?.dest?.mode || "newSheet"} must include at least one select entry.`,
        stepId: step.id,
      });
    }
  });

  return { valid: issues.length === 0, issues };
}
