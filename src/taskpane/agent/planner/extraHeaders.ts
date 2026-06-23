import { AgentPlan } from "../types";

function addHeaderEntry(map: Record<string, string[]>, key: string | undefined, header?: string) {
  if (!key || !header) return;
  const list = map[key] || [];
  if (!list.includes(header)) {
    list.push(header);
    map[key] = list;
  }
}

export function collectExtraHeaders(plan?: AgentPlan): Record<string, string[]> {
  const extras: Record<string, string[]> = {};
  if (!plan) return extras;
  (plan.steps || []).forEach((step: any) => {
    if (!step) return;
    if (step.macro === "write_formula") {
      const target = step.params?.target;
      addHeaderEntry(extras, target?.artifactRef, target?.headerName);
      addHeaderEntry(extras, target?.blockRef, target?.headerName);
    }
    if (step.macro === "join_tables" && step.id) {
      const keys = Array.isArray(step.params?.keys) ? step.params.keys : [];
      keys.forEach((key: any) => {
        addHeaderEntry(extras, step.id, key?.right);
        addHeaderEntry(extras, step.id, key?.left);
      });
    }
  });
  return extras;
}
