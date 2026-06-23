import { WorkbookContextSnapshot } from "../../src/taskpane/context/types";

export function checkBlockRefs(plan: any, snapshot: WorkbookContextSnapshot): string[] {
  if (!plan?.steps) return [];
  const blockIds = new Set<string>((snapshot.sheets || []).flatMap((s) => (s.blocks || []).map((b: any) => b.id)));
  const errs: string[] = [];
  (plan.steps || []).forEach((step: any) => {
    const params = step?.params || {};
    const refs: string[] = [];
    if (params.left?.blockRef) refs.push(params.left.blockRef);
    if (params.right?.blockRef) refs.push(params.right.blockRef);
    if (params.target?.blockRef) refs.push(params.target.blockRef);
    if (params.source?.blockRef) refs.push(params.source.blockRef);
    refs.forEach((r) => {
      if (r && !blockIds.has(r)) errs.push(`blockref_not_found:${r}`);
    });
  });
  return errs;
}
