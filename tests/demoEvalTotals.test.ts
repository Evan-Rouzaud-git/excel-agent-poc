import { describe, test, expect } from "@jest/globals";
import { EvalResult } from "./demoEvalRunner";

function computeTotals(results: EvalResult[]) {
  return {
    retryUsedCount: results.filter((r) => r.plannerRetryUsed).length,
    fallbackUsedCount: results.filter((r) => r.plannerFallbackUsed).length,
    nonJsonCount: results.filter((r) => r.plannerOutputWasNonJson).length,
    sanitizeChangedCount: results.filter((r) => r.plannerSanitizeChangedPlan).length,
  };
}

describe("demoEval totals flags", () => {
  test("counts retries/fallback/nonJson/sanitizeChanged correctly", () => {
    const base: Omit<EvalResult, "id" | "prompt" | "suiteId"> = {
      validPlan: true,
      executedOk: true,
      errors: [],
      steps: [],
      artifacts: 0,
      plannerRetryUsed: false,
      plannerFallbackUsed: false,
      plannerOutputWasNonJson: false,
      plannerSanitizeChangedPlan: false,
      plannerParseError: null,
      plannerSanitizeNotes: [],
      plannerRetryReason: null,
    };
    const sample: EvalResult[] = [
      { id: "a", prompt: "", suiteId: "format", ...base, plannerRetryUsed: true },
      { id: "b", prompt: "", suiteId: "format", ...base, plannerFallbackUsed: true },
      { id: "c", prompt: "", suiteId: "format", ...base, validPlan: false, executedOk: false, errors: ["x"], plannerOutputWasNonJson: true },
      { id: "d", prompt: "", suiteId: "format", ...base, plannerSanitizeChangedPlan: true },
    ];
    const totals = computeTotals(sample);
    expect(totals.retryUsedCount).toBe(1);
    expect(totals.fallbackUsedCount).toBe(1);
    expect(totals.nonJsonCount).toBe(1);
    expect(totals.sanitizeChangedCount).toBe(1);
  });
});
