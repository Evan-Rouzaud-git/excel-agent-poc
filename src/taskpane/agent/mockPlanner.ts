import { AgentPlan } from "./types";

function hasWord(prompt: string, word: string) {
  return prompt.toLowerCase().includes(word.toLowerCase());
}

function containsAny(prompt: string, words: string[]) {
  const lower = prompt.toLowerCase();
  return words.some((w) => lower.includes(w.toLowerCase()));
}

export async function mockPlanner(context: any, userPrompt: string): Promise<AgentPlan> {
  const blockRef = context?.active?.selectionInBlockId || context?.active?.nearestBlockId || context?.sheets?.[0]?.blocks?.[0]?.id;
  const steps: any[] = [];

  const needFormat = containsAny(userPrompt, ["format", "propre", "presentation", "présentation", "entetes", "en-têtes", "gèle", "gele"]);
  const needMarge = userPrompt.toLowerCase().includes("marge");
  const needPourcent = userPrompt.includes("%") || userPrompt.toLowerCase().includes("pourcent");
  const needEcart = containsAny(userPrompt, ["ecart", "écart", "difference", "différence"]);
  const needChart = containsAny(userPrompt, ["graphique", "courbe", "chart"]);

  if (needFormat) {
    steps.push({
      id: "fmt1",
      macro: "apply_format",
      params: { target: { blockRef }, options: { preset: "corporate_blue", freezeHeaderRow: containsAny(userPrompt, ["gele", "gèle"]) } },
    });
  }

  if (needMarge || needEcart) {
    const headerName = needPourcent ? "% Marge" : needEcart ? "Ecart" : "Marge";
    const formula = needPourcent ? "=[@Marge]/[@Revenus]" : "=[@Revenus]-[@Depenses]";
    steps.push({
      id: "calc1",
      macro: "write_formula",
      params: {
        target: { blockRef, writeMode: "newColumnRight", headerName },
        formula,
        fillDown: true,
        ifOverwrite: "ask",
      },
    });
  }

  if (needChart) {
    const yCol = hasWord(userPrompt, "dépenses") || hasWord(userPrompt, "depenses") ? 1 : hasWord(userPrompt, "marge") ? 3 : 2; // assume Revenus col=2, marge maybe new col
    steps.push({
      id: "chart1",
      macro: "create_chart",
      params: {
        source: { blockRef },
        mapping: { xCol: { colIndex: 0 }, yCols: [{ colIndex: yCol }] },
        chartType: hasWord(userPrompt, "ligne") ? "line" : "columnClustered",
        dest: { mode: hasWord(userPrompt, "nouvelle feuille") ? "newSheet" : "right", anchor: { blockRef } },
        titleHint: "Graphique",
      },
    });
  }

  steps.push({ id: "sum1", macro: "summarize_actions", params: {} });

  return {
    version: "1.0",
    goal: userPrompt.slice(0, 80),
    steps,
  };
}
