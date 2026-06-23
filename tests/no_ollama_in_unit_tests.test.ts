import { planWithOllama } from "../src/taskpane/agent/planner/ollamaPlanner";

describe("no Ollama in unit tests", () => {
  const originalDisabled = process.env.OLLAMA_DISABLED;
  beforeAll(() => {
    process.env.OLLAMA_DISABLED = "1";
  });
  afterAll(() => {
    if (typeof originalDisabled === "undefined" || originalDisabled === null) {
      delete process.env.OLLAMA_DISABLED;
    } else {
      process.env.OLLAMA_DISABLED = originalDisabled;
    }
  });

  test("planWithOllama throws when disabled", async () => {
    await expect(
      planWithOllama({
        context: { sheets: [], active: {} },
        userPrompt: "test",
      })
    ).rejects.toThrow("Ollama disabled in unit tests");
  });
});
