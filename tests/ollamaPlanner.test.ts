import { planWithOllama, autoRepairFromIntent, parseJson, extractStrictJSONObject } from "../src/taskpane/agent/planner/ollamaPlanner";

const originalOllamaDisabled = process.env.OLLAMA_DISABLED;
beforeAll(() => {
  process.env.OLLAMA_DISABLED = "0";
});
afterAll(() => {
  if (typeof originalOllamaDisabled === "undefined") {
    delete process.env.OLLAMA_DISABLED;
  } else {
    process.env.OLLAMA_DISABLED = originalOllamaDisabled;
  }
});

const samplePlan = {
  version: "1.0",
  goal: "demo",
  
  steps: [{ id: "s1", macro: "summarize_actions", params: {} }],
};

describe("planner JSON helpers", () => {
  test("extractStrictJSONObject isolates the object amid prose", () => {
    const noisy = "Voici un plan:\n```json\n{\"goal\":\"test\"}\n```\nMerci";
    const strict = extractStrictJSONObject(noisy);
    expect(strict).toContain("{\"goal\":\"test\"}");
  });

  test("parseJson succeeds when JSON is embedded in extra text", () => {
    const noisy = "Plan:\n{\n\"goal\":\"embedded\"\n}\nfin.";
    const parsed = parseJson(noisy);
    expect(parsed.ok).toBe(true);
    if (parsed.ok) expect(parsed.value.goal).toBe("embedded");
  });
});

describe("planWithOllama", () => {
  const originalFetch = global.fetch;

  afterEach(() => {
    jest.clearAllMocks();
    global.fetch = originalFetch as any;
    jest.clearAllTimers();
  });

  test("returns ok for valid JSON plan", async () => {
    global.fetch = jest.fn().mockResolvedValue({
      ok: true,
      text: () => Promise.resolve(JSON.stringify({ message: { content: JSON.stringify(samplePlan) } })),
    }) as any;

    const res = await planWithOllama({ context: { sheets: [], active: {} }, userPrompt: "hello" });
    expect(res.status).toBe("ok");
    if (res.status === "ok") {
      expect(res.plan.goal).toBe("demo");
    }
  });

  test("repairs once when first response invalid and stores rawTextRetry", async () => {
    const bad = "{ not json";
    global.fetch = jest
      .fn()
      .mockResolvedValueOnce({
        ok: true,
        text: () => Promise.resolve(bad),
      })
      .mockResolvedValueOnce({
        ok: true,
        text: () => Promise.resolve(JSON.stringify({ message: { content: JSON.stringify(samplePlan) } })),
      }) as any;

    const res = await planWithOllama({ context: { sheets: [], active: {} }, userPrompt: "hello" });
    expect(res.status).toBe("ok");
    if (res.status === "ok") {
      expect(res.rawTextRetry).toBe(JSON.stringify(samplePlan));
      expect(res.rawText).toBe(JSON.stringify(samplePlan)); // retry becomes source of truth
    }
    expect((global.fetch as any).mock.calls.length).toBe(2);
  });

  test("forbidden keys triggers retry then error if still invalid", async () => {
    const badWithForbidden = JSON.stringify({ action: "do", version: "1.0", goal: "x", steps: [] });
    global.fetch = jest
      .fn()
      .mockResolvedValueOnce({
        ok: true,
        text: () => Promise.resolve(badWithForbidden),
      })
      .mockResolvedValueOnce({
        ok: true,
        text: () => Promise.resolve(badWithForbidden),
      }) as any;

    const res = await planWithOllama({ context: { sheets: [], active: {} }, userPrompt: "hello" });
    expect(res.status).not.toBe("ok");
  });

  test("sanitize rejects empty steps", async () => {
    const emptyPlan = { version: "1.0", goal: "g", steps: [] };
    global.fetch = jest.fn().mockResolvedValue({
      ok: true,
      text: () => Promise.resolve(JSON.stringify({ message: { content: JSON.stringify(emptyPlan) } })),
    }) as any;
    const res = await planWithOllama({ context: { sheets: [], active: {} }, userPrompt: "hello" });
    expect(res.status).not.toBe("ok");
  });

  test("parses JSON with trailing text", async () => {
    const content = `${JSON.stringify(samplePlan)}\nTrailing notes...`;
    global.fetch = jest.fn().mockResolvedValue({
      ok: true,
      text: () => Promise.resolve(JSON.stringify({ message: { content } })),
    }) as any;
    const res = await planWithOllama({ context: { sheets: [], active: {} }, userPrompt: "hello" });
    expect(res.status).toBe("ok");
  });

  test("returns error when fetch aborts immediately", async () => {
    global.fetch = jest.fn().mockRejectedValue(new Error("AbortError")) as any;
    const res = await planWithOllama({ context: { sheets: [], active: {} }, userPrompt: "hello" });
    expect(res.status).toBe("error");
  });

  test("returns invalid_plan when responses remain invalid", async () => {
    global.fetch = jest.fn().mockResolvedValue({
      ok: true,
      text: () => Promise.resolve("not json"),
    }) as any;

    const res = await planWithOllama({ context: {}, userPrompt: "graph" });
    expect(res.status).toBe("invalid_plan");
  });

  test("rejects plans referencing unknown blockRef placeholders", async () => {
    const badPlan = {
      version: "1.0",
      goal: "bad block ref",
      
      steps: [
        {
          id: "tv1",
          macro: "table_view",
          params: {
            source: { blockRef: "context.sheets[0].blocks[0].id" },
            dest: { mode: "newSheet" },
            select: ["Projet"],
          },
        },
      ],
    };
    global.fetch = jest.fn().mockResolvedValue({
      ok: true,
      text: () => Promise.resolve(JSON.stringify({ message: { content: JSON.stringify(badPlan) } })),
    }) as any;

    const res = await planWithOllama({
      context: {
        sheets: [
          {
            name: "Sheet1",
            usedRange: "A1:A1",
            valueBounds: { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0, address: "A1:A1" },
            counts: { tables: 0, charts: 0 },
            tables: [],
            blocks: [
              {
                id: "Sheet1!A1:A1",
                address: "A1:A1",
                kind: "table",
                confidence: 1,
                headerRowIndex: 0,
                headers: ["Projet"],
                columnTypes: ["text"],
                preview: [],
                source: { type: "range" } as any,
              },
            ],
            charts: [],
            limitations: [],
          },
        ],
        active: {},
        capabilities: [],
        limitations: [],
        workbook: { name: "Book", readOnly: false },
        totals: { sheets: 1, tables: 0, charts: 0, blocks: 1, durationMs: 0 },
      },
      userPrompt: "bad block ref",
    });
    expect(res.status).toBe("invalid_plan");
    expect(res.failureStage).toBe("block_ref_invalid");
    if (res.status === "invalid_plan") {
      expect(res.errors.some((err: string) => err.includes("blockref_invalid"))).toBe(true);
    }
  });

  test("fails invariants when wantsFilter but no filter", async () => {
    const badPlan = {
      version: "1.0",
      goal: "filter please",
      
      steps: [{ id: "v1", macro: "table_view", params: { source: { blockRef: "b1" }, dest: { mode: "newSheet" }, select: ["A"] } }],
    };
    global.fetch = jest.fn().mockResolvedValue({
      ok: true,
      text: () => Promise.resolve(JSON.stringify({ message: { content: JSON.stringify(badPlan) } })),
    }) as any;

    const res = await planWithOllama({ context: { sheets: [{ blocks: [{ id: "b1", headers: ["A"] }] }], active: {} }, userPrompt: "filtrer Hab > 25" });
    expect(res.status).toBe("invalid_plan");
  });

  test("returns error when fetch fails", async () => {
    global.fetch = jest.fn().mockRejectedValue(new Error("ECONNREFUSED")) as any;
    const res = await planWithOllama({ context: {}, userPrompt: "x" });
    expect(res.status).toBe("error");
  });
});



describe('autoRepairFromIntent select repair', () => {
  test('fills select when missing and filter present', () => {
    const plan = {
      version: '1.0',
      goal: 'demo',
      
      steps: [
        { id: 'v1', macro: 'table_view', params: { source: { blockRef: 'b1' }, filter: [{ col: 'm2', op: 'gt', value: 5000 }], dest: { mode: 'newSheet' } } },
      ],
    } as any;
    const repaired = autoRepairFromIntent(plan);
    expect(repaired.changed).toBe(true);
    expect(repaired.plan.steps[0].params.select).toEqual(['m2']);
  });
});

