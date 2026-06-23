describe("withTimeout helper", () => {
  const withTimeout = async <T>(p: Promise<T>, ms: number) => {
    return Promise.race([
      p.then((v) => ({ ok: true, v })),
      new Promise<{ ok: false; err: any }>((resolve) => {
        const to = setTimeout(() => resolve({ ok: false, err: new Error("prompt_timeout") }), ms);
        p.finally(() => clearTimeout(to));
      }),
    ]);
  };

  test("returns timeout when promise hangs", async () => {
    const hang = new Promise<string>(() => {});
    const res = await withTimeout(hang, 20);
    expect(res.ok).toBe(false);
    if (!res.ok) expect((res as any).err).toBeInstanceOf(Error);
  });
});
