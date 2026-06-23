import { __dateTest__ } from "../src/taskpane/agent/macros";

describe("parseDateLikeString", () => {
  const { parseDateLikeString, isAllDigits } = __dateTest__ as any;

  test("isAllDigits handles undefined and empty", () => {
    expect(isAllDigits(undefined)).toBe(false);
    expect(isAllDigits("")).toBe(false);
    expect(isAllDigits("123")).toBe(true);
    expect(isAllDigits("12a")).toBe(false);
  });

  test("parseDateLikeString ignores invalid parts safely", () => {
    expect(parseDateLikeString("")).toBeNull();
    expect(parseDateLikeString("10/13/2020")).toBeNull(); // invalid month
    expect(parseDateLikeString("32/01/2020")).toBeNull(); // invalid day
  });

  test("parseDateLikeString parses valid dd/mm/yyyy and yyyy-mm-dd", () => {
    const d1 = parseDateLikeString("05/11/2025");
    expect(d1).not.toBeNull();
    expect((d1 as Date).getFullYear()).toBe(2025);
    const d2 = parseDateLikeString("2027-09-15");
    expect(d2).not.toBeNull();
    expect((d2 as Date).getMonth()).toBe(8);
  });
});
