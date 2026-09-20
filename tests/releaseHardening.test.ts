import { clampViewport, shouldPreferCompactLayout, withinRenderBudget } from "../src/releaseHardening";

describe("release hardening", () => {
  it("normalizes invalid viewport values", () => {
    expect(clampViewport(-2, 300)).toEqual({ width: 0, height: 300 });
    expect(clampViewport(Number.NaN, Number.POSITIVE_INFINITY)).toEqual({ width: 0, height: 0 });
  });
  it("uses deterministic compact thresholds", () => {
    expect(shouldPreferCompactLayout(400, 500)).toBe(true);
    expect(shouldPreferCompactLayout(800, 500)).toBe(false);
  });
  it("checks a render-time budget without throwing", () => {
    expect(withinRenderBudget(100)).toBe(true);
    expect(withinRenderBudget(501)).toBe(false);
  });
});
