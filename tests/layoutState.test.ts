import { normalizeLayoutState, updateLayoutState } from "../src/layoutState";

describe("layout state transaction", () => {
  const fallback = { fontSize: 14, rowHeight: 36, headerBold: true, responsive: true, autoFitColumns: true };
  it("preserves unrelated values when one property changes", () => {
    const next = updateLayoutState(fallback, { rowHeight: 48 });
    expect(next.rowHeight).toBe(48);
    expect(next.fontSize).toBe(14);
    expect(next.headerBold).toBe(true);
  });
  it("clamps invalid persisted numbers", () => {
    const next = normalizeLayoutState({ fontSize: 999, rowHeight: -4 }, fallback);
    expect(next.fontSize).toBe(48);
    expect(next.rowHeight).toBe(22);
  });
  test("font-size updates preserve row height", () => {
    const state = updateLayoutState(fallback, { fontSize: 18.5 });
    expect(state.fontSize).toBe(18.5);
    expect(state.rowHeight).toBe(fallback.rowHeight);
  });

  test("row-height updates preserve font size", () => {
    const state = updateLayoutState(fallback, { rowHeight: 90 });
    expect(state.rowHeight).toBe(90);
    expect(state.fontSize).toBe(fallback.fontSize);
  });
});

