import { pivotReducer, initialPivotState, PivotState, PivotCommand, normalizePivotConfig } from "../src/state/pivotState";

describe("pivotReducer", () => {
    it("starts in table mode with no pivot config", () => {
        expect(initialPivotState).toEqual({ viewMode: "table", pivotConfig: null });
    });

    it("SET_VIEW_MODE switches to a valid mode", () => {
        const next = pivotReducer(initialPivotState, { type: "SET_VIEW_MODE", mode: "pivot" });
        expect(next.viewMode).toBe("pivot");
    });

    it("SET_VIEW_MODE ignores an invalid mode and returns unchanged state", () => {
        const bad = { type: "SET_VIEW_MODE", mode: "bogus" } as unknown as PivotCommand;
        const next = pivotReducer(initialPivotState, bad);
        expect(next).toEqual(initialPivotState);
    });

    it("TOGGLE_VIEW_MODE flips table <-> pivot", () => {
        const once = pivotReducer(initialPivotState, { type: "TOGGLE_VIEW_MODE" });
        expect(once.viewMode).toBe("pivot");
        const twice = pivotReducer(once, { type: "TOGGLE_VIEW_MODE" });
        expect(twice.viewMode).toBe("table");
    });

    it("SET_PIVOT_CONFIG sets a full config", () => {
        const config = { rowField: "Region", columnField: "Quarter", valueField: "Revenue" };
        const next = pivotReducer(initialPivotState, { type: "SET_PIVOT_CONFIG", config });
        expect(next.pivotConfig).toEqual(config);
    });

    it("SET_PIVOT_CONFIG ignores a null/malformed config", () => {
        const bad = { type: "SET_PIVOT_CONFIG", config: null } as unknown as PivotCommand;
        const next = pivotReducer(initialPivotState, bad);
        expect(next.pivotConfig).toBeNull();
    });

    it("UPDATE_PIVOT_CONFIG merges a partial update onto an existing config", () => {
        const base: PivotState = { viewMode: "pivot", pivotConfig: { rowField: "Region", columnField: "Quarter", valueField: "Revenue" } };
        const next = pivotReducer(base, { type: "UPDATE_PIVOT_CONFIG", partial: { valueField: "Cost" } });
        expect(next.pivotConfig).toEqual({ rowField: "Region", columnField: "Quarter", valueField: "Cost" });
    });

    it("UPDATE_PIVOT_CONFIG starts from empty fields when there is no existing config", () => {
        const next = pivotReducer(initialPivotState, { type: "UPDATE_PIVOT_CONFIG", partial: { rowField: "Region" } });
        expect(next.pivotConfig).toEqual({ rowField: "Region", columnField: "", valueField: "" });
    });

    it("RESET_PIVOT_CONFIG clears the config without changing viewMode", () => {
        const base: PivotState = { viewMode: "pivot", pivotConfig: { rowField: "Region", columnField: "Quarter", valueField: "Revenue" } };
        const next = pivotReducer(base, { type: "RESET_PIVOT_CONFIG" });
        expect(next.pivotConfig).toBeNull();
        expect(next.viewMode).toBe("pivot");
    });

    it("HYDRATE applies only well-formed fields from persisted state", () => {
        const next = pivotReducer(initialPivotState, {
            type: "HYDRATE",
            state: { viewMode: "pivot", pivotConfig: { rowField: "Region", columnField: "Quarter", valueField: "Revenue" } }
        });
        expect(next.viewMode).toBe("pivot");
        expect(next.pivotConfig).toEqual({ rowField: "Region", columnField: "Quarter", valueField: "Revenue" });
    });

    it("HYDRATE ignores malformed persisted fields instead of throwing", () => {
        const bad = { type: "HYDRATE", state: { viewMode: "not-a-mode", pivotConfig: { nope: true } } } as unknown as PivotCommand;
        expect(() => pivotReducer(initialPivotState, bad)).not.toThrow();
        const next = pivotReducer(initialPivotState, bad);
        expect(next).toEqual(initialPivotState);
    });

    it("never mutates the input state object", () => {
        const base: PivotState = { viewMode: "table", pivotConfig: null };
        const frozen = Object.freeze(base);
        expect(() => pivotReducer(frozen, { type: "TOGGLE_VIEW_MODE" })).not.toThrow();
    });
});

describe("normalizePivotConfig", () => {
    it("accepts the legacy optional aggregation shape and rejects malformed fields", () => {
        expect(normalizePivotConfig({ rowField: "Region", columnField: "Quarter", valueField: "Revenue" })).toEqual({ rowField: "Region", columnField: "Quarter", valueField: "Revenue" });
        expect(normalizePivotConfig({ rowField: 1, columnField: "Quarter", valueField: "Revenue" })).toBeNull();
        expect(normalizePivotConfig({ rowField: "Region", columnField: "Quarter", valueField: "Revenue", aggregation: "bogus" })).toEqual({ rowField: "Region", columnField: "Quarter", valueField: "Revenue" });
    });
});
