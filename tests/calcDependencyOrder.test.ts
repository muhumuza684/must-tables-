import { orderCalcColumns, CalcColumnLike } from "../src/calcDependencyOrder";

function cols(map: Record<string, string[]>): Map<string, CalcColumnLike> {
    return new Map(Object.entries(map).map(([name, refs]) => [name, { referencedColumns: refs }]));
}

describe("orderCalcColumns", () => {
    it("returns an empty result for no columns", () => {
        const result = orderCalcColumns(new Map());
        expect(result.order).toEqual([]);
        expect(result.cyclic.size).toBe(0);
    });

    it("orders independent columns with no dependencies in any valid order (all present, none cyclic)", () => {
        const result = orderCalcColumns(cols({ A: ["Revenue"], B: ["Cost"] }));
        expect(result.order.sort()).toEqual(["A", "B"]);
        expect(result.cyclic.size).toBe(0);
    });

    it("orders a simple chain so dependencies come before dependents", () => {
        // C depends on B, B depends on A -- A must come first, C last.
        const result = orderCalcColumns(cols({ A: ["Revenue"], B: ["A"], C: ["B"] }));
        expect(result.order.indexOf("A")).toBeLessThan(result.order.indexOf("B"));
        expect(result.order.indexOf("B")).toBeLessThan(result.order.indexOf("C"));
        expect(result.cyclic.size).toBe(0);
    });

    it("ignores references to raw (non-calculated) columns entirely", () => {
        const result = orderCalcColumns(cols({ A: ["Revenue", "Cost", "SomeRawField"] }));
        expect(result.order).toEqual(["A"]);
    });

    it("detects a direct two-column cycle and excludes both from the ordering", () => {
        const result = orderCalcColumns(cols({ A: ["B"], B: ["A"] }));
        expect(result.cyclic.has("A")).toBe(true);
        expect(result.cyclic.has("B")).toBe(true);
        expect(result.order).not.toContain("A");
        expect(result.order).not.toContain("B");
    });

    it("detects a longer transitive cycle (A -> B -> C -> A)", () => {
        const result = orderCalcColumns(cols({ A: ["B"], B: ["C"], C: ["A"] }));
        expect(result.cyclic.size).toBe(3);
        expect(result.order.length).toBe(0);
    });

    it("does not flag a column referencing itself as part of a larger false cycle with unrelated columns", () => {
        const result = orderCalcColumns(cols({ A: ["A"], B: ["Revenue"] }));
        // Self-reference is filtered out by the `ref !== name` guard, so A has no real
        // dependencies and is not cyclic; B is unrelated and also fine.
        expect(result.cyclic.size).toBe(0);
        expect(result.order.sort()).toEqual(["A", "B"]);
    });

    it("correctly orders a diamond dependency (D depends on B and C, both depend on A)", () => {
        const result = orderCalcColumns(cols({ A: [], B: ["A"], C: ["A"], D: ["B", "C"] }));
        expect(result.cyclic.size).toBe(0);
        const pos = (n: string) => result.order.indexOf(n);
        expect(pos("A")).toBeLessThan(pos("B"));
        expect(pos("A")).toBeLessThan(pos("C"));
        expect(pos("B")).toBeLessThan(pos("D"));
        expect(pos("C")).toBeLessThan(pos("D"));
    });

    it("isolates a cycle without breaking ordering of unrelated non-cyclic columns", () => {
        const result = orderCalcColumns(cols({ A: ["B"], B: ["A"], C: ["Revenue"] }));
        expect(result.cyclic.has("A")).toBe(true);
        expect(result.cyclic.has("B")).toBe(true);
        expect(result.cyclic.has("C")).toBe(false);
        expect(result.order).toEqual(["C"]);
    });
});