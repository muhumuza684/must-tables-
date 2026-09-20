import { aggregateRows, groupRows, pivotRows } from "../src/advancedModel";
import type { ITableRow } from "../src/tableRenderer";

function row(values: Record<string, unknown>): ITableRow {
    return { key: JSON.stringify(values), values } as ITableRow;
}

describe("advancedModel resilience against malformed input", () => {
    describe("aggregateRows", () => {
        it("does not throw when rows is null/undefined", () => {
            expect(() => aggregateRows(null as any, ["revenue"])).not.toThrow();
            expect(() => aggregateRows(undefined as any, ["revenue"])).not.toThrow();
        });
        it("does not throw when valueColumns is null/undefined", () => {
            expect(() => aggregateRows([row({ revenue: 5 })], null as any)).not.toThrow();
        });
        it("skips rows with a missing/malformed values property instead of throwing", () => {
            const malformed = [{ key: "bad" } as any, row({ revenue: 10 })];
            const result = aggregateRows(malformed, ["revenue"]);
            expect(result.count).toBe(2);
            expect(result.sums["revenue"]).toBe(10);
        });
    });

    describe("groupRows", () => {
        it("does not throw when rows/groupColumns/valueColumns are null/undefined", () => {
            expect(() => groupRows(null as any, ["region"], ["revenue"])).not.toThrow();
            expect(() => groupRows([row({ region: "East" })], null as any, ["revenue"])).not.toThrow();
            expect(() => groupRows([row({ region: "East" })], ["region"], null as any)).not.toThrow();
        });
        it("skips malformed rows instead of throwing", () => {
            const malformed = [{ key: "bad" } as any, row({ region: "East", revenue: 5 })];
            expect(() => groupRows(malformed, ["region"], ["revenue"])).not.toThrow();
            const result = groupRows(malformed, ["region"], ["revenue"]);
            expect(result.length).toBe(1);
            expect(result[0].key).toBe("East");
        });
    });

    describe("pivotRows", () => {
        it("does not throw when rows is null/undefined", () => {
            expect(() => pivotRows(null as any, "region", "product", "revenue")).not.toThrow();
            expect(() => pivotRows(undefined as any, "region", "product", "revenue")).not.toThrow();
        });
        it("skips malformed rows instead of throwing", () => {
            const malformed = [{ key: "bad" } as any, row({ region: "East", product: "A", revenue: 10 })];
            expect(() => pivotRows(malformed, "region", "product", "revenue")).not.toThrow();
            const result = pivotRows(malformed, "region", "product", "revenue");
            expect(result.rows.length).toBe(1);
        });
        it("returns an empty-but-valid result for an empty dataset", () => {
            const result = pivotRows([], "region", "product", "revenue");
            expect(result.rows).toEqual([]);
            expect(result.columns).toEqual(["region"]);
        });
    });
});