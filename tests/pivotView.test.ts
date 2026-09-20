import { pivotRows } from "../src/advancedModel";
import type { ITableRow } from "../src/tableRenderer";

function row(values: Record<string, unknown>): ITableRow {
    return { key: JSON.stringify(values), values } as ITableRow;
}

describe("pivotRows", () => {
    it("aggregates values into a row x column matrix", () => {
        const rows = [
            row({ region: "East", product: "A", revenue: 10 }),
            row({ region: "East", product: "B", revenue: 5 }),
            row({ region: "West", product: "A", revenue: 7 })
        ];
        const result = pivotRows(rows, "region", "product", "revenue");
        expect(result.columns).toEqual(["region", "revenue \u2014 A", "revenue \u2014 B"]);
        const east = result.rows.find((r) => r.region === "East")!;
        expect(east["revenue \u2014 A"]).toBe(10);
        expect(east["revenue \u2014 B"]).toBe(5);
        const west = result.rows.find((r) => r.region === "West")!;
        expect(west["revenue \u2014 B"]).toBe(0);
    });

    it("handles blank/null dimension values without crashing", () => {
        const rows = [row({ region: null, product: undefined, revenue: 3 })];
        const result = pivotRows(rows, "region", "product", "revenue");
        expect(result.rows[0].region).toBe("(Blank)");
    });

    it("sums repeated row/column combinations rather than overwriting", () => {
        const rows = [
            row({ region: "East", product: "A", revenue: 10 }),
            row({ region: "East", product: "A", revenue: 4 })
        ];
        const result = pivotRows(rows, "region", "product", "revenue");
        expect(result.rows[0]["revenue \u2014 A"]).toBe(14);
    });
});