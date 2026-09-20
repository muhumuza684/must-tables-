import { applyAdvancedFilter, aggregateRows, groupRows, parseAdvancedFilter, pivotRows } from "../src/advancedModel";
import { ITableRow } from "../src/tableRenderer";

function row(key: string, values: Record<string, string | number | null>): ITableRow {
    return { key, values, selectionId: {} as ITableRow["selectionId"] };
}

describe("advanced table model", () => {
    const rows = [
        row("1", { Region: "East", Status: "Open", Revenue: 100, Quarter: "Q1" }),
        row("2", { Region: "East", Status: "Closed", Revenue: 50, Quarter: "Q2" }),
        row("3", { Region: "West", Status: "Open", Revenue: 200, Quarter: "Q1" }),
        row("4", { Region: "West", Status: null, Revenue: 25, Quarter: "Q2" })
    ];

    it("parses and evaluates compound filters", () => {
        const parsed = parseAdvancedFilter("Region = 'East' AND Revenue >= 100");
        expect(parsed.ok).toBe(true);
        if (!parsed.ok) return;
        expect(applyAdvancedFilter(rows, parsed.expression).map((item) => item.key)).toEqual(["1"]);
    });

    it("supports contains, between, in, and null predicates", () => {
        for (const input of ["Status CONTAINS 'pen'", "Revenue BETWEEN 50 AND 150", "Region IN ('East', 'North')", "Status IS NULL"]) {
            const parsed = parseAdvancedFilter(input);
            expect(parsed.ok).toBe(true);
            if (!parsed.ok) continue;
            expect(applyAdvancedFilter(rows, parsed.expression).length).toBeGreaterThan(0);
        }
    });

    it("reports malformed expressions without throwing", () => {
        const parsed = parseAdvancedFilter("Revenue >=");
        expect(parsed.ok).toBe(false);
        if (parsed.ok === false) expect(parsed.error).toMatch(/value|position/i);
    });

    it("computes deterministic group aggregates", () => {
        const grouped = groupRows(rows, ["Region"], ["Revenue"]);
        expect(grouped.map((group) => group.key)).toEqual(["East", "West"]);
        expect(grouped[0].aggregate).toMatchObject({ count: 2, sums: { Revenue: 150 }, minimums: { Revenue: 50 }, maximums: { Revenue: 100 } });
        expect(aggregateRows(rows, ["Revenue"]).sums.Revenue).toBe(375);
    });

    it("pivots a dimension into stable generated value columns", () => {
        const pivot = pivotRows(rows, "Region", "Quarter", "Revenue");
        expect(pivot.columns).toEqual(["Region", "Revenue — Q1", "Revenue — Q2"]);
        expect(pivot.rows).toEqual([
            { Region: "East", "Revenue — Q1": 100, "Revenue — Q2": 50 },
            { Region: "West", "Revenue — Q1": 200, "Revenue — Q2": 25 }
        ]);
    });
});
