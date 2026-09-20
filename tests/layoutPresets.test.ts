import { createLayoutId, normalizeLayoutName, parseNamedLayouts, removeNamedLayout, upsertNamedLayout, LayoutSnapshot } from "../src/layoutPresets";

describe("named layouts", () => {
    const state: LayoutSnapshot = {
        sortColumn: "Revenue",
        sortDirection: "desc" as const,
        columnOrder: ["Region", "Revenue"],
        columnWidths: { Region: 180, Revenue: 120 },
        hiddenColumns: [],
        searchTerm: "East",
        groupExpansion: { East: true },
        fontSize: 12,
        rowHeight: 32,
        headerBold: true
    };

    it("normalizes names and creates stable readable ids", () => {
        expect(normalizeLayoutName("  Regional   review  ")).toBe("Regional review");
        expect(createLayoutId("Regional review", new Date("2026-01-01T00:00:00.000Z"))).toMatch(/^regional-review-/);
    });

    it("upserts, sorts, parses, and removes layouts", () => {
        const first = { id: "b", name: "Zulu", updatedAt: new Date().toISOString(), state };
        const second = { id: "a", name: "Alpha", updatedAt: new Date().toISOString(), state };
        const layouts = upsertNamedLayout(upsertNamedLayout([], first), second);
        expect(layouts.map((item) => item.name)).toEqual(["Alpha", "Zulu"]);
        expect(parseNamedLayouts(layouts).length).toBe(2);
        expect(removeNamedLayout(layouts, "a").map((item) => item.id)).toEqual(["b"]);
    });

    it("rejects malformed layout payloads", () => {
        expect(parseNamedLayouts([{ id: "bad", name: "Bad", state: {} }])).toEqual([]);
    });
});
