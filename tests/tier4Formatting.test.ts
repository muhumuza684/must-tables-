import {
    colorForTier4Value,
    deriveGradientFromPalette,
    matchesTier4Rule,
    normalizeHex,
    paletteColors,
    parseCustomPalette,
    safeTier4ThemeList
} from "../src/tier4Formatting";

describe("Tier 4 formatting primitives", () => {
    it("supports numeric and text rule operators", () => {
        expect(matchesTier4Rule(10, { column: "Revenue", operator: "gt", value: "9", color: "#f00" })).toBe(true);
        expect(matchesTier4Rule("Overdue", { column: "Status", operator: "contains", value: "due", color: "#f00" })).toBe(true);
        expect(colorForTier4Value(10, [{ column: "Revenue", operator: "gte", value: "10", color: "#f00" }])).toBe("#f00");
    });

    it("exposes color-blind-safe presets, consistent with what the toolbar actually renders", () => {
        expect(paletteColors("deuteranopia")).toEqual(["#FDE725", "#440154"]);
        expect(paletteColors("protanopia")).toEqual(["#FDE725", "#31688E"]);
    });

    it("derives a brand-palette gradient from the shipped 9-color swatch set", () => {
        expect(paletteColors("brand")).toEqual(["#FAF623", "#124E9B"]);
    });

    describe("normalizeHex", () => {
        it("accepts hex with or without a leading #, case-insensitively", () => {
            expect(normalizeHex("faf623")).toBe("#FAF623");
            expect(normalizeHex("#FaF623")).toBe("#FAF623");
        });
        it("rejects malformed input", () => {
            expect(normalizeHex("not-a-color")).toBeNull();
            expect(normalizeHex("#fff")).toBeNull();
            expect(normalizeHex("")).toBeNull();
        });
    });

    describe("deriveGradientFromPalette", () => {
        it("picks the lightest and darkest swatch as [min, max]", () => {
            expect(deriveGradientFromPalette(["#124E9B", "#FAF623", "#606E4F"])).toEqual(["#FAF623", "#124E9B"]);
        });
        it("falls back to the default gradient for an empty palette", () => {
            expect(deriveGradientFromPalette([])).toEqual(["#DFFF91", "#0B3A70"]);
        });
        it("returns the same color twice for a single-swatch palette", () => {
            expect(deriveGradientFromPalette(["#3089BB"])).toEqual(["#3089BB", "#3089BB"]);
        });
    });

    describe("parseCustomPalette", () => {
        it("parses a comma-separated hex list", () => {
            expect(parseCustomPalette("faf623,eaec4a,124e9b")).toEqual(["#FAF623", "#EAEC4A", "#124E9B"]);
        });
        it("parses a JSON array of hex strings", () => {
            expect(parseCustomPalette('["faf623","eaec4a","124e9b"]')).toEqual(["#FAF623", "#EAEC4A", "#124E9B"]);
        });
        it("dedupes and caps at 12 colors", () => {
            const many = Array.from({ length: 20 }, (_, i) => `#${(i + 1).toString(16).padStart(6, "0")}`).join(",");
            const result = parseCustomPalette(many);
            expect(result).not.toBeNull();
            expect(result!.length).toBe(12);
        });
        it("returns null for input with no valid hex colors", () => {
            expect(parseCustomPalette("not colors at all")).toBeNull();
            expect(parseCustomPalette("")).toBeNull();
        });
    });

    describe("safeTier4ThemeList", () => {
        it("validates a list of saved themes, dropping malformed entries", () => {
            const result = safeTier4ThemeList([
                { id: "t1", name: "Operations", headerBg: "#0B3A70" },
                { id: "t2", name: "" }, // invalid: empty name
                { name: "No id" }, // invalid: missing id
                "not an object"
            ]);
            expect(result).toEqual([{ id: "t1", name: "Operations", headerBg: "#0B3A70", headerFont: undefined, cellBg: undefined, cellFont: undefined, altRow: undefined, accent: undefined }]);
        });

        it("upgrades a legacy single-theme object into a one-item list", () => {
            const result = safeTier4ThemeList({ id: "legacy", name: "Old Theme" });
            expect(result.length).toBe(1);
            expect(result[0].name).toBe("Old Theme");
        });

        it("returns an empty list for missing or invalid input", () => {
            expect(safeTier4ThemeList(undefined)).toEqual([]);
            expect(safeTier4ThemeList(null)).toEqual([]);
            expect(safeTier4ThemeList("garbage")).toEqual([]);
        });
    });
});
