import { normalizeRenderingFailure } from "../src/renderDiagnostics";

describe("Tier 3 rendering diagnostics", () => {
    it.each([
        [new Error("invalid viewport dimensions"), "rendering"],
        [new Error("dataView data unavailable"), "data"],
        [new Error("formatting state unavailable"), "state"],
        [new Error("fetchMoreData failed"), "fetch"]
    ])("classifies %s", (error, kind) => {
        const result = normalizeRenderingFailure(error);
        expect(result.kind).toBe(kind);
        expect(result.reason).toBe(error.message);
    });

    it("always returns a non-empty reason for non-Error failures", () => {
        expect(normalizeRenderingFailure("unexpected failure")).toEqual({ kind: "unknown", reason: "unexpected failure" });
    });
});


describe("release diagnostics", () => {
    it("accepts standard Power BI table shapes and rejects pathological shapes", () => {
        const { isReasonableTableShape } = require("../src/releaseDiagnostics");
        expect(isReasonableTableShape(30000, 20)).toBe(true);
        expect(isReasonableTableShape(30001, 20)).toBe(false);
        expect(isReasonableTableShape(-1, 20)).toBe(false);
    });
    it("detects common mojibake markers", () => {
        const { hasMojibakeMarkers } = require("../src/releaseDiagnostics");
        expect(hasMojibakeMarkers("plain utf8 text")).toBe(false);
        expect(hasMojibakeMarkers("broken Ã text")).toBe(true);
    });
});
