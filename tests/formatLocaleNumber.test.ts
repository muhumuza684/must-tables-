import { formatLocaleNumber } from "../src/tier4Governance";

describe("formatLocaleNumber", () => {
    it("formats a plain number with the default locale", () => {
        expect(formatLocaleNumber(1234.5)).toContain("1,234.5");
    });

    it("formats a currency value when a currency code is given", () => {
        const result = formatLocaleNumber(99.99, "en-US", "USD");
        expect(result).toMatch(/\$99\.99/);
    });

    it("falls back to a plain formatter instead of throwing on an invalid locale", () => {
        expect(() => formatLocaleNumber(42, "not-a-real-locale-xyz")).not.toThrow();
        const result = formatLocaleNumber(42, "not-a-real-locale-xyz");
        expect(result).toContain("42");
    });

    it("falls back to a plain formatter instead of throwing on an invalid currency code", () => {
        expect(() => formatLocaleNumber(10, "en-US", "NOTREAL")).not.toThrow();
        const result = formatLocaleNumber(10, "en-US", "NOTREAL");
        expect(result).toContain("10");
    });
});