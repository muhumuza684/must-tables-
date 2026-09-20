import { AnalyticsFormatKind, AnalyticsRuleOperator, AnalyticsRuleScope, matchesAnalyticsRule } from "./analyticsFormatting";

export type Tier4PaletteName = "default" | "deuteranopia" | "protanopia" | "brand" | "custom";

export interface ITier4ConditionalRule {
    column: string;
    operator: AnalyticsRuleOperator | "gte" | "lte";
    value: string;
    secondValue?: string;
    color: string;
    /** Row scope remains backward compatible with Tier 4 saved rules. */
    scope?: AnalyticsRuleScope;
    /** Missing format means legacy background-color formatting. */
    format?: AnalyticsFormatKind;
    icon?: "up" | "down" | "dot";
}

/** A complete, click-to-apply visual theme: header/cell colors + accent, independent of the data palette. */
export interface ITier4SavedTheme {
    id: string;
    name: string;
    headerBg?: string;
    headerFont?: string;
    cellBg?: string;
    cellFont?: string;
    altRow?: string;
    accent?: string;
}

const HEX_COLOR = /^#?[0-9a-fA-F]{6}$/;

/** Normalizes a hex string to `#RRGGBB` uppercase, or null if it isn't a valid 6-digit hex color. */
export function normalizeHex(raw: string): string | null {
    const trimmed = raw.trim();
    if (!HEX_COLOR.test(trimmed)) return null;
    const withHash = trimmed.startsWith("#") ? trimmed : `#${trimmed}`;
    return withHash.toUpperCase();
}

/** Relative luminance (WCAG-style, sRGB) used only to order swatches light -> dark for gradient derivation. */
function relativeLuminance(hex: string): number {
    const clean = hex.replace("#", "");
    const r = parseInt(clean.slice(0, 2), 16) / 255;
    const g = parseInt(clean.slice(2, 4), 16) / 255;
    const b = parseInt(clean.slice(4, 6), 16) / 255;
    const lin = (c: number) => (c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4));
    return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
}

/**
 * Derives a [min, max] gradient pair from an arbitrary palette by picking the lightest and
 * darkest swatches. Used for both the built-in "brand" preset and any custom pasted palette,
 * so the data-bar / conditional-format gradient always tracks whatever colors are supplied.
 */
export function deriveGradientFromPalette(colors: string[]): [string, string] {
    const valid = colors.map(normalizeHex).filter((c): c is string => c !== null);
    if (valid.length === 0) return ["#DFFF91", "#0B3A70"];
    if (valid.length === 1) return [valid[0], valid[0]];
    const sorted = [...valid].sort((a, b) => relativeLuminance(b) - relativeLuminance(a));
    return [sorted[0], sorted[sorted.length - 1]];
}

/**
 * Parses a pasted palette in either the "CSV" (`faf623,eaec4a,...`) or JSON array
 * (`["faf623","eaec4a",...]`) shapes. Returns validated, normalized hex colors (dedupe,
 * capped at 12), or null if nothing valid could be parsed.
 */
export function parseCustomPalette(text: string): string[] | null {
    const trimmed = text.trim();
    if (!trimmed) return null;

    let candidates: string[];
    if (trimmed.startsWith("[")) {
        try {
            const parsed = JSON.parse(trimmed);
            if (!Array.isArray(parsed)) return null;
            candidates = parsed.map((v) => String(v));
        } catch {
            return null;
        }
    } else {
        candidates = trimmed.split(/[,\n]/);
    }

    const seen = new Set<string>();
    const result: string[] = [];
    for (const candidate of candidates) {
        const normalized = normalizeHex(candidate);
        if (normalized && !seen.has(normalized)) {
            seen.add(normalized);
            result.push(normalized);
            if (result.length >= 12) break;
        }
    }
    return result.length > 0 ? result : null;
}

/** Fixed palette presets. `brand` is the shipped default palette; `custom` is populated at runtime from user input. */
export const TIER4_PALETTE_SWATCHES: Record<Exclude<Tier4PaletteName, "custom">, string[]> = {
    default: ["#DFFF91", "#0B3A70"],
    deuteranopia: ["#FDE725", "#440154"],
    protanopia: ["#FDE725", "#31688E"],
    brand: ["#FAF623", "#EAEC4A", "#124E9B", "#606E4F", "#3089BB", "#9CD5D6", "#5D6A6F", "#8C8F44", "#C4CB5A"]
};

export function paletteColors(name: Tier4PaletteName, customPalette?: string[]): [string, string] {
    if (name === "custom") return deriveGradientFromPalette(customPalette ?? []);
    const swatches = TIER4_PALETTE_SWATCHES[name] ?? TIER4_PALETTE_SWATCHES.default;
    return deriveGradientFromPalette(swatches);
}

export function matchesTier4Rule(raw: unknown, rule: ITier4ConditionalRule): boolean {
    // Preserve legacy >= / <= operators while routing the Tier 6 operator set through
    // one calculation helper shared by the editor and renderer.
    if (rule.operator === "gte" || rule.operator === "lte") {
        const lhs = Number(raw);
        const rhs = Number(rule.value);
        return Number.isFinite(lhs) && Number.isFinite(rhs) && (rule.operator === "gte" ? lhs >= rhs : lhs <= rhs);
    }
    return matchesAnalyticsRule(raw, { ...rule, operator: rule.operator as Exclude<ITier4ConditionalRule["operator"], "gte" | "lte"> });
}

export function colorForTier4Value(raw: unknown, rules: ITier4ConditionalRule[]): string | null {
    const rule = rules.find((candidate) => candidate.scope !== "row" && (!candidate.format || candidate.format === "background") && matchesTier4Rule(raw, candidate));
    return rule ? rule.color : null;
}

/** Validates one persisted theme object; drops anything malformed rather than throwing. */
function safeTier4Theme(raw: unknown): ITier4SavedTheme | null {
    if (!raw || typeof raw !== "object") return null;
    const value = raw as Partial<ITier4SavedTheme>;
    if (typeof value.name !== "string" || !value.name.trim()) return null;
    if (typeof value.id !== "string" || !value.id.trim()) return null;
    const field = (v: unknown): string | undefined => (typeof v === "string" ? v : undefined);
    return {
        id: value.id,
        name: value.name.slice(0, 80),
        headerBg: field(value.headerBg),
        headerFont: field(value.headerFont),
        cellBg: field(value.cellBg),
        cellFont: field(value.cellFont),
        altRow: field(value.altRow),
        accent: field(value.accent)
    };
}

/** Validates a persisted list of saved themes, dropping malformed entries. Also accepts a single
 *  legacy theme object (pre-gallery format) and upgrades it into a one-item list. */
export function safeTier4ThemeList(raw: unknown): ITier4SavedTheme[] {
    if (Array.isArray(raw)) {
        return raw.map(safeTier4Theme).filter((t): t is ITier4SavedTheme => t !== null);
    }
    const single = safeTier4Theme(raw);
    return single ? [single] : [];
}
