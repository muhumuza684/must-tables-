"use strict";

export type DataLakeStylePreset = "clean" | "classic" | "compact" | "executive" | "highContrast" | "custom";
export type DataLakeDensity = "compact" | "comfortable" | "spacious" | "custom";

export interface IDataLakeStyleTokens {
    background: string;
    surface: string;
    surfaceAlt: string;
    elevatedSurface: string;
    primaryText: string;
    secondaryText: string;
    disabledText: string;
    border: string;
    divider: string;
    accent: string;
    accentHover: string;
    focus: string;
    success: string;
    warning: string;
    error: string;
    selection: string;
    hover: string;
    fontFamily: string;
    fontSize: number;
    labelSize: number;
    rowHeight: number;
    headerHeight: number;
    controlHeight: number;
    spacing: number;
    radius: number;
    shadow: string;
    cellPaddingX: number;
    minColumnWidth: number;
    maxColumnWidth: number;
}

export interface IDataLakeStyleOverrides {
    colors?: Partial<Pick<IDataLakeStyleTokens,
        "background" | "surface" | "surfaceAlt" | "elevatedSurface" | "primaryText" | "secondaryText" |
        "disabledText" | "border" | "divider" | "accent" | "accentHover" | "focus" | "success" |
        "warning" | "error" | "selection" | "hover">>;
    fontFamily?: string;
    fontSize?: number;
    rowHeight?: number;
    headerHeight?: number;
    spacing?: number;
    radius?: number;
    minColumnWidth?: number;
    maxColumnWidth?: number;
}

export interface IDataLakeResolvedStyle {
    preset: DataLakeStylePreset;
    density: DataLakeDensity;
    tokens: IDataLakeStyleTokens;
}

const BASE: IDataLakeStyleTokens = {
    background: "#F6F8FA",
    surface: "#FFFFFF",
    surfaceAlt: "#F8FAFC",
    elevatedSurface: "#FFFFFF",
    primaryText: "#1F2937",
    secondaryText: "#667085",
    disabledText: "#98A2B3",
    border: "#D9E0E7",
    divider: "#E8ECF1",
    accent: "#2563EB",
    accentHover: "#1D4ED8",
    focus: "#2563EB",
    success: "#16803C",
    warning: "#A15C00",
    error: "#B42318",
    selection: "#E8F1FF",
    hover: "#F5F8FC",
    fontFamily: "Segoe UI, sans-serif",
    fontSize: 14,
    labelSize: 12,
    rowHeight: 36,
    headerHeight: 42,
    controlHeight: 32,
    spacing: 8,
    radius: 6,
    shadow: "0 2px 8px rgba(16,24,40,.08)",
    cellPaddingX: 12,
    minColumnWidth: 72,
    maxColumnWidth: 420
};

const PRESETS: Record<Exclude<DataLakeStylePreset, "custom">, Partial<IDataLakeStyleTokens>> = {
    clean: {},
    classic: {
        surfaceAlt: "#F5F6F8", border: "#CDD3DA", divider: "#DDE2E7", accent: "#0078D4", accentHover: "#106EBE",
        primaryText: "#242424", secondaryText: "#5F6368", rowHeight: 36, headerHeight: 40, radius: 4, cellPaddingX: 12
    },
    compact: {
        surfaceAlt: "#F7F8FA", border: "#D7DCE2", divider: "#E6E9ED", accent: "#2563EB", accentHover: "#1D4ED8",
        fontSize: 13, labelSize: 11, rowHeight: 30, headerHeight: 34, controlHeight: 28, spacing: 6, radius: 5, cellPaddingX: 8
    },
    executive: {
        background: "#F4F6F8", surface: "#FFFFFF", surfaceAlt: "#F8FAFC", border: "#D7DEE7", divider: "#E7EBF0",
        primaryText: "#172B4D", secondaryText: "#5B677A", accent: "#0F766E", accentHover: "#0B5F59",
        selection: "#E7F5F3", hover: "#F2F8F7", rowHeight: 38, headerHeight: 44, radius: 6, cellPaddingX: 14
    },
    highContrast: {
        background: "#FFFFFF", surface: "#FFFFFF", surfaceAlt: "#FFFFFF", elevatedSurface: "#FFFFFF", primaryText: "#000000",
        secondaryText: "#202020", disabledText: "#595959", border: "#000000", divider: "#000000", accent: "#005FCC",
        accentHover: "#004799", focus: "#000000", selection: "#D9EAFD", hover: "#EEEEEE", rowHeight: 36, headerHeight: 42,
        radius: 2, cellPaddingX: 10, shadow: "none"
    }
};

function clamp(value: number, min: number, max: number, fallback: number): number {
    return Number.isFinite(value) ? Math.max(min, Math.min(max, value)) : fallback;
}

export function resolveDataLakeStyle(
    preset: DataLakeStylePreset,
    density: DataLakeDensity,
    overrides: IDataLakeStyleOverrides = {}
): IDataLakeResolvedStyle {
    const base = { ...BASE, ...(preset === "custom" ? {} : PRESETS[preset]) };
    const withOverrides: IDataLakeStyleTokens = {
        ...base,
        ...overrides,
        ...(overrides.colors ?? {})
    } as IDataLakeStyleTokens;

    const densityAdjustments: Record<DataLakeDensity, Partial<IDataLakeStyleTokens>> = {
        compact: { fontSize: 13, labelSize: 11, rowHeight: 30, headerHeight: 34, controlHeight: 28, spacing: 6, cellPaddingX: 8 },
        comfortable: {},
        spacious: { rowHeight: 44, headerHeight: 50, controlHeight: 36, spacing: 10, cellPaddingX: 14 },
        custom: {}
    };
    const tokens = { ...withOverrides, ...densityAdjustments[density] };

    tokens.fontSize = clamp(tokens.fontSize, 10, 24, BASE.fontSize);
    tokens.labelSize = clamp(tokens.labelSize, 9, 18, BASE.labelSize);
    tokens.rowHeight = clamp(tokens.rowHeight, 22, 80, BASE.rowHeight);
    tokens.headerHeight = clamp(tokens.headerHeight, 28, 90, BASE.headerHeight);
    tokens.controlHeight = clamp(tokens.controlHeight, 24, 48, BASE.controlHeight);
    tokens.spacing = clamp(tokens.spacing, 4, 16, BASE.spacing);
    tokens.radius = clamp(tokens.radius, 0, 16, BASE.radius);
    tokens.cellPaddingX = clamp(tokens.cellPaddingX, 4, 24, BASE.cellPaddingX);
    tokens.minColumnWidth = clamp(tokens.minColumnWidth, 48, 240, BASE.minColumnWidth);
    tokens.maxColumnWidth = clamp(tokens.maxColumnWidth, tokens.minColumnWidth, 640, BASE.maxColumnWidth);

    return { preset, density, tokens };
}

export function stylePresetItems(): Array<{ value: DataLakeStylePreset; displayName: string }> {
    return [
        { value: "clean", displayName: "Clean" },
        { value: "classic", displayName: "Classic" },
        { value: "compact", displayName: "Compact" },
        { value: "executive", displayName: "Executive" },
        { value: "highContrast", displayName: "High Contrast" },
        { value: "custom", displayName: "Custom" }
    ];
}

export function densityItems(): Array<{ value: DataLakeDensity; displayName: string }> {
    return [
        { value: "compact", displayName: "Compact" },
        { value: "comfortable", displayName: "Comfortable" },
        { value: "spacious", displayName: "Spacious" },
        { value: "custom", displayName: "Custom" }
    ];
}

export function styleMigrationPreset(values: {
    headerBg?: string; headerFont?: string; cellBg?: string; cellFont?: string; altRow?: string;
    fontSize?: number; rowHeight?: number;
}): DataLakeStylePreset {
    const legacyDefaults = ["#F0F2F5", "#333333", "#FFFFFF", "#333333", "#FAFAFA"];
    const current = [values.headerBg, values.headerFont, values.cellBg, values.cellFont, values.altRow];
    const customColors = current.some((value, index) => value && value.toUpperCase() !== legacyDefaults[index]);
    const customSizing = values.fontSize !== undefined && values.fontSize !== 14 || values.rowHeight !== undefined && values.rowHeight !== 36;
    return customColors || customSizing ? "custom" : "clean";
}

export function toCssVariables(style: IDataLakeResolvedStyle): Record<string, string> {
    const t = style.tokens;
    return {
        "--dlt-background": t.background,
        "--dlt-surface": t.surface,
        "--dlt-surface-alt": t.surfaceAlt,
        "--dlt-surface-elevated": t.elevatedSurface,
        "--dlt-text-primary": t.primaryText,
        "--dlt-text-secondary": t.secondaryText,
        "--dlt-text-disabled": t.disabledText,
        "--dlt-border": t.border,
        "--dlt-divider": t.divider,
        "--dlt-accent": t.accent,
        "--dlt-accent-hover": t.accentHover,
        "--dlt-focus": t.focus,
        "--dlt-success": t.success,
        "--dlt-warning": t.warning,
        "--dlt-error": t.error,
        "--dlt-selection": t.selection,
        "--dlt-hover": t.hover,
        "--dlt-font-family": t.fontFamily,
        "--dlt-font-body": `${t.fontSize}px`,
        "--dlt-font-label": `${t.labelSize}px`,
        "--dlt-row-h": `${t.rowHeight}px`,
        "--dlt-header-h": `${t.headerHeight}px`,
        "--dlt-control-h": `${t.controlHeight}px`,
        "--dlt-space": `${t.spacing}px`,
        "--dlt-radius": `${t.radius}px`,
        "--dlt-shadow": t.shadow,
        "--dlt-cell-x": `${t.cellPaddingX}px`,
        "--dlt-min-col-w": `${t.minColumnWidth}px`,
        "--dlt-max-col-w": `${t.maxColumnWidth}px`
    };
}
