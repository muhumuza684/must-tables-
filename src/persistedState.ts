"use strict";

export const USER_CONFIG_SCHEMA_VERSION = 2;
export const SAVED_VIEW_SCHEMA_VERSION = 1;

export interface PersistedUserConfigEnvelope {
    schemaVersion: number;
    calc?: Array<{ name: string; formula: string }>;
    combined?: Array<{ name: string; template: string }>;
    sparklines?: string[];
    dragGroup?: string | null;
    viewMode?: "table" | "pivot";
    pivotConfig?: unknown;
    tier4Rules?: unknown[];
    tier4ColumnColors?: Record<string, string>;
    tier4Palette?: string;
    tier4CustomPalette?: string[];
    tier4SavedThemes?: unknown[];
    tier4SavedTheme?: unknown;
    tier4ActiveThemeId?: string | null;
    namedLayouts?: unknown[];
    announcementDismissed?: boolean;
    layout?: { headerBold?: boolean; fontSize?: number; rowHeight?: number };
    settingsDrawerWidth?: number;
}

function finiteNumber(value: unknown, fallback: number, min: number, max: number): number {
    const n = Number(value);
    return Number.isFinite(n) ? Math.max(min, Math.min(max, n)) : fallback;
}

/**
 * Normalizes the report-persisted userConfig envelope without changing its existing
 * property names. Unknown properties are intentionally ignored rather than copied into
 * the live renderer state.
 */
export function normalizePersistedUserConfig(raw: unknown): PersistedUserConfigEnvelope | null {
    if (!raw || typeof raw !== "object") return null;
    const source = raw as Record<string, unknown>;

    const result: PersistedUserConfigEnvelope = {
        schemaVersion: typeof source.schemaVersion === "number" ? source.schemaVersion : 1
    };

    if (Array.isArray(source.calc)) {
        result.calc = source.calc.filter((item): item is { name: string; formula: string } =>
            !!item && typeof item === "object" && typeof (item as Record<string, unknown>).name === "string" && typeof (item as Record<string, unknown>).formula === "string"
        ).slice(0, 500);
    }
    if (Array.isArray(source.combined)) {
        result.combined = source.combined.filter((item): item is { name: string; template: string } =>
            !!item && typeof item === "object" && typeof (item as Record<string, unknown>).name === "string" && typeof (item as Record<string, unknown>).template === "string"
        ).slice(0, 500);
    }
    if (Array.isArray(source.sparklines)) result.sparklines = source.sparklines.filter((v): v is string => typeof v === "string").slice(0, 500);
    if (typeof source.dragGroup === "string") result.dragGroup = source.dragGroup;
    else if (source.dragGroup === null) result.dragGroup = null;
    if (source.viewMode === "table" || source.viewMode === "pivot") result.viewMode = source.viewMode;
    if (source.pivotConfig && typeof source.pivotConfig === "object") result.pivotConfig = { ...(source.pivotConfig as Record<string, unknown>) };
    if (Array.isArray(source.tier4Rules)) result.tier4Rules = source.tier4Rules.slice(0, 500);
    if (source.tier4ColumnColors && typeof source.tier4ColumnColors === "object") {
        result.tier4ColumnColors = Object.fromEntries(Object.entries(source.tier4ColumnColors as Record<string, unknown>).filter(([, v]) => typeof v === "string").map(([k, v]) => [k, String(v)]).slice(0, 500));
    }
    if (typeof source.tier4Palette === "string") result.tier4Palette = source.tier4Palette;
    if (Array.isArray(source.tier4CustomPalette)) result.tier4CustomPalette = source.tier4CustomPalette.filter((v): v is string => typeof v === "string").slice(0, 12);
    if (Array.isArray(source.tier4SavedThemes)) result.tier4SavedThemes = source.tier4SavedThemes.slice(0, 50);
    if (typeof source.tier4ActiveThemeId === "string") result.tier4ActiveThemeId = source.tier4ActiveThemeId;
    else if (source.tier4ActiveThemeId === null) result.tier4ActiveThemeId = null;
    if (Array.isArray(source.namedLayouts)) result.namedLayouts = source.namedLayouts.slice(0, 50);
    if (typeof source.announcementDismissed === "boolean") result.announcementDismissed = source.announcementDismissed;

    if (source.layout && typeof source.layout === "object") {
        const layout = source.layout as Record<string, unknown>;
        result.layout = {
            headerBold: typeof layout.headerBold === "boolean" ? layout.headerBold : undefined,
            fontSize: finiteNumber(layout.fontSize, 14, 10, 48),
            rowHeight: finiteNumber(layout.rowHeight, 36, 22, 120)
        };
    }
    if (source.settingsDrawerWidth !== undefined) {
        result.settingsDrawerWidth = finiteNumber(source.settingsDrawerWidth, 430, 280, 720);
    }

    return result;
}

export function serializePersistedUserConfig(state: Omit<PersistedUserConfigEnvelope, "schemaVersion">): string {
    return JSON.stringify({ ...state, schemaVersion: USER_CONFIG_SCHEMA_VERSION });
}

export function migrateSavedView(raw: unknown): unknown {
    if (!raw || typeof raw !== "object") return null;
    const source = raw as Record<string, unknown>;
    return {
        schemaVersion: typeof source.schemaVersion === "number" ? source.schemaVersion : SAVED_VIEW_SCHEMA_VERSION,
        sortColumn: typeof source.sortColumn === "string" ? source.sortColumn : null,
        sortDirection: source.sortDirection === "asc" || source.sortDirection === "desc" ? source.sortDirection : "none",
        columnOrder: Array.isArray(source.columnOrder) ? source.columnOrder.filter((v): v is string => typeof v === "string").slice(0, 500) : [],
        columnWidths: source.columnWidths && typeof source.columnWidths === "object"
            ? Object.fromEntries(Object.entries(source.columnWidths as Record<string, unknown>).filter(([, v]) => typeof v === "number" && Number.isFinite(v)).slice(0, 500))
            : {},
        hiddenColumns: Array.isArray(source.hiddenColumns) ? source.hiddenColumns.filter((v): v is string => typeof v === "string").slice(0, 500) : [],
        searchTerm: typeof source.searchTerm === "string" ? source.searchTerm.slice(0, 500) : "",
        groupExpansion: source.groupExpansion && typeof source.groupExpansion === "object"
            ? Object.fromEntries(Object.entries(source.groupExpansion as Record<string, unknown>).filter(([, v]) => typeof v === "boolean").slice(0, 1000))
            : {}
    };
}
