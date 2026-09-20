"use strict";

import { DataLakeDensity } from "./styleSystem";

export type ResponsiveLayoutMode = "compact" | "standard" | "wide";

export interface IResponsiveLayout {
    width: number;
    height: number;
    mode: ResponsiveLayoutMode;
    compact: boolean;
    narrow: boolean;
    toolbarMode: "full" | "condensed";
    contentPadding: number;
    rowHeight: number;
    headerHeight: number;
    drawerMaxWidth: number;
}

export function resolveResponsiveLayout(width: number, height: number, density: DataLakeDensity, rowHeight: number, headerHeight: number): IResponsiveLayout {
    const w = Math.max(0, Number.isFinite(width) ? width : 0);
    const h = Math.max(0, Number.isFinite(height) ? height : 0);
    const compact = w < 420 || h < 240;
    const narrow = w <= 340;
    const mode: ResponsiveLayoutMode = w >= 1000 ? "wide" : compact ? "compact" : "standard";
    const densityScale = density === "compact" ? 0.9 : density === "spacious" ? 1.1 : 1;
    const minRow = compact ? 26 : 28;
    const resolvedRow = Math.max(minRow, Math.round(rowHeight * densityScale));
    const resolvedHeader = Math.max(compact ? 32 : 36, Math.round(headerHeight * densityScale));
    const contentPadding = narrow ? 6 : compact ? 8 : mode === "wide" ? 12 : 10;
    const drawerMaxWidth = narrow ? Math.max(280, w - 12) : Math.min(920, Math.max(360, Math.round(w * 0.78)));
    return {
        width: w,
        height: h,
        mode,
        compact,
        narrow,
        toolbarMode: compact ? "condensed" : "full",
        contentPadding,
        rowHeight: resolvedRow,
        headerHeight: resolvedHeader,
        drawerMaxWidth
    };
}

export function autoFitColumnWidth(columnCount: number, availableWidth: number, minWidth: number, maxWidth: number): number {
    if (columnCount <= 0 || availableWidth <= 0) return minWidth;
    return Math.max(minWidth, Math.min(maxWidth, Math.floor(availableWidth / columnCount)));
}
