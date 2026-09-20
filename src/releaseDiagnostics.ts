"use strict";

export interface ReleaseSizeLimits { maxColumns: number; maxRows: number; }

export function isReasonableTableShape(rows: number, columns: number, limits: ReleaseSizeLimits = { maxRows: 30000, maxColumns: 200 }): boolean {
    return Number.isFinite(rows) && Number.isFinite(columns) && rows >= 0 && columns >= 0 && rows <= limits.maxRows && columns <= limits.maxColumns;
}

export function hasMojibakeMarkers(text: string): boolean {
    return /[\u00C3\u00C2\u00E2\u00EF\uFFFD]/u.test(text);
}
