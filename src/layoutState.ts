"use strict";

export interface ILayoutState {
    fontSize: number;
    rowHeight: number;
    headerBold: boolean;
    responsive: boolean;
    autoFitColumns: boolean;
}

export type LayoutStatePatch = Partial<ILayoutState>;

export function clampLayoutNumber(value: unknown, min: number, max: number, fallback: number): number {
    const n = Number(value);
    if (!Number.isFinite(n)) return fallback;
    return Math.max(min, Math.min(max, n));
}

export function normalizeLayoutState(state: Partial<ILayoutState> | undefined, fallback: ILayoutState): ILayoutState {
    return {
        fontSize: clampLayoutNumber(state?.fontSize, 8, 48, fallback.fontSize),
        rowHeight: clampLayoutNumber(state?.rowHeight, 22, 120, fallback.rowHeight),
        headerBold: typeof state?.headerBold === "boolean" ? state.headerBold : fallback.headerBold,
        responsive: typeof state?.responsive === "boolean" ? state.responsive : fallback.responsive,
        autoFitColumns: typeof state?.autoFitColumns === "boolean" ? state.autoFitColumns : fallback.autoFitColumns,
    };
}

export function updateLayoutState(current: ILayoutState, patch: LayoutStatePatch): ILayoutState {
    return normalizeLayoutState({ ...current, ...patch }, current);
}
