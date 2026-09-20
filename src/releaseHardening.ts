"use strict";

export interface IRenderBudget { rows: number; columns: number; elapsedMs: number; }
export const DEFAULT_RENDER_BUDGET_MS = 500;

export function withinRenderBudget(elapsedMs: number, budgetMs = DEFAULT_RENDER_BUDGET_MS): boolean {
    return Number.isFinite(elapsedMs) && elapsedMs >= 0 && elapsedMs <= budgetMs;
}

export function clampViewport(width: number, height: number): { width: number; height: number } {
    return {
        width: Math.max(0, Number.isFinite(width) ? width : 0),
        height: Math.max(0, Number.isFinite(height) ? height : 0),
    };
}

export function shouldPreferCompactLayout(width: number, height: number): boolean {
    return width < 420 || height < 240;
}
