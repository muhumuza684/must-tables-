// MUST Tables -- standalone error-handling utilities (Tier 5).
//
// This file is intentionally self-contained and not yet imported anywhere. Wiring it into
// tableRenderer.ts / visual.ts needs exact current call-site text (same reason every other
// patch in this project needs a real dump first) -- this is the safe, zero-risk half: the
// utilities exist and are unit-testable on their own before anything depends on them.

/** Parses a value as a finite number, or returns the fallback instead of NaN/undefined
 *  leaking into calculations, sorting, or formatting. Never throws. */
export function safeNumber(value: unknown, fallback = 0): number {
    if (typeof value === "number" && Number.isFinite(value)) return value;
    if (value === null || value === undefined || value === "") return fallback;
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : fallback;
}

/** Parses a value as a valid Date, or returns null instead of an "Invalid Date" object
 *  that silently breaks downstream formatting/sorting. Never throws. */
export function safeDate(value: unknown): Date | null {
    if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
    if (typeof value !== "string" && typeof value !== "number") return null;
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
}

/** Wraps a render step so one bad row/value/rule can't blank the whole table. Logs a
 *  diagnostic (dev tools console only -- never shown to the report viewer) and returns the
 *  fallback instead of letting the exception propagate up through Power BI's own render loop,
 *  which is what causes a visual to go permanently blank with no explanation. */
export function safeRender<T>(context: string, fn: () => T, fallback: T): T {
    try {
        return fn();
    } catch (error) {
        logDiagnostic(context, error);
        return fallback;
    }
}

/** Single place all caught errors funnel through, so the format is consistent and it's easy
 *  to find every catch site later with one search. Never throws itself. */
export function logDiagnostic(context: string, error: unknown): void {
    try {
        console.warn(`MUST Tables: recovered from an error in ${context}.`, error);
    } catch {
        // If even logging fails, do nothing -- never let error handling itself throw.
    }
}

/** Truncates a string for display so an unexpectedly huge cell value can't blow out layout
 *  or freeze rendering. Returns the original value untouched if it's already short enough. */
export function safeTruncate(value: string, maxLength = 500): string {
    if (typeof value !== "string") return "";
    return value.length > maxLength ? `${value.slice(0, maxLength)}...` : value;
}