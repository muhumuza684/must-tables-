/**
 * Gate 1 (first slice): a versioned, typed command/reducer layer for pivot-mode state.
 * This is a proof-of-pattern migration -- most of tableRenderer.ts still manages state
 * as direct field mutation. Pivot mode was chosen as the first subsystem to migrate
 * because it is small, self-contained, and was built recently enough to be well understood.
 */

export type PivotAggregation = "sum" | "avg" | "min" | "max" | "count";

export interface PivotConfig {
    rowField: string;
    columnField: string;
    valueField: string;
    aggregation?: PivotAggregation;
}

export interface PivotState {
    viewMode: "table" | "pivot";
    pivotConfig: PivotConfig | null;
}


export function normalizePivotConfig(config: unknown): PivotConfig | null {
    if (!config || typeof config !== "object") return null;
    const candidate = config as Partial<PivotConfig> & { aggregation?: unknown };
    const validAggregation: PivotAggregation[] = ["sum", "avg", "min", "max", "count"];
    if (typeof candidate.rowField !== "string" || typeof candidate.columnField !== "string" || typeof candidate.valueField !== "string") return null;
    if (candidate.aggregation !== undefined && !validAggregation.includes(candidate.aggregation as PivotAggregation)) {
        return { rowField: candidate.rowField, columnField: candidate.columnField, valueField: candidate.valueField };
    }
    return { rowField: candidate.rowField, columnField: candidate.columnField, valueField: candidate.valueField, ...(candidate.aggregation !== undefined ? { aggregation: candidate.aggregation as PivotAggregation } : {}) };
}

export const initialPivotState: PivotState = {
    viewMode: "table",
    pivotConfig: null
};

export type PivotCommand =
    | { type: "SET_VIEW_MODE"; mode: "table" | "pivot" }
    | { type: "TOGGLE_VIEW_MODE" }
    | { type: "SET_PIVOT_CONFIG"; config: PivotConfig }
    | { type: "UPDATE_PIVOT_CONFIG"; partial: Partial<PivotConfig> }
    | { type: "RESET_PIVOT_CONFIG" }
    | { type: "HYDRATE"; state: Partial<PivotState> };

export function pivotReducer(state: PivotState, command: PivotCommand): PivotState {
    switch (command.type) {
        case "SET_VIEW_MODE":
            if (command.mode !== "table" && command.mode !== "pivot") return state;
            return { ...state, viewMode: command.mode };

        case "TOGGLE_VIEW_MODE":
            return { ...state, viewMode: state.viewMode === "pivot" ? "table" : "pivot" };

        case "SET_PIVOT_CONFIG": {
            const config = normalizePivotConfig(command.config);
            return config ? { ...state, pivotConfig: config } : state;
        }

        case "UPDATE_PIVOT_CONFIG": {
            const base: PivotConfig = state.pivotConfig ?? { rowField: "", columnField: "", valueField: "" };
            const next = { ...base, ...command.partial } as PivotConfig;
            return { ...state, pivotConfig: next };
        }

        case "RESET_PIVOT_CONFIG":
            return { ...state, pivotConfig: null };

        case "HYDRATE": {
            const next: PivotState = { ...state };
            if (command.state.viewMode === "table" || command.state.viewMode === "pivot") {
                next.viewMode = command.state.viewMode;
            }
            const config = normalizePivotConfig(command.state.pivotConfig);
            if (config) next.pivotConfig = config;
            return next;
        }

        default:
            return state;
    }
}