export type AnalyticsRuleOperator = "gt" | "lt" | "equals" | "between" | "contains" | "startsWith" | "endsWith" | "blank" | "notBlank";
export type AnalyticsFormatKind = "background" | "font" | "dataBar" | "icon";
export type AnalyticsRuleScope = "cell" | "row";
export type AnalyticsValueKind = "numeric" | "text" | "blank";

export interface IAnalyticsConditionalRule {
    column: string;
    operator: AnalyticsRuleOperator;
    value?: string;
    secondValue?: string;
    color: string;
    scope?: AnalyticsRuleScope;
    format?: AnalyticsFormatKind;
    icon?: "up" | "down" | "dot";
}

export interface IDataBarLayout {
    visible: boolean;
    leftPct: number;
    widthPct: number;
    negative: boolean;
}

export function inferAnalyticsValueKind(raw: unknown): AnalyticsValueKind {
    if (raw === null || raw === undefined || raw === "") return "blank";
    return typeof raw === "number" && Number.isFinite(raw) ? "numeric" : "text";
}

export function compatibleAnalyticsOperators(kind: AnalyticsValueKind): AnalyticsRuleOperator[] {
    if (kind === "numeric") return ["gt", "lt", "equals", "between", "blank", "notBlank"];
    if (kind === "blank") return ["blank", "notBlank"];
    return ["equals", "contains", "startsWith", "endsWith", "blank", "notBlank"];
}

export function matchesAnalyticsRule(raw: unknown, rule: IAnalyticsConditionalRule): boolean {
    const blank = raw === null || raw === undefined || raw === "";
    if (rule.operator === "blank") return blank;
    if (rule.operator === "notBlank") return !blank;
    if (blank) return false;

    const text = String(raw);
    const needle = String(rule.value ?? "");
    if (rule.operator === "contains") return text.toLocaleLowerCase().includes(needle.toLocaleLowerCase());
    if (rule.operator === "startsWith") return text.toLocaleLowerCase().startsWith(needle.toLocaleLowerCase());
    if (rule.operator === "endsWith") return text.toLocaleLowerCase().endsWith(needle.toLocaleLowerCase());

    const lhs = typeof raw === "number" ? raw : Number(raw);
    const rhs = Number(rule.value);
    const numeric = Number.isFinite(lhs) && Number.isFinite(rhs);
    if (rule.operator === "equals") return numeric ? lhs === rhs : text === needle;
    if (!numeric) return false;
    if (rule.operator === "gt") return lhs > rhs;
    if (rule.operator === "lt") return lhs < rhs;
    if (rule.operator === "between") {
        const rhs2 = Number(rule.secondValue);
        if (!Number.isFinite(rhs2)) return false;
        const low = Math.min(rhs, rhs2);
        const high = Math.max(rhs, rhs2);
        return lhs >= low && lhs <= high;
    }
    return false;
}

export function normalizedScalePosition(value: number, min: number, max: number): number | null {
    if (![value, min, max].every(Number.isFinite) || max <= min) return null;
    return Math.max(0, Math.min(1, (value - min) / (max - min)));
}

export function dataBarLayout(value: number, min: number, max: number): IDataBarLayout {
    if (![value, min, max].every(Number.isFinite) || max <= min) return { visible: false, leftPct: 0, widthPct: 0, negative: false };
    if (min >= 0) return { visible: true, leftPct: 0, widthPct: normalizedScalePosition(value, min, max)! * 100, negative: false };
    if (max <= 0) return { visible: true, leftPct: normalizedScalePosition(value, min, max)! * 100, widthPct: (1 - normalizedScalePosition(value, min, max)!) * 100, negative: true };
    const zero = (-min / (max - min)) * 100;
    if (value >= 0) return { visible: true, leftPct: zero, widthPct: Math.max(0, normalizedScalePosition(value, min, max)! * 100 - zero), negative: false };
    const pos = normalizedScalePosition(value, min, max)! * 100;
    return { visible: true, leftPct: pos, widthPct: Math.max(0, zero - pos), negative: true };
}
