import type { ITableRow } from "./tableRenderer";

export type AdvancedValue = string | number | boolean | Date | null | undefined;

export type FilterOperator =
    | "equals"
    | "notEquals"
    | "greaterThan"
    | "greaterThanOrEqual"
    | "lessThan"
    | "lessThanOrEqual"
    | "contains"
    | "startsWith"
    | "endsWith"
    | "between"
    | "in"
    | "isNull"
    | "isNotNull";

export type FilterExpression =
    | { kind: "comparison"; field: string; operator: FilterOperator; value?: AdvancedValue; secondValue?: AdvancedValue; values?: AdvancedValue[] }
    | { kind: "and" | "or"; left: FilterExpression; right: FilterExpression };

export interface GroupAggregate {
    count: number;
    sums: Record<string, number>;
    minimums: Record<string, number | null>;
    maximums: Record<string, number | null>;
}

export interface GroupResult {
    key: string;
    values: Record<string, AdvancedValue>;
    rows: ITableRow[];
    aggregate: GroupAggregate;
    children: GroupResult[];
}

export interface PivotResult {
    columns: string[];
    rows: Array<Record<string, AdvancedValue>>;
    rowTotals: Record<string, number>;
    columnTotals: Record<string, number>;
    grandTotal: number;
    /** Generated pivot cells with no source rows; renderer may display a custom empty token without changing numeric aggregation semantics. */
    emptyCells: Record<string, boolean>;
}

interface Token {
    type: "word" | "string" | "number" | "operator" | "lparen" | "rparen" | "comma" | "eof";
    text: string;
    position: number;
}

const RESERVED_WORDS = new Set(["AND", "OR", "CONTAINS", "STARTS", "ENDS", "WITH", "BETWEEN", "IN", "IS", "NOT", "NULL"]);

function tokenize(input: string): { tokens: Token[]; error?: string } {
    const tokens: Token[] = [];
    let index = 0;
    while (index < input.length) {
        const char = input[index];
        if (/\s/.test(char)) {
            index++;
            continue;
        }
        if (char === "(") {
            tokens.push({ type: "lparen", text: char, position: index++ });
            continue;
        }
        if (char === ")") {
            tokens.push({ type: "rparen", text: char, position: index++ });
            continue;
        }
        if (char === ",") {
            tokens.push({ type: "comma", text: char, position: index++ });
            continue;
        }
        if (char === "'" || char === '"') {
            const quote = char;
            const start = index++;
            let value = "";
            let closed = false;
            while (index < input.length) {
                if (input[index] === "\\" && index + 1 < input.length) {
                    value += input[index + 1];
                    index += 2;
                } else if (input[index] === quote) {
                    index++;
                    closed = true;
                    break;
                } else {
                    value += input[index++];
                }
            }
            if (!closed) {
                return { tokens, error: `Missing closing quote at position ${start}.` };
            }
            tokens.push({ type: "string", text: value, position: start });
            continue;
        }
        const two = input.slice(index, index + 2);
        if ([">=", "<=", "!=", "=="].includes(two)) {
            tokens.push({ type: "operator", text: two, position: index });
            index += 2;
            continue;
        }
        if ([">", "<", "="].includes(char)) {
            tokens.push({ type: "operator", text: char, position: index++ });
            continue;
        }
        if (/[0-9.-]/.test(char)) {
            const start = index;
            index++;
            while (index < input.length && /[0-9.eE+-]/.test(input[index])) {
                index++;
            }
            const text = input.slice(start, index);
            if (!Number.isFinite(Number(text))) {
                return { tokens, error: `Invalid number "${text}" at position ${start}.` };
            }
            tokens.push({ type: "number", text, position: start });
            continue;
        }
        if (/[A-Za-z_\[]/.test(char)) {
            const start = index;
            if (char === "[") {
                const close = input.indexOf("]", index + 1);
                if (close < 0) {
                    return { tokens, error: `Missing closing ] at position ${start}.` };
                }
                tokens.push({ type: "word", text: input.slice(index + 1, close).trim(), position: start });
                index = close + 1;
                continue;
            }
            index++;
            while (index < input.length && /[A-Za-z0-9_.]/.test(input[index])) {
                index++;
            }
            tokens.push({ type: "word", text: input.slice(start, index).trim(), position: start });
            continue;
        }
        return { tokens, error: `Unexpected character "${char}" at position ${index}.` };
    }
    tokens.push({ type: "eof", text: "", position: input.length });
    return { tokens };
}

class FilterParser {
    private index = 0;
    constructor(private readonly tokens: Token[]) {}

    private peek(): Token { return this.tokens[this.index]; }
    private next(): Token { return this.tokens[this.index++]; }
    private accept(type: Token["type"], text?: string): boolean {
        const token = this.peek();
        if (token.type === type && (text === undefined || token.text.toUpperCase() === text.toUpperCase())) {
            this.index++;
            return true;
        }
        return false;
    }
    private expect(type: Token["type"], text?: string): Token {
        const token = this.peek();
        if (!this.accept(type, text)) {
            throw new Error(`Expected ${text ?? type} near position ${token.position}.`);
        }
        return token;
    }

    parse(): FilterExpression {
        const expression = this.parseOr();
        if (this.peek().type !== "eof") {
            throw new Error(`Unexpected token "${this.peek().text}" near position ${this.peek().position}.`);
        }
        return expression;
    }

    private parseOr(): FilterExpression {
        let expression = this.parseAnd();
        while (this.accept("word", "OR")) {
            expression = { kind: "or", left: expression, right: this.parseAnd() };
        }
        return expression;
    }

    private parseAnd(): FilterExpression {
        let expression = this.parsePrimary();
        while (this.accept("word", "AND")) {
            expression = { kind: "and", left: expression, right: this.parsePrimary() };
        }
        return expression;
    }

    private parsePrimary(): FilterExpression {
        if (this.accept("lparen")) {
            const expression = this.parseOr();
            this.expect("rparen");
            return expression;
        }
        const field = this.expect("word").text;
        const operatorToken = this.next();
        if (operatorToken.type === "operator") {
            const operatorMap: Record<string, FilterOperator> = {
                "=": "equals", "==": "equals", "!=": "notEquals", ">": "greaterThan", ">=": "greaterThanOrEqual", "<": "lessThan", "<=": "lessThanOrEqual"
            };
            const operator = operatorMap[operatorToken.text];
            if (!operator) throw new Error(`Unsupported operator "${operatorToken.text}".`);
            return { kind: "comparison", field, operator, value: this.parseValue() };
        }
        if (operatorToken.type !== "word") {
            throw new Error(`Expected a filter operator near position ${operatorToken.position}.`);
        }
        const keyword = operatorToken.text.toUpperCase();
        if (keyword === "CONTAINS" || keyword === "STARTS" || keyword === "ENDS") {
            if (keyword !== "CONTAINS") this.expect("word", "WITH");
            return { kind: "comparison", field, operator: keyword === "CONTAINS" ? "contains" : keyword === "STARTS" ? "startsWith" : "endsWith", value: this.parseValue() };
        }
        if (keyword === "BETWEEN") {
            const first = this.parseValue();
            this.expect("word", "AND");
            return { kind: "comparison", field, operator: "between", value: first, secondValue: this.parseValue() };
        }
        if (keyword === "IN") {
            this.expect("lparen");
            const values: AdvancedValue[] = [];
            do { values.push(this.parseValue()); } while (this.accept("comma"));
            this.expect("rparen");
            return { kind: "comparison", field, operator: "in", values };
        }
        if (keyword === "IS") {
            const not = this.accept("word", "NOT");
            this.expect("word", "NULL");
            return { kind: "comparison", field, operator: not ? "isNotNull" : "isNull" };
        }
        throw new Error(`Unsupported filter keyword "${operatorToken.text}".`);
    }

    private parseValue(): AdvancedValue {
        const token = this.next();
        if (token.type === "number") return Number(token.text);
        if (token.type === "string") return token.text;
        if (token.type === "word") {
            const upper = token.text.toUpperCase();
            if (upper === "NULL") return null;
            if (upper === "TRUE") return true;
            if (upper === "FALSE") return false;
            return token.text;
        }
        throw new Error(`Expected a value near position ${token.position}.`);
    }
}

export function parseAdvancedFilter(input: string): { ok: true; expression: FilterExpression } | { ok: false; error: string } {
    const text = input.trim();
    if (!text) return { ok: false, error: "Enter a filter expression or leave it blank." };
    const tokenized = tokenize(text);
    if (tokenized.error) return { ok: false, error: tokenized.error };
    try {
        return { ok: true, expression: new FilterParser(tokenized.tokens).parse() };
    } catch (error) {
        return { ok: false, error: error instanceof Error ? error.message : "The filter expression could not be parsed." };
    }
}

function comparable(value: AdvancedValue): string | number | null {
    if (value === null || value === undefined) return null;
    if (value instanceof Date) return value.getTime();
    if (typeof value === "number" || typeof value === "string") return value;
    return String(value);
}

function matchesComparison(actual: AdvancedValue, expression: Extract<FilterExpression, { kind: "comparison" }>): boolean {
    const normalized = comparable(actual);
    if (expression.operator === "isNull") return normalized === null;
    if (expression.operator === "isNotNull") return normalized !== null;
    if (normalized === null) return false;
    const target = comparable(expression.value);
    const text = String(normalized).toLocaleLowerCase();
    const targetText = String(target ?? "").toLocaleLowerCase();
    switch (expression.operator) {
        case "equals": return typeof normalized === "number" && typeof target === "number" ? normalized === target : text === targetText;
        case "notEquals": return !matchesComparison(actual, { ...expression, operator: "equals" });
        case "greaterThan": return normalized > (target as typeof normalized);
        case "greaterThanOrEqual": return normalized >= (target as typeof normalized);
        case "lessThan": return normalized < (target as typeof normalized);
        case "lessThanOrEqual": return normalized <= (target as typeof normalized);
        case "contains": return text.includes(targetText);
        case "startsWith": return text.startsWith(targetText);
        case "endsWith": return text.endsWith(targetText);
        case "between": {
            const second = comparable(expression.secondValue);
            return second !== null && normalized >= (target as typeof normalized) && normalized <= (second as typeof normalized);
        }
        case "in": return (expression.values ?? []).some((value) => matchesComparison(actual, { ...expression, operator: "equals", value }));
        default: return false;
    }
}

export function evaluateFilter(expression: FilterExpression, row: ITableRow): boolean {
    if (expression.kind === "and") return evaluateFilter(expression.left, row) && evaluateFilter(expression.right, row);
    if (expression.kind === "or") return evaluateFilter(expression.left, row) || evaluateFilter(expression.right, row);
    if (expression.kind === "comparison") {
        return matchesComparison(row.values[expression.field] as AdvancedValue, expression);
    }
    return false;
}

export function applyAdvancedFilter(rows: ITableRow[], expression: FilterExpression | null): ITableRow[] {
    return expression ? rows.filter((row) => evaluateFilter(expression, row)) : rows.slice();
}

function numericValue(value: AdvancedValue): number | null {
    if (typeof value === "number" && Number.isFinite(value)) return value;
    if (typeof value === "string" && value.trim() !== "" && Number.isFinite(Number(value))) return Number(value);
    return null;
}

export function aggregateRows(rows: ITableRow[], valueColumns: string[]): GroupAggregate {
    const safeRows = Array.isArray(rows) ? rows : [];
    const safeColumns = Array.isArray(valueColumns) ? valueColumns : [];
    const sums: Record<string, number> = {};
    const minimums: Record<string, number | null> = {};
    const maximums: Record<string, number | null> = {};
    for (const column of safeColumns) {
        sums[column] = 0;
        minimums[column] = null;
        maximums[column] = null;
    }
    for (const row of safeRows) {
        if (!row || typeof row.values !== "object" || row.values === null) continue;
        for (const column of safeColumns) {
            const number = numericValue(row.values[column] as AdvancedValue);
            if (number === null) continue;
            sums[column] += number;
            minimums[column] = minimums[column] === null ? number : Math.min(minimums[column]!, number);
            maximums[column] = maximums[column] === null ? number : Math.max(maximums[column]!, number);
        }
    }
    return { count: safeRows.length, sums, minimums, maximums };
}

export function groupRows(rows: ITableRow[], groupColumns: string[], valueColumns: string[], depth = 0): GroupResult[] {
    const safeRows = Array.isArray(rows) ? rows : [];
    const safeGroupColumns = Array.isArray(groupColumns) ? groupColumns : [];
    const safeValueColumns = Array.isArray(valueColumns) ? valueColumns : [];
    if (depth >= safeGroupColumns.length) return [];
    const field = safeGroupColumns[depth];
    const buckets = new Map<string, ITableRow[]>();
    for (const row of safeRows) {
        if (!row || typeof row.values !== "object" || row.values === null) continue;
        const raw = row.values[field] as AdvancedValue;
        const key = raw === null || raw === undefined || raw === "" ? "(Blank)" : String(raw);
        const bucket = buckets.get(key) ?? [];
        bucket.push(row);
        buckets.set(key, bucket);
    }
    return Array.from(buckets.entries()).map(([key, bucket]) => ({
        key,
        values: { [field]: key },
        rows: bucket,
        aggregate: aggregateRows(bucket, safeValueColumns),
        children: groupRows(bucket, safeGroupColumns, safeValueColumns, depth + 1)
    }));
}

export type PivotAggregation = "sum" | "avg" | "min" | "max" | "count";

export function pivotRows(
    rows: ITableRow[],
    rowField: string,
    columnField: string,
    valueField: string,
    aggregation: PivotAggregation = "sum"
): PivotResult {
    const safeRows = Array.isArray(rows) ? rows : [];
    const generated = new Set<string>();
    const buckets = new Map<string, Map<string, number[]>>();
    for (const row of safeRows) {
        if (!row || typeof row.values !== "object" || row.values === null) continue;
        const rowKey = String(row.values[rowField] ?? "(Blank)");
        const columnKey = String(row.values[columnField] ?? "(Blank)");
        const rowBucket = buckets.get(rowKey) ?? new Map<string, number[]>();
        const values = rowBucket.get(columnKey) ?? [];
        const numeric = numericValue(row.values[valueField] as AdvancedValue);
        if (aggregation === "count") {
            values.push(1);
        } else if (numeric !== null && Number.isFinite(numeric)) {
            values.push(numeric);
        }
        rowBucket.set(columnKey, values);
        buckets.set(rowKey, rowBucket);
        generated.add(columnKey);
    }
    const aggregate = (values: number[]): number => {
        if (values.length === 0) return 0;
        switch (aggregation) {
            case "count": return values.length;
            case "avg": return values.reduce((a, b) => a + b, 0) / values.length;
            case "min": return Math.min(...values);
            case "max": return Math.max(...values);
            case "sum":
            default: return values.reduce((a, b) => a + b, 0);
        }
    };
    const generatedColumns = Array.from(generated).sort((a, b) => a.localeCompare(b));
    const suffix = aggregation === "sum" ? "" : ` (${aggregation === "count" ? "Count" : aggregation.toUpperCase()})`;
    const columns = [rowField, ...generatedColumns.map((column) => `${valueField} — ${column}${suffix}`)];
    const rowTotals: Record<string, number> = {};
    const columnTotals: Record<string, number> = {};
    const emptyCells: Record<string, boolean> = {};
    const allValues: number[] = [];
    const outputRows = Array.from(buckets.entries()).map(([rowKey, values]) => {
        const output: Record<string, AdvancedValue> = { [rowField]: rowKey };
        const rowValues: number[] = [];
        for (const column of generatedColumns) {
            const bucketValues = values.get(column) ?? [];
            const generatedName = `${valueField} — ${column}${suffix}`;
            output[generatedName] = aggregate(bucketValues);
            if (bucketValues.length === 0) emptyCells[`${rowKey}\u241F${generatedName}`] = true;
            rowValues.push(...bucketValues);
            allValues.push(...bucketValues);
        }
        rowTotals[rowKey] = aggregate(rowValues);
        return output;
    });
    for (const column of generatedColumns) {
        const raw: number[] = [];
        buckets.forEach((rowBucket) => raw.push(...(rowBucket.get(column) ?? [])));
        columnTotals[`${valueField} — ${column}${suffix}`] = aggregate(raw);
    }
    return { columns, rows: outputRows, rowTotals, columnTotals, grandTotal: aggregate(allValues), emptyCells };
}
