export interface IFormatColumn { name: string; isMeasure: boolean; formatString?: string; }

export type NumberAlignment = "auto" | "left" | "center" | "right";
export type PerformanceMode = "auto" | "balanced" | "maximum";

export interface IRegionalFormatSettings {
    locale: string;
    currency: string;
    dateFormat: string;
    numberSeparators: "auto" | "locale";
    decimalPrecision: number;
}

export function resolveCellAlignment(column: IFormatColumn, alignment: NumberAlignment): "left" | "center" | "right" {
    if (alignment !== "auto") return alignment;
    if (column.isMeasure) return "right";
    return "left";
}

function modelLooksPercent(format?: string): boolean {
    return !!format && format.indexOf("%") >= 0;
}

function modelLooksCurrency(format?: string): boolean {
    if (!format) return false;
    return /[$€£¥₹]|\b(USD|EUR|GBP|JPY|INR|UGX|KES|TZS|RWF|MWK)\b/i.test(format);
}

function modelDecimalPlaces(format?: string): number | null {
    if (!format) return null;
    const firstSection = format.split(";")[0];
    const match = firstSection.match(/\.([0#]+)/);
    return match ? match[1].length : null;
}

function applyModelNumberFormat(value: number, format: string | undefined, locale: string): string {
    if (!format || /^(General|@)$/i.test(format.trim())) return new Intl.NumberFormat(locale).format(value);
    const percent = modelLooksPercent(format);
    const currency = modelLooksCurrency(format);
    const decimals = modelDecimalPlaces(format);
    const options: Intl.NumberFormatOptions = {};
    if (percent) options.style = "percent";
    if (currency) {
        options.style = "currency";
        const codeMatch = format.match(/\b(USD|EUR|GBP|JPY|INR|UGX|KES|TZS|RWF|MWK)\b/i);
        const symbolMap: Record<string, string> = { "$": "USD", "€": "EUR", "£": "GBP", "¥": "JPY", "₹": "INR" };
        const symbol = Object.keys(symbolMap).find((candidate) => format.indexOf(candidate) >= 0);
        options.currency = codeMatch ? codeMatch[1].toUpperCase() : (symbol ? symbolMap[symbol] : undefined);
    }
    if (decimals !== null) {
        options.minimumFractionDigits = decimals;
        options.maximumFractionDigits = decimals;
    }
    try {
        return new Intl.NumberFormat(locale, options).format(percent ? value : value);
    } catch {
        return String(value);
    }
}

function applyDateFormat(value: Date, format: string | undefined, locale: string): string {
    if (!format) return new Intl.DateTimeFormat(locale).format(value);
    const lower = format.toLowerCase();
    const options: Intl.DateTimeFormatOptions = {};
    if (lower.includes("h") || lower.includes("s")) {
        options.dateStyle = "short";
        options.timeStyle = lower.includes("s") ? "medium" : "short";
    } else if (lower.includes("y") && lower.includes("m") && lower.includes("d")) {
        options.dateStyle = "medium";
    } else {
        return new Intl.DateTimeFormat(locale).format(value);
    }
    try { return new Intl.DateTimeFormat(locale, options).format(value); } catch { return value.toLocaleString(locale); }
}

export function formatDataValue(value: unknown, column: IFormatColumn, regional: IRegionalFormatSettings): string {
    if (value === null || value === undefined || value === "") return "";
    const locale = regional.locale || "en-US";
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        return applyDateFormat(value, regional.dateFormat || column.formatString, locale);
    }
    if (typeof value === "number" && Number.isFinite(value)) {
        const modelFormat = column.formatString;
        if (regional.numberSeparators === "auto" && regional.decimalPrecision < 0 && !regional.currency && !regional.dateFormat) {
            return applyModelNumberFormat(value, modelFormat, locale);
        }
        const percent = modelLooksPercent(modelFormat);
        const currency = modelLooksCurrency(modelFormat) ? undefined : (regional.currency || undefined);
        const decimals = regional.decimalPrecision >= 0 ? regional.decimalPrecision : modelDecimalPlaces(modelFormat);
        const options: Intl.NumberFormatOptions = {};
        if (percent) options.style = "percent";
        if (currency && !percent) { options.style = "currency"; options.currency = currency; }
        if (decimals !== null && decimals !== undefined) {
            options.minimumFractionDigits = decimals;
            options.maximumFractionDigits = decimals;
        }
        try { return new Intl.NumberFormat(locale, options).format(value); } catch { return String(value); }
    }
    return String(value);
}

export function formatNumberForColumn(value: number, column: IFormatColumn, regional: IRegionalFormatSettings): string {
    return formatDataValue(value, column, regional);
}
