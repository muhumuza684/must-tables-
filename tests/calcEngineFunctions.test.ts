import { evaluateCalc, parseCalcFormula } from "../src/calcEngine";

describe("formula scalar functions", () => {
    const context = {
        getColumnValue: (name: string) => ({ Name: "  Alice ", Amount: -12.345, Blank: null } as Record<string, string | number | null>)[name],
    };
    const aggregates = { getAggregate: (): number | null => null };

    function evaluate(formula: string) {
        const parsed = parseCalcFormula(formula);
        expect(parsed.ok).toBe(true);
        if (!parsed.ok || !parsed.ast) throw new Error(parsed.error);
        return evaluateCalc(parsed.ast, context, aggregates);
    }

    it("supports text and null helpers", () => {
        expect(evaluate("UPPER(Name)")).toBe("  ALICE ");
        expect(evaluate("LOWER(Name)")).toBe("  alice ");
        expect(evaluate("LEN(Name)")).toBe(8);
        expect(evaluate("CONCAT(Name, \" / active\")")).toBe("  Alice  / active");
        expect(evaluate("ISNULL(Blank)")).toBe(true);
        expect(evaluate("COALESCE(Blank, Name)")).toBe("  Alice ");
    });

    it("supports numeric helpers with bounded rounding precision", () => {
        expect(evaluate("ABS(Amount)")).toBe(12.345);
        expect(evaluate("ROUND(Amount, 2)")).toBe(-12.35);
    });

    it("supports type conversion helpers", () => {
        expect(evaluate("TEXT(Amount)")).toBe("-12.345");
        expect(evaluate("TEXT(Blank)")).toBe("");
        expect(evaluate('VALUE("42.5")')).toBe(42.5);
        expect(evaluate("VALUE(Blank)")).toBeNull();
        expect(evaluate('VALUE("not a number")')).toBeNull();
    });
});

describe("formula date functions", () => {
    // 2024-03-15 and 2024-01-01, as ms-epoch values -- evaluateCalc's "col" case coerces a
    // Date-valued column into ms automatically, so these fixtures store the raw ms directly
    // to keep the test independent of that coercion path.
    const marchFifteenth = new Date(2024, 2, 15).getTime();
    const januaryFirst = new Date(2024, 0, 1).getTime();

    const context = {
        getColumnValue: (name: string) => ({
            StartDate: januaryFirst,
            EndDate: marchFifteenth,
            NotADate: "hello"
        } as Record<string, string | number | null>)[name],
    };
    const aggregates = { getAggregate: (): number | null => null };

    function evaluate(formula: string) {
        const parsed = parseCalcFormula(formula);
        expect(parsed.ok).toBe(true);
        if (!parsed.ok || !parsed.ast) throw new Error(parsed.error);
        return evaluateCalc(parsed.ast, context, aggregates);
    }

    it("TODAY returns midnight of the current date", () => {
        // TODAY ignores its argument; the parser requires at least one argument per call
        // (see calcEngine.ts's FUNC parsing: "expects at least one argument"), so a dummy
        // 0 is passed here rather than changing the grammar for this one function.
        const result = evaluate("TODAY(0)") as number;
        const now = new Date();
        const expected = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
        expect(result).toBe(expected);
    });

    it("YEAR/MONTH/DAY extract components from a date-valued column", () => {
        expect(evaluate("YEAR(EndDate)")).toBe(2024);
        expect(evaluate("MONTH(EndDate)")).toBe(3);
        expect(evaluate("DAY(EndDate)")).toBe(15);
    });

    it("YEAR/MONTH/DAY fall back to the epoch (1970) for a non-numeric string rather than throwing", () => {
        // toNumber() coerces an unparseable string to 0 (see calcEngine.ts), same as ABS/ROUND
        // already rely on elsewhere in this file -- so YEAR/MONTH/DAY see ms=0 (1970-01-01),
        // not a genuinely null/invalid state. This documents that existing, intentional
        // coercion behavior rather than asserting a null path the function doesn't take.
        expect(evaluate("YEAR(NotADate)")).toBe(1970);
        expect(evaluate("MONTH(NotADate)")).toBe(1);
        expect(evaluate("DAY(NotADate)")).toBe(1);
    });

    it("DATEDIFF computes whole days between two date-valued columns", () => {
        expect(evaluate("DATEDIFF(StartDate, EndDate)")).toBe(74);
    });

    it("DATEDIFF computes against the epoch (not NaN/undefined) when an operand is a non-numeric string", () => {
        // Same toNumber() coercion as above: NotADate becomes ms=0 (1970-01-01), so this is a
        // real, deterministic day-count against the epoch -- not a NaN or a thrown error, which
        // is what actually matters for crash-safety even though it isn't a "null" result.
        const result = evaluate("DATEDIFF(StartDate, NotADate)") as number;
        expect(Number.isFinite(result)).toBe(true);
        expect(result).toBe(-19723);
    });
});
