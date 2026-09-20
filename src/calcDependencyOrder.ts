/**
 * Topological ordering and cycle detection for calculated columns that reference other
 * calculated columns. Built from the referencedColumns list calcEngine.ts's parser already
 * produces for every formula (see ICalcParseResult.referencedColumns).
 *
 * This is a standalone, pure utility -- it does not touch tableRenderer.ts's evaluation loop.
 * Wiring it into recomputeVirtualColumns() is a separate follow-up: that loop's current
 * behavior (including its error-surfacing for parse failures) needs a careful read before
 * changing evaluation order, the same way the Gate 1 reducer migration was done one
 * subsystem at a time rather than all at once.
 */

export interface CalcColumnLike {
    referencedColumns: string[];
}

export interface DependencyOrderResult {
    /** Calculated-column names in a valid evaluation order (dependencies before dependents). */
    order: string[];
    /** Calculated-column names that are part of a dependency cycle and cannot be safely ordered. */
    cyclic: Set<string>;
}

/**
 * Orders calculated columns so that any column referencing another calculated column is
 * evaluated after the column it depends on. References to raw data columns (anything not
 * itself a key in `columns`) are ignored for ordering purposes, since raw values are already
 * present on each row regardless of calculated-column evaluation order.
 */
export function orderCalcColumns(columns: Map<string, CalcColumnLike>): DependencyOrderResult {
    const names = Array.from(columns.keys());
    const calcNameSet = new Set(names);

    const graph = new Map<string, string[]>();
    for (const name of names) {
        const def = columns.get(name)!;
        const deps = (def.referencedColumns ?? []).filter((ref) => calcNameSet.has(ref) && ref !== name);
        graph.set(name, deps);
    }

    const WHITE = 0, GRAY = 1, BLACK = 2;
    const color = new Map<string, number>(names.map((n) => [n, WHITE]));
    const order: string[] = [];
    const cyclic = new Set<string>();

    function visit(name: string, stack: string[]): void {
        color.set(name, GRAY);
        stack.push(name);
        for (const dep of graph.get(name) ?? []) {
            const depColor = color.get(dep);
            if (depColor === GRAY) {
                // Found a cycle: everything currently on the stack from dep onward is cyclic.
                const cycleStart = stack.indexOf(dep);
                const cycleMembers = cycleStart >= 0 ? stack.slice(cycleStart) : [dep, name];
                cycleMembers.forEach((m) => cyclic.add(m));
                continue;
            }
            if (depColor === WHITE) {
                visit(dep, stack);
            }
        }
        stack.pop();
        color.set(name, BLACK);
        if (!cyclic.has(name)) {
            order.push(name);
        }
    }

    for (const name of names) {
        if (color.get(name) === WHITE) {
            visit(name, []);
        }
    }

    return { order, cyclic };
}