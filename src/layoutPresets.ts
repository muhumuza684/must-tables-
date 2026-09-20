export interface LayoutSnapshot {
    sortColumn: string | null;
    sortDirection: "asc" | "desc" | "none";
    columnOrder: string[];
    columnWidths: Record<string, number>;
    hiddenColumns: string[];
    searchTerm: string;
    groupExpansion: Record<string, boolean>;
    fontSize: number;
    rowHeight: number;
    headerBold: boolean;
}

export interface NamedLayout {
    id: string;
    name: string;
    updatedAt: string;
    state: LayoutSnapshot;
}

export function normalizeLayoutName(name: string): string {
    return name.trim().replace(/\s+/g, " ").slice(0, 80);
}

export function createLayoutId(name: string, now = new Date()): string {
    const slug = normalizeLayoutName(name).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "layout";
    return `${slug}-${now.getTime().toString(36)}`;
}

export function upsertNamedLayout(layouts: NamedLayout[], layout: NamedLayout): NamedLayout[] {
    const normalized: NamedLayout = { ...layout, name: normalizeLayoutName(layout.name) };
    const withoutSameId = layouts.filter((item) => item.id !== normalized.id);
    return [...withoutSameId, normalized].sort((a, b) => a.name.localeCompare(b.name));
}

export function removeNamedLayout(layouts: NamedLayout[], id: string): NamedLayout[] {
    return layouts.filter((item) => item.id !== id);
}

export function parseNamedLayouts(value: unknown): NamedLayout[] {
    if (!Array.isArray(value)) return [];
    const output: NamedLayout[] = [];
    for (const item of value) {
        if (!item || typeof item !== "object") continue;
        const candidate = item as Partial<NamedLayout>;
        if (typeof candidate.id !== "string" || typeof candidate.name !== "string" || !candidate.state || typeof candidate.state !== "object") continue;
        const state = candidate.state as Partial<LayoutSnapshot>;
        if (!Array.isArray(state.columnOrder) || !Array.isArray(state.hiddenColumns) || typeof state.searchTerm !== "string") continue;
        output.push({
            id: candidate.id.slice(0, 120),
            name: normalizeLayoutName(candidate.name),
            updatedAt: typeof candidate.updatedAt === "string" ? candidate.updatedAt : new Date(0).toISOString(),
            state: {
                sortColumn: typeof state.sortColumn === "string" ? state.sortColumn : null,
                sortDirection: state.sortDirection === "asc" || state.sortDirection === "desc" ? state.sortDirection : "none",
                columnOrder: state.columnOrder.filter((x): x is string => typeof x === "string").slice(0, 500),
                columnWidths: state.columnWidths && typeof state.columnWidths === "object" ? Object.fromEntries(Object.entries(state.columnWidths).filter(([, width]) => typeof width === "number" && Number.isFinite(width)).slice(0, 500)) : {},
                hiddenColumns: state.hiddenColumns.filter((x): x is string => typeof x === "string").slice(0, 500),
                searchTerm: state.searchTerm.slice(0, 500),
                groupExpansion: state.groupExpansion && typeof state.groupExpansion === "object" ? Object.fromEntries(Object.entries(state.groupExpansion).filter(([, expanded]) => typeof expanded === "boolean").slice(0, 1000)) as Record<string, boolean> : {},
                fontSize: typeof state.fontSize === "number" ? Math.max(10, Math.min(24, state.fontSize)) : 14,
                rowHeight: typeof state.rowHeight === "number" ? Math.max(28, Math.min(72, state.rowHeight)) : 36,
                headerBold: typeof state.headerBold === "boolean" ? state.headerBold : true
            }
        });
    }
    return output.slice(0, 50);
}
