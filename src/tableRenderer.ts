"use strict";

import * as d3 from "d3";
import ExcelJS from "exceljs";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import powerbi from "powerbi-visuals-api";
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import ISelectionId = powerbi.visuals.ISelectionId;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ITooltipService = powerbi.extensibility.ITooltipService;
import VisualTooltipDataItem = powerbi.extensibility.VisualTooltipDataItem;
import ILocalizationManager = powerbi.extensibility.ILocalizationManager;
import ISandboxExtendedColorPalette = powerbi.extensibility.ISandboxExtendedColorPalette;

import {
    parseCalcFormula,
    evaluateCalc,
    ICalcParseResult,
    ICalcRowContext,
    ICalcAggregates,
    CalcValue
} from "./calcEngine";
import { selectColumnsForWidth } from "./mobileLayout";
import {
    colorForTier4Value,
    deriveGradientFromPalette,
    ITier4ConditionalRule,
    ITier4SavedTheme,
    matchesTier4Rule,
    normalizeHex,
    parseCustomPalette,
    Tier4PaletteName,
    TIER4_PALETTE_SWATCHES,
    safeTier4ThemeList
} from "./tier4Formatting";
import { buildWatermarkText, createExportAuditEvent, recordExportAudit } from "./tier4Governance";
import { applyAdvancedFilter, parseAdvancedFilter, aggregateRows, pivotRows, GroupAggregate, PivotResult } from "./advancedModel";
import { PivotState, PivotCommand, PivotAggregation, initialPivotState, pivotReducer } from "./state/pivotState";
import { createLayoutId, normalizeLayoutName, parseNamedLayouts, removeNamedLayout, upsertNamedLayout, NamedLayout, LayoutSnapshot } from "./layoutPresets";
import { updateLayoutState, ILayoutState } from "./layoutState";
import { DataLakeEventBus, DataLakeEventListener } from "./customEvents";
import { orderCalcColumns } from "./calcDependencyOrder";
import { safeNumber, safeTruncate, logDiagnostic } from "./errorHandling";
import { compatibleAnalyticsOperators, dataBarLayout, inferAnalyticsValueKind } from "./analyticsFormatting";
import { formatDataValue, resolveCellAlignment, IRegionalFormatSettings, PerformanceMode } from "./dataFormatting";

/** A single logical column: a plain row dimension, a value measure, a tooltip-only field, or a user-defined virtual column. */
export interface ITableColumn {
    name: string;
    displayName: string;
    formatString?: string;
    isMeasure: boolean;
    isGroupBy: boolean;
    /** True for a user-defined "Calculations" column (Item 1) — behaves like any other measure column. */
    isCalculated?: boolean;
    /** True for a user-defined "Combine columns" column (Item 4) — behaves like a text row column. */
    isCombined?: boolean;
}

interface ICalcColumnDef {
    formula: string;
    parsed: ICalcParseResult;
}

interface ICombinedColumnDef {
    template: string;
}

/** A single flattened data row, keyed by column name. */
export interface ITableRow {
    key: string;
    values: { [columnName: string]: powerbi.PrimitiveValue };
    selectionId: ISelectionId;
}

export interface ITableRendererSettings {
    fontFamily: string;
    fontSize: number;
    fontColor?: string;
    rowHeight: number;
    headerBg: string;
    headerFont: string;
    headerFontFamily?: string;
    headerFontSize?: number;
    headerFontWeight?: string;
    headerAlignment?: "left" | "center" | "right";
    headerWrap?: boolean;
    headerBold: boolean;
    cellBg: string;
    cellFont: string;
    cellFontSize?: number;
    cellFontWeight?: string;
    cellAlignment?: "auto" | "left" | "center" | "right";
    altRow: string;
    enableDataBars: boolean;
    barColor: string;
    negativeBarColor?: string;
    tableSurface?: { grid: "none" | "horizontal" | "vertical" | "both"; headerBorder: string; rowDivider: string; hoverBackground: string; selectedBackground: string };
    showTotals: boolean;
    totalsLabel: string;
    totalsBg: string;
    totalsFont?: string;
    totalsBold?: boolean;
    totalsShowBorder?: boolean;
    totalsAlignment?: "left" | "center" | "right";
    virtualScrollEnabled: boolean;
    virtualScrollRowHeight: number;
    showToolbar: boolean;
    searchEnabled: boolean;
    showSearchClear?: boolean;
    showFilterClear?: boolean;
    allowMultiColumnFiltering?: boolean;
    showFilterIndicator?: boolean;
    showFilteredState?: boolean;
    enableSorting?: boolean;
    showSortIndicator?: boolean;
    enableRowSelection?: boolean;
    selectionMode?: "single" | "multi";
    showClearSelection?: boolean;
    showSelectionHighlight?: boolean;
    enableCopy?: boolean;
    copyCell?: boolean;
    copyRow?: boolean;
    copySelected?: boolean;
    enableColumnFilters: boolean;
    /** Optional author-defined boolean filter expression applied before search and column filters. */
    advancedFilterExpression?: string;
    conditionalFormatEnabled: boolean;
    conditionalFormatMinColor: string;
    conditionalFormatMidpointEnabled?: boolean;
    conditionalFormatMidpointColor?: string;
    conditionalFormatMaxColor: string;
    groupsDefaultExpanded: boolean;
    groupIndentation?: number;
    showGroupCount?: boolean;
    showGroupTotals?: boolean;
    pivotEnabled?: boolean;
    pivotShowRowTotals?: boolean;
    pivotShowColumnTotals?: boolean;
    pivotShowGrandTotal?: boolean;
    pivotEmptyValueDisplay?: string;
    /**
     * Resolved value of the optional "Permissions" data role for the current viewer
     * (e.g. "no-export", "read-only"), or null when the role is left unbound. This is
     * a UI/workflow control enforced client-side in the visual -- it only removes
     * controls from this visual's own toolbar/header. It is a complement to, not a
     * replacement for, the report's actual Row-Level Security (RLS) configuration at
     * the dataset level, which governs the underlying data access itself.
     */
    permission: string | null;
    /** Parsed, validated conditional URL-action rules (Item 9). Empty when unset or malformed. */
    linkActionRules: ILinkActionRule[];
    /** Exact display name of the column that should show the link-action icon. */
    linkActionIconColumn: string;
    /** The report's persisted default view (Item 7), read back from `savedView` object properties, or null if none has been saved yet. */
    savedViewState: ISavedViewState | null;
    /** Allow Interactions compliance (item 10): false in read-only/embedded host contexts. */
    allowInteractions: boolean;
    /** Fetch More Data (A1-A4): true while Power BI still has more row segments beyond what's
     *  currently loaded. Read from dataView.metadata.segment by visual.ts. */
    hasMoreData: boolean;
    exportSettings?: { enabled: boolean; rowScope: "visible" | "filtered" | "selected" | "available"; includeHeaders: boolean; includeTotals: boolean; csv: boolean; excel: boolean; json: boolean; pdf: boolean };
    governance?: { watermarkEnabled: boolean; watermarkText: string; watermarkPlacement: "footer" | "center"; auditEnabled: boolean; };
    regionalFormat?: IRegionalFormatSettings;
    performanceMode?: PerformanceMode;
    density?: import("./styleSystem").DataLakeDensity;
    stylePreset?: import("./styleSystem").DataLakeStylePreset;
    styleTokens?: import("./styleSystem").IDataLakeStyleTokens;
    responsiveLayout?: boolean;
    autoFitColumns?: boolean;
    minColumnWidth?: number;
    maxColumnWidth?: number;
    cellWrap?: boolean;
    textOverflow?: string;
}

type SortDirection = "asc" | "desc" | "none";

interface ISortState {
    column: string | null;
    direction: SortDirection;
}

/**
 * Serializable shape of "the report's default view" (Item 7). Written to the
 * report's own object model via `host.persistProperties` under the `savedView`
 * object -- no external backend, no browser storage -- and read back on the next
 * `update()` so every viewer who opens the report sees the same standard view.
 */
export interface ISavedViewState {
    sortColumn: string | null;
    sortDirection: SortDirection;
    columnOrder: string[];
    columnWidths: { [columnName: string]: number };
    hiddenColumns: string[];
    searchTerm: string;
    groupExpansion: { [groupPath: string]: boolean };
}

/** One conditional URL-action rule (Item 9), as authored in the `linkActions.rules` JSON array. */
export type LinkActionOperator = "equals" | "notEquals" | "gt" | "gte" | "lt" | "lte" | "contains";

export interface ILinkActionRule {
    column: string;
    operator: LinkActionOperator;
    value: string;
    urlTemplate: string;
}

type FilterType = "text" | "number" | "date";
type FilterOperator = "contains" | "equals" | "gt" | "gte" | "lt" | "lte" | "between";

interface IColumnFilter {
    type: FilterType;
    operator: FilterOperator;
    value: string;
    value2?: string;
}

/** One row of the (post-filter, post-sort) flattened render list: a group header, a leaf data row, or an expanded record-detail sub-grid. */
type RenderNode =
    | { kind: "group"; depth: number; path: string; column: ITableColumn; value: powerbi.PrimitiveValue; count: number; sums: Map<string, number>; minimums: Map<string, number | null>; maximums: Map<string, number | null> }
    | { kind: "row"; depth: number; row: ITableRow }
    | { kind: "detail"; depth: number; row: ITableRow };

const DEFAULT_ROW_BUFFER = 6;
const BALANCED_ROW_BUFFER = 4;
const MAXIMUM_ROW_BUFFER = 2; // extra rows rendered above/below viewport to avoid flicker while scrolling

/** Short random id suffix for a saved theme. Uses crypto.getRandomValues rather than Math.random --
 *  Power BI's certification linter (powerbi-visuals/insecure-random) flags Math.random as an error. */
function generateThemeIdSuffix(): string {
    const bytes = new Uint8Array(6);
    (window.crypto ?? (globalThis as any).crypto).getRandomValues(bytes);
    return Array.from(bytes, (b) => b.toString(36)).join("").slice(0, 8);
}

/**
 * Data Lake Tables palettes. Every fixed preset's [min, max] gradient is derived (light -> dark)
 * from its full swatch list via `deriveGradientFromPalette`, so the data-bar/conditional-format
 * gradient and the categorical swatches always agree. "custom" is populated at runtime from
 * whatever a user pastes into the Custom palette box.
 */
// Tier C: trimmed to presets that earn their place -- Standard (baseline), the two
// accessibility palettes (compliance, not optional), Brand (identity), URA (a real named
// client), and High contrast (accessibility). Removed "Ocean blue" and "Teal operations" --
// decorative-only presets a typical client was never going to reach for.
const DLT_PALETTE_PRESETS: Array<{ name: Tier4PaletteName | "highContrast"; label: string; swatches: string[] }> = [
    { name: "default", label: "Standard", swatches: TIER4_PALETTE_SWATCHES.default },
    { name: "deuteranopia", label: "Deuteranopia-friendly", swatches: TIER4_PALETTE_SWATCHES.deuteranopia },
    { name: "protanopia", label: "Protanopia-friendly", swatches: TIER4_PALETTE_SWATCHES.protanopia },
    { name: "brand", label: "Brand palette", swatches: TIER4_PALETTE_SWATCHES.brand },
    { name: "highContrast", label: "High contrast", swatches: ["#FFFFFF", "#000000"] }
];
const DLT_PALETTE_NAMES = DLT_PALETTE_PRESETS.map((preset) => preset.name);
/** The full set of selectable palette identifiers, including the runtime-populated "custom" slot. */
type DltPaletteName = (typeof DLT_PALETTE_PRESETS)[number]["name"] | "custom";
const GROUP_SEP = "\u241F"; // unit separator — safe delimiter for building unique group path keys

/**
 * TableRenderer owns everything that happens inside the scrollable table
 * surface: virtualization, multi-level grouping/drill-down, sorting,
 * per-column + global search filtering, column resize/reorder, cross-filter
 * selection, data bars, conditional (value-based) formatting, smart
 * tooltips, CSV/Excel/PDF export, and the accessibility/certification surface
 * (allow-interactions, theme colors, context menu, high contrast, keyboard
 * navigation, landing page, localization, rendering events hook-in, and
 * multi-visual selection sync).
 */
export class TableRenderer {
    private container: HTMLDivElement;
    private host: IVisualHost;
    private selectionManager: ISelectionManager;
    private tooltipService: ITooltipService;
    private localizationManager: ILocalizationManager;
    private colorPalette: ISandboxExtendedColorPalette;

    private settings!: ITableRendererSettings;
    private columns: ITableColumn[] = [];
    private rowColumns: ITableColumn[] = [];
    private groupColumns: ITableColumn[] = [];
    private valueColumns: ITableColumn[] = [];
    private tooltipColumns: ITableColumn[] = [];
    /** Separate view mode: table (default, grouped/flat rows) or pivot (row x column x value matrix). Not merged into the row tree. */
    private _pivotState: PivotState = initialPivotState;
    private dispatchPivot(command: PivotCommand): void {
        this._pivotState = pivotReducer(this._pivotState, command);
    }

    private _data: ITableRow[] = [];
    private _filteredData: ITableRow[] = [];
    private _renderNodes: RenderNode[] = [];
    private _groupExpansion: Map<string, boolean> = new Map();
    private _sortState: ISortState = { column: null, direction: "none" };
    private _searchTerm: string = "";
    private _columnWidths: Map<string, number> = new Map();
    private _columnWidthModes: Map<string, "auto" | "fit" | "fill" | "fixed" | "custom"> = new Map();
    private _columnMinWidths: Map<string, number> = new Map();
    private _columnMaxWidths: Map<string, number> = new Map();
    private _columnAlignments: Map<string, "auto" | "left" | "center" | "right"> = new Map();
    private _nonSortableColumns: Set<string> = new Set();
    private _nonFilterableColumns: Set<string> = new Set();
    private _hiddenColumns: Set<string> = new Set();
    private _columnOrder: string[] = [];
    private _columnFilters: Map<string, IColumnFilter> = new Map();
    private _advancedFilterError: string | null = null;
    private _columnStats: Map<string, { mean: number; deviation: number }> = new Map();
    private _columnMinMax: Map<string, { min: number; max: number }> = new Map();
    private columnMaxCache: Map<string, number> = new Map();

    // Fetch More Data (D1/D2/A1-A4) -----------------------------------------------------
    private _isFetchingMore = false;
    private _hasMoreData = false;
    private _fetchMoreFailed = false;
    private _forceFetchAllReason: "search" | "export-csv" | "export-excel" | "export-pdf" | null = null;
    private _scrollListenerAttached = false;
    private _scrollRenderFrame: number | null = null;
    private _searchDebounceTimer: number | undefined;

    // Module D / Item 34: mobile support -------------------------------------------------
    private _isNarrow = false;
    private _narrowAvailableWidth = 0;

    // Item 1: calculated columns (name -> formula/parsed AST)
    private _calcColumns: Map<string, ICalcColumnDef> = new Map();
    // Item 4: combined columns (name -> template string)
    private _combinedColumns: Map<string, ICombinedColumnDef> = new Map();
    // Tracks which virtual (calculated/combined) column names are currently materialized into row.values,
    // so they can be cleanly removed/recomputed without leaking stale keys.
    private _virtualColumnNames: Set<string> = new Set();
    // Item 3: sparklines — set of measure column names with the trend indicator turned on
    private _sparklineColumns: Set<string> = new Set();
    private _tier4Rules: ITier4ConditionalRule[] = [];
    private _tier4ColumnColors: Map<string, string> = new Map();
    private _tier4Palette: DltPaletteName = "default";
    /** Hex colors most recently pasted into the "Custom palette" box (validated, deduped, capped at 12). */
    private _tier4CustomPalette: string[] = [];
    /** Every user-saved theme (header/cell/accent bundle), click-to-apply. No author/viewer hierarchy --
     *  every user gets the same flat gallery and picks freely; each user's active pick is stored separately
     *  as part of their own persisted config, not this list. */
    private _tier4SavedThemes: ITier4SavedTheme[] = [];
    /** id of the theme currently applied by THIS user, or null if they're on the report's plain default. */
    private _tier4ActiveThemeId: string | null = null;
    /** Versioned reusable layouts stored in the visual's report-persisted userConfig object. */
    private _namedLayouts: NamedLayout[] = [];
    private readonly _customEvents = new DataLakeEventBus();
    // Item 2: drag-to-pivot — the single row column currently promoted to a group-by, if any
    private _quickGroupColumn: ITableColumn | null = null;
    // Item 5: true drill-down — leaf rows whose full-record detail sub-grid is expanded (keyed by ITableRow.key).
    private _expandedDetailRows: Set<string> = new Set();
    // Per-node pixel offsets into the virtualized body, aligned with _renderNodes; recomputed whenever
    // the node list changes, since detail nodes have a variable height unlike the uniform-height rows/groups.
    private _nodeOffsets: number[] = [];
    private _totalContentHeight = 0;
    private readonly detailFieldRowHeight = 22;

    private _pendingDragGroupName?: string;
    private _hydratedFromPersist = false;
    private _persistDebounce: number | undefined;
    /** One document-level dismissal handler, replaced rather than accumulated on every toolbar rerender. */
    private _toolbarDocumentClickHandler: ((event: MouseEvent) => void) | undefined;
    private _settingsDrawerOpen = false;
    private _settingsDrawerWidth = 430;
    private _settingsDrawerManuallyResized = false;
    /** Layout values edited inside the visual; reapplied after every Power BI update so the host format model cannot snap them back. */
    private _userLayoutOverrides: ILayoutState | null = null;
    // Tracks which settings tab is open so it survives renderToolbar()'s full rebuild --
    // every button in this drawer calls renderToolbar(), which used to always reset back to
    // the first tab, making anything on another tab (e.g. "Add rule") look broken even
    // though it had actually worked.
    private _activeSettingsTab: string = "layout";
    private _announcementDismissed = false;
    private announcementRoot!: HTMLDivElement;

    /** Latest known "report's default view" read back from the report's object model (Item 7). Kept fresh on every setData(); only *applied* to live state once, by the caller. */
    private _persistedViewState: ISavedViewState | null = null;

    private scrollRoot!: HTMLDivElement;
    private headerRoot!: HTMLDivElement;
    private bodyRoot!: HTMLDivElement;
    private toolbarRoot!: HTMLDivElement;
    private searchRoot!: HTMLDivElement;
    private rowCountRoot!: HTMLDivElement;
    private quickGroupChipRoot!: HTMLDivElement;
    private quickGroupDropRoot!: HTMLDivElement;
    private filterChipsRoot!: HTMLDivElement;
    private pivotToggleRoot!: HTMLDivElement;
    private pivotConfigRoot!: HTMLDivElement;

    private defaultRowHeight = 32;
    private reportTitle = "Data Lake Tables";

    constructor(
        container: HTMLDivElement,
        host: IVisualHost,
        selectionManager: ISelectionManager,
        tooltipService: ITooltipService,
        localizationManager: ILocalizationManager,
        colorPalette: ISandboxExtendedColorPalette
    ) {
        this.container = container;
        this.host = host;
        this.selectionManager = selectionManager;
        this.tooltipService = tooltipService;
        this.localizationManager = localizationManager;
        this.colorPalette = colorPalette;

        this.container.classList.add("skiba-table-root");
        this.buildSkeleton();
    }

    /**
     * Multi-language support (item 16): looks up `key` in the current locale's
     * stringResources; falls back to `fallback` if the key is missing (e.g. a locale that
     * hasn't been translated yet, or `en-US` before the resource file loads in the dev
     * server, which doesn't support localization). `args` are substituted for `{0}`, `{1}`,
     * ... placeholders -- ILocalizationManager.getDisplayName has no built-in templating.
     */
    private loc(key: string, fallback: string, ...args: string[]): string {
        let text = fallback;
        try {
            const resolved = this.localizationManager && this.localizationManager.getDisplayName(key);
            if (resolved && resolved !== key) {
                text = resolved;
            }
        } catch {
            // A missing/broken localization manager should never break rendering.
        }
        args.forEach((arg, i) => {
            text = text.replace(`{${i}}`, arg);
        });
        return text;
    }

    /** Allow Interactions compliance (item 10): guards every selection/context-menu call site. */
    private interactionsAllowed(): boolean {
        return !this.settings || this.settings.allowInteractions !== false;
    }

    /** Resize-only update path: avoids reparsing the DataView and rebuilding filters/grouping. */
    public handleViewportResize(): void {
        if (!this.settings) return;
        this.applyThemeVars();
        this.renderHeader();
        this.renderVisibleRows();
    }

    /**
     * Multi-visual selection sync (item 18): re-renders the visible rows so this visual's
     * highlighted rows reflect the current selection state, without a full setData() cycle.
     * Called from visual.ts whenever Power BI reports a selection change originating outside
     * this visual (another visual's cross-filter, a bookmark, the filter pane), and also after
     * this visual's own selectionManager.select()/clear() calls resolve.
     */
    public syncExternalSelection(): void {
        if (this.settings) {
            this.renderVisibleRows();
        }
        this._customEvents.emit("selectionChanged");
    }

    /** Subscribe to the original Data Lake Tables extension event surface. */
    public onCustomEvent(listener: DataLakeEventListener): () => void {
        return this._customEvents.on(listener);
    }

    /**
     * Removes all children of an element without using innerHTML (certification
     * requirement -- assigning to innerHTML is flagged as a potential XSS vector
     * by Power BI's own linter, powerbi-visuals/no-inner-outer-html, even when
     * the value being assigned is always the empty string).
     */
    private clearElement(el: HTMLElement): void {
        while (el.firstChild) {
            el.removeChild(el.firstChild);
        }
    }

    /** Builds the static DOM skeleton once: toolbar, search bar, filter chip strip, header, scroll body. */
    private buildSkeleton(): void {
        this.clearElement(this.container);

        this.announcementRoot = document.createElement("div");
        this.announcementRoot.className = "datalake-announcement-root";
        this.container.appendChild(this.announcementRoot);

        this.toolbarRoot = document.createElement("div");
        this.toolbarRoot.className = "skiba-toolbar";
        this.container.appendChild(this.toolbarRoot);

        this.searchRoot = document.createElement("div");
        this.searchRoot.className = "skiba-search";
        this.container.appendChild(this.searchRoot);

        // Item 23: honest row-count display. Deliberately separate from the search/filter
        // match line inside searchRoot (renderStatusLine()), since this must stay visible
        // even when no search term or column filter is active.
        this.rowCountRoot = document.createElement("div");
        this.rowCountRoot.className = "skiba-row-count";
        this.rowCountRoot.setAttribute("role", "status");
        this.rowCountRoot.setAttribute("aria-live", "polite");
        this.container.appendChild(this.rowCountRoot);

        // Item 2: drag-to-pivot drop target. Invisible until a column header drag starts
        // (progressive disclosure) — see renderHeader()'s dragstart/dragend handlers.
        this.quickGroupDropRoot = document.createElement("div");
        this.quickGroupDropRoot.className = "skiba-quick-group-drop";
        this.quickGroupDropRoot.textContent = "Drop here to group by this column";
        this.quickGroupDropRoot.style.display = "none";
        this.quickGroupDropRoot.setAttribute("aria-hidden", "true");
        this.quickGroupDropRoot.addEventListener("dragover", (evt: DragEvent) => evt.preventDefault());
        this.quickGroupDropRoot.addEventListener("drop", (evt: DragEvent) => {
            evt.preventDefault();
            const draggedName = evt.dataTransfer?.getData("text/skiba-column");
            if (draggedName) {
                this.applyQuickGroup(draggedName);
            }
            this.quickGroupDropRoot.style.display = "none";
        });
        this.container.appendChild(this.quickGroupDropRoot);

        this.quickGroupChipRoot = document.createElement("div");
        this.quickGroupChipRoot.className = "skiba-quick-group-chip-root";
        this.container.appendChild(this.quickGroupChipRoot);

        this.pivotToggleRoot = document.createElement("div");
        this.pivotToggleRoot.className = "skiba-pivot-toggle-root";
        this.container.appendChild(this.pivotToggleRoot);

        this.pivotConfigRoot = document.createElement("div");
        this.pivotConfigRoot.className = "skiba-pivot-config-root";
        this.container.appendChild(this.pivotConfigRoot);

        this.filterChipsRoot = document.createElement("div");
        this.filterChipsRoot.className = "skiba-filter-chips";
        this.container.appendChild(this.filterChipsRoot);

        const tableWrap = document.createElement("div");
        tableWrap.className = "skiba-table";
        this.container.appendChild(tableWrap);

        this.headerRoot = document.createElement("div");
        this.headerRoot.className = "skiba-table__header";
        tableWrap.appendChild(this.headerRoot);

        this.scrollRoot = document.createElement("div");
        this.scrollRoot.className = "skiba-table__scroll";
        tableWrap.appendChild(this.scrollRoot);

        this.bodyRoot = document.createElement("div");
        this.bodyRoot.className = "skiba-table__body";
        this.scrollRoot.appendChild(this.bodyRoot);

        this.scrollRoot.addEventListener("scroll", () => {
            if (this._scrollRenderFrame !== null) return;
            const schedule = typeof window.requestAnimationFrame === "function"
                ? window.requestAnimationFrame.bind(window)
                : (callback: FrameRequestCallback) => window.setTimeout(() => callback(Date.now()), 0);
            this._scrollRenderFrame = schedule(() => {
                this._scrollRenderFrame = null;
                this.renderVisibleRows();
            }) as number;
        });

        // Right-Click Context Menu (item 12): empty-space mode. AppSource certification
        // requires both the empty-space and data-point context menu modes; the data-point
        // mode is wired per-row in renderRow(). Guarded by allowInteractions like every
        // other selection-adjacent interaction.
        this.container.addEventListener("contextmenu", (evt: MouseEvent) => {
            if (!this.interactionsAllowed()) {
                return;
            }
            const target = evt.target as HTMLElement;
            if (target.closest(".skiba-table__row")) {
                return; // handled by the row's own contextmenu listener
            }
            evt.preventDefault();
            this.selectionManager.showContextMenu({}, { x: evt.clientX, y: evt.clientY });
        });
    }

    /** Replaces the dataset and columns, resets derived (filtered/sorted/grouped) state, and renders. */
    public setData(
        rowColumns: ITableColumn[],
        groupColumns: ITableColumn[],
        valueColumns: ITableColumn[],
        tooltipColumns: ITableColumn[],
        data: ITableRow[],
        settings: ITableRendererSettings,
        persistedStateJson?: string,
        reportTitle?: string,
        isSegmentContinuation: boolean = false
    ): void {
        // Landing/empty states clear the container and detach the renderer skeleton.
        // Rebuild it before rendering real data so Report view never renders into
        // detached header/body roots after fields are assigned.
        if (!this.container.contains(this.headerRoot)) {
            this.buildSkeleton();
        }

        this.rowColumns = rowColumns;
        this.groupColumns = groupColumns;
        this.valueColumns = valueColumns;
        this.tooltipColumns = tooltipColumns;
        // Verified against Microsoft's documented fetchMoreData contract (see FINDINGS.md):
        // with aggregateSegments left at its default (true) -- see requestMoreData() below --
        // each continuation's data already contains the full cumulative row set (prior
        // segments + the new one), merged by Power BI itself. So a continuation still does a
        // full REPLACE, not a concat -- concatenating an already-cumulative array here would
        // double-count every row on every scroll-triggered fetch. The one real gap: an empty/
        // failed continuation (isSegmentContinuation true, data empty) must not wipe an
        // already-rendered table -- a genuine empty result from a real filter/search/sort
        // change (isSegmentContinuation false) still legitimately clears.
        if (!(isSegmentContinuation && data.length === 0)) {
            this._data = data;
        }
        this.settings = settings;

        // Tier 3 fix: reapply in-visual layout edits after each Power BI update. The
        // formatting model is repopulated from report state on every update() cycle, and
        // that repopulation can win a race against the async persistProperties() call
        // commitLayoutState() below fires -- without this override layer, adjusting one of
        // fontSize/rowHeight/headerBold could silently revert whichever one was set most
        // recently (or a value applied via a named layout / saved view). All three now go
        // through the SAME override layer, not just headerBold.
        if (this._userLayoutOverrides) {
            this.settings.headerBold = this._userLayoutOverrides.headerBold;
            this.settings.fontSize = this._userLayoutOverrides.fontSize;
            this.settings.rowHeight = this._userLayoutOverrides.rowHeight;
            this.settings.virtualScrollRowHeight = this._userLayoutOverrides.rowHeight;
        }

        // Themes fix: same class of bug as the layout override above -- this.settings just got
        // fully replaced from the report's formatting model, which has no idea a theme was
        // applied in-visual. Without this, clicking a theme "worked" for one render and then
        // silently reverted on the very next update() cycle.
        if (this._tier4ActiveThemeId) {
            const builtinThemes: Array<{ id: string; headerBg: string; headerFont: string; cellBg: string; cellFont: string; altRow: string; barColor: string }> = [
                { id: "builtin-slate", headerBg: "#334155", headerFont: "#FFFFFF", cellBg: "#FFFFFF", cellFont: "#1E293B", altRow: "#F1F5F9", barColor: "#64748B" },
                { id: "builtin-ocean", headerBg: "#0C4A6E", headerFont: "#FFFFFF", cellBg: "#FFFFFF", cellFont: "#0C4A6E", altRow: "#F0F9FF", barColor: "#0EA5E9" },
                { id: "builtin-forest", headerBg: "#14532D", headerFont: "#FFFFFF", cellBg: "#FFFFFF", cellFont: "#14532D", altRow: "#F0FDF4", barColor: "#22C55E" },
                { id: "builtin-sunrise", headerBg: "#C2410C", headerFont: "#FFFFFF", cellBg: "#FFFFFF", cellFont: "#431407", altRow: "#FFF7ED", barColor: "#FB923C" }
            ];
            const builtin = builtinThemes.find((t) => t.id === this._tier4ActiveThemeId);
            if (builtin) {
                this.settings.headerBg = builtin.headerBg;
                this.settings.headerFont = builtin.headerFont;
                this.settings.cellBg = builtin.cellBg;
                this.settings.cellFont = builtin.cellFont;
                this.settings.altRow = builtin.altRow;
                this.settings.barColor = builtin.barColor;
            } else {
                const saved = this._tier4SavedThemes.find((t) => t.id === this._tier4ActiveThemeId);
                if (saved) {
                    if (saved.headerBg) this.settings.headerBg = saved.headerBg;
                    if (saved.headerFont) this.settings.headerFont = saved.headerFont;
                    if (saved.cellBg) this.settings.cellBg = saved.cellBg;
                    if (saved.cellFont) this.settings.cellFont = saved.cellFont;
                    if (saved.altRow) this.settings.altRow = saved.altRow;
                    if (saved.accent) this.settings.barColor = saved.accent;
                }
            }
        }

        this._isFetchingMore = false;
        this._hasMoreData = settings.hasMoreData;
        this._fetchMoreFailed = false;
        this.renderRowCountStatus();
        this.renderLoadingMoreIndicator();

        if (this._forceFetchAllReason) {
            if (this._hasMoreData) {
                this.requestMoreData();
                this.renderForceFetchProgress();
            } else {
                this.completeForceFetchAll();
            }
        }

        if (!this._scrollListenerAttached) {
            this._scrollListenerAttached = true;
            this.scrollRoot.addEventListener("scroll", () => this.maybeRequestMoreDataFromScroll());
        }
        this._persistedViewState = settings.savedViewState;
        this.defaultRowHeight = settings.virtualScrollEnabled ? settings.virtualScrollRowHeight : settings.rowHeight;
        if (reportTitle) {
            this.reportTitle = reportTitle;
        }

        // Calculations/combined columns/sparklines/drag-pivot are saved onto the report so a
        // page refresh or reload never silently discards a user's calculated columns (Items 1 & 4
        // explicitly must not be lost). Hydrate exactly once — subsequent updates keep the live,
        // in-memory state so an in-flight persistProperties() write can't be raced by a redraw.
        if (!this._hydratedFromPersist) {
            if (persistedStateJson) {
                this.hydrateUserConfig(persistedStateJson);
            }
            this._hydratedFromPersist = true;
        }

        // Rebuilds this.columns (rows + values + calculated + combined) and materializes
        // calculated/combined values into each row before anything downstream reads them.
        this.recomputeVirtualColumns();

        if (this._pendingDragGroupName) {
            const restored = this.rowColumns.find((c) => c.name === this._pendingDragGroupName);
            if (restored) {
                this._quickGroupColumn = restored;
            }
            this._pendingDragGroupName = undefined;
        }

        this.computeColumnStats();
        this.applyPipeline();

        if (isSegmentContinuation) {
            this.renderVisibleRows();
            this.renderStatusLine();
        } else {
            this.render();
        }
    }

    /** Full re-render: toolbar, search bar, pivot chip, filter chips, header row, and the virtualized body. */
    private render(): void {
        this.applyThemeVars();
        this.renderAnnouncement();
        this.renderToolbar();
        this.renderPivotToggle();
        this.renderPivotConfigPanel();
        if (this._pivotState.viewMode === "pivot") {
            this.renderPivotView();
            return;
        }
        this.renderSearchBar();
        this.renderQuickGroupChip();
        this.renderFilterChips();
        this.renderHeader();
        this.renderVisibleRows();
    }

    // -----------------------------------------------------------------
    // Fetch More Data (D1, D2, A1-A4)
    // -----------------------------------------------------------------

    /**
     * D1: export restriction is intentionally separate from Fetch More Data.
     * See docs/DECISIONS.md: no-export blocks export but not data loading.
     */
    private isExportRestricted(): boolean {
        return this.settings.permission === "no-export" || this.settings.permission === "read-only";
    }

    /** D1: read-only blocks Fetch More Data; no-export does not. */
    private isFetchMoreDataRestricted(): boolean {
        return this.settings.permission === "read-only";
    }

    public isAwaitingMoreData(): boolean {
        return this._isFetchingMore;
    }

    private maybeRequestMoreDataFromScroll(): void {
        const threshold = this.scrollRoot.clientHeight || 400;
        const distanceFromBottom = this._totalContentHeight - (this.scrollRoot.scrollTop + this.scrollRoot.clientHeight);
        if (distanceFromBottom <= threshold) {
            this.requestMoreData();
        }
    }

    private requestMoreData(): void {
        if (!this._hasMoreData || this._isFetchingMore || this.isFetchMoreDataRestricted()) {
            return;
        }
        if (typeof this.host.fetchMoreData !== "function") {
            return;
        }
        this._fetchMoreFailed = false;
        this.renderFetchMoreFailedIndicator();
        // Explicit `true` (matches the API default) rather than relying on the implicit
        // default, since the whole segment-accumulation design in setData() above depends on
        // Power BI delivering the cumulative merged row set on each continuation, not a raw
        // incremental delta. See FINDINGS.md for the verified source of this contract.
        const accepted = this.host.fetchMoreData(true);
        if (accepted) {
            this._isFetchingMore = true;
            this.renderLoadingMoreIndicator();
        } else {
            // host.fetchMoreData() is synchronous and returns a boolean (verified against the
            // real powerbi-visuals-api .d.ts) -- there is no Promise/rejection path. `false` IS
            // the failure signal (e.g. the host rejected the request), so it's handled inline.
            this._fetchMoreFailed = true;
            this.renderFetchMoreFailedIndicator();
        }
    }

    private renderLoadingMoreIndicator(): void {
        this.container.querySelectorAll(".skiba-fetch-more-indicator").forEach((el) => el.remove());
        // A fresh loading attempt (including a retry) supersedes any previously shown failure.
        this.container.querySelectorAll(".skiba-fetch-more-failed").forEach((el) => el.remove());
        if (!this._isFetchingMore || this._forceFetchAllReason) {
            return;
        }
        const indicator = document.createElement("div");
        indicator.className = "skiba-fetch-more-indicator";
        indicator.setAttribute("role", "status");
        indicator.setAttribute("aria-live", "polite");
        indicator.textContent = this.loc("FetchMore_Loading", "Loading more rows\u2026");
        this.container.appendChild(indicator);
    }

    /**
     * Item 24: explicit retry control for a failed Fetch More Data request. Uses the exact
     * same DOM-creation pattern as renderLoadingMoreIndicator() above -- elements built via
     * document.createElement only, never innerHTML (flagged by the Power BI linter
     * no-inner-outer-html even for an empty-string assignment; see clearElement()) -- and the
     * same keyboard-accessibility pattern used elsewhere in this file (tabIndex + role="button"
     * + Enter/Space activation alongside the native click handler, matching the pivot chip in
     * renderQuickGroupChip() and the group-row disclosure control in renderGroupRow()).
     */
    private renderFetchMoreFailedIndicator(): void {
        this.container.querySelectorAll(".skiba-fetch-more-failed").forEach((el) => el.remove());
        if (!this._fetchMoreFailed || this._forceFetchAllReason) {
            return;
        }

        const wrap = document.createElement("div");
        wrap.className = "skiba-fetch-more-failed";
        wrap.setAttribute("role", "status");
        wrap.setAttribute("aria-live", "assertive");

        const message = document.createElement("span");
        message.textContent = this.loc("Skiba_Visual_FetchMore_Failed", "Couldn't load more rows \u2014");
        wrap.appendChild(message);

        const retryBtn = document.createElement("span");
        retryBtn.className = "skiba-fetch-more-failed__retry";
        const retryLabel = this.loc("Skiba_Visual_FetchMore_Retry", "Retry");
        retryBtn.textContent = retryLabel;
        retryBtn.setAttribute("role", "button");
        retryBtn.tabIndex = 0;
        retryBtn.setAttribute("aria-label", retryLabel);

        const retry = (): void => {
            this.requestMoreData();
        };
        retryBtn.addEventListener("click", retry);
        retryBtn.addEventListener("keydown", (evt: KeyboardEvent) => {
            if (evt.key === "Enter" || evt.key === " ") {
                evt.preventDefault();
                retry();
            }
        });

        wrap.appendChild(document.createTextNode(" "));
        wrap.appendChild(retryBtn);
        this.container.appendChild(wrap);
    }

    /**
     * Item 23: honest row-count display, independent of the search/filter match line in
     * renderStatusLine(). Power BI's Fetch More Data segmentation only ever tells this visual
     * "at least one more segment exists" -- it never provides a true total row count -- so a
     * partial state can only honestly say "{0}+ rows loaded, more available"
     * (Skiba_Visual_RowCount_PartialUnknownTotal), not "{0} of {1}+ loaded"
     * (Skiba_Visual_RowCount_Partial), which would require a second, genuinely-known total this
     * visual does not have. Skiba_Visual_RowCount_Partial is deliberately left unused rather
     * than fed a fabricated second number -- flag this for review if a real total becomes
     * available from elsewhere (e.g. a DAX total-rows measure bound to a new data role).
     */
    private renderRowCountStatus(): void {
        this.clearElement(this.rowCountRoot);
        if (!this.settings) {
            return;
        }
        const count = this._data.length;
        const text = document.createElement("span");
        text.className = "skiba-row-count__text";
        text.textContent = this._hasMoreData
            ? this.loc("Skiba_Visual_RowCount_PartialUnknownTotal", "{0}+ rows loaded, more available", String(count))
            : this.loc("Skiba_Visual_RowCount_Complete", "{0} rows", String(count));
        this.rowCountRoot.appendChild(text);
    }

    private beginForceFetchAll(reason: "search" | "export-csv" | "export-excel" | "export-pdf"): void {
        if (!this._hasMoreData) {
            this.runForceFetchAction(reason);
            return;
        }
        this._forceFetchAllReason = reason;
        this.renderForceFetchProgress();
        this.requestMoreData();
    }

    private completeForceFetchAll(): void {
        const reason = this._forceFetchAllReason;
        this._forceFetchAllReason = null;
        this.renderForceFetchProgress();
        if (reason) {
            this.runForceFetchAction(reason);
        }
    }

    private runForceFetchAction(reason: "search" | "export-csv" | "export-excel" | "export-pdf"): void {
        switch (reason) {
            case "search": this.renderStatusLine(); break;
            case "export-csv": this.exportCSV(); break;
            case "export-excel": this.exportExcel(); break;

        }
    }

    private renderForceFetchProgress(): void {
        this.container.querySelectorAll(".skiba-force-fetch-progress").forEach((el) => el.remove());
        if (!this._forceFetchAllReason) {
            return;
        }
        const progress = document.createElement("div");
        progress.className = "skiba-force-fetch-progress";
        progress.setAttribute("role", "status");
        progress.setAttribute("aria-live", "polite");

        const labelKey = this._forceFetchAllReason === "search" ? "FetchMore_LoadingForSearch" : "FetchMore_LoadingForExport";
        const labelFallback = this._forceFetchAllReason === "search"
            ? "Loading the rest of the dataset to search\u2026 ({0} rows loaded so far)"
            : "Loading the rest of the dataset before exporting\u2026 ({0} rows loaded so far)";

        const spinner = document.createElement("span");
        spinner.className = "skiba-spinner";
        spinner.setAttribute("aria-hidden", "true");
        progress.appendChild(spinner);

        const text = document.createElement("span");
        text.textContent = this.loc(labelKey, labelFallback, String(this._data.length));
        progress.appendChild(text);

        this.container.insertBefore(progress, this.container.firstChild);
    }

    // -----------------------------------------------------------------
    // Item 1 + Item 4: calculated & combined (virtual) columns
    // -----------------------------------------------------------------

    /** Rebuilds this.columns from rows + values + calculated + combined, preserving user column order. */
    private rebuildColumnsList(): void {
        const calcCols: ITableColumn[] = Array.from(this._calcColumns.keys()).map((name) => ({
            name,
            displayName: name,
            isMeasure: true,
            isGroupBy: false,
            isCalculated: true
        }));
        const combinedCols: ITableColumn[] = Array.from(this._combinedColumns.keys()).map((name) => ({
            name,
            displayName: name,
            isMeasure: false,
            isGroupBy: false,
            isCombined: true
        }));

        this.columns = [...this.rowColumns, ...this.valueColumns, ...calcCols, ...combinedCols];

        const knownNames = new Set(this._columnOrder);
        this.columns.forEach((c) => {
            if (!knownNames.has(c.name)) {
                this._columnOrder.push(c.name);
            }
        });
        this._columnOrder = this._columnOrder.filter((name) => this.columns.some((c) => c.name === name));
    }

    private buildAggregateContext(): ICalcAggregates {
        const cache = new Map<string, number>();
        const data = this._data;
        return {
            getAggregate: (fn, columnName) => {
                const key = `${fn}::${columnName}`;
                if (cache.has(key)) {
                    return cache.get(key)!;
                }
                const values = data
                    .map((r) => r.values[columnName])
                    .filter((v): v is number => typeof v === "number");
                if (values.length === 0) {
                    return null;
                }
                let result: number | undefined;
                switch (fn) {
                    case "AVG": result = d3.mean(values); break;
                    case "SUM": result = d3.sum(values); break;
                    case "MIN": result = d3.min(values); break;
                    case "MAX": result = d3.max(values); break;
                }
                if (result !== undefined) {
                    cache.set(key, result);
                    return result;
                }
                return null;
            }
        };
    }

    private rowContext(row: ITableRow): ICalcRowContext {
        return { getColumnValue: (name: string) => row.values[name] };
    }

    /** Clears previously materialized virtual values, then recomputes calculated + combined columns for every row. */
    private recomputeVirtualColumns(): void {
        this._virtualColumnNames.forEach((name) => {
            this._data.forEach((r) => {
                delete r.values[name];
            });
        });
        this._virtualColumnNames.clear();

        const agg = this.buildAggregateContext();

        const dependencyResult = orderCalcColumns(new Map(
            Array.from(this._calcColumns.entries()).map(([name, def]) => [name, { referencedColumns: def.parsed.referencedColumns ?? [] }])
        ));
        const evaluationOrder = [...dependencyResult.order, ...Array.from(dependencyResult.cyclic)];
        evaluationOrder.forEach((name) => {
            const def = this._calcColumns.get(name);
            if (!def) return;
            this._virtualColumnNames.add(name);
            if (!def.parsed.ok || !def.parsed.ast || dependencyResult.cyclic.has(name)) {
                return;
            }
            this._data.forEach((row) => {
                const value: CalcValue = evaluateCalc(def.parsed.ast!, this.rowContext(row), agg);
                row.values[name] = typeof value === "boolean" ? (value ? 1 : 0) : value;
            });
        });

        this._combinedColumns.forEach((def, name) => {
            this._virtualColumnNames.add(name);
            this._data.forEach((row) => {
                row.values[name] = def.template.replace(/\{([^}]+)\}/g, (_match, ref: string) => {
                    const colName = ref.trim();
                    const v = row.values[colName];
                    return v === null || v === undefined ? "" : String(v);
                });
            });
        });

        this.rebuildColumnsList();
    }

    private currentLayoutState(): ILayoutState {
        return {
            fontSize: Number(this.settings.fontSize) || 14,
            rowHeight: Number(this.settings.rowHeight) || 36,
            headerBold: !!this.settings.headerBold,
            responsive: true,
            autoFitColumns: true,
        };
    }

    private commitLayoutState(patch: Partial<ILayoutState>): void {
        const next = updateLayoutState(this.currentLayoutState(), patch);

        this.settings.fontSize = next.fontSize;
        this.settings.rowHeight = next.rowHeight;
        this.settings.headerBold = next.headerBold;
        this.settings.virtualScrollRowHeight = next.rowHeight;
        this.defaultRowHeight = next.rowHeight;
        this.applyThemeVars();
        this.renderHeader();
        this.renderVisibleRows();

        // Tier 3 fix: always capture the FULL layout state into the override layer (not
        // just headerBold) and always persist it via persistUserConfig(). This is what
        // actually stops "adjust row height, font size resets" -- rendering no longer
        // depends on winning a race against the async persistProperties() round-trip
        // below, because setData() now reapplies this override on every update().
        this._userLayoutOverrides = next;
        this.persistUserConfig();

        // Still mirrored into the Format pane's "general" object too, so the pane and a
        // freshly-opened report (before userConfig is parsed) show sensible values.
        const generalProps: { [key: string]: number } = {};
        if (patch.fontSize !== undefined) generalProps.fontSize = next.fontSize;
        if (patch.rowHeight !== undefined) generalProps.rowHeight = next.rowHeight;
        if (Object.keys(generalProps).length > 0) {
            const merge: { objectName: string; selector: null; properties: { [key: string]: number } }[] = [
                { objectName: "general", selector: null, properties: generalProps }
            ];
            // Tier 3 fix: rowHeight was previously mirrored into "general" ONLY -- but when
            // virtual scrolling is enabled, setData() reads defaultRowHeight from
            // "virtualScrolling.rowHeight" instead, a second, separate storage location that
            // never received this update. Keeping both in sync removes that disagreement.
            if (patch.rowHeight !== undefined) {
                merge.push({ objectName: "virtualScrolling", selector: null, properties: { rowHeight: next.rowHeight } });
            }
            this.host.persistProperties({ merge });
        }
    }

    // -----------------------------------------------------------------
    // Persistence — saves calc/combined columns, sparkline toggles, and the
    // drag-to-pivot column onto the report so user work is never silently lost.
    // -----------------------------------------------------------------

    private persistUserConfig(): void {
        const state = {
            calc: Array.from(this._calcColumns.entries()).map(([name, def]) => ({ name, formula: def.formula })),
            combined: Array.from(this._combinedColumns.entries()).map(([name, def]) => ({ name, template: def.template })),
            sparklines: Array.from(this._sparklineColumns),
            dragGroup: this._quickGroupColumn ? this._quickGroupColumn.name : null,
            viewMode: this._pivotState.viewMode,
            pivotConfig: this._pivotState.pivotConfig,
            tier4Rules: this._tier4Rules,
            tier4ColumnColors: Object.fromEntries(this._tier4ColumnColors.entries()),
            tier4Palette: this._tier4Palette,
            tier4CustomPalette: this._tier4CustomPalette,
            tier4SavedThemes: this._tier4SavedThemes,
            tier4ActiveThemeId: this._tier4ActiveThemeId,
            namedLayouts: this._namedLayouts,
            announcementDismissed: this._announcementDismissed,
            layout: { headerBold: this.settings.headerBold, fontSize: this.settings.fontSize, rowHeight: this.settings.rowHeight },
            columnWidthModes: Object.fromEntries(this._columnWidthModes.entries()),
            columnMinWidths: Object.fromEntries(this._columnMinWidths.entries()),
            columnMaxWidths: Object.fromEntries(this._columnMaxWidths.entries()),
            columnAlignments: Object.fromEntries(this._columnAlignments.entries()),
            nonSortableColumns: Array.from(this._nonSortableColumns),
            nonFilterableColumns: Array.from(this._nonFilterableColumns),
            settingsDrawerWidth: this._settingsDrawerWidth,
        };

        // Debounced by a tick so rapid successive edits (e.g. toggling several sparkline
        // checkboxes) collapse into a single write instead of flooding persistProperties.
        if (this._persistDebounce !== undefined) {
            window.clearTimeout(this._persistDebounce);
        }
        this._persistDebounce = window.setTimeout(() => {
            this.host.persistProperties({
                merge: [
                    {
                        objectName: "userConfig",
                        selector: null,
                        properties: { state: JSON.stringify(state) }
                    }
                ]
            });
        }, 0);
    }

    private hydrateUserConfig(json: string): void {
        try {
            const state = JSON.parse(json);
            if (Array.isArray(state.calc)) {
                state.calc.forEach((c: { name?: unknown; formula?: unknown }) => {
                    if (typeof c.name === "string" && typeof c.formula === "string") {
                        this._calcColumns.set(c.name, { formula: c.formula, parsed: parseCalcFormula(c.formula) });
                    }
                });
            }
            if (Array.isArray(state.combined)) {
                state.combined.forEach((c: { name?: unknown; template?: unknown }) => {
                    if (typeof c.name === "string" && typeof c.template === "string") {
                        this._combinedColumns.set(c.name, { template: c.template });
                    }
                });
            }
            if (Array.isArray(state.sparklines)) {
                state.sparklines.forEach((n: unknown) => {
                    if (typeof n === "string") {
                        this._sparklineColumns.add(n);
                    }
                });
            }
            if (state.columnWidthModes && typeof state.columnWidthModes === "object") Object.entries(state.columnWidthModes).forEach(([k, v]) => { if (["auto", "fit", "fill", "fixed", "custom"].includes(String(v))) this._columnWidthModes.set(k, v as "auto" | "fit" | "fill" | "fixed" | "custom"); });
            if (state.columnMinWidths && typeof state.columnMinWidths === "object") Object.entries(state.columnMinWidths).forEach(([k, v]) => { if (Number.isFinite(Number(v))) this._columnMinWidths.set(k, Number(v)); });
            if (state.columnMaxWidths && typeof state.columnMaxWidths === "object") Object.entries(state.columnMaxWidths).forEach(([k, v]) => { if (Number.isFinite(Number(v))) this._columnMaxWidths.set(k, Number(v)); });
            if (state.columnAlignments && typeof state.columnAlignments === "object") Object.entries(state.columnAlignments).forEach(([k, v]) => { if (["auto", "left", "center", "right"].includes(String(v))) this._columnAlignments.set(k, v as "auto" | "left" | "center" | "right"); });
            if (Array.isArray(state.nonSortableColumns)) state.nonSortableColumns.forEach((n: unknown) => { if (typeof n === "string") this._nonSortableColumns.add(n); });
            if (Array.isArray(state.nonFilterableColumns)) state.nonFilterableColumns.forEach((n: unknown) => { if (typeof n === "string") this._nonFilterableColumns.add(n); });
            if (Array.isArray(state.tier4Rules)) {
                this._tier4Rules = state.tier4Rules.filter((r: ITier4ConditionalRule) => r && typeof r.column === "string" && typeof r.value === "string" && typeof r.color === "string");
            }
            if (state.tier4ColumnColors && typeof state.tier4ColumnColors === "object") {
                Object.entries(state.tier4ColumnColors).forEach(([k, v]) => { if (typeof v === "string") this._tier4ColumnColors.set(k, v); });
            }
            if (typeof state.tier4Palette === "string" && (DLT_PALETTE_NAMES.includes(state.tier4Palette) || state.tier4Palette === "custom")) {
                this._tier4Palette = state.tier4Palette as DltPaletteName;
            }
            if (Array.isArray(state.tier4CustomPalette)) {
                const normalized = state.tier4CustomPalette.map((v: unknown) => (typeof v === "string" ? normalizeHex(v) : null)).filter((v: string | null): v is string => v !== null);
                if (normalized.length > 0) this._tier4CustomPalette = normalized.slice(0, 12);
            }
            // Accepts either the new gallery array shape or the old single-theme object shape (pre-upgrade reports).
            this._tier4SavedThemes = safeTier4ThemeList(state.tier4SavedThemes ?? state.tier4SavedTheme);
            if (typeof state.tier4ActiveThemeId === "string") this._tier4ActiveThemeId = state.tier4ActiveThemeId;
            if (typeof state.announcementDismissed === "boolean") this._announcementDismissed = state.announcementDismissed;
            if (Number.isFinite(Number(state.settingsDrawerWidth))) this._settingsDrawerWidth = Number(state.settingsDrawerWidth);
            this._namedLayouts = parseNamedLayouts(state.namedLayouts);
            if (state.layout && typeof state.layout === "object") {
                const layoutPatch: Partial<ILayoutState> = {};
                if (typeof state.layout.headerBold === "boolean") layoutPatch.headerBold = state.layout.headerBold;
                if (typeof state.layout.fontSize === "number") layoutPatch.fontSize = state.layout.fontSize;
                if (typeof state.layout.rowHeight === "number") layoutPatch.rowHeight = state.layout.rowHeight;
                if (Object.keys(layoutPatch).length > 0) {
                    this._userLayoutOverrides = updateLayoutState(this.currentLayoutState(), layoutPatch);
                    this.settings.headerBold = this._userLayoutOverrides.headerBold;
                    this.settings.fontSize = this._userLayoutOverrides.fontSize;
                    this.settings.rowHeight = this._userLayoutOverrides.rowHeight;
                    this.settings.virtualScrollRowHeight = this._userLayoutOverrides.rowHeight;
                }
            }
            if (typeof state.dragGroup === "string") {
                this._pendingDragGroupName = state.dragGroup;
            }
            if (state.viewMode === "table" || state.viewMode === "pivot") {
                this.dispatchPivot({ type: "SET_VIEW_MODE", mode: state.viewMode });
            }
            if (state.pivotConfig && typeof state.pivotConfig.rowField === "string") {
                this.dispatchPivot({ type: "SET_PIVOT_CONFIG", config: { aggregation: "sum", ...state.pivotConfig } });
            }
        } catch (error) {
            // Corrupt or pre-upgrade persisted state — start clean rather than crashing the visual.
            // Tier 5: this used to swallow silently with no trace. Logging here doesn't change
            // the recovery behavior -- it just means a corrupted save is diagnosable instead of
            // invisible if a report's settings ever mysteriously reset.
            logDiagnostic("hydrateUserConfig", error);
        }
    }

    // -----------------------------------------------------------------
    // Item 2: drag-to-pivot
    // -----------------------------------------------------------------

    private captureLayoutSnapshot(): LayoutSnapshot {
        return {
            sortColumn: this._sortState.column,
            sortDirection: this._sortState.direction,
            columnOrder: [...this._columnOrder],
            columnWidths: Object.fromEntries(this._columnWidths.entries()),
            hiddenColumns: Array.from(this._hiddenColumns),
            searchTerm: this._searchTerm,
            groupExpansion: Object.fromEntries(this._groupExpansion.entries()),
            fontSize: Number(this.settings.fontSize) || 14,
            rowHeight: Number(this.settings.rowHeight) || 36,
            headerBold: !!this.settings.headerBold
        };
    }

    private applyNamedLayout(layout: NamedLayout): void {
        const state = layout.state;
        this._sortState = { column: state.sortColumn, direction: state.sortDirection };
        this._columnOrder = [...state.columnOrder];
        this._columnWidths = new Map(Object.entries(state.columnWidths));
        this._hiddenColumns = new Set(state.hiddenColumns);
        this._searchTerm = state.searchTerm;
        this._groupExpansion = new Map(Object.entries(state.groupExpansion));
        this._userLayoutOverrides = updateLayoutState(this.currentLayoutState(), {
            fontSize: state.fontSize, rowHeight: state.rowHeight, headerBold: state.headerBold
        });
        this.settings.fontSize = this._userLayoutOverrides.fontSize;
        this.settings.rowHeight = this._userLayoutOverrides.rowHeight;
        this.settings.headerBold = this._userLayoutOverrides.headerBold;
        this.settings.virtualScrollRowHeight = this._userLayoutOverrides.rowHeight;
        this.defaultRowHeight = this._userLayoutOverrides.rowHeight;
        this.applyThemeVars();
        this.applyPipeline();
        this.renderSearchBar();
        this.renderHeader();
        this.renderFilterChips();
        this.renderVisibleRows();
        this.renderStatusLine();
        this.persistUserConfig();
        this._customEvents.emit("layoutApplied", { id: layout.id, name: layout.name });
        this.renderToolbar();
    }

    private renderNamedLayouts(section: HTMLElement): void {
        const title = document.createElement("div");
        title.className = "datalake-settings-subtitle";
        title.textContent = "Named layouts";
        section.appendChild(title);

        const row = document.createElement("div");
        row.className = "datalake-layout-presets__create";
        const input = document.createElement("input");
        input.type = "text";
        input.maxLength = 80;
        input.placeholder = "e.g. Regional review";
        input.setAttribute("aria-label", "New layout name");
        const save = document.createElement("button");
        save.type = "button";
        save.className = "datalake-settings-secondary";
        save.textContent = "Save layout";
        save.addEventListener("click", () => {
            const name = normalizeLayoutName(input.value);
            if (!name) return;
            const layout: NamedLayout = {
                id: createLayoutId(name),
                name,
                updatedAt: new Date().toISOString(),
                state: this.captureLayoutSnapshot()
            };
            this._namedLayouts = upsertNamedLayout(this._namedLayouts, layout);
            this._customEvents.emit("layoutSaved", { id: layout.id, name: layout.name });
            input.value = "";
            this.persistUserConfig();
            this.renderToolbar();
        });
        row.append(input, save);
        section.appendChild(row);

        const list = document.createElement("div");
        list.className = "datalake-layout-presets";
        if (this._namedLayouts.length === 0) {
            const empty = document.createElement("div");
            empty.className = "datalake-settings-help";
            empty.textContent = "Save a layout to reuse columns, filters, sorting, groups, and density.";
            list.appendChild(empty);
        }
        this._namedLayouts.forEach((layout) => {
            const item = document.createElement("div");
            item.className = "datalake-layout-preset";
            const apply = document.createElement("button");
            apply.type = "button";
            apply.className = "datalake-settings-secondary";
            apply.textContent = layout.name;
            apply.title = `Apply ${layout.name}`;
            apply.addEventListener("click", () => this.applyNamedLayout(layout));
            const remove = document.createElement("button");
            remove.type = "button";
            remove.className = "datalake-settings-secondary";
            remove.textContent = "Remove";
            remove.setAttribute("aria-label", `Remove layout ${layout.name}`);
            remove.addEventListener("click", () => {
                this._namedLayouts = removeNamedLayout(this._namedLayouts, layout.id);
                this._customEvents.emit("layoutRemoved", { id: layout.id, name: layout.name });
                this.persistUserConfig();
                this.renderToolbar();
            });
            item.append(apply, remove);
            list.appendChild(item);
        });
        section.appendChild(list);
    }

    private renderDataLakeLayoutSection(menu: HTMLDivElement): void {
        const section = document.createElement("section");
        section.className = "skiba-toolbar__section datalake-layout-section";
        const title = document.createElement("div");
        title.className = "skiba-toolbar__section-title";
        title.textContent = "Data Lake Tables layout";
        section.appendChild(title);
        const help = document.createElement("div");
        help.className = "datalake-settings-help";
        help.textContent = "Adjust the table density and hierarchy without reopening the formatting pane.";
        section.appendChild(help);
        this.renderNamedLayouts(section);

        const addRange = (labelText: string, min: number, max: number, step: number, current: number, apply: (value: number) => void): void => {
            const row = document.createElement("div");
            row.className = "datalake-layout-control";
            const label = document.createElement("span");
            label.textContent = labelText;
            const input = document.createElement("input");
            input.type = "range";
            input.min = String(min);
            input.max = String(max);
            input.step = String(step);
            input.value = String(current);
            input.setAttribute("aria-label", labelText);
            const number = document.createElement("input");
            number.type = "number";
            number.inputMode = "decimal";
            number.autocomplete = "off";
            number.min = String(min);
            number.max = String(max);
            number.step = String(step);
            number.value = String(current);
            number.className = "datalake-layout-control__number";
            number.setAttribute("aria-label", `${labelText} exact value`);
            const commit = (value: string): void => {
                const n = Number(value);
                if (!Number.isFinite(n)) return;
                const clamped = Math.max(min, Math.min(max, n));
                input.value = String(clamped);
                number.value = String(clamped);
                apply(clamped);
                this.persistUserConfig();
            };
            input.addEventListener("input", () => commit(input.value));
            input.addEventListener("change", () => commit(input.value));
            number.addEventListener("change", () => commit(number.value));
            number.addEventListener("keydown", (evt) => {
                if (evt.key === "Enter") { evt.preventDefault(); commit(number.value); number.blur(); }
            });
            row.append(label, input, number);
            section.appendChild(row);
        };
        addRange("Font size", 8, 48, 0.5, Number(this.settings.fontSize) || 14, (n) => {
            this.commitLayoutState({ fontSize: n });
        });
        addRange("Row height", 22, 120, 1, Number(this.settings.rowHeight) || 36, (n) => {
            this.commitLayoutState({ rowHeight: n });
        });

        const boldLabel = document.createElement("label");
        boldLabel.className = "datalake-layout-check";
        const bold = document.createElement("input");
        bold.type = "checkbox";
        bold.checked = !!this.settings.headerBold;
        bold.addEventListener("change", () => { this.commitLayoutState({ headerBold: bold.checked }); });
        boldLabel.append(bold, document.createTextNode("Emphasize headers"));
        section.appendChild(boldLabel);

        const reset = document.createElement("button");
        reset.type = "button";
        reset.className = "datalake-settings-secondary";
        reset.textContent = "Reset layout";
        reset.addEventListener("click", () => {
            this.commitLayoutState({ fontSize: 14, rowHeight: 36, headerBold: true });
            this.renderToolbar();
        });
        section.appendChild(reset);
        menu.appendChild(section);
    }
    private renderTier4FormattingSection(menu: HTMLDivElement, closeMenu: () => void): void {
        const section = document.createElement("section");
        section.className = "skiba-toolbar__section skiba-toolbar__section--tier4 skiba-format-editor";
        section.setAttribute("aria-label", this.loc("Toolbar_FormattingRules", "In-visual formatting"));

        const heading = document.createElement("div");
        heading.className = "skiba-format-editor__heading";
        const title = document.createElement("strong");
        title.textContent = this.loc("Toolbar_FormattingRules", "In-visual formatting");
        heading.appendChild(title);
        const subtitle = document.createElement("span");
        subtitle.textContent = this.loc("Toolbar_FormattingHelp", "Create rules that apply immediately to table cells.");
        heading.appendChild(subtitle);
        section.appendChild(heading);



        const ruleBuilder = document.createElement("div");
        ruleBuilder.className = "skiba-format-editor__builder";
        const builderTitle = document.createElement("div");
        builderTitle.className = "skiba-format-editor__label";
        builderTitle.textContent = this.loc("Toolbar_NewRule", "New rule");
        ruleBuilder.appendChild(builderTitle);

        const column = document.createElement("select");
        column.className = "skiba-format-editor__select";
        column.setAttribute("aria-label", this.loc("Toolbar_RuleColumn", "Column"));
        this.columns.forEach((c) => {
            const option = document.createElement("option");
            option.value = c.name;
            option.textContent = c.displayName;
            column.appendChild(option);
        });
        ruleBuilder.appendChild(column);

        const operator = document.createElement("select");
        operator.className = "skiba-format-editor__select";
        operator.setAttribute("aria-label", this.loc("Toolbar_RuleOperator", "Operator"));
        const refreshOperators = (): void => {
            const selectedColumn = this.columns.find((c) => c.name === column.value);
            const sample = this._data.find((row) => row.values[column.value] !== null && row.values[column.value] !== undefined)?.values[column.value];
            const kind = selectedColumn?.isMeasure ? "numeric" : inferAnalyticsValueKind(sample);
            const labels: Record<string, string> = { gt: "Greater than", lt: "Less than", equals: "Equal", between: "Between", contains: "Contains", startsWith: "Starts with", endsWith: "Ends with", blank: "Blank", notBlank: "Not blank" };
            const previous = operator.value;
            operator.replaceChildren();
            compatibleAnalyticsOperators(kind).forEach((name) => {
                const opt = document.createElement("option"); opt.value = name; opt.textContent = labels[name]; operator.appendChild(opt);
            });
            if (Array.from(operator.options).some((option) => option.value === previous)) operator.value = previous;
        };
        refreshOperators();
        column.addEventListener("change", refreshOperators);
        ruleBuilder.appendChild(operator);

        const value = document.createElement("input");
        value.className = "skiba-format-editor__input";
        value.type = "text";
        value.placeholder = this.loc("Toolbar_RuleValue", "Value or threshold");
        value.setAttribute("aria-label", this.loc("Toolbar_RuleValue", "Value or threshold"));
        ruleBuilder.appendChild(value);

        const secondValue = document.createElement("input");
        secondValue.className = "skiba-format-editor__input";
        secondValue.type = "text";
        secondValue.placeholder = this.loc("Toolbar_RuleSecondValue", "Upper threshold");
        secondValue.setAttribute("aria-label", this.loc("Toolbar_RuleSecondValue", "Upper threshold"));
        secondValue.style.display = "none";
        operator.addEventListener("change", () => {
            secondValue.style.display = operator.value === "between" ? "" : "none";
            const noValue = operator.value === "blank" || operator.value === "notBlank";
            value.style.display = noValue ? "none" : "";
        });
        ruleBuilder.appendChild(secondValue);

        const format = document.createElement("select");
        format.className = "skiba-format-editor__select";
        format.setAttribute("aria-label", this.loc("Toolbar_RuleFormat", "Format by"));
        [["background", "Background color"], ["font", "Font color"], ["dataBar", "Data bar"], ["icon", "Icon"]].forEach(([v, label]) => { const opt = document.createElement("option"); opt.value = v; opt.textContent = label; format.appendChild(opt); });
        ruleBuilder.appendChild(format);

        const scope = document.createElement("select");
        scope.className = "skiba-format-editor__select";
        scope.setAttribute("aria-label", this.loc("Toolbar_RuleScope", "Apply to"));
        (([["cell", this.loc("Toolbar_RuleScopeCell", "Cell")], ["row", this.loc("Toolbar_RuleScopeRow", "Whole row")]]) as [string, string][]).forEach(([val, label]) => {
            const opt = document.createElement("option");
            opt.value = val;
            opt.textContent = label;
            scope.appendChild(opt);
        });
        ruleBuilder.appendChild(scope);
        format.addEventListener("change", () => {
            const rowAllowed = format.value === "background";
            scope.disabled = !rowAllowed;
            if (!rowAllowed) scope.value = "cell";
        });

        const color = document.createElement("input");
        color.className = "skiba-format-editor__color";
        color.type = "color";
        color.value = "#FAF623";
        color.setAttribute("aria-label", this.loc("Toolbar_RuleColor", "Rule color"));
        ruleBuilder.appendChild(color);

        const add = document.createElement("button");
        add.className = "skiba-format-editor__primary";
        add.type = "button";
        add.textContent = this.loc("Toolbar_AddRule", "Add rule");
        add.addEventListener("click", () => {
            const noValue = operator.value === "blank" || operator.value === "notBlank";
            if (!column.value || (!noValue && !value.value.trim()) || (operator.value === "between" && !secondValue.value.trim())) return;
            this._tier4Rules.push({
                column: column.value,
                operator: operator.value as ITier4ConditionalRule["operator"],
                value: noValue ? "" : value.value.trim(),
                secondValue: operator.value === "between" ? secondValue.value.trim() : undefined,
                color: color.value,
                scope: scope.value as "cell" | "row",
                format: format.value as "background" | "font" | "dataBar" | "icon",
                icon: format.value === "icon" ? "up" : undefined
            });
            value.value = "";
            secondValue.value = "";
            this.persistUserConfig();
            this.renderVisibleRows();
            this.renderToolbar();
        });
        ruleBuilder.appendChild(add);
        section.appendChild(ruleBuilder);

        const rulesTitle = document.createElement("div");
        rulesTitle.className = "skiba-format-editor__label";
        rulesTitle.textContent = `${this.loc("Toolbar_ActiveRules", "Active rules")} (${this._tier4Rules.length})`;
        section.appendChild(rulesTitle);
        const rules = document.createElement("div");
        rules.className = "skiba-format-editor__rules";
        this._tier4Rules.forEach((rule, index) => {
            const card = document.createElement("div");
            card.className = "skiba-format-editor__rule-card";
            const swatch = document.createElement("span");
            swatch.className = "skiba-format-editor__swatch";
            swatch.style.backgroundColor = rule.color;
            card.appendChild(swatch);
            const text = document.createElement("span");
            text.textContent = `${rule.column} ${rule.operator}${rule.value ? ` ${rule.value}` : ""}${rule.secondValue ? ` and ${rule.secondValue}` : ""} · ${rule.format ?? "background"}${rule.scope === "row" ? " · whole row" : ""}`;
            card.appendChild(text);
            const remove = document.createElement("button");
            remove.type = "button";
            remove.className = "skiba-format-editor__icon-button";
            remove.textContent = "×";
            remove.title = this.loc("Toolbar_RemoveRule", "Remove rule");
            remove.setAttribute("aria-label", `${this.loc("Toolbar_RemoveRule", "Remove rule")} ${index + 1}`);
            remove.addEventListener("click", () => {
                this._tier4Rules.splice(index, 1);
                this.persistUserConfig();
                this.renderVisibleRows();
                this.renderToolbar();
            });
            card.appendChild(remove);
            rules.appendChild(card);
        });
        if (this._tier4Rules.length === 0) {
            const empty = document.createElement("span");
            empty.className = "skiba-format-editor__empty";
            empty.textContent = this.loc("Toolbar_NoRules", "No active rules yet.");
            rules.appendChild(empty);
        }
        section.appendChild(rules);

        const columnsTitle = document.createElement("div");
        columnsTitle.className = "skiba-format-editor__label";
        columnsTitle.textContent = this.loc("Toolbar_ColumnOverrides", "Column colors");
        section.appendChild(columnsTitle);
        const overrides = document.createElement("div");
        overrides.className = "skiba-format-editor__overrides";
        this.columns.forEach((c) => {
            const row = document.createElement("label");
            row.className = "skiba-format-editor__override";
            const name = document.createElement("span");
            name.textContent = c.displayName;
            row.appendChild(name);
            const input = document.createElement("input");
            input.type = "color";
            input.value = this._tier4ColumnColors.get(c.name) ?? "#FFFFFF";
            input.title = this.loc("Toolbar_ColumnColorHint", "Click to pick a background color for this column");
            input.setAttribute("aria-label", `Color override for ${c.displayName}`);
            input.addEventListener("change", () => {
                this._tier4ColumnColors.set(c.name, input.value);
                this.persistUserConfig();
                this.renderVisibleRows();
            });
            row.appendChild(input);
            const clear = document.createElement("button");
            clear.type = "button";
            clear.className = "skiba-format-editor__icon-button";
            clear.textContent = "x";
            clear.title = this.loc("Toolbar_ClearColumnColor", "Clear color");
            clear.setAttribute("aria-label", `Clear color for ${c.displayName}`);
            clear.addEventListener("click", (evt) => {
                evt.preventDefault();
                input.value = "#FFFFFF";
                this._tier4ColumnColors.delete(c.name);
                this.persistUserConfig();
                this.renderVisibleRows();
            });
            row.appendChild(clear);
            overrides.appendChild(row);
        });
        section.appendChild(overrides);

        this.renderThemeGallerySection(section);

        const actions = document.createElement("div");
        actions.className = "skiba-format-editor__actions";
        const reset = document.createElement("button");
        reset.type = "button";
        reset.textContent = this.loc("Toolbar_ResetRules", "Reset");
        reset.addEventListener("click", () => {
            this._tier4Rules = [];
            this._tier4ColumnColors.clear();
            this._tier4Palette = "default";
            this.persistUserConfig();
            this.renderVisibleRows();
            this.renderToolbar();
        });
        actions.appendChild(reset);
        const close = document.createElement("button");
        close.type = "button";
        close.textContent = this.loc("Toolbar_Done", "Done");
        close.addEventListener("click", closeMenu);
        actions.appendChild(close);
        section.appendChild(actions);
        menu.appendChild(section);
    }

    /** Applies a saved theme's header/cell/accent colors as this user's active choice, or clears back to
     *  the report's plain built-in colors when `theme` is null. No author/viewer precedence: this is the
     *  same action available to every user, and it only ever affects what this user sees. */
    private applyTier4Theme(theme: ITier4SavedTheme | null): void {
        this._tier4ActiveThemeId = theme ? theme.id : null;
        if (theme) {
            if (theme.headerBg) this.settings.headerBg = theme.headerBg;
            if (theme.headerFont) this.settings.headerFont = theme.headerFont;
            if (theme.cellBg) this.settings.cellBg = theme.cellBg;
            if (theme.cellFont) this.settings.cellFont = theme.cellFont;
            if (theme.altRow) this.settings.altRow = theme.altRow;
            if (theme.accent) this.settings.barColor = theme.accent;
        }
        this.applyThemeVars();
        this.renderHeader();
        this.renderVisibleRows();
        this.persistUserConfig();
        this.renderToolbar();
    }

    /**
     * One-click visual theme gallery (header/cell/accent bundles), independent of the data
     * palette above. Every user — not just the report author — sees the same flat list and can
     * apply, save their own, or delete their own. Backed by `_tier4SavedThemes` /
     * `_tier4ActiveThemeId`, persisted per-user via `persistUserConfig`.
     */
    private renderThemeGallerySection(section: HTMLElement): void {
        const heading = document.createElement("div");
        heading.className = "skiba-format-editor__label";
        heading.textContent = this.loc("Toolbar_VisualTheme", "Themes");
        section.appendChild(heading);

        const gallery = document.createElement("div");
        gallery.className = "skiba-format-editor__theme-gallery";

        const defaultCard = document.createElement("button");
        defaultCard.type = "button";
        defaultCard.className = "skiba-format-editor__theme-card";
        defaultCard.setAttribute("aria-pressed", String(this._tier4ActiveThemeId === null));
        const defaultPreview = document.createElement("span");
        defaultPreview.className = "skiba-format-editor__theme-card-preview";
        defaultPreview.style.background = "linear-gradient(135deg, #124E9B, #FAF623)";
        defaultCard.appendChild(defaultPreview);
        const defaultLabel = document.createElement("span");
        defaultLabel.textContent = this.loc("Toolbar_ThemeDefault", "Default");
        defaultCard.appendChild(defaultLabel);
        defaultCard.addEventListener("click", () => {
            this.settings.headerBg = "#124E9B";
            this.settings.headerFont = "#FFFFFF";
            this.settings.cellBg = "#FFFFFF";
            this.settings.cellFont = "#1A2B1A";
            this.settings.altRow = "#F5F8EE";
            this.settings.barColor = "#3089BB";
            this.applyTier4Theme(null);
        });
        gallery.appendChild(defaultCard);

        const builtins: Array<{ id: string; label: string; headerBg: string; headerFont: string; cellBg: string; cellFont: string; altRow: string; barColor: string }> = [
            { id: "builtin-slate", label: "Slate", headerBg: "#334155", headerFont: "#FFFFFF", cellBg: "#FFFFFF", cellFont: "#1E293B", altRow: "#F1F5F9", barColor: "#64748B" },
            { id: "builtin-ocean", label: "Ocean", headerBg: "#0C4A6E", headerFont: "#FFFFFF", cellBg: "#FFFFFF", cellFont: "#0C4A6E", altRow: "#F0F9FF", barColor: "#0EA5E9" },
            { id: "builtin-forest", label: "Forest", headerBg: "#14532D", headerFont: "#FFFFFF", cellBg: "#FFFFFF", cellFont: "#14532D", altRow: "#F0FDF4", barColor: "#22C55E" },
            { id: "builtin-sunrise", label: "Sunrise", headerBg: "#C2410C", headerFont: "#FFFFFF", cellBg: "#FFFFFF", cellFont: "#431407", altRow: "#FFF7ED", barColor: "#FB923C" }
        ];
        builtins.forEach((theme) => {
            const card = document.createElement("button");
            card.type = "button";
            card.className = "skiba-format-editor__theme-card";
            card.setAttribute("aria-pressed", String(this._tier4ActiveThemeId === theme.id));
            const preview = document.createElement("span");
            preview.className = "skiba-format-editor__theme-card-preview";
            preview.style.background = `linear-gradient(135deg, ${theme.headerBg}, ${theme.barColor})`;
            card.appendChild(preview);
            const label = document.createElement("span");
            label.textContent = theme.label;
            card.appendChild(label);
            card.addEventListener("click", () => {
                this.settings.headerBg = theme.headerBg;
                this.settings.headerFont = theme.headerFont;
                this.settings.cellBg = theme.cellBg;
                this.settings.cellFont = theme.cellFont;
                this.settings.altRow = theme.altRow;
                this.settings.barColor = theme.barColor;
                this._tier4ActiveThemeId = theme.id;
                this.applyThemeVars();
                this.renderHeader();
                this.renderVisibleRows();
                this.persistUserConfig();
                this.renderToolbar();
            });
            gallery.appendChild(card);
        });

        this._tier4SavedThemes.forEach((theme) => {
            const card = document.createElement("div");
            card.className = "skiba-format-editor__theme-card-wrap";
            const applyBtn = document.createElement("button");
            applyBtn.type = "button";
            applyBtn.className = "skiba-format-editor__theme-card";
            applyBtn.setAttribute("aria-pressed", String(this._tier4ActiveThemeId === theme.id));
            const preview = document.createElement("span");
            preview.className = "skiba-format-editor__theme-card-preview";
            preview.style.background = `linear-gradient(135deg, ${theme.headerBg ?? "#124E9B"}, ${theme.accent ?? theme.cellBg ?? "#FAF623"})`;
            applyBtn.appendChild(preview);
            const label = document.createElement("span");
            label.textContent = theme.name;
            applyBtn.appendChild(label);
            applyBtn.addEventListener("click", () => this.applyTier4Theme(theme));
            card.appendChild(applyBtn);

            const remove = document.createElement("button");
            remove.type = "button";
            remove.className = "skiba-format-editor__icon-button";
            remove.textContent = "×";
            remove.title = this.loc("Toolbar_DeleteTheme", "Delete theme");
            remove.setAttribute("aria-label", `${this.loc("Toolbar_DeleteTheme", "Delete theme")}: ${theme.name}`);
            remove.addEventListener("click", () => {
                this._tier4SavedThemes = this._tier4SavedThemes.filter((t) => t.id !== theme.id);
                if (this._tier4ActiveThemeId === theme.id) this._tier4ActiveThemeId = null;
                this.persistUserConfig();
                this.renderToolbar();
            });
            card.appendChild(remove);
            gallery.appendChild(card);
        });
        section.appendChild(gallery);

        // "Save current as new theme" — captures whatever header/cell/accent colors are live
        // right now (whether from the Format pane, a previous theme, or manual tweaks) as a new
        // reusable card, rather than the old one-shot prompt that saved nothing usable.
        const saveRow = document.createElement("div");
        saveRow.className = "skiba-format-editor__save-theme-row";
        const saveToggle = document.createElement("button");
        saveToggle.type = "button";
        saveToggle.className = "skiba-format-editor__button";
        saveToggle.textContent = this.loc("Toolbar_SaveCurrentTheme", "+ Save current colors as theme");
        const saveForm = document.createElement("div");
        saveForm.className = "skiba-calc-form";
        saveForm.style.display = "none";
        const nameInput = document.createElement("input");
        nameInput.type = "text";
        nameInput.className = "skiba-calc-form__input";
        nameInput.placeholder = this.loc("Toolbar_ThemeName", "Theme name, e.g. Field Operations");
        nameInput.setAttribute("aria-label", this.loc("Toolbar_ThemeName", "Theme name"));
        saveForm.appendChild(nameInput);
        const saveConfirm = document.createElement("button");
        saveConfirm.type = "button";
        saveConfirm.className = "skiba-toolbar__button skiba-toolbar__button--primary";
        saveConfirm.textContent = this.loc("Toolbar_Save", "Save");
        saveConfirm.addEventListener("click", (evt) => {
            evt.stopPropagation();
            const name = nameInput.value.trim();
            if (!name) return;
            const theme: ITier4SavedTheme = {
                id: `theme_${Date.now()}_${generateThemeIdSuffix()}`,
                name,
                headerBg: this.settings.headerBg,
                headerFont: this.settings.headerFont,
                cellBg: this.settings.cellBg,
                cellFont: this.settings.cellFont,
                altRow: this.settings.altRow,
                accent: this.settings.barColor
            };
            this._tier4SavedThemes.push(theme);
            this._tier4ActiveThemeId = theme.id;
            nameInput.value = "";
            saveForm.style.display = "none";
            this.persistUserConfig();
            this.renderToolbar();
        });
        saveForm.appendChild(saveConfirm);
        saveToggle.addEventListener("click", (evt) => {
            evt.stopPropagation();
            saveForm.style.display = saveForm.style.display === "none" ? "block" : "none";
        });
        saveRow.appendChild(saveToggle);
        saveRow.appendChild(saveForm);
        section.appendChild(saveRow);
    }
    /** The group-by columns actually used for grouping: the real "Group by" role, plus any drag-pivoted column first. */
    private effectiveGroupColumns(): ITableColumn[] {
        return this._quickGroupColumn ? [this._quickGroupColumn, ...this.groupColumns] : this.groupColumns;
    }

    private applyQuickGroup(columnName: string): void {
        const col = this.rowColumns.find((c) => c.name === columnName);
        if (!col || col.isMeasure) {
            // Only plain dimension columns make sense to group by; silently ignore a measure drop
            // rather than surfacing an error for something that was never going to work.
            return;
        }
        this._quickGroupColumn = col;
        this._groupExpansion.clear();
        this.applyPipeline();
        this.render();
        this.persistUserConfig();
    }

    /** One click, no confirmation: grouping is non-destructive and instantly reversible. */
    private removeQuickGroup(): void {
        this._quickGroupColumn = null;
        this._groupExpansion.clear();
        this.applyPipeline();
        this.render();
        this.persistUserConfig();
    }

    private renderQuickGroupChip(): void {
        this.clearElement(this.quickGroupChipRoot);
        if (!this._quickGroupColumn) {
            this.quickGroupChipRoot.style.display = "none";
            return;
        }
        this.quickGroupChipRoot.style.display = "";

        const chip = document.createElement("span");
        chip.className = "skiba-quick-group-chip";
        chip.textContent = `Grouped by ${this._quickGroupColumn.displayName} \u2014 click to remove`;
        chip.setAttribute("role", "button");
        chip.tabIndex = 0;
        chip.addEventListener("click", () => this.removeQuickGroup());
        chip.addEventListener("keydown", (evt: KeyboardEvent) => {
            if (evt.key === "Enter" || evt.key === " ") {
                evt.preventDefault();
                this.removeQuickGroup();
            }
        });
        this.quickGroupChipRoot.appendChild(chip);
    }

    private measureColumnNames(): string[] {
        return this.valueColumns.map((c) => c.name);
    }

    private renderPivotToggle(): void {
        this.clearElement(this.pivotToggleRoot);
        const button = document.createElement("button");
        if (this.settings.pivotEnabled === false) {
            this.pivotToggleRoot.style.display = "none";
            if (this._pivotState.viewMode === "pivot") this.dispatchPivot({ type: "SET_VIEW_MODE", mode: "table" });
            return;
        }
        this.pivotToggleRoot.style.display = "";
        button.className = "skiba-pivot-toggle" + (this._pivotState.viewMode === "pivot" ? " skiba-pivot-toggle--active" : "");
        button.textContent = this.loc("Toolbar_PivotMode", "Pivot");
        button.setAttribute("aria-pressed", String(this._pivotState.viewMode === "pivot"));
        button.addEventListener("click", () => {
            this.dispatchPivot({ type: "TOGGLE_VIEW_MODE" });
            this.persistUserConfig();
            this.render();
        });
        this.pivotToggleRoot.appendChild(button);
    }

    private renderPivotConfigPanel(): void {
        this.clearElement(this.pivotConfigRoot);
        if (this.settings.pivotEnabled === false || this._pivotState.viewMode !== "pivot") {
            this.pivotConfigRoot.style.display = "none";
            return;
        }
        this.pivotConfigRoot.style.display = "";
        const dimensionColumns = this.columns.filter((c) => !c.isMeasure);
        const measureColumns = this.columns.filter((c) => c.isMeasure);
        const config = {
            rowField: this._pivotState.pivotConfig?.rowField ?? dimensionColumns[0]?.name ?? "",
            columnField: this._pivotState.pivotConfig?.columnField ?? dimensionColumns[1]?.name ?? dimensionColumns[0]?.name ?? "",
            valueField: this._pivotState.pivotConfig?.valueField ?? measureColumns[0]?.name ?? "",
            aggregation: this._pivotState.pivotConfig?.aggregation ?? "sum" as PivotAggregation
        };
        if (!this._pivotState.pivotConfig && config.rowField && config.columnField && config.valueField) {
            this.dispatchPivot({ type: "SET_PIVOT_CONFIG", config });
        }
        const card = document.createElement("div");
        card.className = "skiba-pivot-config-card";
        const title = document.createElement("div");
        title.className = "skiba-pivot-config-card__title";
        title.textContent = this.loc("Pivot_ConfigTitle", "Pivot configuration");
        card.appendChild(title);
        const subtitle = document.createElement("div");
        subtitle.className = "skiba-pivot-config-card__subtitle";
        subtitle.textContent = this.loc("Pivot_ConfigSubtitle", "Build a row-grouped matrix with an optional aggregation on the value field.");
        card.appendChild(subtitle);
        const fields = document.createElement("div");
        fields.className = "skiba-pivot-config-card__fields";

        const makeSelect = (label: string, options: ITableColumn[], current: string, onChange: (name: string) => void): void => {
            const wrap = document.createElement("label");
            wrap.className = "skiba-pivot-field";
            const labelEl = document.createElement("span");
            labelEl.textContent = label;
            const select = document.createElement("select");
            select.setAttribute("aria-label", label);
            const placeholder = document.createElement("option");
            placeholder.value = "";
            placeholder.textContent = options.length ? this.loc("Pivot_SelectPlaceholder", "Choose a field") : this.loc("Pivot_NoField", "No compatible fields");
            placeholder.disabled = options.length > 0;
            select.appendChild(placeholder);
            for (const col of options) {
                const opt = document.createElement("option");
                opt.value = col.name;
                opt.textContent = col.displayName;
                opt.selected = col.name === current;
                select.appendChild(opt);
            }
            select.addEventListener("change", () => onChange(select.value));
            wrap.appendChild(labelEl);
            wrap.appendChild(select);
            fields.appendChild(wrap);
        };

        const commit = (partial: Partial<typeof config>): void => {
            this.dispatchPivot({ type: "UPDATE_PIVOT_CONFIG", partial });
            this.persistUserConfig();
            this.render();
        };
        makeSelect(this.loc("Pivot_RowGroups", "Row Groups"), dimensionColumns, config.rowField, (v) => commit({ rowField: v }));
        makeSelect(this.loc("Pivot_ColumnLabels", "Column Labels"), dimensionColumns, config.columnField, (v) => commit({ columnField: v }));
        makeSelect(this.loc("Pivot_Value", "Values"), measureColumns, config.valueField, (v) => commit({ valueField: v }));

        const aggWrap = document.createElement("label");
        aggWrap.className = "skiba-pivot-field";
        const aggLabel = document.createElement("span");
        aggLabel.textContent = this.loc("Pivot_Aggregation", "Aggregation");
        const agg = document.createElement("select");
        agg.setAttribute("aria-label", this.loc("Pivot_Aggregation", "Aggregation"));
        const aggregations: Array<[PivotAggregation, string]> = [
            ["sum", "Sum"], ["avg", "Average"], ["min", "Minimum"], ["max", "Maximum"], ["count", "Count"]
        ];
        aggregations.forEach(([value, label]) => {
            const opt = document.createElement("option");
            opt.value = value; opt.textContent = label; opt.selected = config.aggregation === value; agg.appendChild(opt);
        });
        agg.addEventListener("change", () => commit({ aggregation: agg.value as PivotAggregation }));
        aggWrap.appendChild(aggLabel); aggWrap.appendChild(agg); fields.appendChild(aggWrap);

        card.appendChild(fields);
        const hint = document.createElement("div");
        hint.className = "skiba-pivot-config-card__hint";
        hint.textContent = dimensionColumns.length && measureColumns.length
            ? this.loc("Pivot_ConfigHint", "Choose Row Groups, Column Labels and Values. The aggregation picker controls how each cell is calculated.")
            : this.loc("Pivot_NoFields", "Add at least one dimension and one measure to use Pivot mode.");
        card.appendChild(hint);
        this.pivotConfigRoot.appendChild(card);
    }

    private renderPivotView(): void {
        if (!this._pivotState.pivotConfig) {
            this.clearElement(this.bodyRoot);
            const msg = document.createElement("div");
            msg.className = "skiba-pivot-empty";
            msg.textContent = this.loc("Pivot_NoFields", "Add at least one dimension and one measure to use Pivot mode.");
            this.bodyRoot.appendChild(msg);
            return;
        }
        const { rowField, columnField, valueField, aggregation = "sum" } = this._pivotState.pivotConfig;
        let result: PivotResult;
        try {
            result = pivotRows(this._filteredData, rowField, columnField, valueField, aggregation);
        } catch (error) {
            console.error("Data Lake Tables: pivot computation failed.", error);
            this.clearElement(this.bodyRoot);
            const errMsg = document.createElement("div");
            errMsg.className = "skiba-pivot-empty";
            errMsg.textContent = this.loc("Pivot_Error", "This pivot configuration could not be applied to the current data. Try different fields.");
            this.bodyRoot.appendChild(errMsg);
            return;
        }
        this.clearElement(this.bodyRoot);
        const table = document.createElement("table");
        table.className = "skiba-pivot-table";
        const thead = document.createElement("thead");
        const caption = document.createElement("caption");
        caption.className = "skiba-pivot-caption skiba-visually-hidden";
        caption.textContent = this.loc(
            "Pivot_Caption",
            "Pivot table: rows by {0}, columns by {1}, values from {2}",
            rowField,
            columnField,
            valueField
        );
        table.appendChild(caption);

        const headRow = document.createElement("tr");
        for (const col of result.columns) {
            const th = document.createElement("th");
            th.setAttribute("scope", "col");
            th.textContent = col;
            headRow.appendChild(th);
        }
        if (this.settings.pivotShowRowTotals === true || this.settings.pivotShowGrandTotal === true) {
            const th = document.createElement("th");
            th.setAttribute("scope", "col");
            th.textContent = this.settings.totalsLabel || this.loc("Pivot_Total", "Total");
            headRow.appendChild(th);
        }
        thead.appendChild(headRow);
        table.appendChild(thead);
        const tbody = document.createElement("tbody");
        for (const row of result.rows) {
            const tr = document.createElement("tr");
            for (const col of result.columns) {
                const td = document.createElement("td");
                const raw = row[col];
                const rowKey = String(row[rowField] ?? "(Blank)");
                const isGeneratedEmpty = result.emptyCells[`${rowKey}\u241F${col}`] === true;
                td.textContent = isGeneratedEmpty || raw === null || raw === undefined || raw === "" ? (this.settings.pivotEmptyValueDisplay ?? "") : String(raw);
                tr.appendChild(td);
            }
            if (this.settings.pivotShowRowTotals === true || this.settings.pivotShowGrandTotal === true) {
                const td = document.createElement("td");
                td.className = "skiba-pivot-table__total";
                td.textContent = this.settings.pivotShowRowTotals === true
                    ? this.formatNumber(result.rowTotals[String(row[rowField] ?? "(Blank)")] ?? 0)
                    : "";
                tr.appendChild(td);
            }
            tbody.appendChild(tr);
        }
        if (this.settings.pivotShowColumnTotals === true || this.settings.pivotShowGrandTotal === true) {
            const totalRow = document.createElement("tr");
            totalRow.className = "skiba-pivot-table__totals";
            result.columns.forEach((col, index) => {
                const td = document.createElement("td");
                if (index === 0) td.textContent = this.settings.totalsLabel || this.loc("Pivot_Total", "Total");
                else if (this.settings.pivotShowColumnTotals === true) td.textContent = this.formatNumber(result.columnTotals[col] ?? 0);
                totalRow.appendChild(td);
            });
            if (this.settings.pivotShowRowTotals === true || this.settings.pivotShowGrandTotal === true) {
                const td = document.createElement("td");
                if (this.settings.pivotShowGrandTotal === true) td.textContent = this.formatNumber(result.grandTotal);
                totalRow.appendChild(td);
            }
            tbody.appendChild(totalRow);
        }
        table.appendChild(tbody);
        this.bodyRoot.appendChild(table);
    }

    private applyThemeVars(): void {
        const root = this.container;

        // High-Contrast Accessibility Mode (item 13): when Power BI's host reports high
        // contrast is active, override every rendered color with the host-provided
        // foreground/background/foregroundSelected/hyperlink set instead of the normal
        // theme/user-configured colors, per Microsoft's high-contrast guidance. The
        // `skiba-high-contrast` class also switches the stylesheet to solid borders/outlines
        // since the default subtle rgba borders aren't reliably visible in high contrast.
        const isHighContrast = !!this.colorPalette && this.colorPalette.isHighContrast;
        root.classList.toggle("skiba-high-contrast", isHighContrast);

        if (isHighContrast) {
            const palette = this.colorPalette;
            const foreground = palette.foreground ? palette.foreground.value : "#FFFFFF";
            const background = palette.background ? palette.background.value : "#000000";
            const foregroundSelected = palette.foregroundSelected ? palette.foregroundSelected.value : foreground;
            const hyperlink = palette.hyperlink ? palette.hyperlink.value : foreground;

            root.style.setProperty("--skiba-font-family", this.settings.fontFamily);
            root.style.setProperty("--skiba-font-size", `${this.settings.fontSize}px`);
            root.style.setProperty("--skiba-font-color", foreground);
            root.style.setProperty("--skiba-header-family", this.settings.headerFontFamily ?? this.settings.fontFamily);
            root.style.setProperty("--skiba-header-size", `${this.settings.headerFontSize ?? this.settings.fontSize}px`);
            root.style.setProperty("--skiba-header-weight", this.settings.headerFontWeight ?? (this.settings.headerBold ? "600" : "400"));
            root.style.setProperty("--skiba-cell-size", `${this.settings.cellFontSize ?? this.settings.fontSize}px`);
            root.style.setProperty("--skiba-cell-weight", this.settings.cellFontWeight ?? "400");
            root.style.setProperty("--skiba-row-height", `${this.defaultRowHeight}px`);
            root.style.setProperty("--skiba-header-bg", background);
            root.style.setProperty("--skiba-header-font", foreground);
            root.style.setProperty("--skiba-header-weight", this.settings.headerBold ? "600" : "400");
            root.style.setProperty("--skiba-cell-bg", background);
            root.style.setProperty("--skiba-cell-font", foreground);
            root.style.setProperty("--skiba-alt-row", background);
            root.style.setProperty("--skiba-bar-color", foreground);
            root.style.setProperty("--skiba-negative-bar-color", foreground);
            root.style.setProperty("--skiba-totals-bg", background);
            root.style.setProperty("--skiba-hc-border", foreground);
            root.style.setProperty("--skiba-hc-selected", foregroundSelected);
            root.style.setProperty("--skiba-hc-hyperlink", hyperlink);
            return;
        }

        root.style.setProperty("--skiba-font-family", this.settings.fontFamily);
        root.style.setProperty("--skiba-font-size", `${this.settings.fontSize}px`);
        root.style.setProperty("--skiba-row-height", `${this.defaultRowHeight}px`);
        root.style.setProperty("--skiba-header-bg", this.settings.headerBg);
        root.style.setProperty("--skiba-header-font", this.settings.headerFont);
        root.style.setProperty("--skiba-header-weight", this.settings.headerBold ? "600" : "400");
        root.style.setProperty("--skiba-cell-bg", this.settings.cellBg);
        root.style.setProperty("--skiba-cell-font", this.settings.cellFont);
        root.style.setProperty("--skiba-alt-row", this.settings.altRow);
        root.style.setProperty("--skiba-bar-color", this.settings.barColor);
        root.style.setProperty("--skiba-negative-bar-color", (this.settings.negativeBarColor ?? "#C50F1F"));
        root.style.setProperty("--skiba-totals-bg", this.settings.totalsBg);
        const surface = this.settings.tableSurface;
        if (surface) {
            root.style.setProperty("--dlt-header-border", surface.headerBorder);
            root.style.setProperty("--dlt-grid-horizontal", surface.grid === "horizontal" || surface.grid === "both" ? surface.rowDivider : "transparent");
            root.style.setProperty("--dlt-grid-vertical", surface.grid === "vertical" || surface.grid === "both" ? surface.rowDivider : "transparent");
            root.style.setProperty("--dlt-hover-bg", surface.hoverBackground);
            root.style.setProperty("--dlt-selected-bg", surface.selectedBackground);
            root.classList.remove("dlt-grid-none", "dlt-grid-horizontal", "dlt-grid-vertical", "dlt-grid-both");
            root.classList.add(`dlt-grid-${surface.grid}`);
        }
    }

    // -----------------------------------------------------------------
    // Toolbar (minimal floating menu — progressive disclosure)
    // -----------------------------------------------------------------

    /** The host owns the outer visual tile; the settings drawer derives its width from the live viewport. */
    private getSettingsDrawerWidthForViewport(): number {
        const cssWidth = Number.parseFloat(this.container.style.getPropertyValue("--dlt-viewport-width"));
        const width = Math.max(1, cssWidth || this.container.clientWidth || 900);
        if (width <= 560) return Math.max(280, Math.round(width - 14));
        return Math.max(420, Math.min(920, Math.round(width * 0.75)));
    }

    private clampSettingsDrawerWidth(rawWidth: number): number {
        const cssWidth = Number.parseFloat(this.container.style.getPropertyValue("--dlt-viewport-width"));
        const width = Math.max(1, cssWidth || this.container.clientWidth || 900);
        const max = width <= 560 ? Math.max(280, Math.round(width - 14)) : Math.min(920, Math.max(420, Math.round(width * 0.92)));
        const min = width <= 560 ? Math.min(280, max) : Math.min(420, max);
        return Math.max(min, Math.min(max, Math.round(rawWidth)));
    }

    private renderToolbar(): void {
        if (this._toolbarDocumentClickHandler) {
            document.removeEventListener("pointerdown", this._toolbarDocumentClickHandler, true);
            this._toolbarDocumentClickHandler = undefined;
        }
        this.clearElement(this.toolbarRoot);
        // Bug fix: the drawer + backdrop live in this.container, not this.toolbarRoot, so
        // clearElement(toolbarRoot) above never removed them. Every settings interaction
        // (adding a rule, picking a theme, anything that calls renderToolbar() again) was
        // leaving the OLD drawer/backdrop in the DOM and stacking a brand new one on top --
        // that's the duplicate-panels / "can't minimize" / "opens by default" bug.
        this.container.querySelectorAll(".datalake-settings-drawer, .datalake-settings-backdrop").forEach((el) => el.remove());
        this.toolbarRoot.style.display = this.settings.showToolbar ? "" : "none";
        if (!this.settings.showToolbar) return;

        const toggle = document.createElement("button");
        toggle.type = "button";
        toggle.className = "datalake-drawer-toggle skiba-hamburger datalake-tables-settings-button";
        toggle.setAttribute("aria-label", this.loc("Toolbar_TableOptions", "Open Data Lake Tables settings"));
        toggle.setAttribute("aria-expanded", String(this._settingsDrawerOpen));
        // Tier 1 fix: a real settings gear icon instead of a hamburger glyph, pinned to a
        // fixed corner with its own z-index (above the drawer's 320 / backdrop's 315) so it
        // never gets covered by, or moves with, the drawer when it opens or expands.
        const gear = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        gear.setAttribute("width", "16");
        gear.setAttribute("height", "16");
        gear.setAttribute("viewBox", "0 0 24 24");
        gear.setAttribute("fill", "none");
        gear.setAttribute("stroke", "currentColor");
        gear.setAttribute("stroke-width", "2");
        gear.setAttribute("stroke-linecap", "round");
        gear.setAttribute("stroke-linejoin", "round");
        const gearCircle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
        gearCircle.setAttribute("cx", "12"); gearCircle.setAttribute("cy", "12"); gearCircle.setAttribute("r", "3");
        const gearPath = document.createElementNS("http://www.w3.org/2000/svg", "path");
        gearPath.setAttribute("d", "M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z");
        gear.append(gearCircle, gearPath);
        toggle.appendChild(gear);
        toggle.style.setProperty("position", "absolute", "important");
        toggle.style.setProperty("top", "10px", "important");
        toggle.style.setProperty("inset-inline-end", "10px", "important");
        toggle.style.setProperty("z-index", "330", "important");
        toggle.style.display = "flex";
        toggle.style.alignItems = "center";
        toggle.style.justifyContent = "center";
        toggle.title = this.loc("Toolbar_TableOptions", "Data Lake Tables settings");
        this.toolbarRoot.appendChild(toggle);

        const drawer = document.createElement("aside");
        drawer.className = "datalake-settings-drawer";
        drawer.setAttribute("role", "dialog");
        drawer.setAttribute("aria-label", this.loc("Toolbar_TableOptions", "Data Lake Tables settings"));
        // Tier 2 fix: this.container is the visual's own full-size root (see
        // buildSkeleton()) -- appending the drawer there instead of under the tiny
        // absolutely-positioned toolbar icon box (see below) means these !important
        // inline styles resolve against the actual visual canvas, not an 8px icon.
        drawer.style.setProperty("position", "absolute", "important");
        drawer.style.setProperty("top", "0", "important");
        drawer.style.setProperty("bottom", "0", "important");
        drawer.style.setProperty("inset-inline-end", "0", "important");
        drawer.style.setProperty("height", "100%", "important");
        drawer.style.setProperty("max-height", "none", "important");
        drawer.style.setProperty("margin", "0", "important");
        drawer.style.setProperty("z-index", "320", "important");
        this._settingsDrawerWidth = this.clampSettingsDrawerWidth(this._settingsDrawerOpen ? this._settingsDrawerWidth : this.getSettingsDrawerWidthForViewport());
        drawer.style.width = `${this._settingsDrawerWidth}px`;
        drawer.style.display = this._settingsDrawerOpen ? "flex" : "none";

        const header = document.createElement("div");
        header.className = "datalake-settings-drawer__header";
        const titleWrap = document.createElement("div");
        const title = document.createElement("div");
        title.className = "datalake-settings-drawer__title";
        title.textContent = this.loc("Toolbar_TableOptions", "Data Lake Tables settings");
        const subtitle = document.createElement("div");
        subtitle.className = "datalake-settings-drawer__subtitle";
        subtitle.textContent = this.loc("Toolbar_SettingsSubtitle", "Configure layout, calculations, formatting and report actions.");
        titleWrap.appendChild(title); titleWrap.appendChild(subtitle);
        const close = document.createElement("button");
        close.type = "button"; close.className = "datalake-settings-drawer__close";
        close.textContent = "×";
        close.setAttribute("aria-label", this.loc("Toolbar_CloseSettings", "Close settings"));
        close.addEventListener("click", (evt) => { evt.stopPropagation(); closeDrawer(); });
        header.appendChild(titleWrap); header.appendChild(close);
        drawer.appendChild(header);

        const body = document.createElement("div");
        body.className = "datalake-settings-drawer__body";
        const nav = document.createElement("nav"); nav.className="datalake-settings-drawer__nav"; nav.setAttribute("aria-label",this.loc("Toolbar_SettingsNavigation","Settings sections"));
        [["layout","Layout"],["display","Data Display"],["experience","Filters"],["styling","Styling"],["analysis","Analysis"]].forEach(([id,label],i)=>{ const b=document.createElement("button"); b.type="button"; b.className="datalake-settings-drawer__nav-item"; b.textContent=label; b.setAttribute("aria-controls",`datalake-settings-section-${id}`); if(id===this._activeSettingsTab){b.classList.add("is-active");b.setAttribute("aria-current","page");} b.addEventListener("click",()=>{this._activeSettingsTab=id;nav.querySelectorAll(".datalake-settings-drawer__nav-item").forEach(e=>{e.classList.remove("is-active");e.removeAttribute("aria-current");});b.classList.add("is-active");b.setAttribute("aria-current","page");body.querySelectorAll("[data-settings-section]").forEach((el)=>{ (el as HTMLElement).style.display = (el as HTMLElement).dataset.settingsSection===id ? "block" : "none"; });}); nav.appendChild(b);});
        drawer.appendChild(nav); drawer.appendChild(body);

        const footer = document.createElement("div");
        footer.className = "datalake-settings-drawer__footer";
        const footerText = document.createElement("span");
        footerText.className = "datalake-settings-drawer__attribution-button";
        footerText.textContent = this.loc("Landing_Attribution", "Built by Simon KP and Bryt Ma Tech UG");
        footerText.style.fontWeight = "700";
        /* attribution is static text: no popup */
        footer.appendChild(footerText);
        drawer.appendChild(footer);

        const resizer = document.createElement("div");
        resizer.className = "datalake-settings-drawer__resizer";
        resizer.setAttribute("role", "separator");
        resizer.setAttribute("aria-orientation", "vertical");
        resizer.tabIndex = 0;
        drawer.appendChild(resizer);

        // Tier 2 fix: translucent click-to-close backdrop -- spec'd earlier this session,
        // never actually shipped. Appended to this.container for the same containing-block
        // reason as the drawer above: it must cover the visual's full box.
        const backdrop = document.createElement("div");
        backdrop.className = "datalake-settings-backdrop";
        backdrop.style.setProperty("position", "absolute", "important");
        backdrop.style.setProperty("inset", "0", "important");
        backdrop.style.setProperty("z-index", "315", "important");
        backdrop.style.display = "none";
        this.container.appendChild(backdrop);
        this.container.appendChild(drawer);

        const closeDrawer = (): void => {
            this._settingsDrawerOpen = false;
            drawer.style.display = "none";
            backdrop.style.display = "none";
            toggle.setAttribute("aria-expanded", "false");
        };
        const openDrawer = (): void => {
            this._settingsDrawerOpen = true;
            // Always recompute proportionally from the live viewport unless the user has
            // manually dragged the resizer this session -- otherwise a stale persisted
            // pixel width silently overrides the intended "75% of the visual" sizing
            // every single time the drawer opens.
            if (!this._settingsDrawerManuallyResized) {
                this._settingsDrawerWidth = this.getSettingsDrawerWidthForViewport();
            }
            this._settingsDrawerWidth = this.clampSettingsDrawerWidth(this._settingsDrawerWidth);
            drawer.style.width = `${this._settingsDrawerWidth}px`;
            drawer.style.display = "flex";
            backdrop.style.display = "block";
            toggle.setAttribute("aria-expanded", "true");
        };
        toggle.addEventListener("click", (evt) => { evt.stopPropagation(); this._settingsDrawerOpen ? closeDrawer() : openDrawer(); });
        backdrop.addEventListener("click", () => closeDrawer());
        drawer.addEventListener("click", (evt) => evt.stopPropagation());
        drawer.addEventListener("keydown", (evt) => {
            if (evt.key === "Escape") { evt.stopPropagation(); closeDrawer(); toggle.focus(); }
        });

        let resizing = false;
        let startX = 0;
        let startWidth = this._settingsDrawerWidth;
        const stopResize = (): void => {
            if (!resizing) return;
            resizing = false;
            drawer.classList.remove("is-resizing");
            document.body.style.cursor = "";
            window.removeEventListener("pointermove", onResize);
            window.removeEventListener("pointerup", stopResize);
            this.persistUserConfig();
        };
        const onResize = (evt: PointerEvent): void => {
            if (!resizing) return;
            const delta = startX - evt.clientX;
            this._settingsDrawerWidth = this.clampSettingsDrawerWidth(startWidth + delta);
            drawer.style.width = `${this._settingsDrawerWidth}px`;
        };
        resizer.addEventListener("pointerdown", (evt) => {
            resizing = true; startX = evt.clientX; startWidth = this._settingsDrawerWidth;
            this._settingsDrawerManuallyResized = true;
            drawer.classList.add("is-resizing"); document.body.style.cursor = "col-resize";
            window.addEventListener("pointermove", onResize);
            window.addEventListener("pointerup", stopResize);
            evt.preventDefault();
        });
        resizer.addEventListener("keydown", (evt) => {
            if (evt.key === "ArrowLeft" || evt.key === "ArrowRight") {
                evt.preventDefault();
                const delta = evt.key === "ArrowLeft" ? 24 : -24;
                this._settingsDrawerManuallyResized = true;
                this._settingsDrawerWidth = this.clampSettingsDrawerWidth(this._settingsDrawerWidth + delta);
                drawer.style.width = `${this._settingsDrawerWidth}px`;
                this.persistUserConfig();
            }
        });

        // Four real panels. Navigation targets the actual panel instead of a decorative heading,
        // so the user's mental model and the DOM structure stay aligned.
        const makePanel = (id: string, titleText: string, help: string): HTMLDivElement => {
            const panel = document.createElement("section");
            panel.id = `datalake-settings-section-${id}`;
            panel.dataset.settingsSection = id;
            panel.className = "datalake-settings-panel";
            const heading = document.createElement("div");
            heading.className = "datalake-settings-panel__heading";
            const title = document.createElement("strong");
            title.textContent = titleText;
            const description = document.createElement("span");
            description.textContent = help;
            heading.append(title, description);
            panel.appendChild(heading);
            const content = document.createElement("div");
            content.className = "datalake-settings-panel__content";
            panel.appendChild(content);
            body.appendChild(panel);
            return content;
        };

        const layoutContent = makePanel("layout", "Layout", "Control density, sizing and how the table fits the available space.");
        this.renderDataLakeLayoutSection(layoutContent);
        layoutContent.appendChild(this.makeMenuButton(this.loc("Toolbar_ResetColumnWidths", "Reset column widths"), () => { this.resetColumnWidths(); this.persistUserConfig(); }));
        layoutContent.appendChild(this.makeMenuButton(this.loc("Toolbar_ResetColumnOrder", "Reset column order"), () => { this._columnOrder = this.columns.map((c) => c.name); this.renderHeader(); this.renderVisibleRows(); this.persistUserConfig(); }));

        const displayContent = makePanel("display", "Data Display", "Choose which fields appear and how values are presented.");
        const columnsSection = document.createElement("div");
        columnsSection.className = "skiba-toolbar__section";
        const columnsTitle = document.createElement("div");
        columnsTitle.className = "skiba-toolbar__section-title";
        columnsTitle.textContent = this.loc("Toolbar_ShowColumns", "Show columns");
        columnsSection.appendChild(columnsTitle);
        this.columns.forEach((col) => {
            const rowWrap = document.createElement("div"); rowWrap.className = "skiba-toolbar__column-row";
            const label = document.createElement("label"); label.className = "skiba-toolbar__checkbox";
            const checkbox = document.createElement("input"); checkbox.type = "checkbox"; checkbox.checked = !this._hiddenColumns.has(col.name);
            checkbox.setAttribute("aria-label", `Show ${col.displayName} column`);
            checkbox.addEventListener("change", () => {
                if (checkbox.checked) this._hiddenColumns.delete(col.name); else this._hiddenColumns.add(col.name);
                this.renderHeader(); this.renderVisibleRows(); this.persistUserConfig();
            });
            label.appendChild(checkbox); label.appendChild(document.createTextNode(col.displayName)); rowWrap.appendChild(label);
            if (col.isMeasure) {
                const sparkLabel = document.createElement("label"); sparkLabel.className = "skiba-toolbar__checkbox skiba-toolbar__checkbox--sparkline";
                const sparkCheckbox = document.createElement("input"); sparkCheckbox.type = "checkbox"; sparkCheckbox.checked = this._sparklineColumns.has(col.name);
                sparkCheckbox.setAttribute("aria-label", `Show sparkline trend for ${col.displayName}`);
                sparkCheckbox.addEventListener("change", () => {
                    if (sparkCheckbox.checked) this._sparklineColumns.add(col.name); else this._sparklineColumns.delete(col.name);
                    this.renderVisibleRows(); this.persistUserConfig();
                });
                sparkLabel.append(sparkCheckbox, document.createTextNode("Trend")); rowWrap.appendChild(sparkLabel);
            }
            columnsSection.appendChild(rowWrap);
        });
        displayContent.appendChild(columnsSection);
        this.renderColumnControlsSection(displayContent);

        const experienceContent = makePanel("experience", "Filters", "Filter table values live, per column. Click a column to set or change its filter.");

        const filterListSection = document.createElement("div");
        filterListSection.className = "skiba-toolbar__section datalake-filter-list";
        const filterListTitle = document.createElement("div");
        filterListTitle.className = "skiba-toolbar__section-title";
        filterListTitle.textContent = this.loc("Toolbar_ColumnFilters", "Column filters");
        filterListSection.appendChild(filterListTitle);
        this.columns.forEach((col) => {
            const row = document.createElement("div");
            row.className = "datalake-filter-list__row";
            const label = document.createElement("span");
            label.className = "datalake-filter-list__name";
            label.textContent = col.displayName;
            row.appendChild(label);
            const active = this._columnFilters.has(col.name);
            const btn = document.createElement("button");
            btn.type = "button";
            btn.className = "datalake-filter-list__button" + (active ? " is-active" : "");
            btn.textContent = active ? this.loc("Toolbar_FilterActive", "Filtering") : this.loc("Toolbar_FilterSet", "Set filter");
            btn.addEventListener("click", (evt) => {
                evt.stopPropagation();
                this.openFilterPopover(col, btn);
            });
            row.appendChild(btn);
            filterListSection.appendChild(row);
        });
        experienceContent.appendChild(filterListSection);

        if (this.effectiveGroupColumns().length > 0) {
            experienceContent.appendChild(this.makeMenuButton(this.loc("Toolbar_ExpandAllGroups", "Expand all groups"), () => { this.expandAllGroups(); closeDrawer(); }));
            experienceContent.appendChild(this.makeMenuButton(this.loc("Toolbar_CollapseAllGroups", "Collapse all groups"), () => { this.collapseAllGroups(); closeDrawer(); }));
        }
        if (this.settings.enableSorting) {
            experienceContent.appendChild(this.makeMenuButton(this.loc("Toolbar_ResetSorts", "Reset sorts"), () => { this.resetSorts(); this.persistUserConfig(); }));
        }
        if (this.settings.showFilterClear) {
            experienceContent.appendChild(this.makeMenuButton(this.loc("Toolbar_ResetFilters", "Reset filters"), () => { this._columnFilters.clear(); this.commitFilterChange(); this.persistUserConfig(); }));
        }
        if (this.settings.enableRowSelection && this.settings.showClearSelection) {
            experienceContent.appendChild(this.makeMenuButton(this.loc("Toolbar_ClearSelection", "Clear selection"), () => {
                if (!this.interactionsAllowed()) return;
                this.selectionManager.clear().then(() => this.syncExternalSelection());
            }));
        }
        if (this.settings.enableCopy && this.settings.copySelected && this.canCopyToClipboard()) {
            experienceContent.appendChild(this.makeMenuButton(this.loc("Toolbar_CopySelected", "Copy selected"), () => this.copySelectedRows()));
        }

        const stylingContent = makePanel("styling", "Styling", "Data colors, conditional rules and table theme.");
        this.renderTier4FormattingSection(stylingContent, closeDrawer);

        const analysisContent = makePanel("analysis", "Analysis & Export", "Pivot, calculations, saved layouts and governed exports.");
        this.renderCalculationsSection(analysisContent, closeDrawer);
        this.renderCombineColumnsSection(analysisContent, closeDrawer);
        this.renderExportControls(analysisContent, closeDrawer);
        const isReadOnly = this.settings.permission === "read-only";
        if (!isReadOnly) analysisContent.appendChild(this.makeSaveDefaultViewButton(closeDrawer));
        analysisContent.appendChild(this.makeMenuButton("Reset to default view", () => { this.resetToDefaultView(); closeDrawer(); }));

        // Tier A fix: real tabs. Only the first panel starts visible; the nav click handler
        // above now shows/hides the matching [data-settings-section] panel instead of
        // scrolling to it, so all four tabs are properly self-contained instead of one long
        // shared scrolling page.
        body.querySelectorAll("[data-settings-section]").forEach((el) => {
            (el as HTMLElement).style.display = (el as HTMLElement).dataset.settingsSection === this._activeSettingsTab ? "block" : "none";
        });

        this._toolbarDocumentClickHandler = (event: MouseEvent): void => {
            const target = event.target as Node | null;
            // Tier 2 fix: the drawer now lives in this.container, not this.toolbarRoot --
            // it must be checked separately, or every click inside the drawer would look
            // like an "outside click" and close it instantly.
            if (this._settingsDrawerOpen && target && !this.toolbarRoot.contains(target) && !drawer.contains(target)) closeDrawer();
        };
        document.addEventListener("pointerdown", this._toolbarDocumentClickHandler, true);
    }

    private closeAnnouncementDetails(): void {
        // Return from the full-bleed product-details surface to the underlying
        // announcement card without dismissing the card itself. Dismissal is a
        // separate, explicit user action on the card's own X control.
        this.announcementRoot.style.removeProperty("inset");
        this.announcementRoot.style.removeProperty("width");
        this.announcementRoot.style.removeProperty("height");
        this.announcementRoot.style.removeProperty("max-width");
        this.announcementRoot.style.removeProperty("z-index");
        this.renderAnnouncement();
    }

    private openAnnouncementDetails(): void {
        // Product details open in the same right-hand panel as FAQ / Support / About,
        // so every information surface looks and behaves the same.
        this.renderLandingInfoModal(
            "Data Lake Tables",
            "ABOUT THE PRODUCT",
            "A focused Power BI table workspace for exploring, analysing, presenting and exporting report data.",
            [
                "Explore: search, filter, sort, hide and reorder columns, with grouping and drill-down interactions.",
                "Analyse: switch to pivot mode, create calculations, combine columns and apply conditional formatting.",
                "Present: use exact font and row-height controls with viewport-driven sizing that follows the Power BI tile.",
                "Format: apply conditional formatting, data bars and saved color themes to your table.",
                "Getting started: add fields to Rows and Values, then use the settings button in the visual to tune the table to your report.",
                "Built by Simon KP and Bryt Ma Tech UG"
            ]
        );
    }

    private openAnnouncementDetailsLegacy(): void {
        this.clearElement(this.announcementRoot);
        this.announcementRoot.style.display = "";
        // Expand like the settings drawer does -- full-bleed over the visual's own box
        // instead of staying trapped inside the small corner card.
        this.announcementRoot.style.setProperty("inset", "0", "important");
        this.announcementRoot.style.setProperty("width", "100%", "important");
        this.announcementRoot.style.setProperty("height", "100%", "important");
        this.announcementRoot.style.setProperty("max-width", "none", "important");
        this.announcementRoot.style.setProperty("z-index", "310", "important");

        const dialog = document.createElement("section");
        dialog.className = "datalake-announcement-details";
        dialog.setAttribute("role", "dialog");
        dialog.setAttribute("aria-modal", "true");
        dialog.setAttribute("aria-label", "Data Lake Tables product details");

        const head = document.createElement("div");
        head.className = "datalake-announcement-details__head";
        const titleWrap = document.createElement("div");
        const eyebrow = document.createElement("span");
        eyebrow.className = "datalake-announcement-details__eyebrow";
        eyebrow.textContent = "Simon KP · Bryt Ma Tech UG";
        const title = document.createElement("strong");
        title.textContent = "Data Lake Tables";
        titleWrap.append(eyebrow, title);
        const close = document.createElement("button");
        close.type = "button";
        close.className = "datalake-announcement-details__close";
        close.textContent = "×";
        close.setAttribute("aria-label", "Close product details");
        close.addEventListener("click", (evt) => {
            evt.preventDefault();
            evt.stopPropagation();
            this.closeAnnouncementDetails();
        });
        head.append(titleWrap, close);
        dialog.appendChild(head);

        const lead = document.createElement("p");
        lead.className = "datalake-announcement-details__lead";
        lead.textContent = "A focused Power BI table workspace for exploring, analysing, presenting and exporting report data.";
        dialog.appendChild(lead);

        const features: Array<[string, string]> = [
            ["Explore", "Search, filter, sort, hide and reorder columns, with grouping and drill-down interactions."],
            ["Analyse", "Switch to pivot mode, create calculations, combine columns and apply conditional formatting."],
            ["Present", "Use exact font and row-height controls with viewport-driven sizing that follows the Power BI tile."],
            ["Format", "Apply conditional formatting, data bars and saved color themes to your table."]
        ];
        const grid = document.createElement("div");
        grid.className = "datalake-announcement-details__grid";
        features.forEach(([name, text]) => {
            const card = document.createElement("div");
            card.className = "datalake-announcement-details__item";
            const itemTitle = document.createElement("strong"); itemTitle.textContent = name;
            const itemText = document.createElement("span"); itemText.textContent = text;
            card.append(itemTitle, itemText); grid.appendChild(card);
        });
        dialog.appendChild(grid);

        const start = document.createElement("div");
        start.className = "datalake-announcement-details__start";
        const startTitle = document.createElement("strong");
        startTitle.textContent = "Getting started";
        const startText = document.createElement("span");
        startText.textContent = "Add fields to Rows and Values, then use the settings button in the visual to tune the table to your report.";
        start.append(startTitle, startText);
        dialog.appendChild(start);

        const attribution = document.createElement("span");
        attribution.className = "datalake-announcement-details__attribution datalake-settings-drawer__attribution-button";
        attribution.textContent = "Built by Simon KP and Bryt Ma Tech UG";
        /* attribution is static text: no popup */
        dialog.appendChild(attribution);

        const back = document.createElement("button");
        back.type = "button";
        back.className = "datalake-announcement-details__back";
        back.textContent = "Back to table";
        back.addEventListener("click", () => this.closeAnnouncementDetails());
        dialog.appendChild(back);

        this.announcementRoot.appendChild(dialog);
        window.setTimeout(() => close.focus(), 0);
    }

    /** Compact, branded in-report product card. Click opens product details; X dismisses it. */
    private renderAnnouncement(): void {
        this.clearElement(this.announcementRoot);
        // Minimize back to the small corner card if openAnnouncementDetails() had expanded
        // this to full-bleed.
        this.announcementRoot.style.removeProperty("inset");
        this.announcementRoot.style.removeProperty("width");
        this.announcementRoot.style.removeProperty("height");
        this.announcementRoot.style.removeProperty("max-width");
        this.announcementRoot.style.removeProperty("z-index");
        if (this._announcementDismissed || !this.settings || this.settings.permission === "read-only") {
            this.announcementRoot.style.display = "none";
            return;
        }
        this.announcementRoot.style.display = "";

        const banner = document.createElement("aside");
        banner.className = "datalake-announcement";
        banner.setAttribute("role", "button");
        banner.tabIndex = 0;
        banner.setAttribute("aria-haspopup", "dialog");
        banner.setAttribute("aria-label", this.loc("Announcement_AriaLabel", "Learn about Data Lake Tables"));

        // The in-report advert stays a lightweight native product card.
        // Showcase artwork belongs to the landing/welcome surface; keeping the advert text-first
        // prevents a large embedded poster from competing with the report data or collapsing
        // on small Power BI tiles.
        const head = document.createElement("div");
        head.className = "datalake-announcement__head";
        const brand = document.createElement("div");
        brand.className = "datalake-announcement__brand";
        const mark = document.createElement("span");
        mark.className = "datalake-announcement__mark";
        mark.textContent = "DLT";
        const brandText = document.createElement("span");
        brandText.textContent = "Simon KP · Bryt Ma Tech UG";
        brand.append(mark, brandText);
        const badge = document.createElement("span");
        badge.className = "datalake-announcement__badge";
        badge.textContent = "Power BI visual";
        head.append(brand, badge);

        const copy = document.createElement("div");
        copy.className = "datalake-announcement__copy";
        const title = document.createElement("strong");
        title.textContent = "Data Lake Tables";
        const text = document.createElement("span");
        text.textContent = "A cleaner workspace for search, grouping, pivoting and table analysis.";
        copy.append(title, text);

        const features = document.createElement("div");
        features.className = "datalake-announcement__features";
        ["Exact sizing", "Pivot + calc", "Formatting"].forEach((label) => {
            const chip = document.createElement("span");
            chip.textContent = label;
            features.appendChild(chip);
        });

        const cta = document.createElement("div");
        cta.className = "datalake-announcement__cta";
        cta.textContent = "View product details";
        const arrow = document.createElement("span");
        arrow.textContent = "→";
        cta.appendChild(arrow);

        const close = document.createElement("button");
        close.type = "button";
        close.className = "datalake-announcement__close";
        close.textContent = "×";
        close.setAttribute("aria-label", this.loc("Announcement_Dismiss", "Dismiss announcement"));
        close.setAttribute("title", this.loc("Announcement_Dismiss", "Dismiss announcement"));
        close.addEventListener("click", (evt) => {
            evt.preventDefault();
            evt.stopPropagation();
            this._announcementDismissed = true;
            this.announcementRoot.style.display = "none";
            this.persistUserConfig();
        });

        banner.append(head, copy, features, cta, close);
        banner.addEventListener("click", (evt) => {
            if ((evt.target as HTMLElement | null)?.closest(".datalake-announcement__close")) return;
            this.openAnnouncementDetails();
        });
        banner.addEventListener("keydown", (evt) => {
            if (evt.key === "Enter" || evt.key === " ") {
                evt.preventDefault();
                this.openAnnouncementDetails();
            }
        });

        this.announcementRoot.appendChild(banner);
    }

    /**
     * "Save current view as default" — a normal button unless a default view
     * already exists, in which case a click swaps the button for a brief inline
     * confirmation (not a browser confirm() dialog) before overwriting it, since
     * doing so affects every other viewer of this shared report.
     */
    private makeSaveDefaultViewButton(closeMenu: () => void): HTMLDivElement {
        const wrapper = document.createElement("div");

        const btn = this.makeMenuButton("Save current view as default", () => {
            if (!this._persistedViewState) {
                this.saveCurrentViewAsDefault();
                closeMenu();
                return;
            }

            wrapper.replaceChildren();
            const confirmBox = document.createElement("div");
            confirmBox.className = "skiba-toolbar__confirm";

            const msg = document.createElement("div");
            msg.className = "skiba-toolbar__confirm-msg";
            msg.textContent = "This replaces the current default view for everyone who opens this report — save?";
            confirmBox.appendChild(msg);

            const actions = document.createElement("div");
            actions.className = "skiba-toolbar__confirm-actions";

            const cancelBtn = document.createElement("button");
            cancelBtn.textContent = "Cancel";
            cancelBtn.addEventListener("click", (evt) => {
                evt.stopPropagation();
                closeMenu();
            });

            const confirmBtn = document.createElement("button");
            confirmBtn.textContent = "Save";
            confirmBtn.addEventListener("click", (evt) => {
                evt.stopPropagation();
                this.saveCurrentViewAsDefault();
                closeMenu();
            });

            actions.appendChild(cancelBtn);
            actions.appendChild(confirmBtn);
            confirmBox.appendChild(actions);
            wrapper.appendChild(confirmBox);
        });

        wrapper.appendChild(btn);
        return wrapper;
    }

    // -----------------------------------------------------------------
    // Item 1: "Calculations" toolbar section — add a calculated column
    // -----------------------------------------------------------------

    private renderCalculationsSection(menu: HTMLDivElement, closeMenu: () => void): void {
        const section = document.createElement("div");
        section.className = "skiba-toolbar__section";

        const title = document.createElement("div");
        title.className = "skiba-toolbar__section-title";
        title.textContent = "Calculations";
        section.appendChild(title);

        this._calcColumns.forEach((_def, name) => {
            section.appendChild(this.makeVirtualColumnListItem(name, `Remove the calculated column "${name}"? This can't be undone.`, () => {
                this._calcColumns.delete(name);
                this.recomputeVirtualColumns();
                this.computeColumnStats();
                this.applyPipeline();
                this.render();
                this.persistUserConfig();
            }));
        });

        const addBtn = document.createElement("button");
        addBtn.className = "skiba-toolbar__button";
        addBtn.textContent = "+ Add a calculated column";
        section.appendChild(addBtn);

        const form = document.createElement("div");
        form.className = "skiba-calc-form";
        form.style.display = "none";

        const nameInput = document.createElement("input");
        nameInput.type = "text";
        nameInput.placeholder = "Column name, e.g. Margin";
        nameInput.className = "skiba-calc-form__input";
        nameInput.setAttribute("aria-label", "Calculated column name");
        form.appendChild(nameInput);

        const formulaInput = document.createElement("textarea");
        formulaInput.placeholder = "e.g. Revenue - Cost";
        formulaInput.className = "skiba-calc-form__formula";
        formulaInput.rows = 2;
        formulaInput.setAttribute("aria-label", "Calculation formula");
        form.appendChild(formulaInput);

        const errorMsg = document.createElement("div");
        errorMsg.className = "skiba-calc-form__error";
        errorMsg.setAttribute("role", "alert");
        form.appendChild(errorMsg);

        const applyBtn = document.createElement("button");
        applyBtn.className = "skiba-toolbar__button skiba-toolbar__button--primary";
        applyBtn.textContent = "Add column";
        applyBtn.addEventListener("click", (evt) => {
            evt.stopPropagation();
            errorMsg.textContent = "";

            const name = nameInput.value.trim();
            if (name.length === 0) {
                errorMsg.textContent = "Give the calculated column a name.";
                return;
            }
            if (this.columns.some((c) => c.name === name)) {
                errorMsg.textContent = `"${name}" is already a column name \u2014 choose a different name.`;
                return;
            }

            const parsed = parseCalcFormula(formulaInput.value);
            if (!parsed.ok) {
                errorMsg.textContent = parsed.error ?? "Check your formula \u2014 for example: Revenue - Cost";
                return;
            }
            const unknown = parsed.referencedColumns.filter(
                (ref) => !this.columns.some((c) => c.displayName === ref || c.name === ref)
            );
            if (unknown.length > 0) {
                errorMsg.textContent = `Check your formula \u2014 I don't see a column called "${unknown[0]}".`;
                return;
            }

            this._calcColumns.set(name, { formula: formulaInput.value.trim(), parsed });
            this.recomputeVirtualColumns();
            this.computeColumnStats();
            this.applyPipeline();
            this.render();
            this.persistUserConfig();
            closeMenu();
        });
        form.appendChild(applyBtn);

        addBtn.addEventListener("click", (evt) => {
            evt.stopPropagation();
            form.style.display = form.style.display === "none" ? "block" : "none";
        });
        [nameInput, formulaInput].forEach((el) => el.addEventListener("click", (evt) => evt.stopPropagation()));

        section.appendChild(form);
        menu.appendChild(section);
        menu.appendChild(this.makeDivider());
    }

    // -----------------------------------------------------------------
    // Item 4: "Combine columns" toolbar section
    // -----------------------------------------------------------------

    private renderCombineColumnsSection(menu: HTMLDivElement, closeMenu: () => void): void {
        const section = document.createElement("div");
        section.className = "skiba-toolbar__section";

        const title = document.createElement("div");
        title.className = "skiba-toolbar__section-title";
        title.textContent = "Combined columns";
        section.appendChild(title);

        this._combinedColumns.forEach((_def, name) => {
            section.appendChild(this.makeVirtualColumnListItem(name, `Remove the combined column "${name}"? This can't be undone.`, () => {
                this._combinedColumns.delete(name);
                this.recomputeVirtualColumns();
                this.applyPipeline();
                this.render();
                this.persistUserConfig();
            }));
        });

        const addBtn = document.createElement("button");
        addBtn.className = "skiba-toolbar__button";
        addBtn.textContent = "+ Combine columns";
        section.appendChild(addBtn);

        const form = document.createElement("div");
        form.className = "skiba-calc-form";
        form.style.display = "none";

        const nameInput = document.createElement("input");
        nameInput.type = "text";
        nameInput.placeholder = "Column name, e.g. Full Name";
        nameInput.className = "skiba-calc-form__input";
        nameInput.setAttribute("aria-label", "Combined column name");
        form.appendChild(nameInput);

        const templateInput = document.createElement("input");
        templateInput.type = "text";
        templateInput.placeholder = "e.g. {FirstName} {LastName}";
        templateInput.className = "skiba-calc-form__input";
        templateInput.setAttribute("aria-label", "Combine template");
        form.appendChild(templateInput);

        const errorMsg = document.createElement("div");
        errorMsg.className = "skiba-calc-form__error";
        errorMsg.setAttribute("role", "alert");
        form.appendChild(errorMsg);

        const applyBtn = document.createElement("button");
        applyBtn.className = "skiba-toolbar__button skiba-toolbar__button--primary";
        applyBtn.textContent = "Combine";
        applyBtn.addEventListener("click", (evt) => {
            evt.stopPropagation();
            errorMsg.textContent = "";

            const name = nameInput.value.trim();
            if (name.length === 0) {
                errorMsg.textContent = "Give the combined column a name.";
                return;
            }
            if (this.columns.some((c) => c.name === name)) {
                errorMsg.textContent = `"${name}" is already a column name \u2014 choose a different name.`;
                return;
            }

            const template = templateInput.value;
            const refs = Array.from(template.matchAll(/\{([^}]+)\}/g)).map((m) => m[1].trim());
            if (refs.length < 2) {
                errorMsg.textContent = "Reference at least 2 columns, e.g. {FirstName} {LastName}.";
                return;
            }
            const unknown = refs.filter((ref) => !this.columns.some((c) => c.displayName === ref || c.name === ref));
            if (unknown.length > 0) {
                errorMsg.textContent = `I don't see a column called "${unknown[0]}" \u2014 check the spelling.`;
                return;
            }

            this._combinedColumns.set(name, { template });
            this.recomputeVirtualColumns();
            this.applyPipeline();
            this.render();
            this.persistUserConfig();
            closeMenu();
        });
        form.appendChild(applyBtn);

        addBtn.addEventListener("click", (evt) => {
            evt.stopPropagation();
            form.style.display = form.style.display === "none" ? "block" : "none";
        });
        [nameInput, templateInput].forEach((el) => el.addEventListener("click", (evt) => evt.stopPropagation()));

        section.appendChild(form);
        menu.appendChild(section);
        menu.appendChild(this.makeDivider());
    }

    /** Shared list-item row (name + delete) used by both the Calculations and Combined columns sections. */
    private makeVirtualColumnListItem(name: string, confirmMessage: string, onDelete: () => void): HTMLDivElement {
        const row = document.createElement("div");
        row.className = "skiba-calc-item";

        const label = document.createElement("span");
        label.className = "skiba-calc-item__name";
        label.textContent = name;
        row.appendChild(label);

        const del = document.createElement("button");
        del.className = "skiba-calc-item__delete";
        del.textContent = "\u2715";
        del.setAttribute("aria-label", `Remove ${name}`);
        del.addEventListener("click", (evt) => {
            evt.stopPropagation();
            // Destructive (discards a saved calculation) — confirm before applying, per the design charter.
            if (window.confirm(confirmMessage)) {
                onDelete();
            }
        });
        row.appendChild(del);
        return row;
    }

    private makeDivider(): HTMLDivElement {
        const divider = document.createElement("div");
        divider.className = "skiba-toolbar__divider";
        return divider;
    }

    private makeMenuButton(label: string, onClick: () => void): HTMLButtonElement {
        const btn = document.createElement("button");
        btn.className = "skiba-toolbar__button";
        btn.textContent = label;
        btn.setAttribute("role", "menuitem");
        btn.addEventListener("click", (evt) => {
            evt.stopPropagation();
            onClick();
        });
        return btn;
    }

    // -----------------------------------------------------------------
    // Search
    // -----------------------------------------------------------------

    private renderSearchBar(): void {
        this.clearElement(this.searchRoot);
        this.searchRoot.style.display = this.settings.searchEnabled ? "" : "none";
        if (!this.settings.searchEnabled) {
            return;
        }

        const input = document.createElement("input");
        input.type = "text";
        input.className = "skiba-search__input";
        input.placeholder = this.loc("Search_Placeholder", "Search this table");
        input.setAttribute("aria-label", this.loc("Search_AriaLabel", "Search this table"));
        input.value = this._searchTerm;
        let clear: HTMLButtonElement | null = null;
        if (this.settings.showSearchClear) {
            clear = document.createElement("button");
            clear.type = "button";
            clear.className = "skiba-search__clear";
            clear.textContent = "Clear";
            clear.disabled = this._searchTerm.length === 0;
            clear.setAttribute("aria-label", this.loc("Search_ClearAriaLabel", "Clear search"));
            clear.addEventListener("click", () => {
                this._searchTerm = "";
                input.value = "";
                clear!.disabled = true;
                this.applyPipeline();
                this.renderVisibleRows();
                this.renderStatusLine();
            });
        }
        input.addEventListener("input", () => {
            this._searchTerm = input.value;
            if (clear) clear.disabled = this._searchTerm.length === 0;
            this.applyPipeline();
            this.renderVisibleRows();
            this.renderStatusLine();
        });
        this.searchRoot.appendChild(input);
        if (clear) this.searchRoot.appendChild(clear);

        const status = document.createElement("span");
        status.className = "skiba-search__status";
        this.searchRoot.appendChild(status);
        this.renderStatusLine();
    }

    private renderStatusLine(): void {
        const status = this.searchRoot.querySelector<HTMLSpanElement>(".skiba-search__status");
        if (!status) {
            return;
        }
        this.clearElement(status);

        if (this._advancedFilterError) {
            const error = document.createElement("span");
            error.className = "skiba-search__advanced-filter-error";
            error.setAttribute("role", "alert");
            error.textContent = this.loc("Filter_AdvancedInvalid", "Advanced filter ignored: {0}", this._advancedFilterError);
            status.appendChild(error);
            if (typeof this.host.displayWarningIcon === "function") {
                this.host.displayWarningIcon(
                    this.loc("Filter_AdvancedInvalidTitle", "Advanced filter ignored"),
                    this._advancedFilterError
                );
            }
        }

        if (this._searchTerm.trim().length === 0 && this._columnFilters.size === 0) {
            return;
        }

        const matchText = document.createElement("span");
        matchText.textContent = this.loc(
            "Search_StatusMatch",
            "{0} of {1} rows match",
            String(this._filteredData.length),
            String(this._data.length)
        );
        status.appendChild(matchText);

        if (this._filteredData.length === 0) {
            const empty = document.createElement("span");
            empty.className = "skiba-search__empty-state";
            empty.textContent = this._searchTerm.trim().length > 0
                ? this.loc("Search_NoMatchesSearch", "No records match your search.")
                : this.loc("Search_NoMatchesFilter", "No records match the current filters.");
            status.appendChild(empty);
            const reset = document.createElement("button");
            reset.type = "button";
            reset.className = "skiba-search__reset";
            reset.textContent = this.loc("Search_Reset", "Reset");
            reset.addEventListener("click", () => {
                this._searchTerm = "";
                this._columnFilters.clear();
                this.commitFilterChange();
                this.renderSearchBar();
            });
            status.appendChild(reset);
        }

        if (this._searchTerm.trim().length > 0 && this._hasMoreData && !this.isFetchMoreDataRestricted()) {
            const link = document.createElement("button");
            link.type = "button";
            link.className = "skiba-search__full-dataset-link";
            link.textContent = this.loc("Search_FullDataset", "Search all available rows");
            link.setAttribute("aria-label", this.loc("Search_FullDatasetAriaLabel", "Search rows currently available to this visual"));
            link.addEventListener("click", () => this.beginForceFetchAll("search"));
            status.appendChild(document.createTextNode(" \u2014 "));
            status.appendChild(link);
        }
    }

    // -----------------------------------------------------------------
    // Per-column filters (header-driven popover) + filter chip strip
    // -----------------------------------------------------------------

    private inferFilterType(col: ITableColumn): FilterType {
        const sample = this._data.find((r) => r.values[col.name] !== null && r.values[col.name] !== undefined);
        const v = sample ? sample.values[col.name] : undefined;
        if (typeof v === "number") {
            return "number";
        }
        if (v instanceof Date) {
            return "date";
        }
        return "text";
    }

    private openFilterPopover(col: ITableColumn, anchor: HTMLElement): void {
        this.container.querySelectorAll(".skiba-filter-popover").forEach((el) => el.remove());

        const type = this.inferFilterType(col);
        const existing = this._columnFilters.get(col.name);
        const popover = document.createElement("div");
        popover.className = "skiba-filter-popover";

        const applyAndCommit = (filter: IColumnFilter | null): void => {
            if (filter) {
                this._columnFilters.set(col.name, filter);
            } else {
                this._columnFilters.delete(col.name);
            }
            this.commitFilterChange();
        };

        if (type === "text") {
            const input = document.createElement("input");
            input.type = "text";
            input.placeholder = this.loc("Filter_ContainsPlaceholder", "Contains...");
            input.value = existing?.value ?? "";
            popover.appendChild(input);

            const apply = (): void => {
                applyAndCommit(input.value.trim().length === 0 ? null : { type: "text", operator: "contains", value: input.value });
            };
            input.addEventListener("keydown", (e) => {
                if (e.key === "Enter") {
                    apply();
                    popover.remove();
                }
            });
            popover.appendChild(this.filterActionsRow(apply, () => applyAndCommit(null), popover));
        } else if (type === "number") {
            const opSelect = document.createElement("select");
            const opLabels: Record<FilterOperator, string> = {
                equals: this.loc("Filter_OpEquals", "="),
                gt: this.loc("Filter_OpGt", ">"),
                gte: this.loc("Filter_OpGte", "\u2265"),
                lt: this.loc("Filter_OpLt", "<"),
                lte: this.loc("Filter_OpLte", "\u2264"),
                between: this.loc("Filter_OpBetween", "between"),
                contains: this.loc("Filter_OpContains", "contains")
            };
            (["equals", "gt", "gte", "lt", "lte", "between"] as FilterOperator[]).forEach((op) => {
                const o = document.createElement("option");
                o.value = op;
                o.textContent = opLabels[op];
                if (existing?.operator === op) {
                    o.selected = true;
                }
                opSelect.appendChild(o);
            });
            popover.appendChild(opSelect);

            const val1 = document.createElement("input");
            val1.type = "number";
            val1.value = existing?.value ?? "";
            popover.appendChild(val1);

            const val2 = document.createElement("input");
            val2.type = "number";
            val2.placeholder = this.loc("Filter_AndPlaceholder", "and");
            val2.value = existing?.value2 ?? "";
            val2.style.display = opSelect.value === "between" ? "" : "none";
            popover.appendChild(val2);

            opSelect.addEventListener("change", () => {
                val2.style.display = opSelect.value === "between" ? "" : "none";
            });

            const apply = (): void => {
                applyAndCommit(val1.value.trim().length === 0 ? null : {
                    type: "number",
                    operator: opSelect.value as FilterOperator,
                    value: val1.value,
                    value2: val2.value
                });
            };
            popover.appendChild(this.filterActionsRow(apply, () => applyAndCommit(null), popover));
        } else {
            const from = document.createElement("input");
            from.type = "date";
            from.value = existing?.value ?? "";
            popover.appendChild(from);
            const to = document.createElement("input");
            to.type = "date";
            to.value = existing?.value2 ?? "";
            popover.appendChild(to);

            const apply = (): void => {
                applyAndCommit((!from.value && !to.value) ? null : { type: "date", operator: "between", value: from.value, value2: to.value });
            };
            popover.appendChild(this.filterActionsRow(apply, () => applyAndCommit(null), popover));
        }

        const rect = anchor.getBoundingClientRect();
        const containerRect = this.container.getBoundingClientRect();
        const popoverWidth = 236;
        const popoverHeight = 150;
        const gap = 6;
        const containerWidth = Math.max(0, containerRect.width);
        const containerHeight = Math.max(0, containerRect.height);
        const rawLeft = rect.left - containerRect.left;
        const left = Math.max(8, Math.min(rawLeft, Math.max(8, containerWidth - popoverWidth - 8)));
        const below = rect.bottom - containerRect.top + gap;
        const above = rect.top - containerRect.top - popoverHeight - gap;
        const top = (below + popoverHeight <= containerHeight - 8 || above < 8) ? below : above;
        popover.style.left = `${left}px`;
        popover.style.top = `${Math.max(8, top)}px`;
        popover.style.zIndex = "1000";
        this.container.appendChild(popover);

        const dismiss = (evt: MouseEvent): void => {
            if (!popover.contains(evt.target as Node) && evt.target !== anchor) {
                popover.remove();
                document.removeEventListener("click", dismiss);
                document.removeEventListener("keydown", onKeyDown);
            }
        };
        const onKeyDown = (evt: KeyboardEvent): void => {
            if (evt.key === "Escape") {
                evt.preventDefault();
                popover.remove();
                document.removeEventListener("click", dismiss);
                document.removeEventListener("keydown", onKeyDown);
            }
        };
        setTimeout(() => {
            document.addEventListener("click", dismiss);
            document.addEventListener("keydown", onKeyDown);
            const firstControl = popover.querySelector("input, select, button") as HTMLElement | null;
            firstControl?.focus();
        }, 0);
    }

    private filterActionsRow(onApply: () => void, onClear: () => void, popover: HTMLDivElement): HTMLDivElement {
        const row = document.createElement("div");
        row.className = "skiba-filter-popover__actions";
        const clearBtn = document.createElement("button");
        clearBtn.textContent = this.loc("Filter_Clear", "Clear");
        clearBtn.addEventListener("click", () => {
            onClear();
            popover.remove();
        });
        const applyBtn = document.createElement("button");
        applyBtn.textContent = this.loc("Filter_Apply", "Apply");
        applyBtn.addEventListener("click", () => {
            onApply();
            popover.remove();
        });
        row.appendChild(clearBtn);
        row.appendChild(applyBtn);
        return row;
    }

    private commitFilterChange(): void {
        this._customEvents.emit("filtersChanged", { columnFilterCount: this._columnFilters.size });
        this.applyPipeline();
        this.renderHeader();
        this.renderFilterChips();
        this.renderVisibleRows();
        this.renderStatusLine();
    }

    private renderFilterChips(): void {
        this.clearElement(this.filterChipsRoot);
        if (this._columnFilters.size === 0) {
            this.filterChipsRoot.style.display = "none";
            return;
        }
        this.filterChipsRoot.style.display = "";

        this._columnFilters.forEach((filter, colName) => {
            const col = this.columns.find((c) => c.name === colName);
            const label = col ? col.displayName : colName;
            const opLabels: Record<FilterOperator, string> = {
                equals: this.loc("Filter_OpEquals", "="),
                gt: this.loc("Filter_OpGt", ">"),
                gte: this.loc("Filter_OpGte", "\u2265"),
                lt: this.loc("Filter_OpLt", "<"),
                lte: this.loc("Filter_OpLte", "\u2264"),
                between: this.loc("Filter_OpBetween", "between"),
                contains: this.loc("Filter_OpContains", "contains")
            };
            const desc = filter.type === "text"
                ? `${opLabels.contains} "${filter.value}"`
                : filter.type === "date"
                    ? `${filter.value || "..."} \u2192 ${filter.value2 || "..."}`
                    : `${opLabels[filter.operator]} ${filter.value}${filter.operator === "between" ? ` ${this.loc("Filter_AndPlaceholder", "and")} ${filter.value2}` : ""}`;

            const chip = document.createElement("span");
            chip.className = "skiba-filter-chip";
            chip.textContent = `${label}: ${desc}`;

            const remove = document.createElement("button");
            remove.className = "skiba-filter-chip__remove";
            remove.textContent = "\u00D7";
            remove.setAttribute("aria-label", this.loc("Filter_RemoveAriaLabel", "Remove filter on {0}", label));
            remove.addEventListener("click", () => {
                this._columnFilters.delete(colName);
                this.commitFilterChange();
            });
            chip.appendChild(remove);
            this.filterChipsRoot.appendChild(chip);
        });
    }

    // -----------------------------------------------------------------
    // Header (sorting, resizing, drag-to-reorder, filter icon)
    // -----------------------------------------------------------------

    /**
     * Item 34 (Module D): called from visual.ts on every update() with the visual's own
     * viewport width. Only triggers a re-render when the narrow state (or, while narrow,
     * the available width) actually changed, so a same-size update doesn't do extra work.
     * Storing rather than immediately rendering on first call is deliberate -- this can run
     * before `settings` exists (resizeViewport runs early in updateInternal()); the value
     * is simply picked up by the first real render() once settings are populated.
     */
    public setNarrowLayout(isNarrow: boolean, availableWidth: number): void {
        const changed = isNarrow !== this._isNarrow || (isNarrow && availableWidth !== this._narrowAvailableWidth);
        this._isNarrow = isNarrow;
        this._narrowAvailableWidth = availableWidth;
        // Bugfix (item 34 loop): resizeViewport() -> setNarrowLayout() runs at the TOP of
        // updateInternal(), before this cycle's data has been parsed and handed to
        // setData(). Without this guard, a narrow-state flip on the very first render (or
        // any render where width crosses the breakpoint) forces an immediate re-render
        // against stale/empty `_renderNodes` from the previous cycle -- which can itself
        // trigger another host update(), producing an infinite update/resize loop. Only
        // re-render here if there's actually data to render; the real table render later
        // in this same updateInternal() cycle will pick up the new narrow state anyway.
        if (changed && this.settings && this._renderNodes && this._renderNodes.length > 0) {
            this.renderHeader();
            this.renderVisibleRows();
        }
    }

    private visibleColumns(): ITableColumn[] {
        const pivotName = this._quickGroupColumn?.name;
        const base = this._columnOrder
            .map((name) => this.columns.find((c) => c.name === name))
            .filter((c): c is ITableColumn => !!c && !this._hiddenColumns.has(c.name) && c.name !== pivotName);

        if (!this._isNarrow) {
            return base;
        }
        // Item 1 (Module D): on a phone-narrow canvas, auto-select a reduced, leading
        // subset of columns that actually fits, on top of whatever the user has already
        // hidden via the toolbar. This is purely a *display* selection layered on top of
        // `_hiddenColumns` -- it never mutates it, so widening the tile back out restores
        // every column with no user action, and the toolbar's own column checkboxes still
        // work exactly as before underneath it.
        return selectColumnsForWidth(base, (c) => this.columnWidth(c), this._narrowAvailableWidth);
    }

    private columnWidth(col: ITableColumn): number {
        const min = Math.max(48, this._columnMinWidths.get(col.name) ?? 80);
        const max = Math.max(min, this._columnMaxWidths.get(col.name) ?? 520);
        const mode = this._columnWidthModes.get(col.name) ?? "auto";
        let width = this._columnWidths.get(col.name) ?? 150;
        if (mode === "fixed") width = 150;
        if (mode === "fill") {
            const count = Math.max(1, this.visibleColumns().length);
            const available = this.container.clientWidth || 900;
            width = Math.floor(available / count);
        }
        if (mode === "fit") {
            // Bounded heuristic fit: sample loaded values instead of measuring thousands of DOM nodes.
            const samples = this._data.slice(0, 200).map((r) => formatDataValue(r.values[col.name], col, this.settings.regionalFormat ?? { locale: "en-US", currency: "", dateFormat: "", numberSeparators: "auto", decimalPrecision: -1 }));
            const maxChars = Math.max(col.displayName.length, ...samples.map((v) => v.length), 0);
            width = Math.min(max, Math.max(min, maxChars * 7 + 28));
        }
        return Math.max(min, Math.min(max, Math.round(width)));
    }

    private renderHeader(): void {
        this.clearElement(this.headerRoot);
        const row = document.createElement("div");
        row.className = "skiba-table__row skiba-table__row--header";
        row.setAttribute("role", "row");

        this.visibleColumns().forEach((col) => {
            const th = document.createElement("div");
            th.className = "skiba-table__cell skiba-table__cell--header";
            th.setAttribute("role", "columnheader");
            th.style.width = `${this.columnWidth(col)}px`;
            th.style.justifyContent = this.settings.headerAlignment === "right" ? "flex-end" : this.settings.headerAlignment === "center" ? "center" : "flex-start";
            th.style.whiteSpace = this.settings.headerWrap ? "normal" : "nowrap";
            th.tabIndex = 0;
            th.setAttribute("aria-sort", this.ariaSortFor(col.name));
            th.draggable = true;

            th.addEventListener("dragstart", (evt: DragEvent) => {
                evt.dataTransfer?.setData("text/skiba-column", col.name);
                th.classList.add("skiba-table__cell--dragging");
                // Item 2: reveal the "Group by" drop target only while a header is actually being dragged.
                if (!col.isMeasure) {
                    this.quickGroupDropRoot.style.display = "flex";
                }
            });
            th.addEventListener("dragend", () => {
                th.classList.remove("skiba-table__cell--dragging");
                this.quickGroupDropRoot.style.display = "none";
            });
            th.addEventListener("dragover", (evt: DragEvent) => evt.preventDefault());
            th.addEventListener("drop", (evt: DragEvent) => {
                evt.preventDefault();
                const draggedName = evt.dataTransfer?.getData("text/skiba-column");
                if (!draggedName || draggedName === col.name) {
                    return;
                }
                this.reorderColumn(draggedName, col.name);
            });

            const label = document.createElement("span");
            label.className = "skiba-table__header-label";
            label.textContent = col.displayName;
            th.appendChild(label);

            if (this.settings.showSortIndicator && this._sortState.column === col.name && this._sortState.direction !== "none") {
                const arrow = document.createElement("span");
                arrow.className = "skiba-table__sort-arrow";
                arrow.textContent = this._sortState.direction === "asc" ? "\u25B2" : "\u25BC";
                th.appendChild(arrow);
            }

            if (this.settings.enableColumnFilters && this.settings.showFilterIndicator && !this._nonFilterableColumns.has(col.name)) {
                const filterBtn = document.createElement("button");
                filterBtn.type = "button";
                filterBtn.className = "skiba-filter-icon";
                filterBtn.classList.toggle("skiba-filter-icon--active", this._columnFilters.has(col.name));
                filterBtn.textContent = "\u25BE";
                filterBtn.setAttribute("aria-label", this.loc("Filter_IconAriaLabel", "Filter {0}", col.displayName));
                filterBtn.setAttribute("title", this.loc("Filter_IconAriaLabel", "Filter {0}", col.displayName));
                filterBtn.addEventListener("click", (evt) => {
                    evt.preventDefault();
                    evt.stopPropagation();
                    this.openFilterPopover(col, filterBtn);
                });
                filterBtn.addEventListener("pointerdown", (evt) => evt.stopPropagation());
                th.appendChild(filterBtn);
            }

            // Full Keyboard Navigation (item 14): header cells are already tabbable and
            // Enter/Space cycles sort, matching the mouse click.
            const activate = (): void => { if (this.settings.enableSorting && !this._nonSortableColumns.has(col.name)) this.cycleSort(col.name); };
            label.addEventListener("click", activate);
            th.addEventListener("keydown", (evt: KeyboardEvent) => {
                if (evt.key === "Enter" || evt.key === " ") {
                    evt.preventDefault();
                    activate();
                }
            });

            // Item 8: "read-only" viewers lose the resize handle entirely (not just
            // disabled) alongside the export buttons and the save-default-view button.
            if (this.settings.permission !== "read-only") {
                const resizer = document.createElement("div");
                resizer.className = "skiba-resizer";
                resizer.setAttribute("aria-hidden", "true");
                this.attachResizeDrag(resizer, col);
                th.appendChild(resizer);
            }

            row.appendChild(th);
        });

        this.headerRoot.appendChild(row);
    }

    /** Moves `draggedName` to sit at `targetName`'s current position. */
    private reorderColumn(draggedName: string, targetName: string): void {
        const from = this._columnOrder.indexOf(draggedName);
        const to = this._columnOrder.indexOf(targetName);
        if (from === -1 || to === -1) {
            return;
        }
        this._columnOrder.splice(from, 1);
        this._columnOrder.splice(to, 0, draggedName);
        this.renderHeader();
        this.renderVisibleRows();
    }

    private ariaSortFor(columnName: string): "ascending" | "descending" | "none" {
        if (this._sortState.column !== columnName) {
            return "none";
        }
        if (this._sortState.direction === "asc") {
            return "ascending";
        }
        if (this._sortState.direction === "desc") {
            return "descending";
        }
        return "none";
    }

    /** Clicking a header cycles Ascending -> Descending -> None, with a visible arrow at every step. */
    private cycleSort(columnName: string): void {
        if (this._sortState.column !== columnName) {
            this._sortState = { column: columnName, direction: "asc" };
        } else if (this._sortState.direction === "asc") {
            this._sortState = { column: columnName, direction: "desc" };
        } else if (this._sortState.direction === "desc") {
            this._sortState = { column: null, direction: "none" };
        } else {
            this._sortState = { column: columnName, direction: "asc" };
        }
        this.applyPipeline();
        this.renderHeader();
        this.renderVisibleRows();
    }

    private resetSorts(): void {
        this._sortState = { column: null, direction: "none" };
        this.applyPipeline();
        this.renderHeader();
        this.renderVisibleRows();
    }

    private resetColumnWidths(): void {
        this._columnWidths.clear();
        this._columnWidthModes.clear();
        this._columnMinWidths.clear();
        this._columnMaxWidths.clear();
        this.renderHeader();
        this.renderVisibleRows();
    }

    private attachResizeDrag(handle: HTMLDivElement, col: ITableColumn): void {
        d3.select(handle).call(
            d3
                .drag<HTMLDivElement, unknown>()
                .on("start", (event: d3.D3DragEvent<HTMLDivElement, unknown, unknown>) => {
                    (event.sourceEvent as Event).stopPropagation();
                    handle.classList.add("skiba-resizer--active");
                })
                .on("drag", (event: d3.D3DragEvent<HTMLDivElement, unknown, unknown>) => {
                    const current = this.columnWidth(col);
                    const next = Math.max(60, current + event.dx);
                    this._columnWidths.set(col.name, next);
                    this._columnWidthModes.set(col.name, "custom");
                    this.renderHeader();
                    this.renderVisibleRows();
                })
                .on("end", () => {
                    handle.classList.remove("skiba-resizer--active");
                })
        );
    }

    // -----------------------------------------------------------------
    // Filter / search / sort / group pipeline
    // -----------------------------------------------------------------

    private matchesColumnFilter(row: ITableRow, colName: string, filter: IColumnFilter): boolean {
        const raw = row.values[colName];

        if (filter.type === "text") {
            if (raw === null || raw === undefined) {
                return false;
            }
            return String(raw).toLowerCase().includes(filter.value.toLowerCase());
        }

        if (filter.type === "number") {
            if (typeof raw !== "number") {
                return false;
            }
            const v1 = parseFloat(filter.value);
            switch (filter.operator) {
                case "between": {
                    const v2 = parseFloat(filter.value2 ?? filter.value);
                    return raw >= Math.min(v1, v2) && raw <= Math.max(v1, v2);
                }
                case "equals": return raw === v1;
                case "gt": return raw > v1;
                case "gte": return raw >= v1;
                case "lt": return raw < v1;
                case "lte": return raw <= v1;
                default: return true;
            }
        }

        // date
        const rawDate = raw instanceof Date ? raw : (typeof raw === "string" ? new Date(raw) : null);
        if (!rawDate || isNaN(rawDate.getTime())) {
            return false;
        }
        if (filter.value) {
            const from = new Date(filter.value);
            if (rawDate < from) {
                return false;
            }
        }
        if (filter.value2) {
            const to = new Date(filter.value2);
            if (rawDate > to) {
                return false;
            }
        }
        return true;
    }

    private applyPipeline(): void {
        let rows = this._data;

        const advancedText = (this.settings.advancedFilterExpression ?? "").trim();
        if (advancedText.length > 0) {
            const parsed = parseAdvancedFilter(advancedText);
            if (parsed.ok) {
                this._advancedFilterError = null;
                rows = applyAdvancedFilter(rows, parsed.expression);
            } else if (parsed.ok === false) {
                this._advancedFilterError = parsed.error;
            }
        } else {
            this._advancedFilterError = null;
        }

        this._columnFilters.forEach((filter, colName) => {
            rows = rows.filter((row) => this.matchesColumnFilter(row, colName, filter));
        });

        const term = this._searchTerm.trim().toLowerCase();
        if (term.length > 0) {
            rows = rows.filter((row) =>
                this.columns.some((col) => {
                    const v = row.values[col.name];
                    return v !== null && v !== undefined && String(v).toLowerCase().includes(term);
                })
            );
        }

        if (this._sortState.column && this._sortState.direction !== "none") {
            const col = this._sortState.column;
            const dir = this._sortState.direction === "asc" ? 1 : -1;
            rows = [...rows].sort((a, b) => {
                const av = a.values[col];
                const bv = b.values[col];
                if (av === null || av === undefined) return 1;
                if (bv === null || bv === undefined) return -1;
                if (typeof av === "number" && typeof bv === "number") {
                    return (av - bv) * dir;
                }
                return String(av).localeCompare(String(bv)) * dir;
            });
        }

        this._filteredData = rows;
        let baseNodes: RenderNode[];
        try {
            baseNodes = this.effectiveGroupColumns().length > 0
                ? this.buildGroupedNodes(rows)
                : rows.map((r) => ({ kind: "row", depth: 0, row: r } as RenderNode));
        } catch (error) {
            // Grouping failed on malformed data -- degrade to a flat, ungrouped view rather than
            // letting the exception reach Power BI's update() cycle and blank the whole tile.
            console.error("Data Lake Tables: grouping failed, falling back to flat rows.", error);
            baseNodes = rows.map((r) => ({ kind: "row", depth: 0, row: r } as RenderNode));
        }

        this._renderNodes = this.insertDetailNodes(baseNodes);
        this.computeNodeOffsets();
    }

    // -----------------------------------------------------------------
    // True drill-down: per-record detail sub-grid (Item 5)
    //
    // Distinct from group expand/collapse above: clicking the disclosure
    // control on a *leaf* row expands an inline sub-grid listing every
    // field of that underlying record -- including columns hidden via the
    // column-visibility toggle and Tooltip-role fields -- not just the
    // columns currently visible in the main grid.
    // -----------------------------------------------------------------

    /** Walks a flat node list and splices in a "detail" node immediately after each row node that's expanded. */
    private insertDetailNodes(nodes: RenderNode[]): RenderNode[] {
        if (this._expandedDetailRows.size === 0) {
            return nodes;
        }
        const out: RenderNode[] = [];
        nodes.forEach((n) => {
            out.push(n);
            if (n.kind === "row" && this._expandedDetailRows.has(n.row.key)) {
                out.push({ kind: "detail", depth: n.depth, row: n.row });
            }
        });
        return out;
    }

    /** Every field available for a record's detail view: all manageable columns plus Tooltip-role fields, deduped. */
    private allDetailColumns(): ITableColumn[] {
        const seen = new Set<string>();
        const list: ITableColumn[] = [];
        [...this.columns, ...this.tooltipColumns].forEach((c) => {
            if (!seen.has(c.name)) {
                seen.add(c.name);
                list.push(c);
            }
        });
        return list;
    }

    private detailRowHeight(): number {
        return this.allDetailColumns().length * this.detailFieldRowHeight + 16;
    }

    private toggleDetail(rowKey: string): void {
        if (this._expandedDetailRows.has(rowKey)) {
            this._expandedDetailRows.delete(rowKey);
        } else {
            this._expandedDetailRows.add(rowKey);
        }
        this.applyPipeline();
        this.renderVisibleRows();
    }

    /**
     * Cumulative pixel offset of every node in `_renderNodes`, since detail nodes break the
     * "every node is `defaultRowHeight` tall" assumption the virtual scroller otherwise relies on.
     */
    private computeNodeOffsets(): void {
        const offsets: number[] = [];
        let acc = 0;
        this._renderNodes.forEach((n) => {
            offsets.push(acc);
            acc += n.kind === "detail" ? this.detailRowHeight() : this.defaultRowHeight;
        });
        this._nodeOffsets = offsets;
        this._totalContentHeight = acc;
    }

    /** Binary search: index of the last node whose offset is <= target. */
    private findNodeIndexAtOffset(target: number): number {
        const offsets = this._nodeOffsets;
        let lo = 0;
        let hi = offsets.length - 1;
        let ans = 0;
        while (lo <= hi) {
            const mid = (lo + hi) >> 1;
            if (offsets[mid] <= target) {
                ans = mid;
                lo = mid + 1;
            } else {
                hi = mid - 1;
            }
        }
        return ans;
    }

    // -----------------------------------------------------------------
    // Grouping / drill-down
    // -----------------------------------------------------------------

    private bucketRows(rows: ITableRow[], col: ITableColumn): Map<string, ITableRow[]> {
        const buckets = new Map<string, ITableRow[]>();
        rows.forEach((r) => {
            const raw = r.values[col.name];
            const key = raw === null || raw === undefined ? "(blank)" : String(raw);
            const arr = buckets.get(key);
            if (arr) {
                arr.push(r);
            } else {
                buckets.set(key, [r]);
            }
        });
        return buckets;
    }

    /** Recursively groups by each "Group by" role column, in order, producing a flat list of group + row nodes. */
    private buildGroupedNodes(rows: ITableRow[]): RenderNode[] {
        const nodes: RenderNode[] = [];
        const groupCols = this.effectiveGroupColumns();

        const recurse = (subRows: ITableRow[], depth: number, prefix: string): void => {
            if (depth >= groupCols.length) {
                subRows.forEach((r) => nodes.push({ kind: "row", depth, row: r }));
                return;
            }

            const col = groupCols[depth];
            const buckets = this.bucketRows(subRows, col);

            buckets.forEach((bucketRows, key) => {
                const path = prefix + GROUP_SEP + col.name + "=" + key;

                const sums = new Map<string, number>();
                const minimums = new Map<string, number | null>();
                const maximums = new Map<string, number | null>();
                const aggregate = aggregateRows(bucketRows, this.valueColumns.map((vc) => vc.name));
                this.valueColumns.forEach((vc) => {
                    sums.set(vc.name, aggregate.sums[vc.name] ?? 0);
                    minimums.set(vc.name, aggregate.minimums[vc.name] ?? null);
                    maximums.set(vc.name, aggregate.maximums[vc.name] ?? null);
                });

                if (!this._groupExpansion.has(path)) {
                    this._groupExpansion.set(path, this.settings.groupsDefaultExpanded);
                }

                const rawValue = bucketRows[0].values[col.name];
                nodes.push({ kind: "group", depth, path, column: col, value: rawValue, count: bucketRows.length, sums, minimums, maximums });

                if (this._groupExpansion.get(path)) {
                    recurse(bucketRows, depth + 1, path);
                }
            });
        };

        recurse(rows, 0, "");
        return nodes;
    }

    private toggleGroup(path: string): void {
        const current = this._groupExpansion.get(path) ?? this.settings.groupsDefaultExpanded;
        this._groupExpansion.set(path, !current);
        this.applyPipeline();
        this.renderVisibleRows();
    }

    /** Walks the full (unfiltered-by-collapse) group tree, forcing every path's expansion state. */
    private setAllGroupsExpansion(expanded: boolean): void {
        const groupCols = this.effectiveGroupColumns();
        const walk = (rows: ITableRow[], depth: number, prefix: string): void => {
            if (depth >= groupCols.length) {
                return;
            }
            const col = groupCols[depth];
            const buckets = this.bucketRows(rows, col);
            buckets.forEach((bucketRows, key) => {
                const path = prefix + GROUP_SEP + col.name + "=" + key;
                this._groupExpansion.set(path, expanded);
                walk(bucketRows, depth + 1, path);
            });
        };
        walk(this._filteredData, 0, "");
        this.applyPipeline();
        this.renderVisibleRows();
    }

    private expandAllGroups(): void {
        this.setAllGroupsExpansion(true);
    }

    private collapseAllGroups(): void {
        this.setAllGroupsExpansion(false);
    }

    // -----------------------------------------------------------------
    // Saved views -- "the report's default view" (Item 7)
    //
    // Persisted via host.persistProperties() into the `savedView` object,
    // which travels with the .pbix automatically (no backend, no browser
    // storage). Saving is always an explicit user action; loading the saved
    // view onto a freshly opened report happens exactly once, driven by
    // visual.ts calling applyPersistedSavedViewIfPresent() after the first
    // setData().
    // -----------------------------------------------------------------

    /** Snapshots the renderer's current live state into a plain, JSON-serializable object. */
    private buildViewStateFromCurrent(): ISavedViewState {
        const columnWidths: { [columnName: string]: number } = {};
        this._columnWidths.forEach((width, name) => {
            columnWidths[name] = width;
        });

        const groupExpansion: { [groupPath: string]: boolean } = {};
        this._groupExpansion.forEach((expanded, path) => {
            groupExpansion[path] = expanded;
        });

        return {
            sortColumn: this._sortState.column,
            sortDirection: this._sortState.direction,
            columnOrder: [...this._columnOrder],
            columnWidths,
            hiddenColumns: [...this._hiddenColumns],
            searchTerm: this._searchTerm,
            groupExpansion
        };
    }

    /** Applies a saved view snapshot onto live state and does a full re-render. Defensive against columns that no longer exist. */
    private applyViewState(state: ISavedViewState): void {
        const knownNames = new Set(this.columns.map((c) => c.name));

        const restoredOrder = (state.columnOrder || []).filter((name) => knownNames.has(name));
        this.columns.forEach((c) => {
            if (restoredOrder.indexOf(c.name) === -1) {
                restoredOrder.push(c.name);
            }
        });
        this._columnOrder = restoredOrder;

        this._columnWidths = new Map(
            Object.keys(state.columnWidths || {})
                .filter((name) => knownNames.has(name))
                .map((name) => [name, state.columnWidths[name]] as [string, number])
        );

        this._hiddenColumns = new Set((state.hiddenColumns || []).filter((name) => knownNames.has(name)));

        this._sortState = state.sortColumn && knownNames.has(state.sortColumn)
            ? { column: state.sortColumn, direction: state.sortDirection }
            : { column: null, direction: "none" };

        this._searchTerm = state.searchTerm || "";

        this._groupExpansion = new Map(Object.entries(state.groupExpansion || {}));

        this.applyPipeline();
        this.render();
    }

    /** Called once by visual.ts on the first update() after a report opens, if a default view was saved. No-op otherwise. */
    public applyPersistedSavedViewIfPresent(): void {
        if (this._persistedViewState) {
            this.applyViewState(this._persistedViewState);
        }
    }

    /** Persists the current live state as the report's default view. An explicit user action only -- never automatic. */
    private saveCurrentViewAsDefault(): void {
        const state = this.buildViewStateFromCurrent();
        this._persistedViewState = state;
        this.host.persistProperties({
            merge: [
                {
                    objectName: "savedView",
                    selector: null,
                    properties: { state: JSON.stringify(state) }
                }
            ]
        });
    }

    /** Discards live customizations and restores the persisted default view (or the visual's original defaults if none was ever saved). Always recoverable -- no confirmation needed. */
    private resetToDefaultView(): void {
        if (this._persistedViewState) {
            this.applyViewState(this._persistedViewState);
            return;
        }
        this.applyViewState({
            sortColumn: null,
            sortDirection: "none",
            columnOrder: this.columns.map((c) => c.name),
            columnWidths: {},
            hiddenColumns: [],
            searchTerm: "",
            groupExpansion: {}
        });
    }

    // -----------------------------------------------------------------
    // Virtual scrolling body
    // -----------------------------------------------------------------

    private renderVisibleRows(): void {
        const rowHeight = this.defaultRowHeight;
        const nodes = this._renderNodes;
        const totalRows = nodes.length;

        if (totalRows === 0) {
            this.clearElement(this.bodyRoot);
            const filtered = this._searchTerm.trim().length > 0 || this._columnFilters.size > 0 || !!this._advancedFilterError;
            const card = document.createElement("div");
            card.className = "skiba-table__empty-card";
            const title = document.createElement("strong");
            title.textContent = filtered
                ? this.loc("Empty_NoRowsTitle", "No results found")
                : this.loc("Empty_NoDataTitle", "No data to display");
            const detail = document.createElement("span");
            detail.textContent = filtered
                ? this.loc("Empty_NoRowsMatch", "Your current search or filters returned 0 records.")
                : this.loc("Empty_NoDataHelp", "Add fields to Rows or Values to begin exploring your data.");
            card.append(title, detail);
            if (filtered) {
                const reset = document.createElement("button");
                reset.type = "button";
                reset.className = "skiba-search__reset";
                reset.textContent = this.loc("Search_Reset", "Reset filters");
                reset.addEventListener("click", () => {
                    this._searchTerm = "";
                    this._columnFilters.clear();
                    this.commitFilterChange();
                });
                card.appendChild(reset);
            }
            this.bodyRoot.appendChild(card);
            this.bodyRoot.style.height = "auto";
            return;
        }

        const viewportHeight = this.scrollRoot.clientHeight || 400;
        const scrollTop = this.scrollRoot.scrollTop;

        // Node heights aren't uniform once a detail sub-grid (Item 5) is expanded, so start/end
        // indices come from the precomputed offset table (computeNodeOffsets) rather than a
        // simple scrollTop / rowHeight division.
        let startIndex = 0;
        let endIndex = totalRows;
        if (this.settings.virtualScrollEnabled) {
            const buffer = this.settings.performanceMode === "maximum" ? MAXIMUM_ROW_BUFFER : this.settings.performanceMode === "balanced" ? BALANCED_ROW_BUFFER : DEFAULT_ROW_BUFFER;
            startIndex = Math.max(0, this.findNodeIndexAtOffset(scrollTop) - Math.floor(buffer / 2));
            endIndex = Math.min(totalRows, this.findNodeIndexAtOffset(scrollTop + viewportHeight) + Math.ceil(buffer / 2) + 1);
        }

        const topSpacerHeight = this._nodeOffsets[startIndex] ?? 0;
        const bottomSpacerHeight = this._totalContentHeight - (endIndex < totalRows ? this._nodeOffsets[endIndex] : this._totalContentHeight);

        this.clearElement(this.bodyRoot);
        this.bodyRoot.style.position = "relative";

        const topSpacer = document.createElement("div");
        topSpacer.style.height = `${topSpacerHeight}px`;
        topSpacer.style.flexShrink = "0";
        this.bodyRoot.appendChild(topSpacer);

        const visibleColumns = this.visibleColumns();
        const selectedIds = this.selectionManager.getSelectionIds() as ISelectionId[];

        for (let i = startIndex; i < endIndex; i++) {
            const node = nodes[i];
            if (node.kind === "group") {
                this.bodyRoot.appendChild(this.renderGroupRow(node, visibleColumns, rowHeight));
            } else if (node.kind === "detail") {
                this.bodyRoot.appendChild(this.renderDetailRow(node.row, node.depth));
            } else {
                const isDetailExpanded = this._expandedDetailRows.has(node.row.key);
                this.bodyRoot.appendChild(this.renderRow(node.row, node.depth, i, visibleColumns, selectedIds, rowHeight, isDetailExpanded));
            }
        }

        const bottomSpacer = document.createElement("div");
        bottomSpacer.style.height = `${bottomSpacerHeight}px`;
        bottomSpacer.style.flexShrink = "0";
        this.bodyRoot.appendChild(bottomSpacer);

        if (this.settings.showTotals) {
            this.bodyRoot.appendChild(this.renderTotalsRow(visibleColumns, rowHeight));
        }
    }

    private renderGroupRow(node: Extract<RenderNode, { kind: "group" }>, visibleColumns: ITableColumn[], rowHeight: number): HTMLDivElement {
        const rowEl = document.createElement("div");
        rowEl.className = "skiba-table__row skiba-table__row--group";
        rowEl.style.height = `${rowHeight}px`;
        rowEl.setAttribute("role", "row");

        const isExpanded = this._groupExpansion.get(node.path) ?? this.settings.groupsDefaultExpanded;

        const chevron = document.createElement("span");
        chevron.className = "skiba-group__chevron";
        chevron.textContent = isExpanded ? "\u25BC" : "\u25B6";
        chevron.style.marginLeft = `${node.depth * Math.max(0, this.settings.groupIndentation ?? 16)}px`;

        const label = document.createElement("span");
        label.className = "skiba-group__label";
        const valueText = node.value === null || node.value === undefined ? this.loc("Group_Blank", "(blank)") : String(node.value);
        label.textContent = `${node.column.displayName}: ${valueText}${this.settings.showGroupCount !== false ? ` (${node.count})` : ""}`;

        const head = document.createElement("div");
        head.className = "skiba-table__cell skiba-table__cell--group-label";
        head.appendChild(chevron);
        head.appendChild(label);
        rowEl.appendChild(head);

        // Item 5: the group row is the disclosure control for its detail sub-grid — keyboard
        // accessible and ARIA-labelled, matching the pattern already used by header sort buttons.
        // Full Keyboard Navigation (item 14): group rows toggle expand/collapse via Enter/Space,
        // like the header sort cells and data rows below. Local UI state (not a cross-filter
        // selection), so it stays available even when allowInteractions is false.
        rowEl.tabIndex = 0;
        rowEl.setAttribute("role", "button");
        rowEl.setAttribute("aria-expanded", String(isExpanded));
        rowEl.setAttribute("aria-label", `${node.column.displayName}: ${valueText}, ${node.count} records, ${isExpanded ? "expanded" : "collapsed"}`);
        rowEl.addEventListener("click", () => this.toggleGroup(node.path));
        rowEl.addEventListener("keydown", (evt: KeyboardEvent) => {
            if (evt.key === "Enter" || evt.key === " ") {
                evt.preventDefault();
                this.toggleGroup(node.path);
            }
        });

        // Aggregate sums for measure columns, aligned like a mini totals strip on the group row.
        if (this.settings.showGroupTotals !== false) {
            visibleColumns.filter((c) => c.isMeasure).forEach((col) => {
                const cell = document.createElement("div");
                cell.className = "skiba-table__cell skiba-table__cell--group-sum";
                cell.style.width = `${this.columnWidth(col)}px`;
                const sum = node.sums.get(col.name) ?? 0;
                cell.textContent = this.formatNumber(sum, col);
                rowEl.appendChild(cell);
            });
        }

        return rowEl;
    }

    private renderRow(
        row: ITableRow,
        depth: number,
        index: number,
        visibleColumns: ITableColumn[],
        selectedIds: ISelectionId[],
        rowHeight: number,
        isDetailExpanded: boolean = false
    ): HTMLDivElement {
        const rowEl = document.createElement("div");
        rowEl.className = "skiba-table__row";
        rowEl.style.height = `${rowHeight}px`;
        rowEl.setAttribute("role", "row");
        rowEl.classList.toggle("skiba-table__row--alt", index % 2 === 1);

        const rowColor = this.tier4RowColor(row);
        if (rowColor) {
            rowEl.style.backgroundColor = rowColor;
        }

        const isSelected = selectedIds.some((id) => id.equals(row.selectionId));
        rowEl.classList.toggle("skiba-table__row--selected", this.settings.showSelectionHighlight && isSelected);

        // Full Keyboard Navigation (item 14): rows are focusable and Enter/Space selects,
        // mirroring the click behavior below.
        rowEl.tabIndex = 0;
        rowEl.setAttribute("aria-selected", String(isSelected));

        const selectRow = (multiSelect: boolean): void => {
            if (!this.interactionsAllowed() || !this.settings.enableRowSelection) {
                return;
            }
            const useMulti = this.settings.selectionMode === "multi" && multiSelect;
            this.selectionManager.select(row.selectionId, useMulti).then(() => {
                this.renderVisibleRows();
            });
        };

        rowEl.addEventListener("click", (evt: MouseEvent) => {
            selectRow(evt.ctrlKey || evt.metaKey);
        });
        rowEl.addEventListener("keydown", (evt: KeyboardEvent) => {
            if (evt.key === "Enter" || evt.key === " ") {
                evt.preventDefault();
                selectRow(evt.ctrlKey || evt.metaKey);
            }
        });

        // Right-Click Context Menu (item 12): data-point mode. Native Power BI menu
        // (Include/Exclude/etc), surfaced via selectionManager -- not custom menu items.
        rowEl.addEventListener("contextmenu", (evt: MouseEvent) => {
            if (!this.interactionsAllowed() || !this.settings.enableRowSelection) {
                return;
            }
            evt.preventDefault();
            evt.stopPropagation();
            this.selectionManager.showContextMenu(row.selectionId, { x: evt.clientX, y: evt.clientY });
        });

        rowEl.addEventListener("mouseenter", (evt: MouseEvent) => this.showRowTooltip(row, evt));
        rowEl.addEventListener("mousemove", (evt: MouseEvent) => this.moveTooltip(evt));
        rowEl.addEventListener("mouseleave", () => this.hideTooltip());

        visibleColumns.forEach((col, idx) => {
            const cell = this.renderCell(row, col);
            if (idx === 0) {
                if (depth > 0) {
                    cell.style.paddingLeft = `${depth * Math.max(0, this.settings.groupIndentation ?? 16) + 8}px`;
                }
                cell.insertBefore(this.renderDetailToggle(row, isDetailExpanded), cell.firstChild);
            }
            rowEl.appendChild(cell);
        });
        this.appendCopyRowAction(rowEl, row);

        return rowEl;
    }

    /**
     * Disclosure control for a leaf row's full-record detail sub-grid (Item 5). Nested inside
     * the row's own focusable/clickable region, so it needs its own tabIndex/role and must stop
     * propagation on click and Enter/Space -- otherwise toggling detail would also fire the row's
     * cross-filter selection handler.
     */
    private renderDetailToggle(row: ITableRow, isExpanded: boolean): HTMLSpanElement {
        const toggle = document.createElement("span");
        toggle.className = "skiba-table__row-toggle";
        toggle.textContent = isExpanded ? "\u25BC" : "\u25B6";
        toggle.tabIndex = 0;
        toggle.setAttribute("role", "button");
        toggle.setAttribute("aria-expanded", String(isExpanded));
        toggle.setAttribute(
            "aria-label",
            isExpanded
                ? this.loc("Detail_Collapse", "Collapse record details")
                : this.loc("Detail_Expand", "Expand record details")
        );

        toggle.addEventListener("click", (evt: MouseEvent) => {
            evt.stopPropagation();
            this.toggleDetail(row.key);
        });
        toggle.addEventListener("keydown", (evt: KeyboardEvent) => {
            if (evt.key === "Enter" || evt.key === " ") {
                evt.preventDefault();
                evt.stopPropagation();
                this.toggleDetail(row.key);
            }
        });

        return toggle;
    }

    /**
     * The record-detail sub-grid itself: every field of the underlying row (including columns
     * hidden from the main grid and Tooltip-role fields) as a two-column Field/Value list.
     */
    private renderDetailRow(row: ITableRow, depth: number): HTMLDivElement {
        const wrap = document.createElement("div");
        wrap.className = "skiba-table__row skiba-table__row--detail";
        wrap.style.height = `${this.detailRowHeight()}px`;
        wrap.style.paddingLeft = `${depth * Math.max(0, this.settings.groupIndentation ?? 16) + 24}px`;
        wrap.setAttribute("role", "row");

        const grid = document.createElement("div");
        grid.className = "skiba-detail-grid";

        this.allDetailColumns().forEach((col) => {
            const field = document.createElement("div");
            field.className = "skiba-detail-grid__field";

            const label = document.createElement("span");
            label.className = "skiba-detail-grid__label";
            label.textContent = col.displayName;

            const raw = row.values[col.name];
            const value = document.createElement("span");
            value.className = "skiba-detail-grid__value";
            value.textContent = formatDataValue(raw, col, this.settings.regionalFormat ?? { locale: "en-US", currency: "", dateFormat: "", numberSeparators: "auto", decimalPrecision: -1 });

            field.appendChild(label);
            field.appendChild(value);
            grid.appendChild(field);
        });

        wrap.appendChild(grid);
        return wrap;
    }

    private tier4CellColor(columnName: string, rawValue: unknown): string | null {
        const columnColor = this._tier4ColumnColors.get(columnName);
        let result = columnColor ?? null;
        this._tier4Rules.forEach((rule) => {
            if (rule.column !== columnName || rule.scope === "row" || (rule.format && rule.format !== "background")) return;
            if (matchesTier4Rule(rawValue, rule)) result = rule.color;
        });
        return result;
    }
    private tier4RowColor(row: ITableRow): string | null {
        let result: string | null = null;
        this._tier4Rules.forEach((rule) => {
            if (rule.scope !== "row" || (rule.format && rule.format !== "background")) return;
            if (matchesTier4Rule(row.values[rule.column], rule)) {
                result = rule.color;
            }
        });
        return result;
    }

    /** A fully custom, self-contained dropdown -- button + listbox we render and position
     *  ourselves, so it can never escape the visual's own bounding box the way a native
     *  <select> popup can (that popup is rendered by the browser/OS, entirely outside our
     *  CSS's control). Returns { root, getValue } -- append `root` where the old <select>
     *  went, and read the current value via getValue() anywhere the old code read `.value`. */
    private createBoundedSelect(
        optionsList: Array<{ value: string; label: string }>,
        initialValue: string,
        ariaLabel: string,
        onChange: (value: string) => void
    ): { root: HTMLElement; getValue: () => string } {
        let currentValue = initialValue;
        const root = document.createElement("div");
        root.className = "skiba-bounded-select";

        const button = document.createElement("button");
        button.type = "button";
        button.className = "skiba-format-editor__select skiba-bounded-select__trigger";
        button.setAttribute("aria-haspopup", "listbox");
        button.setAttribute("aria-expanded", "false");
        button.setAttribute("aria-label", ariaLabel);
        const initial = optionsList.find((o) => o.value === currentValue) || optionsList[0];
        button.textContent = initial ? initial.label : "";
        root.appendChild(button);

        const list = document.createElement("ul");
        list.className = "skiba-bounded-select__list";
        list.setAttribute("role", "listbox");
        list.style.display = "none";
        optionsList.forEach((opt) => {
            const item = document.createElement("li");
            item.setAttribute("role", "option");
            item.className = "skiba-bounded-select__option";
            item.textContent = opt.label;
            item.dataset.value = opt.value;
            item.addEventListener("click", (evt) => {
                evt.stopPropagation();
                currentValue = opt.value;
                button.textContent = opt.label;
                list.style.display = "none";
                button.setAttribute("aria-expanded", "false");
                onChange(opt.value);
            });
            list.appendChild(item);
        });
        root.appendChild(list);

        button.addEventListener("click", (evt) => {
            evt.stopPropagation();
            const isOpen = list.style.display !== "none";
            list.style.display = isOpen ? "none" : "block";
            button.setAttribute("aria-expanded", String(!isOpen));
        });
        document.addEventListener("click", () => {
            list.style.display = "none";
            button.setAttribute("aria-expanded", "false");
        });

        return { root, getValue: () => currentValue };
    }
    private canCopyToClipboard(): boolean {
        return typeof navigator !== "undefined" && !!navigator.clipboard && typeof navigator.clipboard.writeText === "function";
    }

    private copyText(text: string): void {
        if (!this.canCopyToClipboard()) return;
        navigator.clipboard.writeText(text).then(() => {
            this._customEvents.emit("copyRequested", { kind: "text", length: text.length });
        }).catch(() => { /* Host/browser clipboard permissions can deny writes; keep UI quiet. */ });
    }

    private rowCopyText(row: ITableRow): string {
        return this.visibleColumns().map((col) => {
            const value = row.values[col.name];
            return value === null || value === undefined ? "" : String(value);
        }).join("\t");
    }

    private selectedRows(): ITableRow[] {
        const ids = this.selectionManager.getSelectionIds() as ISelectionId[];
        return this._data.filter((row) => ids.some((id) => id.equals(row.selectionId)));
    }

    private copySelectedRows(): void {
        const rows = this.selectedRows();
        if (rows.length === 0) return;
        this.copyText(rows.map((row) => this.rowCopyText(row)).join("\n"));
    }

    private appendCopyCellAction(cell: HTMLDivElement, row: ITableRow, col: ITableColumn): void {
        if (!this.settings.enableCopy || !this.settings.copyCell || !this.canCopyToClipboard()) return;
        const copy = document.createElement("button");
        copy.type = "button";
        copy.className = "skiba-copy-action skiba-copy-action--cell";
        copy.textContent = "Copy";
        copy.setAttribute("aria-label", this.loc("Copy_CellAriaLabel", "Copy {0}", col.displayName));
        copy.addEventListener("click", (evt) => {
            evt.preventDefault(); evt.stopPropagation();
            const value = row.values[col.name];
            this.copyText(value === null || value === undefined ? "" : String(value));
        });
        cell.appendChild(copy);
    }

    private appendCopyRowAction(rowEl: HTMLDivElement, row: ITableRow): void {
        if (!this.settings.enableCopy || !this.settings.copyRow || !this.canCopyToClipboard()) return;
        const firstCell = rowEl.querySelector<HTMLElement>(".skiba-table__cell");
        if (!firstCell) return;
        const copy = document.createElement("button");
        copy.type = "button"; copy.className = "skiba-copy-action skiba-copy-action--row"; copy.textContent = "Copy row";
        copy.setAttribute("aria-label", this.loc("Copy_RowAriaLabel", "Copy row"));
        copy.addEventListener("click", (evt) => { evt.preventDefault(); evt.stopPropagation(); this.copyText(this.rowCopyText(row)); });
        firstCell.appendChild(copy);
    }

    private renderCell(row: ITableRow, col: ITableColumn): HTMLDivElement {
        const cell = document.createElement("div");
        cell.className = "skiba-table__cell";
        cell.style.width = `${this.columnWidth(col)}px`;
        const columnAlignment = this._columnAlignments.get(col.name) ?? this.settings.cellAlignment ?? "auto";
        const resolvedAlignment = resolveCellAlignment(col, columnAlignment);
        cell.style.justifyContent = resolvedAlignment === "right" ? "flex-end" : resolvedAlignment === "center" ? "center" : "flex-start";
        cell.setAttribute("role", "cell");

        const rawValue = row.values[col.name];
        // Tier 5: `typeof rawValue === "number"` is true for NaN too -- that let malformed
        // numeric data (bad source values, failed calc-column results) reach the gradient
        // math and data-bar width calculation below as NaN, producing broken CSS instead of
        // a plain cell. isFiniteMeasure gates all three branches on a real, finite number.
        const isFiniteMeasure = col.isMeasure && typeof rawValue === "number" && Number.isFinite(rawValue);
        const text = document.createElement("span");
        text.className = "skiba-table__cell-text";
        text.textContent = safeTruncate(formatDataValue(rawValue, col, this.settings.regionalFormat ?? { locale: "en-US", currency: "", dateFormat: "", numberSeparators: "auto", decimalPrecision: -1 }));
        cell.appendChild(text);

        if (this.settings.performanceMode !== "maximum" && this._sparklineColumns.has(col.name) && isFiniteMeasure) {
            const spark = this.renderSparkline(col, row);
            if (spark) {
                cell.classList.add("skiba-table__cell--with-sparkline");
                cell.appendChild(spark);
            }
        }

        if (this.settings.conditionalFormatEnabled && isFiniteMeasure) {
            const range = this._columnMinMax.get(col.name);
            if (range && range.max > range.min) {
                const t = (rawValue - range.min) / (range.max - range.min);
                const scaleColor = this.settings.conditionalFormatMidpointEnabled === true
                    ? (t <= 0.5
                        ? d3.interpolateRgb(this.settings.conditionalFormatMinColor, (this.settings.conditionalFormatMidpointColor ?? "#FFF4CE"))(t * 2)
                        : d3.interpolateRgb(this.settings.conditionalFormatMidpointColor ?? "#FFF4CE", this.settings.conditionalFormatMaxColor)((t - 0.5) * 2))
                    : d3.interpolateRgb(this.settings.conditionalFormatMinColor, this.settings.conditionalFormatMaxColor)(t);
                cell.style.backgroundColor = this._tier4ColumnColors.get(col.name) ?? colorForTier4Value(rawValue, this._tier4Rules) ?? scaleColor;
            }
        }

        const tier4Color = this.tier4CellColor(col.name, rawValue);
        if (tier4Color) cell.style.backgroundColor = tier4Color;

        const matchingRules = this._tier4Rules.filter((rule) => rule.column === col.name && rule.scope !== "row" && matchesTier4Rule(rawValue, rule));
        const fontRule = matchingRules.find((rule) => rule.format === "font");
        if (fontRule) cell.style.color = fontRule.color;
        const iconRule = matchingRules.find((rule) => rule.format === "icon");
        if (iconRule) {
            const icon = document.createElement("span");
            icon.className = "skiba-table__condition-icon";
            icon.textContent = iconRule.icon === "down" ? "▼" : iconRule.icon === "dot" ? "●" : "▲";
            icon.style.color = iconRule.color;
            icon.setAttribute("aria-hidden", "true");
            cell.insertBefore(icon, text);
        }

        const wantsRuleBar = matchingRules.some((rule) => rule.format === "dataBar");
        if ((this.settings.enableDataBars || wantsRuleBar) && isFiniteMeasure) {
            const range = this._columnMinMax.get(col.name);
            if (range) {
                const layout = dataBarLayout(rawValue, range.min, range.max);
                if (layout.visible && layout.widthPct > 0) {
                    const bar = document.createElement("div");
                    bar.className = `skiba-table__data-bar${layout.negative ? " skiba-table__data-bar--negative" : ""}`;
                    bar.style.left = `${layout.leftPct}%`;
                    bar.style.width = `${layout.widthPct}%`;
                    const ruleBar = matchingRules.find((rule) => rule.format === "dataBar");
                    if (ruleBar) bar.style.backgroundColor = ruleBar.color;
                    cell.insertBefore(bar, text);
                }
            }
        }

        this.appendLinkActionIcon(cell, col, row);

        this.appendCopyCellAction(cell, row, col);
        return cell;
    }

    // -----------------------------------------------------------------
    // Item 3: sparklines
    // -----------------------------------------------------------------

    /**
     * Trend for a numeric column leading up to and including this row, using the trailing
     * window of rows in the current search/sort/filter order (falls back to the row's own
     * position when no "Group by" dimension is present to bucket a series by).
     */
    private renderSparkline(col: ITableColumn, row: ITableRow): SVGSVGElement | null {
        const idx = this._filteredData.indexOf(row);
        if (idx === -1) {
            return null;
        }

        const windowSize = 10;
        const start = Math.max(0, idx - windowSize + 1);
        const series = this._filteredData
            .slice(start, idx + 1)
            .map((r) => r.values[col.name])
            .filter((v): v is number => typeof v === "number");

        if (series.length < 2) {
            return null;
        }

        const w = 56;
        const h = 18;
        const min = Math.min(...series);
        const max = Math.max(...series);
        const range = max - min || 1;
        const denom = series.length - 1;

        const points = series
            .map((v, i) => {
                const x = (i / denom) * (w - 2) + 1;
                const y = h - 1 - ((v - min) / range) * (h - 2);
                return `${x.toFixed(1)},${y.toFixed(1)}`;
            })
            .join(" ");

        const svgNS = "http://www.w3.org/2000/svg";
        const svg = document.createElementNS(svgNS, "svg") as SVGSVGElement;
        svg.setAttribute("class", "skiba-sparkline");
        svg.setAttribute("viewBox", `0 0 ${w} ${h}`);
        svg.setAttribute("width", String(w));
        svg.setAttribute("height", String(h));
        svg.setAttribute("aria-hidden", "true"); // decorative — the exact value is already in the cell text
        svg.setAttribute("focusable", "false");

        const polyline = document.createElementNS(svgNS, "polyline");
        polyline.setAttribute("points", points);
        polyline.setAttribute("class", "skiba-sparkline__line");
        svg.appendChild(polyline);

        const lastValue = series[series.length - 1];
        const dot = document.createElementNS(svgNS, "circle");
        dot.setAttribute("cx", String(w - 1));
        dot.setAttribute("cy", String(h - 1 - ((lastValue - min) / range) * (h - 2)));
        dot.setAttribute("r", "1.5");
        dot.setAttribute("class", "skiba-sparkline__dot");
        svg.appendChild(dot);

        return svg;
    }

    // -----------------------------------------------------------------
    // Conditional URL actions (Item 9)
    //
    // Rules are authored as a JSON array in the formatting pane (Power BI's
    // format-pane API doesn't support arbitrary user-added rows in a repeating
    // UI), evaluated top-to-bottom, first match wins. Every resolved URL is
    // hard-checked to http(s) only and every substituted value is
    // encodeURIComponent()-escaped before host.launchUrl() is ever called --
    // launchUrl() is the supported Power BI URL-launch API path for a Power BI visual to
    // open a link and needs no special privilege declaration.
    // -----------------------------------------------------------------

    /** If `col` is the designated link-icon column and a rule matches this row, appends a small clickable link icon to `cell`. */
    private appendLinkActionIcon(cell: HTMLDivElement, col: ITableColumn, row: ITableRow): void {
        const iconColumn = (this.settings.linkActionIconColumn || "").trim();
        if (!iconColumn || col.name !== iconColumn) {
            return;
        }

        const rule = this.findMatchingLinkRule(row);
        if (!rule) {
            return;
        }

        const url = this.resolveLinkActionUrl(rule, row);
        if (!url) {
            return;
        }

        const linkBtn = document.createElement("button");
        linkBtn.className = "skiba-link-icon";
        linkBtn.textContent = "\u2197"; // ↗ — reuses the same plain Unicode-glyph icon pattern as the gear/chevron/sort-arrow icons elsewhere in this file, no new icon font/library
        linkBtn.setAttribute("aria-label", "Open linked page");
        linkBtn.title = "Open linked page";
        linkBtn.addEventListener("click", (evt) => {
            evt.stopPropagation();
            this.host.launchUrl(url);
        });
        cell.appendChild(linkBtn);
    }

    /** Evaluates configured link-action rules top-to-bottom against a row's values; returns the first match, or null. */
    private findMatchingLinkRule(row: ITableRow): ILinkActionRule | null {
        const rules = this.settings.linkActionRules || [];
        for (let i = 0; i < rules.length; i++) {
            if (this.evaluateLinkActionRule(rules[i], row)) {
                return rules[i];
            }
        }
        return null;
    }

    private evaluateLinkActionRule(rule: ILinkActionRule, row: ITableRow): boolean {
        const raw = row.values[rule.column];
        if (raw === null || raw === undefined) {
            return false;
        }

        if (rule.operator === "equals") {
            return String(raw) === rule.value;
        }
        if (rule.operator === "notEquals") {
            return String(raw) !== rule.value;
        }
        if (rule.operator === "contains") {
            return String(raw).toLowerCase().indexOf(rule.value.toLowerCase()) !== -1;
        }

        // gt / gte / lt / lte — numeric comparison
        const numRaw = typeof raw === "number" ? raw : parseFloat(String(raw));
        const numRule = parseFloat(rule.value);
        if (isNaN(numRaw) || isNaN(numRule)) {
            return false;
        }
        switch (rule.operator) {
            case "gt": return numRaw > numRule;
            case "gte": return numRaw >= numRule;
            case "lt": return numRaw < numRule;
            case "lte": return numRaw <= numRule;
            default: return false;
        }
    }

    /**
     * Resolves `{ColumnName}` placeholders in a rule's urlTemplate against this
     * row's values, URL-encoding every substituted value. Returns null (never
     * throws) if the resolved string's scheme isn't exactly http:// or https://,
     * which is a hard mandatory safety check, not a best-effort one.
     */
    private resolveLinkActionUrl(rule: ILinkActionRule, row: ITableRow): string | null {
        const resolved = rule.urlTemplate.replace(/\{([^{}]+)\}/g, (_match, columnName: string) => {
            const v = row.values[columnName];
            const strVal = v === null || v === undefined ? "" : String(v);
            return encodeURIComponent(strVal);
        });

        if (!/^https:\/\//i.test(resolved) && !/^http:\/\//i.test(resolved)) {
            return null;
        }
        return resolved;
    }

    private columnStatsMax(columnName: string): number {
        if (this.columnMaxCache.has(columnName)) {
            return this.columnMaxCache.get(columnName)!;
        }
        let max = 0;
        this._data.forEach((r) => {
            const v = r.values[columnName];
            if (typeof v === "number") {
                max = Math.max(max, Math.abs(v));
            }
        });
        this.columnMaxCache.set(columnName, max);
        return max;
    }

    private renderTotalsRow(visibleColumns: ITableColumn[], rowHeight: number): HTMLDivElement {
        const rowEl = document.createElement("div");
        rowEl.className = "skiba-table__row skiba-table__row--totals";
        rowEl.style.height = `${rowHeight}px`;
        rowEl.style.color = this.settings.totalsFont ?? this.settings.cellFont;
        rowEl.style.fontWeight = this.settings.totalsBold !== false ? "600" : "400";
        rowEl.style.borderTop = this.settings.totalsShowBorder !== false ? "2px solid rgba(0,0,0,.12)" : "0";

        visibleColumns.forEach((col, idx) => {
            const cell = document.createElement("div");
            cell.className = "skiba-table__cell";
            cell.style.width = `${this.columnWidth(col)}px`;
            cell.style.justifyContent = this.settings.totalsAlignment === "right" ? "flex-end" : this.settings.totalsAlignment === "center" ? "center" : "flex-start";

            if (idx === 0) {
                cell.textContent = this.settings.totalsLabel;
            } else if (col.isMeasure) {
                const sum = d3.sum(this._filteredData, (r) => {
                    const v = r.values[col.name];
                    return typeof v === "number" ? v : 0;
                });
                cell.textContent = this.formatNumber(sum, col);
            }
            rowEl.appendChild(cell);
        });

        return rowEl;
    }

    private formatNumber(value: number, column?: ITableColumn): string {
        const target = column ?? { name: "", displayName: "", isMeasure: true, isGroupBy: false };
        return formatDataValue(value, target, this.settings.regionalFormat ?? { locale: "en-US", currency: "", dateFormat: "", numberSeparators: "auto", decimalPrecision: -1 });
    }

    // -----------------------------------------------------------------
    // Smart tooltips (mean / deviation) — insight without extra UI
    // -----------------------------------------------------------------

    private computeColumnStats(): void {
        this._columnStats.clear();
        this._columnMinMax.clear();
        this.columnMaxCache.clear();
        this.columns.filter((c) => c.isMeasure).forEach((col) => {
            const values = this._data
                .map((r) => r.values[col.name])
                .filter((v): v is number => typeof v === "number");
            if (values.length === 0) {
                return;
            }
            const mean = d3.mean(values) ?? 0;
            const deviation = d3.deviation(values) ?? 0;
            this._columnStats.set(col.name, { mean, deviation });
            this._columnMinMax.set(col.name, { min: d3.min(values) ?? 0, max: d3.max(values) ?? 0 });
        });
    }

    /**
     * Native Tooltip Registration (item 19): smart mean/deviation tooltips for numeric
     * measure columns (unchanged), plus a simple value tooltip for Tooltip-role fields
     * (item 19's "Tooltips" data role) that aren't already covered above -- including
     * non-numeric ones, which the original implementation silently dropped.
     */
    private showRowTooltip(row: ITableRow, evt: MouseEvent): void {
        if (!this.tooltipService.enabled()) {
            return;
        }

        const items: VisualTooltipDataItem[] = [];
        const covered = new Set<string>();

        this.columns.filter((c) => c.isMeasure).forEach((col) => {
            const raw = row.values[col.name];
            if (typeof raw !== "number") {
                return;
            }
            covered.add(col.name);
            const stats = this._columnStats.get(col.name);
            let detail = this.formatNumber(raw);
            if (stats) {
                const variancePct = stats.mean !== 0 ? ((raw - stats.mean) / stats.mean) * 100 : 0;
                const sign = variancePct >= 0 ? "+" : "";
                const avgLabel = this.loc("Tooltip_Avg", "avg");
                const vsAvgLabel = this.loc("Tooltip_VsAvg", "vs avg");
                detail += ` (${avgLabel} ${this.formatNumber(stats.mean)}, ${sign}${variancePct.toFixed(1)}% ${vsAvgLabel}, \u03C3 ${this.formatNumber(stats.deviation)})`;
            }
            items.push({ displayName: col.displayName, value: detail });
        });

        // Tooltip-role fields: simple value tooltip, numeric or not, skipping any column
        // already covered by the smart numeric tooltip above to avoid duplicate lines.
        this.tooltipColumns.forEach((col) => {
            if (covered.has(col.name)) {
                return;
            }
            const raw = row.values[col.name];
            if (raw === null || raw === undefined) {
                return;
            }
            const value = typeof raw === "number" ? this.formatNumber(raw) : String(raw);
            items.push({ displayName: col.displayName, value });
        });

        if (items.length === 0) {
            return;
        }

        this.tooltipService.show({
            coordinates: [evt.clientX, evt.clientY],
            isTouchEvent: false,
            dataItems: items,
            identities: [row.selectionId]
        });
    }

    private moveTooltip(evt: MouseEvent): void {
        this.tooltipService.move({
            coordinates: [evt.clientX, evt.clientY],
            isTouchEvent: false,
            dataItems: [],
            identities: []
        });
    }

    private hideTooltip(): void {
        this.tooltipService.hide({ immediately: true, isTouchEvent: false });
    }

    private renderColumnControlsSection(menu: HTMLDivElement): void {
        const section = document.createElement("section");
        section.className = "skiba-toolbar__section datalake-column-controls";
        const title = document.createElement("div");
        title.className = "skiba-toolbar__section-title";
        title.textContent = "Column controls";
        section.appendChild(title);
        const help = document.createElement("div");
        help.className = "datalake-settings-help";
        help.textContent = "Widths use bounded calculations; Fit content samples at most 200 loaded values and never measures thousands of cells in the DOM.";
        section.appendChild(help);
        this.columns.forEach((col) => {
            const row = document.createElement("div");
            row.className = "datalake-column-control";
            const label = document.createElement("span");
            label.className = "datalake-column-control__name";
            label.textContent = col.displayName;
            row.appendChild(label);
            const width = this.createBoundedSelect([
                { value: "auto", label: "Auto" }, { value: "fit", label: "Fit content" }, { value: "fill", label: "Fill available" }, { value: "fixed", label: "Fixed" }, { value: "custom", label: "Custom" }
            ], this._columnWidthModes.get(col.name) ?? "auto", `Width mode for ${col.displayName}`, (value) => {
                this._columnWidthModes.set(col.name, value as "auto" | "fit" | "fill" | "fixed" | "custom");
                this.renderHeader(); this.renderVisibleRows(); this.persistUserConfig();
            });
            row.appendChild(width.root);
            const alignment = this.createBoundedSelect([
                { value: "auto", label: "Auto" }, { value: "left", label: "Left" }, { value: "center", label: "Center" }, { value: "right", label: "Right" }
            ], this._columnAlignments.get(col.name) ?? "auto", `Alignment for ${col.displayName}`, (value) => {
                this._columnAlignments.set(col.name, value as "auto" | "left" | "center" | "right");
                this.renderVisibleRows(); this.persistUserConfig();
            });
            row.appendChild(alignment.root);
            const min = document.createElement("input"); min.type = "number"; min.min = "48"; min.max = "800"; min.step = "4"; min.value = String(this._columnMinWidths.get(col.name) ?? 80); min.title = "Minimum width";
            const max = document.createElement("input"); max.type = "number"; max.min = "48"; max.max = "1200"; max.step = "4"; max.value = String(this._columnMaxWidths.get(col.name) ?? 520); max.title = "Maximum width";
            const commitBounds = (): void => {
                const minValue = Math.max(48, Number(min.value) || 80); const maxValue = Math.max(minValue, Number(max.value) || 520);
                this._columnMinWidths.set(col.name, minValue); this._columnMaxWidths.set(col.name, maxValue); min.value = String(minValue); max.value = String(maxValue);
                this.renderHeader(); this.renderVisibleRows(); this.persistUserConfig();
            };
            min.addEventListener("change", commitBounds); max.addEventListener("change", commitBounds);
            row.append(min, max);
            const sortable = document.createElement("input"); sortable.type = "checkbox"; sortable.checked = !this._nonSortableColumns.has(col.name); sortable.title = "Sortable";
            sortable.addEventListener("change", () => { if (sortable.checked) this._nonSortableColumns.delete(col.name); else this._nonSortableColumns.add(col.name); this.renderHeader(); this.persistUserConfig(); });
            const filterable = document.createElement("input"); filterable.type = "checkbox"; filterable.checked = !this._nonFilterableColumns.has(col.name); filterable.title = "Filterable";
            filterable.addEventListener("change", () => { if (filterable.checked) this._nonFilterableColumns.delete(col.name); else this._nonFilterableColumns.add(col.name); this.renderHeader(); this.persistUserConfig(); });
            row.append(sortable, filterable);
            section.appendChild(row);
        });
        menu.appendChild(section);
    }

    // -----------------------------------------------------------------
    // Export
    // -----------------------------------------------------------------

    private exportRows(): ITableRow[] {
        const scope = this.settings.exportSettings?.rowScope ?? "filtered";
        if (scope === "available") return [...this._data];
        if (scope === "selected") {
            const selectedIds = this.selectionManager.getSelectionIds() as ISelectionId[];
            return this._data.filter((row) => selectedIds.some((id) => id.equals(row.selectionId)));
        }
        if (scope === "visible") {
            const visibleKeys = new Set(this._renderNodes.filter((n): n is Extract<RenderNode, { kind: "row" }> => n.kind === "row").map((n) => n.row.key));
            return this._filteredData.filter((row) => visibleKeys.has(row.key));
        }
        return [...this._filteredData];
    }

    private exportColumns(): ITableColumn[] {
        return this.visibleColumns();
    }

    private exportAudit(kind: "csv" | "excel" | "json" | "pdf", rowCount: number): void {
        if (this.settings.governance?.auditEnabled === false) return;
        // Deliberately not a user identity. The visual cannot prove a manually entered identity.
        recordExportAudit(createExportAuditEvent(kind, "not-authenticated", rowCount));
    }

    private exportBlockedMessage(status: powerbi.PrivilegeStatus): [string, string] {
        switch (status) {
            case powerbi.PrivilegeStatus.DisabledByAdmin:
                return [
                    "Downloads disabled by your organization",
                    "Your Power BI administrator has disabled file downloads for visuals in this tenant. Contact your admin if you need this enabled."
                ];
            case powerbi.PrivilegeStatus.NotSupported:
                return [
                    "Downloads not available here",
                    "File export isn't available in this Power BI environment (for example, some embedded or non-Desktop/Service contexts). Try opening the report in Power BI Desktop or the Power BI service."
                ];
            case powerbi.PrivilegeStatus.NotDeclared:
                return [
                    "Export not enabled in this visual",
                    "This visual build hasn't requested export permission. This is a visual configuration issue, not a Power BI restriction."
                ];
            default:
                return [
                    "Export unavailable",
                    "Power BI has not allowed file downloads for this visual in the current environment."
                ];
        }
    }

    private async downloadContent(content: string, filename: string, fileType: string, description: string, fallbackBlob?: Blob): Promise<boolean> {
        const service = this.host.downloadService;
        if (service && typeof service.exportVisualsContentExtended === "function") {
            try {
                const status = await service.exportStatus();
                if (status !== powerbi.PrivilegeStatus.Allowed) {
                    if (typeof this.host.displayWarningIcon === "function") {
                        const [title, detail] = this.exportBlockedMessage(status);
                        this.host.displayWarningIcon(title, detail);
                    }
                    return false;
                }
                const result = await service.exportVisualsContentExtended(content, filename, fileType, description);
                return result.downloadCompleted;
            } catch (error) {
                console.warn("Data Lake Tables: Power BI download API failed.", error);
                return false;
            }
        }
        // Test/legacy-host fallback only. Production Power BI 5.3 uses the privileged download API above.
        if (fallbackBlob) {
            this.downloadBlob(fallbackBlob, filename);
            return true;
        }
        return false;
    }

    private rowsForExportTable(): { columns: ITableColumn[]; rows: string[][]; totals: string[] | null } {
        const columns = this.exportColumns();
        const rows = this.exportRows().map((row) => columns.map((col) => formatDataValue(row.values[col.name], col, this.settings.regionalFormat ?? {
            locale: "en-US", currency: "", dateFormat: "", numberSeparators: "auto", decimalPrecision: -1
        })));
        let totals: string[] | null = null;
        if (this.settings.exportSettings?.includeTotals && columns.length > 0) {
            totals = columns.map((col, index) => {
                if (index === 0) return this.settings.totalsLabel;
                if (!col.isMeasure) return "";
                const sum = d3.sum(this.exportRows(), (row) => {
                    const value = row.values[col.name];
                    return typeof value === "number" && Number.isFinite(value) ? value : 0;
                });
                return formatDataValue(sum, col, this.settings.regionalFormat ?? {
                    locale: "en-US", currency: "", dateFormat: "", numberSeparators: "auto", decimalPrecision: -1
                });
            });
        }
        return { columns, rows, totals };
    }

    private async exportCSV(): Promise<void> {
        if (!this.settings.exportSettings?.enabled || !this.settings.exportSettings.csv || this.isExportRestricted()) return;
        const table = this.rowsForExportTable();
        const header = this.settings.exportSettings.includeHeaders ? [table.columns.map((c) => c.displayName)] : [];
        const rows = table.totals ? [...table.rows, table.totals] : table.rows;
        const csv = d3.csvFormatRows([...header, ...rows]);
        const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
        const ok = await this.downloadContent(csv, "data-lake-tables-export.csv", "text/csv", "Data Lake Tables CSV export", blob);
        if (ok) this.exportAudit("csv", table.rows.length);
    }

    private async exportExcel(): Promise<void> {
        if (!this.settings.exportSettings?.enabled || !this.settings.exportSettings.excel || this.isExportRestricted()) return;
        const table = this.rowsForExportTable();
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Data");
        if (this.settings.exportSettings.includeHeaders) worksheet.columns = table.columns.map((c) => ({ header: c.displayName, key: c.name }));
        table.rows.forEach((row) => worksheet.addRow(row));
        if (table.totals) worksheet.addRow(table.totals);
        const wbout = await workbook.xlsx.writeBuffer();
        const bytes = wbout instanceof ArrayBuffer ? new Uint8Array(wbout) : new Uint8Array(wbout as ArrayBuffer);
        const binary = Array.from(bytes, (byte) => String.fromCharCode(byte)).join("");
        const base64 = btoa(binary);
        const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const ok = await this.downloadContent(base64, "data-lake-tables-export.xlsx", "base64", "Data Lake Tables Excel export", blob);
        if (ok) this.exportAudit("excel", table.rows.length);
    }

    private async exportJSON(): Promise<void> {
        if (!this.settings.exportSettings?.enabled || !this.settings.exportSettings.json || this.isExportRestricted()) return;
        const table = this.rowsForExportTable();
        const rows = this.exportRows().map((row) => Object.fromEntries(table.columns.map((col) => [col.displayName, row.values[col.name] ?? null])));
        const payload = JSON.stringify({ columns: table.columns.map((c) => c.displayName), rows, totals: table.totals }, null, 2);
        const blob = new Blob([payload], { type: "application/json;charset=utf-8;" });
        const ok = await this.downloadContent(payload, "data-lake-tables-export.json", "application/json", "Data Lake Tables JSON export", blob);
        if (ok) this.exportAudit("json", rows.length);
    }

    /** Multi-page PDF export using the same filtered/selected/visible row scope as CSV/Excel/JSON. */
    private async exportPDF(): Promise<void> {
        if (!this.settings.exportSettings?.enabled || !this.settings.exportSettings.pdf || this.isExportRestricted()) return;
        const table = this.rowsForExportTable();
        const head = this.settings.exportSettings.includeHeaders ? [table.columns.map((c) => c.displayName)] : [];
        const bodyRows = [...table.rows];
        if (table.totals) bodyRows.push(table.totals);
        const doc = new jsPDF({ orientation: table.columns.length > 6 ? "landscape" : "portrait", unit: "pt" });
        const exportDate = new Date().toLocaleString(this.settings.regionalFormat?.locale || "en-US");
        autoTable(doc, {
            head,
            body: bodyRows,
            startY: 56,
            margin: { top: 56 },
            styles: { font: "helvetica", fontSize: 8, textColor: this.settings.cellFont, fillColor: this.settings.cellBg, lineColor: "#e6e6e6", lineWidth: 0.5, cellPadding: 4 },
            headStyles: { fillColor: this.settings.headerBg, textColor: this.settings.headerFont, fontStyle: this.settings.headerBold ? "bold" : "normal" },
            alternateRowStyles: { fillColor: this.settings.altRow },
            didDrawPage: (data) => {
                const left = data.settings.margin.left;
                doc.setFontSize(12); doc.setTextColor("#333333"); doc.text(this.reportTitle, left, 24);
                doc.setFontSize(8); doc.setTextColor("#888888"); doc.text(`Exported ${exportDate}`, left, 38);
            },
            showHead: "everyPage"
        });
        const governance = this.settings.governance;
        if (governance?.watermarkEnabled) {
            const text = buildWatermarkText({ enabled: true, watermarkText: governance.watermarkText, locale: this.settings.regionalFormat?.locale ?? "en-US", currency: this.settings.regionalFormat?.currency ?? "", username: "not-authenticated" });
            if (text) {
                doc.setTextColor("#9CA3AF"); doc.setFontSize(9);
                if (governance.watermarkPlacement === "center") {
                    doc.text(text, doc.internal.pageSize.getWidth() / 2, doc.internal.pageSize.getHeight() / 2, { align: "center", angle: 0 });
                } else {
                    doc.text(text, doc.internal.pageSize.getWidth() / 2, doc.internal.pageSize.getHeight() - 18, { align: "center" });
                }
            }
        }
        const arrayBuffer = doc.output("arraybuffer") as ArrayBuffer;
        const bytes = new Uint8Array(arrayBuffer);
        const binary = Array.from(bytes, (byte) => String.fromCharCode(byte)).join("");
        const base64 = btoa(binary);
        const blob = new Blob([bytes], { type: "application/pdf" });
        const ok = await this.downloadContent(base64, "data-lake-tables-export.pdf", "base64", "Data Lake Tables PDF export", blob);
        if (ok) this.exportAudit("pdf", table.rows.length);
    }

    private renderExportControls(section: HTMLDivElement, closeDrawer: () => void): void {
        const settings = this.settings.exportSettings;
        if (!settings || !settings.enabled || this.isExportRestricted()) return;
        const title = document.createElement("div");
        title.className = "datalake-settings-subtitle";
        title.textContent = "Download";
        section.appendChild(title);
        const help = document.createElement("div");
        help.className = "datalake-settings-help";
        help.textContent = "Downloads use Power BI's governed file-download API when available.";
        section.appendChild(help);
        const make = (label: string, enabled: boolean, action: () => void): void => {
            if (!enabled) return;
            section.appendChild(this.makeMenuButton(label, () => { action(); closeDrawer(); }));
        };
        make("Download CSV", settings.csv, () => void this.exportCSV());
        make("Download Excel", settings.excel, () => void this.exportExcel());
        make("Download JSON", settings.json, () => void this.exportJSON());
        make("Download PDF", settings.pdf, () => void this.exportPDF());
    }

    private downloadBlob(blob: Blob, filename: string): void {
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.setTimeout(() => URL.revokeObjectURL(url), 10000);
    }

    // -----------------------------------------------------------------
    // Empty state & landing page
    // -----------------------------------------------------------------

    /** Renders a calm, informational empty state for fields-assigned/no-row results. */
    public renderEmptyState(): void {
        this.clearElement(this.container);
        const wrap = document.createElement("div");
        wrap.className = "skiba-empty-state skiba-empty-state--data";
        const title = document.createElement("strong");
        title.textContent = this.loc("EmptyState_Title", "No data to display");
        const helper = document.createElement("span");
        helper.textContent = this.loc("EmptyState_Helper", "Add fields to Rows and Values, or adjust the current filters.");
        wrap.append(title, helper);
        this.container.appendChild(wrap);
    }

    /**
     * Landing / Welcome Page (item 15): shown only before any fields have ever been
     * assigned to the visual (distinct from renderEmptyState above). Reuses the existing
     * `.skiba-empty-state` branding classes rather than a second, inconsistent design, with
     * an added plain-language description of what Data Lake Tables does.
     */
    public renderLandingPage(): void {
        this.clearElement(this.container);

        const wrap = document.createElement("div");
        wrap.className = "skiba-empty-state skiba-landing-page";
        wrap.setAttribute("data-dlt-surface", "landing");
        wrap.setAttribute("data-dlt-tier", "2");

        const top = document.createElement("header");
        top.className = "skiba-landing-page__top";
        const brandGroup = document.createElement("div");
        brandGroup.className = "skiba-landing-page__brand-group";
        const mark = document.createElement("div");
        mark.className = "skiba-landing-page__mark";
        mark.textContent = "DLT";
        mark.setAttribute("aria-hidden", "true");
        const brandStack = document.createElement("div");
        brandStack.className = "skiba-landing-page__brand-stack";
        const brand = document.createElement("span");
        brand.className = "skiba-landing-page__brand";
        brand.textContent = this.loc("Landing_Brand", "DATA LAKE TABLES");
        const attribution = document.createElement("span");
        attribution.className = "skiba-landing-page__attribution";
        attribution.textContent = this.loc("Landing_Attribution", "SIMON KP · BRYT MA TECH UG");
        brandStack.append(brand, attribution);
        brandGroup.append(mark, brandStack);
        top.append(brandGroup);

        const main = document.createElement("main");
        main.className = "skiba-landing-page__main";

        const hero = document.createElement("section");
        hero.className = "skiba-landing-page__editorial-hero";
        const kicker = document.createElement("span");
        kicker.className = "skiba-landing-page__editorial-kicker";
        kicker.textContent = this.loc("Landing_Editorial_Kicker", "Before you add data");
        const title = document.createElement("h1");
        title.className = "skiba-landing-page__editorial-title";
        const landingTitleText = this.loc("Landing_Editorial_Title", "Make the table do more.");
        const landingTitleWords = landingTitleText.split(" ");
        const landingTitleLast = landingTitleWords.pop() || "";
        if (landingTitleWords.length > 0) { title.appendChild(document.createTextNode(landingTitleWords.join(" ") + " ")); }
        const landingTitleAccent = document.createElement("span");
        landingTitleAccent.className = "skiba-landing-page__editorial-accent";
        landingTitleAccent.textContent = landingTitleLast;
        title.appendChild(landingTitleAccent);
        const description = document.createElement("p");
        description.className = "skiba-landing-page__editorial-description";
        description.textContent = this.loc("Landing_Editorial_Description", "Explore the data already supplied to your visual, then move naturally into deeper analysis and presentation.");

        const actions = document.createElement("div");
        actions.className = "skiba-landing-page__editorial-actions";
        const makeAction = (label: string, variant: "primary" | "secondary", onClick: () => void): HTMLButtonElement => {
            const b = document.createElement("button");
            b.type = "button";
            b.className = `skiba-landing-page__editorial-action skiba-landing-page__editorial-action--${variant}`;
            b.textContent = label;
            b.addEventListener("click", (evt) => { evt.preventDefault(); evt.stopPropagation(); onClick(); });
            return b;
        };
        actions.append(
            makeAction(this.loc("Landing_GetStarted_Cta", "Add your fields"), "primary", () => this.renderLandingInfoModal(
                "Set up your first table", "QUICK SETUP",
                "Use the Power BI field wells to give Data Lake Tables the columns and measures you want to work with.",
                [
                    "Rows — add the fields you want to appear as row-level table columns.",
                    "Values — add measures or value fields you want to analyze in the table.",
                    "After fields are supplied, the visual switches from the landing experience to the data table.",
                    "Use the in-visual controls after import to search, group, pivot, calculate, format and export where enabled."
                ]
            )),
            makeAction(this.loc("Landing_NewHere_Cta", "Learn more"), "secondary", () => this.renderLandingInfoModal(
                "How Data Lake Tables works", "GETTING STARTED",
                "A practical workflow from an empty visual to a focused analytical table.",
                [
                    "Start by supplying the fields you want the visual to display.",
                    "Use the table surface to search, filter, sort and select the rows currently supplied to the visual.",
                    "Use grouping, pivoting and calculations when you need a deeper analytical view.",
                    "Use layout, typography, column and formatting controls to prepare a readable presentation.",
                    "Use export actions only when the visual configuration and Power BI host permit them."
                ]
            ))
        );
        hero.append(kicker, title, description, actions, this.makeLandingSvg("skiba-landing-page__wave", "0 0 400 90", [
            ["M0 60C80 40 160 80 400 30V90H0Z", "#104F9F"],
            ["M0 74C110 56 230 96 400 52V64C250 104 120 68 0 88Z", "#FAF623"],
            ["M0 90V78C120 66 260 100 400 70V90Z", "#308ABE"],
            ["M0 84C120 72 260 106 400 78V82C260 110 120 76 0 88Z", "#EC0A8C"]
        ], "none"));

        const info = document.createElement("section");
        info.className = "skiba-landing-page__editorial-info";
        info.append(this.makeLandingPreview());
        main.append(hero, info);

        const footer = document.createElement("footer");
        footer.className = "skiba-landing-page__footer";
        footer.setAttribute("aria-label", this.loc("Landing_FooterAria", "Data Lake Tables information"));
        const footerCopy = document.createElement("span");
        footerCopy.className = "skiba-landing-page__editorial-footer-copy";
        footerCopy.textContent = this.loc("Landing_Footer", "Special thanks to Mr Simon KP for the guidance behind the project.");
        const links = document.createElement("nav");
        links.className = "skiba-landing-page__links";
        links.setAttribute("aria-label", this.loc("Landing_InfoNavigation", "Information"));
        const makeInfoButton=(label:string,titleText:string,eyebrowText:string,leadText:string,bullets:string[]):HTMLButtonElement=>{const b=document.createElement("button");b.type="button";b.className="skiba-landing-page__link";b.textContent=label;b.addEventListener("click",(evt)=>{evt.preventDefault();evt.stopPropagation();this.renderLandingInfoModal(titleText,eyebrowText,leadText,bullets);});return b;};
        links.append(
            makeInfoButton("Support","Support","SUPPORT","For product assistance, use the published contact options below.",["WhatsApp: +256 759 621 612","Email: muhumuzabright26@gmail.com","Phone: 0759 621 612"]),
            makeInfoButton("FAQ","Frequently asked questions","FAQ","Practical answers for setup, everyday table work and the supported workflow.",[]),
            makeInfoButton("About","About Data Lake Tables","ABOUT THE PRODUCT","Data Lake Tables is a Power BI custom visual focused on analytical table presentation, exploration and report-ready formatting.",["Built by Simon KP and Bryt Ma Tech UG.","Core stack: Power BI Visuals API 5.3.0 and TypeScript.","Interactive/data presentation: D3.js.","Exports: ExcelJS, jsPDF and jsPDF-AutoTable.","Styling: LESS / CSS with a dedicated visual presentation layer.","Verification: Jest, ESLint, TypeScript and Power BI custom-visual tooling.","Special thanks to Mr Simon KP for the product guidance, direction and feedback that helped shape this visual."])
        );
        footer.append(footerCopy, links);
        wrap.append(
            this.makeLandingSvg("skiba-landing-page__swoosh skiba-landing-page__swoosh--top", "0 0 240 90", [
                ["M60 0H240V70C190 40 120 30 60 0Z", "#FAF623"],
                ["M110 0H240V46C200 26 150 16 110 0Z", "#104F9F"],
                ["M85 0H100C130 22 190 34 240 52V58C185 40 120 28 85 0Z", "#EC0A8C"]
            ]),
            this.makeLandingSvg("skiba-landing-page__swoosh skiba-landing-page__swoosh--bottom", "0 0 300 70", [
                ["M0 70C90 60 180 30 300 0V70Z", "#104F9F"],
                ["M40 70C130 62 220 36 300 14V20C220 44 130 68 60 70Z", "#EC0A8C"],
                ["M60 70C140 62 220 42 300 22V34C230 56 150 70 100 70Z", "#FAF623"]
            ]),
            top, main, footer
        );
        this.container.appendChild(wrap);
    }

    private makeLandingSvg(className: string, viewBox: string, paths: Array<[string, string]>, preserve?: string): SVGSVGElement {
        const ns = "http://www.w3.org/2000/svg";
        const svg = document.createElementNS(ns, "svg");
        svg.setAttribute("class", className);
        svg.setAttribute("viewBox", viewBox);
        svg.setAttribute("aria-hidden", "true");
        svg.setAttribute("focusable", "false");
        if (preserve) { svg.setAttribute("preserveAspectRatio", preserve); }
        paths.forEach(([d, fill]) => {
            const p = document.createElementNS(ns, "path");
            p.setAttribute("d", d);
            p.setAttribute("fill", fill);
            svg.appendChild(p);
        });
        return svg;
    }

    private makeLandingPreview(): HTMLElement {
        const card = document.createElement("article");
        card.className = "skiba-landing-page__preview";
        card.setAttribute("aria-hidden", "true");
        const head = document.createElement("div");
        head.className = "skiba-landing-page__preview-head";
        const heading = document.createElement("strong");
        heading.textContent = "Sample preview";
        const tools = document.createElement("div");
        tools.className = "skiba-landing-page__preview-tools";
        ["Filter", "Group", "Pivot"].forEach((label) => {
            const t = document.createElement("span");
            t.textContent = label;
            tools.appendChild(t);
        });
        head.append(heading, tools);

        const search = document.createElement("div");
        search.className = "skiba-landing-page__preview-search";
        const searchGlass = document.createElement("span");
        searchGlass.className = "skiba-landing-page__preview-search-icon";
        searchGlass.setAttribute("aria-hidden", "true");
        const searchText = document.createElement("span");
        searchText.className = "skiba-landing-page__preview-search-text";
        searchText.textContent = "Search tables, fields, or keywords...";
        search.append(searchGlass, searchText);

        const table = document.createElement("table");
        table.className = "skiba-landing-page__preview-table";
        const thead = document.createElement("thead");
        const headRow = document.createElement("tr");
        ["Table", "Source", "Records", "Status"].forEach((label) => {
            const th = document.createElement("th");
            th.textContent = label;
            headRow.appendChild(th);
        });
        thead.appendChild(headRow);
        const tbody = document.createElement("tbody");
        const rows: Array<[string, string, string, string, string]> = [
            ["Taxpayer_Registry", "OLTP", "1,248,532", "Ready", "ready"],
            ["Payments", "Payments DB", "856,210", "Ready", "ready"],
            ["Assessments", "Warehouse", "512,443", "Ready", "ready"],
            ["Compliance_Log", "Warehouse", "1,842,205", "Ready", "ready"],
            ["Exemptions", "OLTP", "76,540", "Ready", "ready"],
            ["Customs", "Customs DB", "298,771", "Processing", "processing"],
            ["Refunds", "Payments DB", "143,908", "Processing", "processing"],
            ["Audit_Reports", "Warehouse", "92,317", "Loading", "loading"]
        ];
        rows.forEach(([name, source, records, status, kind]) => {
            const tr = document.createElement("tr");
            tr.className = "skiba-landing-page__preview-row skiba-landing-page__preview-row--" + kind;
            [name, source, records].forEach((value) => {
                const td = document.createElement("td");
                td.textContent = value;
                tr.appendChild(td);
            });
            const statusCell = document.createElement("td");
            const chip = document.createElement("span");
            chip.className = "skiba-landing-page__preview-chip skiba-landing-page__preview-chip--" + kind;
            chip.textContent = status;
            statusCell.appendChild(chip);
            tr.appendChild(statusCell);
            tbody.appendChild(tr);
        });
        table.append(thead, tbody);

        const stats = document.createElement("div");
        stats.className = "skiba-landing-page__preview-stats";
        [["8", "tables"], ["4.1M", "rows"], ["Live", "sync"]].forEach(([value, label]) => {
            const stat = document.createElement("span");
            stat.className = "skiba-landing-page__preview-stat";
            const v = document.createElement("b");
            v.textContent = value;
            stat.append(v, document.createTextNode(" " + label));
            stats.appendChild(stat);
        });
        const distBar = document.createElement("div");
        distBar.className = "skiba-landing-page__preview-distbar";
        distBar.setAttribute("role", "img");
        distBar.setAttribute("aria-label", "6 ready, 2 processing, 1 loading");
        ([[6, "ready"], [2, "processing"], [1, "loading"]] as Array<[number, string]>).forEach(([count, kind]) => {
            const seg = document.createElement("span");
            seg.className = "skiba-landing-page__preview-distbar-seg skiba-landing-page__preview-distbar-seg--" + kind;
            seg.style.flexGrow = String(count);
            distBar.appendChild(seg);
        });

        const note = document.createElement("span");
        note.className = "skiba-landing-page__preview-note";
        note.textContent = "Sample data. Your fields replace this once added.";
        card.append(head, search, table, stats, distBar, note);
        return card;
    }

    private renderLandingInfoModal(
        titleText: string,
        eyebrowText: string,
        leadText: string,
        bullets: string[]
    ): void {
        this.container.querySelectorAll(".datalake-landing-info-backdrop").forEach((el) => el.remove());

        const backdrop = document.createElement("div");
        backdrop.className = "datalake-landing-info-backdrop";
        backdrop.setAttribute("role", "presentation");

        const close = (): void => {
            backdrop.remove();
            document.removeEventListener("keydown", onKeyDown);
        };

        const onKeyDown = (evt: KeyboardEvent): void => {
            if (evt.key === "Escape") {
                evt.preventDefault();
                close();
            }
        };

        document.addEventListener("keydown", onKeyDown);
        backdrop.addEventListener("click", (evt) => {
            if (evt.target === backdrop) close();
        });

        const dialog = document.createElement("section");
        dialog.className = "datalake-landing-info";
        dialog.setAttribute("role", "dialog");
        dialog.setAttribute("aria-modal", "true");
        dialog.setAttribute("aria-label", titleText);

        const closeButton = document.createElement("button");
        closeButton.type = "button";
        closeButton.className = "datalake-landing-info__close";
        closeButton.textContent = "×";
        closeButton.setAttribute("aria-label", this.loc("Landing_InfoClose", "Close"));
        closeButton.addEventListener("click", (evt) => {
            evt.stopPropagation();
            close();
        });

        const header = document.createElement("div");
        header.className = "datalake-landing-info__header";
        const eyebrow = document.createElement("span");
        eyebrow.className = "datalake-landing-info__eyebrow";
        eyebrow.textContent = eyebrowText;
        const title = document.createElement("h2");
        title.textContent = titleText;
        const lead = document.createElement("p");
        lead.className = "datalake-landing-info__lead";
        lead.textContent = leadText;
        header.append(eyebrow, title, lead);

        const content = document.createElement("div");
        content.className = "datalake-landing-info__content";

        if (titleText === "Frequently asked questions") {
            const faq = [
                ["How do I start?", "Add the fields you want into the Rows and Values wells in Power BI. Once usable fields are supplied, the landing page gives way to the working table."],
                ["What does Rows do?", "Rows defines the row-level fields that become the main columns and grouping dimensions of the table."],
                ["What does Values do?", "Values supplies measures or value fields used in the analytical table and its calculations."],
                ["Can I search and filter?", "Yes. The visual provides search and column-filter controls for the rows currently available to the visual."],
                ["Can I group and pivot?", "Yes. Grouping and pivoting are available when the required data is supplied and the relevant visual controls are enabled."],
                ["Can I customize the table?", "Yes. Layout, typography, column behavior, conditional formatting and totals are part of the visual's configuration surface."],
                ["Can I export?", "Available export actions depend on the visual configuration and the Power BI host's policies. The visual does not assume export permission is always present."]
            ];
            const faqGrid = document.createElement("div");
            faqGrid.className = "datalake-landing-info__faq";
            faq.forEach(([question, answer]) => {
                const item = document.createElement("details");
                item.className = "datalake-landing-info__faq-item";
                const summary = document.createElement("summary");
                summary.textContent = question;
                const answerEl = document.createElement("p");
                answerEl.textContent = answer;
                item.append(summary, answerEl);
                faqGrid.appendChild(item);
            });
            content.appendChild(faqGrid);
        } else if (titleText === "About Data Lake Tables") {
            const brandBlock = document.createElement("div");
            brandBlock.className = "datalake-landing-info__brand-block";
            const brandTitle = document.createElement("strong");
            brandTitle.textContent = "Simon KP";
            const brandStudio = document.createElement("strong");
            brandStudio.textContent = "Bryt Ma Tech UG";
            const brandCopy = document.createElement("p");
            brandCopy.textContent = "The people and studio behind the Data Lake Tables product experience.";
            brandBlock.append(brandTitle, brandStudio, brandCopy);

            const stackTitle = document.createElement("h3");
            stackTitle.textContent = "Technology stack";
            const stack = document.createElement("div");
            stack.className = "datalake-landing-info__stack";
            [
                "Power BI Visuals API 5.3.0",
                "TypeScript",
                "D3.js",
                "ExcelJS",
                "jsPDF + AutoTable",
                "LESS / CSS",
                "Jest + ESLint",
                "Power BI custom-visual tooling"
            ].forEach((itemText) => {
                const chip = document.createElement("span");
                chip.className = "datalake-landing-info__stack-chip";
                chip.textContent = itemText;
                stack.appendChild(chip);
            });

            const focusTitle = document.createElement("h3");
            focusTitle.textContent = "Product focus";
            const focus = document.createElement("div");
            focus.className = "datalake-landing-info__focus";
            [
                "Readable analytical tables",
                "Search, filters, sorting and grouping",
                "Pivoting and calculations",
                "Formatting, totals and saved views",
                "Governed export actions where enabled"
            ].forEach((itemText) => {
                const row = document.createElement("div");
                row.className = "datalake-landing-info__focus-row";
                row.textContent = itemText;
                focus.appendChild(row);
            });

            const thanks = document.createElement("div");
            thanks.className = "datalake-landing-info__thanks";
            const thanksTitle = document.createElement("strong");
            thanksTitle.textContent = "Special thanks";
            const thanksCopy = document.createElement("p");
            thanksCopy.textContent = "Special thanks to Mr Simon KP for the product guidance, direction and feedback that helped shape this visual.";
            thanks.append(thanksTitle, thanksCopy);

            content.append(brandBlock, stackTitle, stack, focusTitle, focus, thanks);
        } else {
            const list = document.createElement("div");
            list.className = "datalake-landing-info__list";
            bullets.forEach((itemText) => {
                const row = document.createElement("div");
                row.className = "datalake-landing-info__item";
                const dot = document.createElement("span");
                dot.className = "datalake-landing-info__dot";
                dot.setAttribute("aria-hidden", "true");
                const copy = document.createElement("span");
                copy.textContent = itemText;
                row.append(dot, copy);
                list.appendChild(row);
            });
            content.appendChild(list);

            if (titleText === "How Data Lake Tables works" || titleText === "Set up your first table") {
                const note = document.createElement("div");
                note.className = "datalake-landing-info__note";
                note.textContent = titleText === "How Data Lake Tables works"
                    ? "The workflow stays inside your Power BI report: supply fields, shape the table, then use the in-visual controls that are available for the supplied data."
                    : "Tip: start small. Add one row field and one value field first, confirm the table is behaving as expected, then add more fields and analysis options. ";
                content.appendChild(note);
            }
        }

        // Support actions are real host-launchable links, not decorative text.
        if (titleText === "Support") {
            const actions = document.createElement("div");
            actions.className = "datalake-landing-info__actions";
            const supportLinks: Array<[string, string]> = [
                ["WhatsApp", "https://wa.me/256759621612"],
                ["Email", "mailto:muhumuzabright26@gmail.com"]
            ];
            supportLinks.forEach(([label, url]) => {
                const button = document.createElement("button");
                button.type = "button";
                button.className = "datalake-landing-info__action";
                button.textContent = label;
                button.addEventListener("click", (evt) => {
                    evt.stopPropagation();
                    if (typeof this.host.launchUrl === "function") this.host.launchUrl(url);
                });
                actions.appendChild(button);
            });
            content.appendChild(actions);
        }

        dialog.append(closeButton, header, content);
        backdrop.appendChild(dialog);
        this.container.appendChild(backdrop);
        closeButton.focus();
    }

    /**
     * Tier 3 -- contact poster, opened from the settings-drawer footer attribution
     * button (see renderToolbar). Deliberately mirrors the dark announcement-card
     * look so the drawer, the advert card and this poster read as one product.
     *
     * Contact rows use host.launchUrl rather than plain anchors: a Power BI custom
     * visual runs in a sandboxed iframe where ordinary navigation is blocked, and
     * launchUrl is the supported way to open an external link (same approach as
     * appendLinkActionIcon).
     */
    private renderPosterModal(): void {
        this.container.querySelectorAll(".datalake-poster, .datalake-poster-backdrop").forEach((el) => el.remove());

        const backdrop = document.createElement("div");
        backdrop.className = "datalake-poster-backdrop";
        backdrop.style.setProperty("position", "absolute", "important");
        backdrop.style.setProperty("inset", "0", "important");
        backdrop.style.setProperty("z-index", "400", "important");
        backdrop.style.background = "rgba(10,22,38,0.55)";
        backdrop.style.display = "flex";
        backdrop.style.alignItems = "center";
        backdrop.style.justifyContent = "center";

        const closePoster = (): void => {
            backdrop.remove();
            document.removeEventListener("keydown", onKeyDown);
        };
        const onKeyDown = (evt: KeyboardEvent): void => {
            if (evt.key === "Escape") { evt.stopPropagation(); closePoster(); }
        };
        document.addEventListener("keydown", onKeyDown);
        backdrop.addEventListener("click", (evt) => { if (evt.target === backdrop) closePoster(); });

        const poster = document.createElement("div");
        poster.className = "datalake-poster";
        poster.setAttribute("role", "dialog");
        poster.setAttribute("aria-label", this.loc("Poster_Label", "Contact the makers of Data Lake Tables"));
        poster.style.position = "relative";
        poster.style.width = "320px";
        poster.style.maxWidth = "86%";
        poster.style.maxHeight = "86%";
        poster.style.overflowY = "auto";
        poster.style.background = "linear-gradient(160deg,#0B2033,#123049)";
        poster.style.color = "#ffffff";
        poster.style.borderRadius = "16px";
        poster.style.padding = "22px";
        poster.style.boxShadow = "0 30px 70px rgba(4,16,28,0.5)";
        poster.addEventListener("click", (evt) => evt.stopPropagation());

        const closeBtn = document.createElement("button");
        closeBtn.type = "button";
        closeBtn.className = "datalake-announcement__close";
        closeBtn.textContent = "\u00D7";
        closeBtn.setAttribute("aria-label", this.loc("Poster_Close", "Close"));
        closeBtn.addEventListener("click", (evt) => { evt.stopPropagation(); closePoster(); });
        poster.appendChild(closeBtn);

        const eyebrow = document.createElement("span");
        eyebrow.className = "datalake-announcement-details__eyebrow";
        eyebrow.textContent = this.loc("Poster_Eyebrow", "BUILT BY");
        poster.appendChild(eyebrow);

        const title = document.createElement("div");
        title.textContent = "Simon KP";
        title.style.fontSize = "19px";
        title.style.fontWeight = "900";
        title.style.marginBottom = "2px";
        poster.appendChild(title);

        const studio = document.createElement("div");
        studio.textContent = "Bryt Ma Tech UG";
        studio.style.fontSize = "13px";
        studio.style.fontWeight = "700";
        studio.style.color = "#7fd7ff";
        studio.style.marginBottom = "10px";
        poster.appendChild(studio);

        const blurb = document.createElement("div");
        blurb.textContent = this.loc("Poster_Blurb", "Custom Power BI visuals, data pipelines and analytics tooling.");
        blurb.style.fontSize = "11.5px";
        blurb.style.lineHeight = "1.45";
        blurb.style.color = "#c6d6e3";
        blurb.style.marginBottom = "16px";
        poster.appendChild(blurb);

        const contacts: Array<{ label: string; value: string; url: string | null }> = [
            { label: this.loc("Poster_WhatsApp", "WhatsApp"), value: "+256 759 621 612", url: "https://wa.me/256759621612" },
            { label: this.loc("Poster_Email", "Email"), value: "muhumuzabright26@gmail.com", url: "mailto:muhumuzabright26@gmail.com" },
            { label: this.loc("Poster_Phone", "Phone"), value: "0759 621 612", url: null }
        ];

        contacts.forEach((contact) => {
            const row = document.createElement(contact.url ? "button" : "div");
            if (contact.url) {
                (row as HTMLButtonElement).type = "button";
                row.className = "datalake-poster__contact";
                row.addEventListener("click", (evt) => {
                    evt.stopPropagation();
                    if (typeof this.host.launchUrl === "function" && contact.url) {
                        this.host.launchUrl(contact.url);
                    }
                });
                row.style.cursor = "pointer";
                row.style.textAlign = "start";
                row.style.font = "inherit";
            }
            row.style.display = "block";
            row.style.width = "100%";
            row.style.marginBottom = "8px";
            row.style.padding = "9px 11px";
            row.style.border = "1px solid rgba(255,255,255,0.10)";
            row.style.borderRadius = "10px";
            row.style.background = "rgba(255,255,255,0.045)";
            row.style.color = "#ffffff";

            const k = document.createElement("div");
            k.textContent = contact.label;
            k.style.fontSize = "9px";
            k.style.fontWeight = "900";
            k.style.letterSpacing = "0.09em";
            k.style.color = "#8fa6bb";
            k.style.marginBottom = "2px";

            const v = document.createElement("div");
            v.textContent = contact.value;
            v.style.fontSize = "13px";
            v.style.fontWeight = "700";

            row.appendChild(k);
            row.appendChild(v);
            poster.appendChild(row);
        });

        backdrop.appendChild(poster);
        this.container.appendChild(backdrop);
        closeBtn.focus();
    }
}

