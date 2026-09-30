"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import { densityItems, stylePresetItems } from "./styleSystem";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

/**
 * General card: font family, font size, base row height.
 */
class GeneralSettingsCard extends FormattingSettingsCard {
    fontFamily = new formattingSettings.FontPicker({
        name: "fontFamily",
        displayName: "Table font",
        displayNameKey: "Prop_FontFamily",
        value: "Segoe UI"
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Table font size",
        displayNameKey: "Prop_FontSize",
        value: 14
    });

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "Table font color",
        value: { value: "#333333" }
    });

    rowHeight = new formattingSettings.NumUpDown({
        name: "rowHeight",
        displayName: "Row height",
        displayNameKey: "Prop_RowHeight",
        value: 36
    });

    name: string = "general";
    displayName: string = "General";
    displayNameKey: string = "Object_General";
    slices: FormattingSettingsSlice[] = [this.fontFamily, this.fontSize, this.fontColor, this.rowHeight];
}

/**
 * Header card: background, text color, bold.
 */
class HeaderSettingsCard extends FormattingSettingsCard {
    fontFamily = new formattingSettings.FontPicker({ name: "fontFamily", displayName: "Font family", value: "Segoe UI" });
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Font size", value: 13 });
    fontWeight = new formattingSettings.ItemDropdown({
        name: "fontWeight", displayName: "Weight",
        items: [{ value: "400", displayName: "Regular" }, { value: "500", displayName: "Medium" }, { value: "600", displayName: "Semibold" }, { value: "700", displayName: "Bold" }],
        value: { value: "600", displayName: "Semibold" }
    });
    alignment = new formattingSettings.ItemDropdown({
        name: "alignment", displayName: "Alignment",
        items: [{ value: "left", displayName: "Left" }, { value: "center", displayName: "Center" }, { value: "right", displayName: "Right" }],
        value: { value: "left", displayName: "Left" }
    });
    wrap = new formattingSettings.ToggleSwitch({ name: "wrap", displayName: "Wrap text", value: false });

    bgColor = new formattingSettings.ColorPicker({
        name: "bgColor",
        displayName: "Background color",
        displayNameKey: "Prop_BgColor",
        value: { value: "#F0F2F5" }
    });

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "Font color",
        displayNameKey: "Prop_FontColor",
        value: { value: "#333333" }
    });

    bold = new formattingSettings.ToggleSwitch({
        name: "bold",
        displayName: "Bold",
        displayNameKey: "Prop_Bold",
        value: true
    });

    name: string = "header";
    displayName: string = "Header";
    displayNameKey: string = "Object_Header";
    slices: FormattingSettingsSlice[] = [this.fontFamily, this.fontSize, this.fontWeight, this.alignment, this.wrap, this.bgColor, this.fontColor, this.bold];
}

/**
 * Cells card: background, text color, alternate (zebra) row color.
 */
class CellsSettingsCard extends FormattingSettingsCard {
    fontSize = new formattingSettings.NumUpDown({ name: "fontSize", displayName: "Font size", value: 13 });
    fontWeight = new formattingSettings.ItemDropdown({
        name: "fontWeight", displayName: "Weight",
        items: [{ value: "400", displayName: "Regular" }, { value: "500", displayName: "Medium" }, { value: "600", displayName: "Semibold" }, { value: "700", displayName: "Bold" }],
        value: { value: "400", displayName: "Regular" }
    });
    alignment = new formattingSettings.ItemDropdown({
        name: "alignment", displayName: "Alignment",
        items: [{ value: "auto", displayName: "Auto" }, { value: "left", displayName: "Left" }, { value: "center", displayName: "Center" }, { value: "right", displayName: "Right" }],
        value: { value: "auto", displayName: "Auto" }
    });
    bgColor = new formattingSettings.ColorPicker({
        name: "bgColor",
        displayName: "Background color",
        displayNameKey: "Prop_BgColor",
        value: { value: "#FFFFFF" }
    });

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "Font color",
        displayNameKey: "Prop_FontColor",
        value: { value: "#333333" }
    });

    alternateRowColor = new formattingSettings.ColorPicker({
        name: "alternateRowColor",
        displayName: "Alternate row color",
        displayNameKey: "Prop_AlternateRowColor",
        value: { value: "#FAFAFA" }
    });

    name: string = "cells";
    displayName: string = "Cells";
    displayNameKey: string = "Object_Cells";
    slices: FormattingSettingsSlice[] = [this.fontSize, this.fontWeight, this.alignment, this.bgColor, this.fontColor, this.alternateRowColor];
}

/**
 * In-cell data bars card.
 */
class DataBarsSettingsCard extends FormattingSettingsCard {
    enableDataBars = new formattingSettings.ToggleSwitch({
        name: "enableDataBars",
        displayName: "Show data bars",
        displayNameKey: "Prop_EnableDataBars",
        value: false
    });

    barColor = new formattingSettings.ColorPicker({
        name: "barColor",
        displayName: "Positive bar color",
        displayNameKey: "Prop_BarColor",
        value: { value: "#0078D4" }
    });

    negativeBarColor = new formattingSettings.ColorPicker({
        name: "negativeBarColor",
        displayName: "Negative bar color",
        value: { value: "#C50F1F" }
    });

    name: string = "formatting";
    displayName: string = "Data bars";
    displayNameKey: string = "Object_DataBars";
    slices: FormattingSettingsSlice[] = [this.enableDataBars, this.barColor, this.negativeBarColor];
}

/**
 * Totals row card.
 */
class TableSurfaceSettingsCard extends FormattingSettingsCard {
    headerBorder = new formattingSettings.ColorPicker({ name: "headerBorder", displayName: "Header divider", value: { value: "#E5E7EB" } });
    rowDivider = new formattingSettings.ColorPicker({ name: "rowDivider", displayName: "Row divider", value: { value: "#EEF1F4" } });
    hoverBackground = new formattingSettings.ColorPicker({ name: "hoverBackground", displayName: "Hover background", value: { value: "#F5F9FC" } });
    selectedBackground = new formattingSettings.ColorPicker({ name: "selectedBackground", displayName: "Selected background", value: { value: "#E8F2FF" } });
    grid = new formattingSettings.ItemDropdown({ name: "grid", displayName: "Grid", items: [{ value: "none", displayName: "None" }, { value: "horizontal", displayName: "Horizontal" }, { value: "vertical", displayName: "Vertical" }, { value: "both", displayName: "Both" }], value: { value: "horizontal", displayName: "Horizontal" } });
    name: string = "tableSurface";
    displayName: string = "Table surface";
    slices: FormattingSettingsSlice[] = [this.headerBorder, this.rowDivider, this.hoverBackground, this.selectedBackground, this.grid];
}

class TotalsSettingsCard extends FormattingSettingsCard {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show totals row",
        displayNameKey: "Prop_ShowTotals",
        value: false
    });

    label = new formattingSettings.TextInput({
        name: "label",
        displayName: "Label",
        displayNameKey: "Prop_TotalsLabel",
        placeholder: "Total",
        value: "Total"
    });

    bgColor = new formattingSettings.ColorPicker({
        name: "bgColor",
        displayName: "Background color",
        displayNameKey: "Prop_BgColor",
        value: { value: "#F0F2F5" }
    });

    fontColor = new formattingSettings.ColorPicker({ name: "fontColor", displayName: "Text color", value: { value: "#333333" } });
    bold = new formattingSettings.ToggleSwitch({ name: "bold", displayName: "Bold text", value: true });
    showBorder = new formattingSettings.ToggleSwitch({ name: "showBorder", displayName: "Top border", value: true });
    alignment = new formattingSettings.ItemDropdown({
        name: "alignment", displayName: "Alignment",
        items: [{ value: "left", displayName: "Left" }, { value: "center", displayName: "Center" }, { value: "right", displayName: "Right" }],
        value: { value: "left", displayName: "Left" }
    });

    name: string = "totals";
    displayName: string = "Totals";
    displayNameKey: string = "Object_Totals";
    slices: FormattingSettingsSlice[] = [this.show, this.label, this.bgColor, this.fontColor, this.bold, this.showBorder, this.alignment];
}

/**
 * Virtual scrolling card. On by default -- performance-as-default, never
 * something the user has to discover or configure to get a smooth table.
 */
class VirtualScrollingSettingsCard extends FormattingSettingsCard {
    enabled = new formattingSettings.ToggleSwitch({
        name: "enabled",
        displayName: "Enabled",
        displayNameKey: "Prop_Enabled",
        value: true
    });

    rowHeight = new formattingSettings.NumUpDown({
        name: "rowHeight",
        displayName: "Row height",
        displayNameKey: "Prop_RowHeight",
        value: 36
    });

    name: string = "virtualScrolling";
    displayName: string = "Smooth scrolling";
    displayNameKey: string = "Object_VirtualScrolling";
    slices: FormattingSettingsSlice[] = [this.enabled, this.rowHeight];
}

/**
 * Toolbar visibility card.
 */
class ToolbarSettingsCard extends FormattingSettingsCard {
    showMenu = new formattingSettings.ToggleSwitch({
        name: "showMenu",
        displayName: "Show toolbar",
        displayNameKey: "Prop_ShowMenu",
        value: true
    });

    name: string = "toolbar";
    displayName: string = "Toolbar";
    displayNameKey: string = "Object_Toolbar";
    slices: FormattingSettingsSlice[] = [this.showMenu];
}

/**
 * Search visibility card.
 */
class SearchSettingsCard extends FormattingSettingsCard {
    enabled = new formattingSettings.ToggleSwitch({
        name: "enabled",
        displayName: "Enable search",
        displayNameKey: "Prop_EnableSearch",
        value: true
    });

    name: string = "search";
    displayName: string = "Search";
    displayNameKey: string = "Object_Search";
    slices: FormattingSettingsSlice[] = [this.enabled];
}

/**
 * Grouping card: whether multi-level groups (from the "Group by" role)
 * start expanded or collapsed.
 */
class GroupingSettingsCard extends FormattingSettingsCard {
    defaultExpanded = new formattingSettings.ToggleSwitch({ name: "defaultExpanded", displayName: "Expanded by default", displayNameKey: "Prop_DefaultExpanded", value: true });
    indentation = new formattingSettings.NumUpDown({ name: "indentation", displayName: "Indentation (px)", value: 16 });
    showCount = new formattingSettings.ToggleSwitch({ name: "showCount", displayName: "Show group count", value: true });
    showTotals = new formattingSettings.ToggleSwitch({ name: "showTotals", displayName: "Show group totals", value: true });

    name: string = "grouping";
    displayName: string = "Grouping";
    displayNameKey: string = "Object_Grouping";
    slices: FormattingSettingsSlice[] = [this.defaultExpanded, this.indentation, this.showCount, this.showTotals];
}

class PivotSettingsCard extends FormattingSettingsCard {
    enabled = new formattingSettings.ToggleSwitch({ name: "enabled", displayName: "Enable Pivot mode", value: true });
    showRowTotals = new formattingSettings.ToggleSwitch({ name: "showRowTotals", displayName: "Row totals", value: false });
    showColumnTotals = new formattingSettings.ToggleSwitch({ name: "showColumnTotals", displayName: "Column totals", value: false });
    showGrandTotal = new formattingSettings.ToggleSwitch({ name: "showGrandTotal", displayName: "Grand total", value: false });
    emptyValueDisplay = new formattingSettings.TextInput({ name: "emptyValueDisplay", displayName: "Empty value display", placeholder: "", value: "" });

    name: string = "pivot";
    displayName: string = "Pivot";
    slices: FormattingSettingsSlice[] = [this.enabled, this.showRowTotals, this.showColumnTotals, this.showGrandTotal, this.emptyValueDisplay];
}

/**
 * Column filters card: toggles the small filter icon in each header cell
 * that opens a type-aware filter popover (text / number / date).
 */
class FiltersSettingsCard extends FormattingSettingsCard {
    advancedExpression = new formattingSettings.TextArea({
        name: "advancedExpression",
        displayName: "Advanced filter expression",
        displayNameKey: "Prop_AdvancedFilterExpression",
        placeholder: "Region = 'East' AND Revenue >= 100",
        value: ""
    });

    showIcons = new formattingSettings.ToggleSwitch({
        name: "showIcons",
        displayName: "Show column filter icons",
        displayNameKey: "Prop_ShowFilterIcons",
        value: true
    });

    name: string = "filters";
    displayName: string = "Column filters";
    displayNameKey: string = "Object_Filters";
    slices: FormattingSettingsSlice[] = [this.advancedExpression, this.showIcons];
}

/**
 * Conditional formatting card: a two-color scale applied to measure cells
 * based on that column's min/max across the full (unfiltered) dataset.
 */
class ConditionalFormattingSettingsCard extends FormattingSettingsCard {
    enabled = new formattingSettings.ToggleSwitch({
        name: "enabled",
        displayName: "Enable color scale",
        displayNameKey: "Prop_EnableColorScale",
        value: false
    });

    minColor = new formattingSettings.ColorPicker({
        name: "minColor",
        displayName: "Low color",
        displayNameKey: "Prop_MinColor",
        value: { value: "#FDE2E2" }
    });

    midpointEnabled = new formattingSettings.ToggleSwitch({ name: "midpointEnabled", displayName: "Use midpoint", value: false });
    midpointColor = new formattingSettings.ColorPicker({ name: "midpointColor", displayName: "Midpoint color", value: { value: "#FFF4CE" } });

    maxColor = new formattingSettings.ColorPicker({
        name: "maxColor",
        displayName: "High color",
        displayNameKey: "Prop_MaxColor",
        value: { value: "#2E7D32" }
    });

    name: string = "conditionalFormatting";
    displayName: string = "Conditional formatting";
    displayNameKey: string = "Object_ConditionalFormatting";
    slices: FormattingSettingsSlice[] = [this.enabled, this.minColor, this.midpointEnabled, this.midpointColor, this.maxColor];
}

/**
 * Conditional URL actions card. Rules are authored as a JSON array (Power BI's
 * formatting pane doesn't support arbitrary user-added rows in a repeating UI),
 * evaluated top-to-bottom per row, first match wins. See ILinkActionRule in
 * tableRenderer.ts for the exact shape.
 *
 * `validationMessage` is a read-only slice that visual.ts toggles visible only
 * when the JSON in `rules` fails to parse or doesn't match the expected shape --
 * malformed input never throws or breaks rendering, it just quietly disables the
 * link-action feature and surfaces this one plain-language note in the pane.
 */
class ExportSettingsCard extends FormattingSettingsCard {
    enabled = new formattingSettings.ToggleSwitch({ name: "enabled", displayName: "Enable export", value: true });
    rowScope = new formattingSettings.ItemDropdown({
        name: "rowScope", displayName: "Rows to export",
        items: [
            { value: "visible", displayName: "Visible rows" },
            { value: "filtered", displayName: "Filtered rows" },
            { value: "selected", displayName: "Selected rows" },
            { value: "available", displayName: "Available rows" }
        ], value: { value: "filtered", displayName: "Filtered rows" }
    });
    includeHeaders = new formattingSettings.ToggleSwitch({ name: "includeHeaders", displayName: "Include headers", value: true });
    includeTotals = new formattingSettings.ToggleSwitch({ name: "includeTotals", displayName: "Include totals", value: false });
    csv = new formattingSettings.ToggleSwitch({ name: "csv", displayName: "CSV", value: true });
    excel = new formattingSettings.ToggleSwitch({ name: "excel", displayName: "Excel", value: false });
    json = new formattingSettings.ToggleSwitch({ name: "json", displayName: "JSON", value: false });
    pdf = new formattingSettings.ToggleSwitch({ name: "pdf", displayName: "PDF", value: true });
    name: string = "export";
    displayName: string = "Export";
    slices: FormattingSettingsSlice[] = [this.enabled, this.rowScope, this.includeHeaders, this.includeTotals, this.csv, this.excel, this.json, this.pdf];
}

class LegacyExportGovernanceSettingsCard extends FormattingSettingsCard {
    enabled = new formattingSettings.ToggleSwitch({ name: "enabled", displayName: "Legacy watermark", value: false, visible: false });
    watermarkText = new formattingSettings.TextInput({ name: "watermarkText", displayName: "Legacy watermark text", placeholder: "CONFIDENTIAL", value: "CONFIDENTIAL", visible: false });
    locale = new formattingSettings.TextInput({ name: "locale", displayName: "Legacy locale", placeholder: "en-UG", value: "en-UG", visible: false });
    currency = new formattingSettings.TextInput({ name: "currency", displayName: "Legacy currency", placeholder: "UGX", value: "UGX", visible: false });
    username = new formattingSettings.TextInput({ name: "username", displayName: "Legacy audit username", placeholder: "Not authenticated identity", value: "", visible: false });
    name: string = "exportGovernance";
    displayName: string = "Legacy export governance";
    visible: boolean = false;
    slices: FormattingSettingsSlice[] = [this.enabled, this.watermarkText, this.locale, this.currency, this.username];
}

class GovernanceSettingsCard extends FormattingSettingsCard {
    watermarkEnabled = new formattingSettings.ToggleSwitch({ name: "watermarkEnabled", displayName: "Enable watermark", value: false });
    watermarkText = new formattingSettings.TextInput({ name: "watermarkText", displayName: "Watermark text", placeholder: "CONFIDENTIAL", value: "CONFIDENTIAL" });
    watermarkPlacement = new formattingSettings.ItemDropdown({
        name: "watermarkPlacement", displayName: "Watermark placement",
        items: [{ value: "footer", displayName: "Footer" }, { value: "center", displayName: "Center" }],
        value: { value: "footer", displayName: "Footer" }
    });
    auditEnabled = new formattingSettings.ToggleSwitch({ name: "auditEnabled", displayName: "Record export audit events", value: true });
    legacyUsername = new formattingSettings.TextInput({ name: "username", displayName: "Audit username (legacy)", placeholder: "Not authenticated identity", value: "", visible: false });
    legacyLocale = new formattingSettings.TextInput({ name: "locale", displayName: "Legacy locale", placeholder: "en-UG", value: "en-UG", visible: false });
    legacyCurrency = new formattingSettings.TextInput({ name: "currency", displayName: "Legacy currency", placeholder: "UGX", value: "UGX", visible: false });
    name: string = "governance";
    displayName: string = "Governance";
    slices: FormattingSettingsSlice[] = [this.watermarkEnabled, this.watermarkText, this.watermarkPlacement, this.auditEnabled, this.legacyUsername, this.legacyLocale, this.legacyCurrency];
}

class RegionalFormatSettingsCard extends FormattingSettingsCard {
    locale = new formattingSettings.TextInput({ name: "locale", displayName: "Locale", placeholder: "en-UG", value: "en-UG" });
    currency = new formattingSettings.TextInput({ name: "currency", displayName: "Currency", placeholder: "UGX", value: "" });
    dateFormat = new formattingSettings.TextInput({ name: "dateFormat", displayName: "Date format", placeholder: "Model-defined", value: "" });
    numberSeparators = new formattingSettings.ItemDropdown({
        name: "numberSeparators", displayName: "Number separators",
        items: [{ value: "auto", displayName: "Auto / model" }, { value: "locale", displayName: "Locale" }],
        value: { value: "auto", displayName: "Auto / model" }
    });
    decimalPrecision = new formattingSettings.NumUpDown({ name: "decimalPrecision", displayName: "Decimal precision", value: -1 });
    name: string = "regionalFormat";
    displayName: string = "Regional format";
    slices: FormattingSettingsSlice[] = [this.locale, this.currency, this.dateFormat, this.numberSeparators, this.decimalPrecision];
}

class AdvancedSettingsCard extends FormattingSettingsCard {
    performanceMode = new formattingSettings.ItemDropdown({
        name: "performanceMode", displayName: "Performance mode",
        items: [{ value: "auto", displayName: "Automatic" }, { value: "balanced", displayName: "Balanced" }, { value: "maximum", displayName: "Maximum performance" }],
        value: { value: "auto", displayName: "Automatic" }
    });
    name: string = "advanced";
    displayName: string = "Advanced";
    slices: FormattingSettingsSlice[] = [this.performanceMode];
}

class LinkActionsSettingsCard extends FormattingSettingsCard {
    rules = new formattingSettings.TextArea({
        name: "rules",
        displayName: "Link rules (JSON)",
        placeholder: '[{"column":"Status","operator":"equals","value":"Overdue","urlTemplate":"https://portal.example.com/case/{CaseID}"}]',
        value: ""
    });

    iconColumn = new formattingSettings.TextInput({
        name: "iconColumn",
        displayName: "Icon column",
        placeholder: "Exact column name to show the link icon in",
        value: ""
    });

    validationMessage = new formattingSettings.ReadOnlyText({
        name: "validationMessage",
        displayName: "",
        value: "Check the formatting of your link rules",
        visible: false
    });

    name: string = "linkActions";
    displayName: string = "Link actions";
    slices: FormattingSettingsSlice[] = [this.rules, this.iconColumn, this.validationMessage];
}

/**
 * Top level formatting settings model, aggregating every card above.
 * Consumed by FormattingSettingsService in visual.ts.
 */
class StyleSettingsCard extends FormattingSettingsCard {
    preset = new formattingSettings.ItemDropdown({
        name: "preset",
        displayName: "Style preset",
        value: { value: "clean", displayName: "Clean", displayNameKey: "Preset_Clean" },
        items: stylePresetItems().map((item) => ({ value: item.value, displayName: item.displayName, displayNameKey: item.displayName }))
    });

    guidance = new formattingSettings.ReadOnlyText({
        name: "guidance",
        displayName: "",
        value: "Built-in presets apply as a coherent system. Choose Custom to keep your existing color and typography controls as overrides.",
        visible: true
    });

    name: string = "style";
    displayName: string = "Data Lake Tables style";
    displayNameKey: string = "Object_Style";
    slices: FormattingSettingsSlice[] = [this.preset, this.guidance];
}

/** Tier 3: one authoritative layout/density contract driven by the Power BI viewport. */


class LayoutSettingsCard extends FormattingSettingsCard {
    density = new formattingSettings.ItemDropdown({
        name: "density",
        displayName: "Density",
        value: { value: "comfortable", displayName: "Comfortable", displayNameKey: "Density_Comfortable" },
        items: densityItems().map((item) => ({ value: item.value, displayName: item.displayName, displayNameKey: item.displayName }))
    });

    responsive = new formattingSettings.ToggleSwitch({
        name: "responsive",
        displayName: "Responsive layout",
        value: true
    });

    autoFitColumns = new formattingSettings.ToggleSwitch({
        name: "autoFitColumns",
        displayName: "Auto-fit columns",
        value: true
    });

    minColumnWidth = new formattingSettings.NumUpDown({
        name: "minColumnWidth",
        displayName: "Minimum column width",
        value: 72
    });

    maxColumnWidth = new formattingSettings.NumUpDown({
        name: "maxColumnWidth",
        displayName: "Maximum column width",
        value: 420
    });

    headerWrap = new formattingSettings.ToggleSwitch({
        name: "headerWrap",
        displayName: "Wrap headers",
        value: false
    });

    cellWrap = new formattingSettings.ToggleSwitch({
        name: "cellWrap",
        displayName: "Wrap cells",
        value: false
    });

    textOverflow = new formattingSettings.ItemDropdown({
        name: "textOverflow",
        displayName: "Cell overflow",
        value: { value: "ellipsis", displayName: "Ellipsis", displayNameKey: "Overflow_Ellipsis" },
        items: [
            { value: "ellipsis", displayName: "Ellipsis" },
            { value: "clip", displayName: "Clip" }
        ]
    });

    name: string = "layout";
    displayName: string = "Responsive layout";
    displayNameKey: string = "Object_Layout";
    slices: FormattingSettingsSlice[] = [this.density, this.responsive, this.autoFitColumns, this.minColumnWidth, this.maxColumnWidth, this.headerWrap, this.cellWrap, this.textOverflow];
}

/**
 * Top level formatting settings model, aggregating every card above.
 * Consumed by FormattingSettingsService in visual.ts.
 */


class InteractionSettingsCard extends FormattingSettingsCard {
    enableSearch = new formattingSettings.ToggleSwitch({
        name: "enableSearch", displayName: "Enable search", displayNameKey: "Prop_EnableSearch", value: true
    });
    searchClear = new formattingSettings.ToggleSwitch({
        name: "searchClear", displayName: "Show clear search action", displayNameKey: "Prop_SearchClear", value: true
    });

    showFilterControls = new formattingSettings.ToggleSwitch({
        name: "showFilterControls", displayName: "Show filter controls", displayNameKey: "Prop_ShowFilterControls", value: true
    });
    multiColumnFiltering = new formattingSettings.ToggleSwitch({
        name: "multiColumnFiltering", displayName: "Allow multi-column filtering", displayNameKey: "Prop_MultiColumnFiltering", value: true
    });
    clearFilters = new formattingSettings.ToggleSwitch({
        name: "clearFilters", displayName: "Show clear filters action", displayNameKey: "Prop_ClearFilters", value: true
    });
    filterIndicator = new formattingSettings.ToggleSwitch({
        name: "filterIndicator", displayName: "Show filter indicator", displayNameKey: "Prop_FilterIndicator", value: true
    });
    filteredState = new formattingSettings.ToggleSwitch({
        name: "filteredState", displayName: "Show filtered-state indication", displayNameKey: "Prop_FilteredState", value: true
    });

    enableSorting = new formattingSettings.ToggleSwitch({
        name: "enableSorting", displayName: "Enable sorting", displayNameKey: "Prop_EnableSorting", value: true
    });
    sortIndicator = new formattingSettings.ToggleSwitch({
        name: "sortIndicator", displayName: "Show sort indicators", displayNameKey: "Prop_SortIndicator", value: true
    });

    enableRowSelection = new formattingSettings.ToggleSwitch({
        name: "enableRowSelection", displayName: "Enable row selection", displayNameKey: "Prop_EnableRowSelection", value: true
    });
    selectionMode = new formattingSettings.ItemDropdown({
        name: "selectionMode", displayName: "Selection mode", displayNameKey: "Prop_SelectionMode",
        items: [
            { value: "single", displayName: "Single" },
            { value: "multi", displayName: "Multiple" }
        ],
        value: { value: "multi", displayName: "Multiple" }
    });
    clearSelection = new formattingSettings.ToggleSwitch({
        name: "clearSelection", displayName: "Show clear selection action", displayNameKey: "Prop_ClearSelection", value: true
    });
    selectionHighlight = new formattingSettings.ToggleSwitch({
        name: "selectionHighlight", displayName: "Show selection highlight", displayNameKey: "Prop_SelectionHighlight", value: true
    });

    enableCopy = new formattingSettings.ToggleSwitch({
        name: "enableCopy", displayName: "Enable copy", displayNameKey: "Prop_EnableCopy", value: false
    });
    copyCell = new formattingSettings.ToggleSwitch({
        name: "copyCell", displayName: "Copy cell", displayNameKey: "Prop_CopyCell", value: true
    });
    copyRow = new formattingSettings.ToggleSwitch({
        name: "copyRow", displayName: "Copy row", displayNameKey: "Prop_CopyRow", value: true
    });
    copySelected = new formattingSettings.ToggleSwitch({
        name: "copySelected", displayName: "Copy selected", displayNameKey: "Prop_CopySelected", value: true
    });

    name: string = "interaction";
    displayName: string = "Interaction";
    displayNameKey: string = "Object_Interaction";
    slices: FormattingSettingsSlice[] = [
        this.enableSearch, this.searchClear,
        this.showFilterControls, this.multiColumnFiltering, this.clearFilters, this.filterIndicator, this.filteredState,
        this.enableSorting, this.sortIndicator,
        this.enableRowSelection, this.selectionMode, this.clearSelection, this.selectionHighlight,
        this.enableCopy, this.copyCell, this.copyRow, this.copySelected
    ];
}

/**
 * Grouping card: whether multi-level groups (from the "Group by" role)
 * start expanded or collapsed.
 */


export class VisualSettingsModel extends FormattingSettingsModel {
    general = new GeneralSettingsCard();
    style = new StyleSettingsCard();
    layout = new LayoutSettingsCard();
    interaction = new InteractionSettingsCard();
    header = new HeaderSettingsCard();
    cells = new CellsSettingsCard();
    formatting = new DataBarsSettingsCard();
    totals = new TotalsSettingsCard();
    tableSurface = new TableSurfaceSettingsCard();
    virtualScrolling = new VirtualScrollingSettingsCard();
    toolbar = new ToolbarSettingsCard();
    search = new SearchSettingsCard();
    grouping = new GroupingSettingsCard();
    pivot = new PivotSettingsCard();
    filters = new FiltersSettingsCard();
    conditionalFormatting = new ConditionalFormattingSettingsCard();
    exportSettings = new ExportSettingsCard();
    governance = new GovernanceSettingsCard();
    legacyExportGovernance = new LegacyExportGovernanceSettingsCard();
    regionalFormat = new RegionalFormatSettingsCard();
    advanced = new AdvancedSettingsCard();
    linkActions = new LinkActionsSettingsCard();

    cards: FormattingSettingsCard[] = [
        this.style,
        this.layout,
        this.general,
        this.header,
        this.cells,
        this.tableSurface,
        this.totals,
        this.formatting,
        this.interaction,
        this.search,
        this.filters,
        this.grouping,
        this.pivot,
        this.conditionalFormatting,
        this.exportSettings,
        this.governance,
        this.regionalFormat,
        this.virtualScrolling,
        this.toolbar,
        this.linkActions,
        this.advanced,
        this.legacyExportGovernance
    ];
}
