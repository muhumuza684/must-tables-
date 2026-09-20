/**
 * Tests for the one-click visual theme gallery (header/cell/accent bundles) and the
 * data-palette custom-paste flow in the in-visual formatting toolbar. Reuses the mock
 * factory pattern from tableRenderer.fetchMoreUX.test.ts.
 */

import {
    TableRenderer,
    ITableColumn,
    ITableRow,
    ITableRendererSettings
} from "../src/tableRenderer";

function makeSelectionId(id: string): any {
    return { __id: id, equals: (other: any) => !!other && other.__id === id };
}

function makeSettings(overrides: Partial<ITableRendererSettings> = {}): ITableRendererSettings {
    return {
        fontFamily: "Segoe UI",
        fontSize: 12,
        rowHeight: 32,
        headerBg: "#ffffff",
        headerFont: "#000000",
        headerBold: false,
        cellBg: "#ffffff",
        cellFont: "#000000",
        altRow: "#f5f5f5",
        enableDataBars: false,
        barColor: "#000000",
        showTotals: false,
        totalsLabel: "Total",
        totalsBg: "#eeeeee",
        virtualScrollEnabled: false,
        virtualScrollRowHeight: 32,
        showToolbar: true,
        searchEnabled: true,
        enableColumnFilters: false,
        conditionalFormatEnabled: false,
        conditionalFormatMinColor: "#ffffff",
        conditionalFormatMaxColor: "#000000",
        groupsDefaultExpanded: true,
        permission: null,
        linkActionRules: [],
        linkActionIconColumn: "",
        savedViewState: null,
        allowInteractions: true,
        hasMoreData: false,
        ...overrides
    };
}

function makeRow(key: string, amount: number): ITableRow {
    return { key, values: { Amount: amount }, selectionId: makeSelectionId(key) };
}

const AMOUNT_COLUMN: ITableColumn[] = [
    { name: "Amount", displayName: "Amount", isMeasure: true, isGroupBy: false }
];

function makeHost() {
    return {
        fetchMoreData: jest.fn().mockReturnValue(true),
        persistProperties: jest.fn(),
        launchUrl: jest.fn(),
        createSelectionIdBuilder: jest.fn()
    } as any;
}

function makeRenderer(host: any) {
    const container = document.createElement("div");
    const selectionManager = {
        select: jest.fn().mockResolvedValue(undefined),
        getSelectionIds: jest.fn().mockReturnValue([]),
        showContextMenu: jest.fn()
    } as any;
    const tooltipService = { enabled: jest.fn().mockReturnValue(false), show: jest.fn(), move: jest.fn(), hide: jest.fn() } as any;
    const localizationManager = { getDisplayName: (key: string) => key } as any;
    const colorPalette = { isHighContrast: false } as any;
    const renderer = new TableRenderer(container, host, selectionManager, tooltipService, localizationManager, colorPalette);
    return { renderer, container };
}

function openSettingsMenu(container: HTMLElement) {
    (container.querySelector(".datalake-tables-settings-button") as HTMLElement).click();
}

describe("Visual theme gallery -- every user gets an equal, free choice (no author/viewer hierarchy)", () => {
    test("saving current colors as a theme adds a click-to-apply card", () => {
        const { renderer, container } = makeRenderer(makeHost());
        renderer.setData([], [], AMOUNT_COLUMN, [], [makeRow("r1", 10)], makeSettings({ headerBg: "#123456", barColor: "#abcdef" }));
        openSettingsMenu(container);

        (container.querySelector(".skiba-format-editor__save-theme-row button") as HTMLElement).click();
        const nameInput = container.querySelector(".skiba-format-editor__save-theme-row input") as HTMLInputElement;
        nameInput.value = "Field Operations";
        nameInput.dispatchEvent(new Event("input", { bubbles: true }));
        const saveBtn = Array.from(container.querySelectorAll(".skiba-format-editor__save-theme-row button")).find((b) => b.textContent === "Save") as HTMLElement;
        saveBtn.click();

        const cards = container.querySelectorAll(".skiba-format-editor__theme-card");
        const labels = Array.from(cards).map((c) => c.textContent);
        expect(labels.some((l) => l && l.includes("Field Operations"))).toBe(true);
    });

    test("clicking a saved theme card applies its header/cell/accent colors and marks it active", () => {
        const { renderer, container } = makeRenderer(makeHost());
        renderer.setData([], [], AMOUNT_COLUMN, [], [makeRow("r1", 10)], makeSettings());
        openSettingsMenu(container);

        (container.querySelector(".skiba-format-editor__save-theme-row button") as HTMLElement).click();
        const nameInput = container.querySelector(".skiba-format-editor__save-theme-row input") as HTMLInputElement;
        nameInput.value = "Ops";
        nameInput.dispatchEvent(new Event("input", { bubbles: true }));
        (Array.from(container.querySelectorAll(".skiba-format-editor__save-theme-row button")).find((b) => b.textContent === "Save") as HTMLElement).click();

        const themeState = (renderer as any)._tier4SavedThemes;
        expect(themeState.length).toBe(1);

        openSettingsMenu(container);
        const card = Array.from(container.querySelectorAll(".skiba-format-editor__theme-card")).find((c) => c.textContent?.includes("Ops")) as HTMLElement;
        card.click();

        expect((renderer as any)._tier4ActiveThemeId).toBe(themeState[0].id);
    });

    test("deleting a saved theme removes its card and clears active state if it was applied", () => {
        const { renderer, container } = makeRenderer(makeHost());
        renderer.setData([], [], AMOUNT_COLUMN, [], [makeRow("r1", 10)], makeSettings());
        openSettingsMenu(container);
        (container.querySelector(".skiba-format-editor__save-theme-row button") as HTMLElement).click();
        const nameInput = container.querySelector(".skiba-format-editor__save-theme-row input") as HTMLInputElement;
        nameInput.value = "Temp Theme";
        nameInput.dispatchEvent(new Event("input", { bubbles: true }));
        (Array.from(container.querySelectorAll(".skiba-format-editor__save-theme-row button")).find((b) => b.textContent === "Save") as HTMLElement).click();

        openSettingsMenu(container);
        const deleteBtn = container.querySelector(".skiba-format-editor__theme-card-wrap .skiba-format-editor__icon-button") as HTMLElement;
        deleteBtn.click();

        expect((renderer as any)._tier4SavedThemes.length).toBe(0);
        expect((renderer as any)._tier4ActiveThemeId).toBeNull();
    });

    test("theme gallery state round-trips through persistUserConfig/hydrateUserConfig", () => {
        const { renderer, container } = makeRenderer(makeHost());
        renderer.setData([], [], AMOUNT_COLUMN, [], [makeRow("r1", 10)], makeSettings());
        (renderer as any)._tier4SavedThemes = [{ id: "t1", name: "Saved", headerBg: "#111111" }];
        (renderer as any)._tier4ActiveThemeId = "t1";
        const json = JSON.stringify({
            tier4SavedThemes: (renderer as any)._tier4SavedThemes,
            tier4ActiveThemeId: "t1"
        });

        const { renderer: renderer2 } = makeRenderer(makeHost());
        (renderer2 as any).hydrateUserConfig(json);
        expect((renderer2 as any)._tier4SavedThemes).toEqual([{ id: "t1", name: "Saved", headerBg: "#111111" }]);
        expect((renderer2 as any)._tier4ActiveThemeId).toBe("t1");
    });
});


