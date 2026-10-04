import * as fs from "fs";
import * as path from "path";

const ROOT = path.resolve(__dirname, "..");
const renderer = (): string => fs.readFileSync(path.join(ROOT, "src", "tableRenderer.ts"), "utf8");
const visual = (): string => fs.readFileSync(path.join(ROOT, "src", "visual.ts"), "utf8");
const less = (): string => fs.readFileSync(path.join(ROOT, "style", "visual.less"), "utf8");

describe("post-import visual refinement contracts", () => {
    test("filter control is a real button with a robust click target and popover", () => {
        const text = renderer();
        expect(text).toContain('filterBtn.type = "button"');
        expect(text).toContain("evt.preventDefault();");
        expect(text).toContain("this.openFilterPopover(col, filterBtn);");
        expect(text).toContain('document.addEventListener("keydown", onKeyDown);');
    });

    test("host viewport is consumed without attempting to resize the Power BI tile", () => {
        const text = visual();
        expect(text).toContain('const viewportWidth = options.viewport.width;');
        expect(text).toContain('const viewportHeight = options.viewport.height;');
        expect(text).toContain('this.resizeViewport(viewportWidth, viewportHeight);');
        expect(text).toContain('this.tableContainer.style.setProperty("--dlt-viewport-width", `${width}px`);');
        expect(text).toContain('--dlt-viewport-width');
        expect(text).toContain('--dlt-viewport-height');
        expect(text).toContain("data-lake-tables-visual--compact");
    });

    test("in-report advert is an anchored, dismissible native product card", () => {
        const text = renderer();
        const styles = less();
        expect(text).toContain('banner.className = "datalake-announcement"');
        expect(text).toContain('close.className = "datalake-announcement__close"');
        expect(text).toContain('this._announcementDismissed = true;');
        expect(text).toContain('banner.setAttribute("aria-haspopup", "dialog");');
        expect(text).toContain('text.textContent = "A cleaner workspace for search, grouping, pivoting and table analysis.";');
        expect(text).toContain('cta.textContent = "View product details";');
        expect(text).not.toContain('image.src = DATA_LAKE_SHOWCASE_POSTERS[0].dataUri;');
        expect(styles).toContain('.datalake-announcement-root {');
        expect(styles).toContain('inset-inline-end:14px;');
        expect(styles).toContain('inset-block-end:14px;');
        expect(styles).toContain('.datalake-announcement__close {');
    });

    test("settings drawer keeps its actual close and resize wiring", () => {
        const text = renderer();
        expect(text).toContain('close.textContent = "×";');
        expect(text).toContain('this._settingsDrawerWidth = this.clampSettingsDrawerWidth(startWidth + delta);');
        expect(text).toContain('resizer.addEventListener("pointerdown"');
        expect(text).toContain('resizer.addEventListener("keydown"');
    });

    test("in-visual layout overrides survive Power BI updates and virtual scrolling", () => {
        const text = renderer();
        expect(text).toContain("private _userLayoutOverrides: ILayoutState | null = null;");
        expect(text).toContain("this.settings.virtualScrollRowHeight = this._userLayoutOverrides.rowHeight;");
        expect(text).toContain('objectName: "general", selector: null, properties: generalProps');
        expect(text).toContain('addRange("Font size", 8, 48, 0.5');
        expect(text).toContain('addRange("Row height", 22, 120, 1');
        expect(text).toContain("datalake-layout-control__number");
    });

    test("settings drawer uses live visual width and preserves user resizing", () => {
        const text = renderer();
        const styles = less();
        expect(text).toContain("getSettingsDrawerWidthForViewport");
        expect(text).toContain("clampSettingsDrawerWidth");
        expect(text).toContain("settingsDrawerWidth");
        expect(text).toContain("Math.round(width * 0.75)");
        expect(styles).toContain("width:75%;");
        expect(styles).toContain("height:100% !important;");
        expect(styles).not.toContain("width:75% !important;");
    });

    test("advert opens accessible product details with the required attribution", () => {
        const text = renderer();
        expect(text).toContain("openAnnouncementDetails");
        expect(text).toContain("aria-haspopup");
        expect(text).toContain("DELIGHT BI · Data Guided Decisions");
        expect(text).toContain("Back to table");
        expect(text).toContain("No results found");
        expect(text).toContain('makePanel("layout"');
        expect(text).toContain('makePanel("display"');
        expect(text).toContain('makePanel("experience"');
        expect(text).toContain('makePanel("analysis"');
    });

    test("exact font size and row height are final CSS authorities, not viewport clamps", () => {
        const styles = less();
        expect(styles).toContain("font-size:var(--skiba-font-size,14px) !important;");
        expect(styles).toContain("font-size:var(--skiba-font-size,14px) !important; }");
        expect(styles).toContain("height:var(--skiba-row-height,36px);");
        expect(styles).toContain("min-height:var(--skiba-row-height,36px);");
        expect(styles).not.toContain("font-size: clamp(14px, calc(var(--dlt-viewport-width, 900px) / 56), 18px);");
    });

    test("landing page uses the full product attribution and feature rail", () => {
        const text = renderer();
        const styles = less();
        expect(text).toContain("skiba-landing-page__editorial-hero");
        expect(text).toContain("DELIGHT BI · Data Guided Decisions");
        expect(styles).toContain(".skiba-landing-page__editorial-hero");
    });
});



