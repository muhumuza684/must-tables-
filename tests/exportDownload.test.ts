/**
 * Coverage for the export/download functions (csv/excel/json/pdf) and the
 * privilege-status messaging split added alongside them. Before this file,
 * nothing in the suite exercised exportCSV/exportExcel/exportJSON/exportPDF
 * or downloadContent()'s privilege-status branch directly -- this closes
 * that gap so a regression here is caught automatically instead of only in
 * a live Power BI Desktop check.
 */
import powerbi from "powerbi-visuals-api";
import { TableRenderer, ITableColumn, ITableRow, ITableRendererSettings } from "../src/tableRenderer";
import {
    makeFakeHost,
    makeFakeSelectionManager,
    makeFakeTooltipService,
    makeFakeLocalizationManager,
    makeFakeColorPalette
} from "./mocks/powerbiMocks";
import { col, row, makeSettings } from "./mocks/fixtures";

const VALUE_COLUMN: ITableColumn = col("Revenue", { isMeasure: true });
const ROW_COLUMN: ITableColumn = col("Region");

function makeRows(): ITableRow[] {
    return [
        row({ Region: "East", Revenue: 100 }),
        row({ Region: "West", Revenue: 200 })
    ];
}

function buildRenderer(hostOverrides: any = {}, settingsOverrides: Partial<ITableRendererSettings> = {}) {
    const host = { ...makeFakeHost(), ...hostOverrides };
    const container = document.createElement("div");
    const renderer = new TableRenderer(
        container,
        host,
        makeFakeSelectionManager(),
        makeFakeTooltipService(),
        makeFakeLocalizationManager(),
        makeFakeColorPalette()
    );
    const settings = makeSettings({
        exportSettings: {
            enabled: true,
            rowScope: "filtered",
            includeHeaders: true,
            includeTotals: false,
            csv: true,
            excel: true,
            json: true,
            pdf: true
        },
        ...settingsOverrides
    });
    renderer.setData([ROW_COLUMN], [], [VALUE_COLUMN], [], makeRows(), settings);
    return { renderer, host };
}

describe("export/download functions", () => {
    test("exportCSV sends real, non-empty CSV content through the privileged download API when allowed", async () => {
        const exportVisualsContentExtended = jest.fn().mockResolvedValue({ downloadCompleted: true });
        const { renderer, host } = buildRenderer({
            downloadService: {
                exportStatus: jest.fn().mockResolvedValue(powerbi.PrivilegeStatus.Allowed),
                exportVisualsContentExtended
            }
        });

        await (renderer as any).exportCSV();

        expect(exportVisualsContentExtended).toHaveBeenCalledTimes(1);
        const [content, filename, fileType] = exportVisualsContentExtended.mock.calls[0];
        expect(filename).toBe("data-lake-tables-export.csv");
        expect(fileType).toBe("text/csv");
        // Real header + both data rows must actually be present in the exported text,
        // not a placeholder -- this is the "is it really a CSV" check.
        expect(content).toContain("Region");
        expect(content).toContain("Revenue");
        expect(content).toContain("East");
        expect(content).toContain("West");
        expect(content).toContain("100");
        expect(content).toContain("200");
    });

    test("exportPDF sends real base64-encoded PDF bytes through the privileged download API when allowed", async () => {
        const exportVisualsContentExtended = jest.fn().mockResolvedValue({ downloadCompleted: true });
        const { renderer } = buildRenderer({
            downloadService: {
                exportStatus: jest.fn().mockResolvedValue(powerbi.PrivilegeStatus.Allowed),
                exportVisualsContentExtended
            }
        });

        await (renderer as any).exportPDF();

        expect(exportVisualsContentExtended).toHaveBeenCalledTimes(1);
        const [content, filename, fileType] = exportVisualsContentExtended.mock.calls[0];
        expect(filename).toBe("data-lake-tables-export.pdf");
        // The codebase's own convention: "base64" is passed as fileType for
        // binary exports (PDF/Excel) so Power BI knows to decode the content
        // string before writing bytes -- see the matching csv/xlsx calls.
        expect(fileType).toBe("base64");
        // A real PDF's base64 body decodes to bytes starting with "%PDF-" (the
        // standard PDF file signature) -- this is what actually distinguishes
        // "a real PDF was built" from "a placeholder string was sent".
        const decoded = Buffer.from(content, "base64").toString("latin1");
        expect(decoded.startsWith("%PDF-")).toBe(true);
    });

    test("a disabled format's button-driving toggle stops that export before any download call", async () => {
        const exportVisualsContentExtended = jest.fn().mockResolvedValue({ downloadCompleted: true });
        const { renderer } = buildRenderer(
            {
                downloadService: {
                    exportStatus: jest.fn().mockResolvedValue(powerbi.PrivilegeStatus.Allowed),
                    exportVisualsContentExtended
                }
            },
            {
                exportSettings: {
                    enabled: true,
                    rowScope: "filtered",
                    includeHeaders: true,
                    includeTotals: false,
                    csv: true,
                    excel: false, // matches the new "Excel off by default" simplification
                    json: false,
                    pdf: true
                }
            }
        );

        await (renderer as any).exportExcel();

        expect(exportVisualsContentExtended).not.toHaveBeenCalled();
    });

    describe("privilege-status messaging (downloadContent)", () => {
        const cases: Array<[powerbi.PrivilegeStatus, string]> = [
            [powerbi.PrivilegeStatus.DisabledByAdmin, "organization"],
            [powerbi.PrivilegeStatus.NotSupported, "not available here"],
            [powerbi.PrivilegeStatus.NotDeclared, "not enabled in this visual"]
        ];

        test.each(cases)("status %p produces a distinct, non-generic warning", async (status, expectedFragment) => {
            const displayWarningIcon = jest.fn();
            const { renderer } = buildRenderer({
                downloadService: {
                    exportStatus: jest.fn().mockResolvedValue(status),
                    exportVisualsContentExtended: jest.fn()
                },
                displayWarningIcon
            });

            const ok = await (renderer as any).exportCSV().then(() => true);
            expect(ok).toBe(true); // exportCSV itself never throws
            expect(displayWarningIcon).toHaveBeenCalledTimes(1);
            const [title, detail] = displayWarningIcon.mock.calls[0];
            expect(`${title} ${detail}`.toLowerCase()).toContain(expectedFragment.toLowerCase());
        });

        test("the three blocked-status messages are not identical to each other", async () => {
            const seen: string[] = [];
            for (const status of [
                powerbi.PrivilegeStatus.DisabledByAdmin,
                powerbi.PrivilegeStatus.NotSupported,
                powerbi.PrivilegeStatus.NotDeclared
            ]) {
                const displayWarningIcon = jest.fn();
                const { renderer } = buildRenderer({
                    downloadService: {
                        exportStatus: jest.fn().mockResolvedValue(status),
                        exportVisualsContentExtended: jest.fn()
                    },
                    displayWarningIcon
                });
                await (renderer as any).exportCSV();
                seen.push(JSON.stringify(displayWarningIcon.mock.calls[0]));
            }
            expect(new Set(seen).size).toBe(3);
        });
    });
});
