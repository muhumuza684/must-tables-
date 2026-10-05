# How MUST Tables is built

This guide explains how the visual works and how to change it safely. Read it top to bottom once; after that use it as a reference.

## 1. What this project is

A **Power BI custom visual**: TypeScript for the logic, LESS for the styling, packaged by `powerbi-visuals-tools` into a single `.pbiviz` file. Power BI hands the visual data and settings; the visual draws everything itself in plain DOM (no framework).

## 2. What you need

- Node.js 20 and Git
- Power BI Desktop (to try the result)
- Nothing else to install globally - the packaging tool is a dev dependency

## 3. Build and try it

```
npm ci              # exact versions from package-lock.json
npm run ci          # type-check -> tests -> lint -> package
```

`npm run ci` is exactly what GitHub Actions runs, so if it passes locally the CI build passes too.

Useful single steps: `npm run typecheck`, `npm test`, `npm run lint`, `npm run package`. `npm start` is an optional live-reload server for the Power BI service developer visual; Desktop users just package and import.

**Try it in Power BI Desktop:** Visualizations > ... > Import a visual from a file > `dist/<name>.pbiviz`, then add the visual to a page.

> **Always raise the version before re-importing.** Power BI caches a visual by its GUID and version. Change `visual.version` in `pbiviz.json` (for example `1.0.0.50` -> `1.0.0.51`). Never change the GUID: that would make every existing report lose the visual. If you still see the old look, close Power BI completely, reopen, import, and add the visual to a new page.

## 4. Project map

| Path | What it is |
|---|---|
| `pbiviz.json` | Visual identity: display name, GUID, version, links, icon, style entry |
| `capabilities.json` | What Power BI may give the visual: data roles, data mapping, every format-pane setting |
| `src/visual.ts` | The entry point. Receives data in `update()`, builds the format pane in `getFormattingModel()` |
| `src/tableRenderer.ts` | The big one. Draws the table, toolbar, settings drawer, landing page, info panels and the in-report ad card |
| `src/visualSettings.ts` | The format-pane model (matches `capabilities.json`) |
| `src/styleSystem.ts` | Style presets and density, turned into CSS variables |
| `src/calcEngine.ts`, `advancedModel.ts`, `calcDependencyOrder.ts` | Calculations and the model behind them |
| `src/state/pivotState.ts` | Pivot mode state |
| `src/persistedState.ts`, `savedViewState.ts`, `layoutState.ts`, `layoutPresets.ts` | Saved views and layouts |
| `src/tier4Formatting.ts`, `tier4Governance.ts`, `dataFormatting.ts`, `analyticsFormatting.ts` | Number/colour formatting and export governance |
| `src/responsiveLayout.ts`, `mobileLayout.ts`, `releaseHardening.ts` | Small-tile and mobile behaviour |
| `src/errorHandling.ts`, `renderDiagnostics.ts`, `releaseDiagnostics.ts` | Safe failure messages |
| `style/visual.less` | Main stylesheet (see section 7) |
| `style/presentationConsolidated.less` | Second stylesheet, loaded after the first |
| `stringResources/en-US/resources.resjson` | Translatable texts |
| `assets/icon.png` | The icon in the Visualizations pane |
| `tests/` | 30 Jest suites (jsdom) |
| `.github/workflows/must-tables-ci.yml` | The clean build and release pipeline |

## 5. How data flows

1. In Power BI the report author drops fields into **Rows**, **Values**, **Group by**, **Tooltips** or **Permissions** (the data roles in `capabilities.json`). Power BI sends up to 30,000 rows per window and more through `fetchMoreData`.
2. `visual.ts` `update()` reads the data and the settings, resolves the style tokens (`styleSystem.ts`) and calls `TableRenderer.setData(...)`.
3. `TableRenderer` virtualises the rows, applies search, filters, sorting, grouping or pivot, and draws only what is visible.
4. With **no fields** assigned (`supportsLandingPage`) it draws the landing page instead: `renderLandingPage()`.
5. User choices (layout, columns, saved views) are written back through Power BI properties so they survive saving the report.

## 6. The landing page, ad card and panels

All three live in `tableRenderer.ts`:

- `renderLandingPage()` and `makeLandingPreview()` - the landing page and its sample table
- `renderAnnouncement()` - the small in-report ad card; clicking it opens a panel
- `renderLandingInfoModal()` - the FAQ, Support, About and product-details panels
- Texts come from `this.loc("Key", "default text")`: the default is in the code, a translation can be added in `resources.resjson`.

**Rebranding checklist** (what was changed for MUST Tables): `displayName` in `pbiviz.json`; the visual and settings names in `capabilities.json` and `visualSettings.ts`; texts in `resources.resjson` and `tableRenderer.ts`; `assets/icon.png`; and the tests that assert those texts. Do not change the GUID, `name` or `visualClassName`.

## 7. Styling - how it is organised and how to change it

`visual.less` grew in layers over many releases. The **final, authoritative rules are three marked blocks at the end of `visual.less`**:

| Block (marker comment) | Styles |
|---|---|
| `DLT-LANDING-PREMIUM` | The landing page: layout, cards, sample table, buttons, footer |
| `DLT-PANELS-AD` | The ad card and the FAQ / Support / About / product-details panels |
| `DLT-READABILITY` | Text sizes and colours in the settings drawer, toolbar and pivot helper text, empty states, and the landing feature line |

**Rules for working with them**

1. **Change the block, do not add a competing rule somewhere else.** Competing rules were the cause of most of the earlier "my change did nothing" problems.
2. **Selectors repeat `.data-lake-tables-visual` three times** (the landing page also repeats `.skiba-landing-page` three times plus `[data-dlt-tier="2"]`) and every declaration is `!important`. That is what lets the blocks win over older layers. Keep the same prefix when you add a rule to a block.
3. **Size things with container units** (`cqmin`) inside `clamp(min, value, max)` so text scales with the visual's own size, not the browser window. The visual root is a size container.
4. **Set text sizes on the element itself.** A size set on a parent only reaches a child by inheritance, and any rule that targets the child directly beats it. (This is why the landing feature line stayed small until the size was put on each `<li>`.)
5. **Table data and header cells take their font and size from the report author's formatting settings.** Do not override them in CSS.
6. After editing, compile the styles: `node node_modules/less/bin/lessc style/visual.less out.css`. Then run `npm run ci`.

**Palette and text scale**

| Use | Colour |
|---|---|
| Brand navy (headings, buttons) | `#104F9F` |
| Magenta accent | `#EC0A8C` |
| Yellow accent | `#FAF623` |
| Cyan | `#308ABE` (large text only) |
| Dark cyan (small text) | `#1F6F9F` |
| Body text / muted text | `#33475F` / `#5B6E85` |
| Lines / soft background | `#DCE6F1` / `#F3F8FD` |

Keep text at 12px or larger and contrast at 4.5:1 or better. The settings drawer uses body 16-20px, labels 15.5-19px and helper text 14.5-18px.

## 8. Adding or changing a setting

A format-pane setting needs three things that must agree:

1. the object and property in `capabilities.json`
2. the matching card and slice in `src/visualSettings.ts`
3. use of the value in `visual.ts` / `tableRenderer.ts`, plus a display name in `resources.resjson`

Then add or extend a test and run `npm run ci`.

## 9. Tests

`npm test` runs 30 suites with Jest and jsdom, straight on the TypeScript (no build step). They cover calculations, pivoting, formatting, saved views, export, `fetchMoreData` segments, layout and a regression matrix of saved reports. Some tests read source files as text and assert that certain class names and strings exist, so renaming a class or a user-visible text means updating its test too.

## 10. Releasing

1. Raise `visual.version` in `pbiviz.json`.
2. Commit and push to `main`. GitHub Actions runs the full gate and stores the build as the `must-tables-pbiviz` artifact (30 days).
3. For a permanent file, push a tag that matches the version: `git tag -a v1.0.0.51 -m "MUST Tables v1.0.0.51"` then `git push origin v1.0.0.51`. A GitHub Release is created with the `.pbiviz` attached.

## 11. Troubleshooting

| Symptom | Fix |
|---|---|
| Power BI still shows the old look | Raise the version, close Power BI completely, reopen, import, add the visual to a **new** page |
| A CSS change does nothing | Find the rule that wins (browser dev tools on the visual), then edit the final block (section 7) instead of adding another rule |
| `No tests found` on Windows | The project is inside OneDrive and files are cloud placeholders: right-click the folder > Always keep on this device, or move the project out of OneDrive |
| `Connection was reset` on `git push` | The network is blocking GitHub: retry, or use another network |
| `LF will be replaced by CRLF` | Harmless Windows line-ending notice |
| `npm ci` fails in CI but `npm install` works | `package.json` and `package-lock.json` are out of sync: run `npm install`, commit both |
