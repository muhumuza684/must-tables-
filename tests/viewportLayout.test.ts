describe("enterprise viewport layout contracts", () => {
  it("uses viewport custom property names rather than fixed visual dimensions", () => {
    const css = require("fs").readFileSync("style/visual.less", "utf8");
    expect(css).toContain("--dlt-viewport-width");
    expect(css).toContain("width: 100%;");
    expect(css).toContain("height: 100%;");
    expect(css).toContain("@container (max-width: 760px)");
    expect(css).toContain(".data-lake-tables-visual");
  });

  it("reads the current Power BI viewport on every update", () => {
    const text = require("fs").readFileSync("src/visual.ts", "utf8");
    expect(text).toContain("const viewportWidth = options.viewport.width;");
    expect(text).toContain("const viewportHeight = options.viewport.height;");
    expect(text).toContain("this.resizeViewport(viewportWidth, viewportHeight);");
  });
});
