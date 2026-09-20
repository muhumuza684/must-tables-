describe("product surface contracts", () => {
  it("uses exactly four learnable settings sections", () => {
    const src = require("fs").readFileSync("src/tableRenderer.ts", "utf8");
    expect(src).toContain("Layout");
    expect(src).toContain("Data Display");
    expect(src).toContain("User Experience");
    expect(src).toContain("Analysis & Export");
    expect(src).toContain("announcementDismissed");
  });
});
