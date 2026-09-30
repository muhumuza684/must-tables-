describe("product surface contracts", () => {
  it("uses the current learnable settings sections and persists announcement state", () => {
    const src = require("fs").readFileSync("src/tableRenderer.ts", "utf8");

    expect(src).toContain("Layout");
    expect(src).toContain("Data Display");
    expect(src).toContain("Filters");
    expect(src).toContain("Styling");
    expect(src).toContain("Analysis & Export");
    expect(src).toContain("announcementDismissed");
  });
});
