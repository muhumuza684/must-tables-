import * as fs from "fs";
import * as path from "path";

const ROOT = path.resolve(__dirname, "..");

const MOJIBAKE_PATTERNS = [
    /Ã/g,
    /Â/g,
    /â€™/g,
    /â€œ/g,
    /â€\u009D/g,
    /â€”/g,
    /â€“/g,
    /â€¦/g,
    /ðŸ/g,
    /ï»¿/g,
    /\uFFFD/g
];

function filesUnder(dir: string): string[] {
    if (!fs.existsSync(dir)) {
        return [];
    }

    const out: string[] = [];
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) {
            out.push(...filesUnder(full));
        } else {
            out.push(full);
        }
    }
    return out;
}

function isProductionTextFile(file: string): boolean {
    if (file.includes(`${path.sep}node_modules${path.sep}`)) return false;
    if (/\.bak(?:\d+)?$/i.test(file)) return false;
    if (file.includes(`${path.sep}_backups${path.sep}`)) return false;
    if (file.includes(`${path.sep}backup-`)) return false;
    return /\.(ts|tsx|js|jsx|less|css|json|resjson)$/.test(file);
}

describe("production source encoding integrity", () => {
    test("contains no obvious mojibake or replacement characters", () => {
        const roots = [path.join(ROOT, "src"), path.join(ROOT, "style"), path.join(ROOT, "stringResources")];
        const failures: string[] = [];

        for (const root of roots) {
            for (const file of filesUnder(root).filter(isProductionTextFile)) {
                const text = fs.readFileSync(file, "utf8");
                const relative = path.relative(ROOT, file);
                for (const pattern of MOJIBAKE_PATTERNS) {
                    if (pattern.test(text)) {
                        failures.push(`${relative} matches ${pattern}`);
                        pattern.lastIndex = 0;
                    }
                }
            }
        }

        expect(failures).toEqual([]);
    });

    test("keeps the known UI glyphs as real Unicode source characters", () => {
        const renderer = fs.readFileSync(path.join(ROOT, "src", "tableRenderer.ts"), "utf8");
        expect(renderer).toContain("datalake-tables-settings-button");
expect(renderer).not.toContain("toggle.innerHTML =");
        expect(renderer).toContain('document.createElementNS("http://www.w3.org/2000/svg", "svg")');
        expect(renderer).toContain('datalake-tables-settings-button');
        expect(renderer).toContain('remove.textContent = "×"');
    });
});
