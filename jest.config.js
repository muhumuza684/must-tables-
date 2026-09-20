/** Jest config for Skiba Tables' automated regression suite (T3 / E2).
 *  Runs against src/*.ts directly via ts-jest Ã¢â‚¬â€ no build step required.
 *  jsdom is required because TableRenderer manipulates `document` directly
 *  (it predates any virtual-DOM abstraction).
 */
module.exports = {
    preset: "ts-jest",
    testEnvironment: "jsdom",
    roots: ["<rootDir>/src", "<rootDir>/tests"],
    testMatch: ["**/tests/**/*.test.ts"],
    setupFilesAfterEnv: ["<rootDir>/tests/jest.setup.ts"],
    moduleFileExtensions: ["ts", "js"],
    transform: {
        "^.+\\.ts$": ["ts-jest", { tsconfig: "<rootDir>/tsconfig.test.json" }]
    },
    // d3's package.json points "main"/"module" at its ESM source (src/index.js, which
    // re-exports a couple dozen d3-* ESM sub-packages) -- ts-jest/CommonJS can't load that
    // directly. d3 also ships a pre-bundled UMD build (dist/d3.js) specifically for
    // non-ESM consumers; redirect the "d3" specifier there instead of fighting
    // transformIgnorePatterns across every d3-* transitive sub-package.
    moduleNameMapper: {
        "^d3$": "<rootDir>/node_modules/d3/dist/d3.js",
        "^jspdf$": "<rootDir>/node_modules/jspdf/dist/jspdf.umd.js",
        "^uuid$": "<rootDir>/node_modules/uuid/dist/index.js",
        "\\.(less|css)$": "<rootDir>/tests/mocks/styleMock.js"
    },
    clearMocks: true
};
