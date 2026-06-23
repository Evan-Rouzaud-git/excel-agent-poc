/** @type {import('ts-jest').JestConfigWithTsJest} */
module.exports = {
  preset: "ts-jest",
  testEnvironment: "node",
  roots: ["<rootDir>/tests"],
  testRegex: "(/__tests__/.*|(\\.|/)(test|spec|jtest))\\.ts$",
  moduleFileExtensions: ["ts", "js", "json"],
  moduleNameMapper: {
    "\\.(css|less|scss)$": "identity-obj-proxy",
    "^office-js$": "<rootDir>/tests/mocks/officejs"
  },
  setupFilesAfterEnv: ["<rootDir>/tests/setupTests.ts"]
};
