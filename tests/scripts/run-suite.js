#!/usr/bin/env node
const { spawnSync } = require("child_process");
const path = require("path");

const suiteArg = (process.argv[2] || "all").toLowerCase();
const extraArgs = process.argv.slice(3).map((arg) => arg.trim()).filter(Boolean);
const demoModeCandidates = new Set(["mock", "ollama"]);
let demoModeArg;
let demoSuiteFilter;

extraArgs.forEach((raw) => {
  const arg = raw.toLowerCase();
  if (demoModeCandidates.has(arg)) {
    demoModeArg = arg;
    return;
  }
  if (!demoSuiteFilter) demoSuiteFilter = arg;
});

const env = { ...process.env };
env.TEST_MODE = (env.TEST_MODE || env.AGENT_MODE || "mock").toLowerCase();
const runDemoEvalFlag = (env.RUN_DEMOEVAL || "0") === "1";

const macroPatterns = {
  format: "apply_format",
  chart: "create_chart",
  formula: "write_formula",
  join: "join_tables",
  multi: "multi_etapes",
  validate: "validate_data",
  view: "table_view",
  all: "",
};

const isWin = process.platform === "win32";
const jestBin = path.join(__dirname, "..", "..", "node_modules", ".bin", isWin ? "jest.cmd" : "jest");
const baseArgs = ["--runInBand"];

const run = (args, envVars) => {
  const cmd = isWin ? "cmd.exe" : jestBin;
  const spawnArgs = isWin ? ["/c", jestBin, ...args] : [...args];
  return spawnSync(cmd, spawnArgs, { stdio: "inherit", env: envVars });
};

const runDemoEval = (demoEnv) => {
  console.log(
    `[run-suite] launching demoEval suite=${demoEnv.DEMO_EVAL_SUITE || "all"} mode=${demoEnv.DEMO_EVAL_MODE}`
  );
  const res = run([...baseArgs, "tests/demoEval.e2e.test.ts"], demoEnv);
  return res.status === null ? 1 : res.status;
};

if (suiteArg === "demoeval") {
  const mode = demoModeArg || (env.DEMO_EVAL_MODE || "mock");
  env.DEMO_EVAL_MODE = mode;
  env.DEMO_EVAL_SUITE = demoSuiteFilter || env.DEMO_EVAL_SUITE || "all";
  env.OLLAMA_DISABLED = mode === "mock" ? "1" : "0";
  console.log(`[run-suite] suite=demoeval mode=${env.DEMO_EVAL_MODE} suite=${env.DEMO_EVAL_SUITE}`);
  const code = runDemoEval(env);
  process.exit(code);
}

env.DEMO_EVAL_MODE = "mock";
env.OLLAMA_DISABLED = "1";
console.log(`[run-suite] suite=${suiteArg} mode=mock`);

const pattern = macroPatterns[suiteArg] ?? macroPatterns.all;
const args = [...baseArgs, "--testPathIgnorePatterns", "demoEval\\.e2e\\.test\\.ts$"];
if (pattern) args.push("--testPathPattern", pattern);
const mainRun = run(args, env);
let exitCode = mainRun.status === null ? 1 : mainRun.status;

if (suiteArg === "all" && runDemoEvalFlag) {
  const demoEnv = {
    ...env,
    DEMO_EVAL_MODE: env.DEMO_EVAL_MODE || "mock",
  };
  demoEnv.OLLAMA_DISABLED = demoEnv.DEMO_EVAL_MODE === "ollama" ? "0" : "1";
  exitCode = exitCode || runDemoEval(demoEnv);
}

process.exit(exitCode);
