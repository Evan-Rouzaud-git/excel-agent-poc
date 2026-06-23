import { expect } from "@jest/globals";
import { runDemoEvalSuites, pickSuites } from "./demoEvalRunner";

const MODE = process.env.DEMO_EVAL_MODE === "mock" ? "mock" : "ollama";

const suiteEnv = process.env.DEMO_EVAL_SUITE;
const suites = pickSuites(suiteEnv ? suiteEnv.split(",") : undefined);

runDemoEvalSuites(suites, { mode: MODE });
