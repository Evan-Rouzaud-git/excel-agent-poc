/* global Office */

import { getWorkbookContext } from "./context/getWorkbookContext";
import { WorkbookContextSnapshot, ContextOptions, ContextLogger } from "./context/types";
import { executePlan } from "./agent/executor";
import { validatePlan } from "./agent/planSchema";
import { AgentPlan, ExecutionResult } from "./agent/types";
import { planWithOllama } from "./agent/planner/ollamaPlanner";
import { applyConfirmationsToPlan } from "./agent/applyConfirmations";
import { repairPlanWriteFormulas } from "./agent/planRepairer";
import { validatePlanInvariants } from "./agent/planInvariants";
import { canonicalizePlan } from "./agent/canonicalizePlan";

const runtimeEnv = typeof process !== "undefined" ? (process as any).env : undefined;
const runtimeWindow = typeof window !== "undefined" ? (window as any) : undefined;
const runtimeConfig = {
  ollamaTimeoutMs: Number(runtimeWindow?.OLLAMA_TIMEOUT_MS || runtimeEnv?.OLLAMA_TIMEOUT_MS || 90000),
  ollamaModel: runtimeWindow?.OLLAMA_MODEL || runtimeEnv?.OLLAMA_MODEL || "qwen3:8b",
};
const MODEL_STORAGE_KEY = "ollama_model_choice";
try {
  if (typeof window !== "undefined" && window.localStorage) {
    const stored = window.localStorage.getItem(MODEL_STORAGE_KEY);
    if (stored) runtimeConfig.ollamaModel = stored;
  }
} catch {
  // storage not available
}

let lastSnapshot: WorkbookContextSnapshot | null = null;
let isRefreshing = false;
let agentRunning = false;
let pendingDecisions: Record<string, string> = {};
let confirmationAttempts = 0;
let lastPlanWasRepaired = false;
let lastRepairNotes: string[] = [];

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const sideload = document.getElementById("sideload-msg");
    const appBody = document.getElementById("app-body");
    if (sideload) sideload.style.display = "none";
    if (appBody) appBody.style.display = "flex";
    wireUi();
  }
});

function wireUi(): void {
  wireTabs();
  wireContextUi();
  wireAgentUi();
  wirePlannerUi();
}

function wireTabs() {
  const tabButtons = Array.from(document.querySelectorAll(".tab-button"));
  tabButtons.forEach((btn) =>
    btn.addEventListener("click", () => {
      const target = btn.getAttribute("data-tab");
      tabButtons.forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      const panels = Array.from(document.querySelectorAll(".tab-panel"));
      panels.forEach((p) => p.classList.remove("active"));
      const targetPanel = document.getElementById(target || "");
      if (targetPanel) targetPanel.classList.add("active");
    })
  );
}

function wireContextUi() {
  const refreshBtn = document.getElementById("refresh-context") as HTMLButtonElement | null;
  const copyBtn = document.getElementById("copy-json") as HTMLButtonElement | null;
  const clearLogsBtn = document.getElementById("clear-logs") as HTMLButtonElement | null;

  if (refreshBtn) refreshBtn.onclick = () => refreshContext();
  if (copyBtn) copyBtn.onclick = () => copyJson();
  if (clearLogsBtn) clearLogsBtn.onclick = () => clearLogs();
}

function wireAgentUi() {
  const runBtn = document.getElementById("run-plan") as HTMLButtonElement | null;
  const clearBtn = document.getElementById("clear-agent-logs") as HTMLButtonElement | null;
  const copyPlanBtn = document.getElementById("copy-agent-plan") as HTMLButtonElement | null;
  if (runBtn) runBtn.onclick = () => runPlan();
  if (clearBtn) clearBtn.onclick = () => clearAgentLogs();
  if (copyPlanBtn) copyPlanBtn.onclick = () => copyNodeText(document.getElementById("plan-input"));
}

function wirePlannerUi() {
  const writeBtn = document.getElementById("btnWritePlan") as HTMLButtonElement | null;
  const writeRunBtn = document.getElementById("btnWriteAndRun") as HTMLButtonElement | null;
  const copyPlanBtn = document.getElementById("copy-llm-plan") as HTMLButtonElement | null;
  const copyRawBtn = document.getElementById("copy-llm-raw") as HTMLButtonElement | null;
  const copyErrBtn = document.getElementById("copy-llm-errors") as HTMLButtonElement | null;
  initModelSelector();
  if (writeBtn) writeBtn.onclick = () => writePlan(false);
  if (writeRunBtn) writeRunBtn.onclick = () => writePlan(true);
  if (copyPlanBtn) copyPlanBtn.onclick = () => copyNodeText(document.getElementById("llmPlanOutput"));
  if (copyRawBtn) copyRawBtn.onclick = () => copyNodeText(document.getElementById("llmRawOutput"));
  if (copyErrBtn) copyErrBtn.onclick = () => copyNodeText(document.getElementById("llmErrors"));
}

function readOptionsFromUi(): ContextOptions {
  const logToggle = document.getElementById("toggle-log") as HTMLInputElement | null;
  const verboseToggle = document.getElementById("toggle-verbose") as HTMLInputElement | null;
  return {
    logTimings: logToggle ? logToggle.checked : false,
    verboseLogs: verboseToggle ? verboseToggle.checked : false,
  };
}

function initModelSelector() {
  const select = document.getElementById("ollama-model") as HTMLSelectElement | null;
  if (!select) return;
  const stored = ((): string | null => {
    try {
      return window.localStorage?.getItem(MODEL_STORAGE_KEY) || null;
    } catch {
      return null;
    }
  })();
  const initial = stored || runtimeConfig.ollamaModel || "qwen3:8b";
  select.value = initial;
  runtimeConfig.ollamaModel = initial;
  select.onchange = () => {
    const val = select.value || "qwen3:8b";
    runtimeConfig.ollamaModel = val;
    try {
      window.localStorage?.setItem(MODEL_STORAGE_KEY, val);
    } catch {
      // ignore
    }
  };
}

function getSelectedModel(): string {
  const select = document.getElementById("ollama-model") as HTMLSelectElement | null;
  if (select && select.value) return select.value;
  return runtimeConfig.ollamaModel || "qwen3:8b";
}

async function refreshContext(): Promise<void> {
  if (isRefreshing) return;
  isRefreshing = true;
  setStatus("Refreshing context...", "info");
  toggleRefreshDisabled(true);
  try {
    const options = readOptionsFromUi();
    const logger = createUiLogger(options.verboseLogs ?? false);
    const snapshot = await getWorkbookContext({ ...options, logger });
    lastSnapshot = snapshot;
    renderSummary(snapshot);
    renderJson(snapshot);
    setStatus(`Context refreshed (${snapshot.totals.durationMs} ms)`, "success");
  } catch (err: any) {
    console.error(err);
    setStatus(`Refresh failed: ${err?.message || err}`, "error");
  } finally {
    toggleRefreshDisabled(false);
    isRefreshing = false;
  }
}

async function refreshContextWithResult(): Promise<WorkbookContextSnapshot | null> {
  await refreshContext();
  return lastSnapshot;
}

function toggleRefreshDisabled(disabled: boolean) {
  const btn = document.getElementById("refresh-context") as HTMLButtonElement | null;
  if (btn) {
    btn.disabled = disabled;
    btn.textContent = disabled ? "Refreshing..." : "Refresh Context";
  }
}

function renderSummary(snapshot: WorkbookContextSnapshot): void {
  const summary = document.getElementById("summary");
  if (!summary) return;
  summary.textContent = `Sheets: ${snapshot.totals.sheets} | Blocks: ${snapshot.totals.blocks} | Tables: ${snapshot.totals.tables} | Charts: ${snapshot.totals.charts} | Duration: ${snapshot.totals.durationMs} ms | selectionInBlockId: ${snapshot.active.selectionInBlockId ?? "null"} | nearestBlockId: ${snapshot.active.nearestBlockId ?? "null"}`;
}

function renderJson(snapshot: WorkbookContextSnapshot): void {
  const jsonNode = document.getElementById("context-json");
  if (!jsonNode) return;
  jsonNode.textContent = JSON.stringify(snapshot, null, 2);
}

function setStatus(message: string, tone: "info" | "success" | "error") {
  const status = document.getElementById("status");
  if (!status) return;
  status.textContent = message;
  status.setAttribute("data-tone", tone);
}

function setPlannerStatus(message: string, tone: "info" | "success" | "error") {
  const status = document.getElementById("planner-status");
  if (!status) return;
  status.textContent = message;
  status.setAttribute("data-tone", tone);
}

async function copyJson(): Promise<void> {
  const jsonNode = document.getElementById("context-json");
  if (!jsonNode) return;
  try {
    await navigator.clipboard.writeText(jsonNode.textContent || "");
    setStatus("JSON copied to clipboard", "success");
  } catch (err) {
    console.warn("Clipboard copy failed", err);
    setStatus("Clipboard not available in this environment", "error");
  }
}

function createUiLogger(verbose: boolean): ContextLogger {
  return {
    info: (step, message, sheet) => logLine("info", step, message, sheet),
    warn: (step, message, sheet, err) => {
      if (verbose) logLine("warn", step, message, sheet, err);
    },
    error: (step, message, sheet, err) => logLine("error", step, message, sheet, err),
  };
}

function logLine(level: "info" | "warn" | "error", step: string, message: string, sheet?: string, err?: any) {
  const container = document.getElementById("log-lines");
  if (!container) return;
  const ts = new Date().toISOString().split("T")[1].replace("Z", "");
  const textParts = [`[${ts}]`, step, sheet ? `(${sheet})` : "", "-", message];
  if (err) {
    textParts.push("-", err?.name || "", err?.message || "");
  }
  const line = document.createElement("div");
  line.className = `log-line ${level}`;
  line.textContent = textParts.filter(Boolean).join(" ");
  if (err?.stack && level !== "info") {
    line.title = err.stack;
  }
  container.appendChild(line);
  container.scrollTop = container.scrollHeight;
}

function clearLogs() {
  const container = document.getElementById("log-lines");
  if (container) container.innerHTML = "";
}

async function copyNodeText(node: HTMLElement | null) {
  if (!node) return;
  const text = (node as HTMLTextAreaElement).value !== undefined ? (node as HTMLTextAreaElement).value : node.textContent || "";
  try {
    await navigator.clipboard.writeText(text);
    setStatus("Copied", "success");
  } catch (err) {
    setStatus("Clipboard not available", "error");
  }
}

// ---------- Agent mode ----------

function setAgentStatus(message: string, tone: "info" | "success" | "error" | "warn") {
  const node = document.getElementById("agent-status");
  if (!node) return;
  node.textContent = message;
  node.setAttribute("data-tone", tone === "warn" ? "error" : tone);
}

function clearAgentLogs() {
  const container = document.getElementById("agent-log-lines");
  if (container) container.innerHTML = "";
}

function pushAgentLog(level: "info" | "warn" | "error", message: string) {
  const container = document.getElementById("agent-log-lines");
  if (!container) return;
  const ts = new Date().toISOString().split("T")[1].replace("Z", "");
  const line = document.createElement("div");
  line.className = `log-line ${level}`;
  line.textContent = `[${ts}] ${message}`;
  container.appendChild(line);
  container.scrollTop = container.scrollHeight;
}

function renderAgentLogs(result: ExecutionResult) {
  const container = document.getElementById("agent-log-lines");
  if (!container) return;
  container.innerHTML = "";
  result.logs.forEach((l) => {
    const line = document.createElement("div");
    line.className = `log-line ${l.level}`;
    line.textContent = `[${l.ts}] ${l.stepId ?? "-"} ${l.macro ? `[${l.macro}]` : ""} - ${l.message}`;
    container.appendChild(line);
  });
  container.scrollTop = container.scrollHeight;
}

function renderConfirmations(conf: NonNullable<ExecutionResult["confirmationsRequested"]>) {
  const host = document.getElementById("agent-confirmations");
  if (!host) return;
  host.innerHTML = "";
  if (!conf || conf.length === 0) return;
  conf.forEach((c) => {
    const block = document.createElement("div");
    block.className = "confirmation-block";
    const q = document.createElement("div");
    q.textContent = c.question;
    const actions = document.createElement("div");
    actions.className = "confirmation-actions";
    c.choices.forEach((choice) => {
      const btn = document.createElement("button");
      btn.className = "ms-Button ms-Button--small";
      btn.textContent = choice.label;
      btn.onclick = () => {
        pendingDecisions[c.id] = choice.id;
        confirmationAttempts += 1;
        runPlan();
      };
      actions.appendChild(btn);
    });
    block.appendChild(q);
    block.appendChild(actions);
    host.appendChild(block);
  });
}

function renderArtifacts(list: ExecutionResult["artifacts"]) {
  const host = document.getElementById("agent-artifacts");
  if (!host) return;
  host.innerHTML = "";
  if (!list || list.length === 0) return;
  const title = document.createElement("div");
  title.textContent = "Artifacts";
  const ul = document.createElement("ul");
  ul.className = "artifact-list";
  list.forEach((a) => {
    const li = document.createElement("li");
    li.textContent = `${a.type} @ ${a.sheet}!${a.anchor}`;
    ul.appendChild(li);
  });
  host.appendChild(title);
  host.appendChild(ul);
}

// ---------- LLM Planner ----------

async function writePlan(executeAfter: boolean): Promise<void> {
  lastPlanWasRepaired = false;
  lastRepairNotes = [];
  renderRepairerNotes([]);
  const promptNode = document.getElementById("llmPrompt") as HTMLTextAreaElement | null;
  const planOutput = document.getElementById("llmPlanOutput");
  const rawOutput = document.getElementById("llmRawOutput");
  const errNode = document.getElementById("llmErrors");
  const agentPlanNode = document.getElementById("plan-input") as HTMLTextAreaElement | null;
  if (!promptNode || !planOutput || !rawOutput || !errNode) return;
  const snapshot = await refreshContextWithResult();
  if (!snapshot) {
    setPlannerStatus("Refresh context failed", "error");
    errNode.textContent = "Impossible de rafraichir le contexte. Réessayez.";
    errNode.setAttribute("data-tone", "error");
    return;
  }
  const userPrompt = promptNode.value.trim();
  if (!userPrompt) {
    setPlannerStatus("Prompt vide", "error");
    errNode.textContent = "Ajoutez une demande utilisateur.";
    errNode.setAttribute("data-tone", "error");
    return;
  }

  setPlannerStatus("Appel LLM...", "info");
  const plannerTimeout = runtimeConfig.ollamaTimeoutMs;
  const plannerModel = getSelectedModel();
  pushAgentLog("info", `planner: model=${plannerModel} timeoutMs=${plannerTimeout}`);
  errNode.textContent = "";
  errNode.setAttribute("data-tone", "info");
  planOutput.textContent = "";
  rawOutput.textContent = "";

  const plannerLogger = (event: string, data?: any) => {
    const msg = data ? `${event} ${JSON.stringify(data)}` : event;
    pushAgentLog(event === "plan_invalid" || event === "planner_retry_failed" ? "warn" : "info", msg);
  };
  const res = await planWithOllama({ context: snapshot, userPrompt, model: plannerModel, timeoutMs: plannerTimeout, logger: plannerLogger });
  if (res.status === "ok") {
    setPlannerStatus("Réparation des formules...", "info");
    const { repairedPlan, notes } = await repairPlanWriteFormulas(res.plan as AgentPlan, snapshot, userPrompt);
    lastPlanWasRepaired = true;
    lastRepairNotes = notes;
    const pretty = JSON.stringify(repairedPlan, null, 2);
    planOutput.textContent = pretty;
    rawOutput.textContent = res.rawText || pretty;
    if (agentPlanNode) agentPlanNode.value = pretty;
    renderRepairerNotes(notes);
    setPlannerStatus("Plan généré + réparé", "success");
    if (executeAfter) {
      await runPlan({ skipRepair: true });
    }
  } else if (res.status === "invalid_plan") {
    lastPlanWasRepaired = false;
    setPlannerStatus("Plan invalide", "error");
    rawOutput.textContent = res.rawText || "";
    errNode.textContent = `Erreurs: ${(res.errors || []).join(" | ")}`;
    errNode.setAttribute("data-tone", "error");
  } else {
    lastPlanWasRepaired = false;
    setPlannerStatus("Erreur LLM", "error");
    errNode.textContent = res.error;
    errNode.setAttribute("data-tone", "error");
  }
}

async function runPlan(opts: { skipRepair?: boolean } = {}): Promise<void> {
  if (agentRunning) return;
  if (!opts.skipRepair) {
    lastPlanWasRepaired = false;
  }
  const textarea = document.getElementById("plan-input") as HTMLTextAreaElement | null;
  if (!textarea) return;
  let plan: AgentPlan;
  try {
    plan = canonicalizePlan(JSON.parse(textarea.value || "{}"));
  } catch (err: any) {
    setAgentStatus(`Plan invalide: ${err?.message || err}`, "error");
    return;
  }

  const validation = validatePlan(plan);
  if (!validation.valid) {
    setAgentStatus("Plan invalide (schema)", "error");
    renderAgentLogs({ logs: [], artifacts: [], status: "error" });
    return;
  }

  const invariant = validatePlanInvariants(plan);
  if (!invariant.valid) {
    const errMsg = (invariant.issues || []).map((i) => i.message).join(" | ") || "Plan invalide (invariants)";
    setAgentStatus(errMsg, "error");
    renderAgentLogs({
      logs: (invariant.issues || []).map((i) => ({ ts: new Date().toISOString(), level: "error", message: i.message } as any)),
      artifacts: [],
      status: "error",
    });
    return;
  }

  agentRunning = true;
  setAgentStatus("Execution en cours...", "info");
  try {
    if (!lastSnapshot) {
      lastSnapshot = await getWorkbookContext();
      renderSummary(lastSnapshot);
      renderJson(lastSnapshot);
    }
    plan = applyConfirmationsToPlan(plan, pendingDecisions, lastSnapshot);
    const result = await Excel.run(async (excelCtx) => {
      return executePlan(plan, lastSnapshot as WorkbookContextSnapshot, excelCtx, {
        confirmationDecisions: pendingDecisions,
        attempt: confirmationAttempts + 1,
      });
    });

    renderAgentLogs(result);
    renderArtifacts(result.artifacts);

    if (result.status === "need_user_confirmation" && result.confirmationsRequested?.length) {
      renderConfirmations(result.confirmationsRequested);
      setAgentStatus("Confirmation requise", "warn");
      return;
    } else {
      confirmationAttempts = 0;
      pendingDecisions = {};
      renderConfirmations([]);
    }

    if (result.status === "ok") setAgentStatus("Plan execute", "success");
    else if (result.status === "error") setAgentStatus("Erreur pendant le plan", "error");
    else setAgentStatus(`Statut: ${result.status}`, "info");
  } catch (err: any) {
    console.error(err);
    setAgentStatus(`Echec: ${err?.message || err}`, "error");
  } finally {
    agentRunning = false;
  }
}

function renderRepairerNotes(notes: string[]) {
  const host = document.getElementById("repairer-notes");
  if (!host) return;
  host.innerHTML = "";
  if (!notes.length) {
    host.textContent = "Aucune modification par le réparateur";
    return;
  }
  const ul = document.createElement("ul");
  notes.forEach((n) => {
    const li = document.createElement("li");
    li.textContent = n;
    ul.appendChild(li);
  });
  host.appendChild(ul);
}
