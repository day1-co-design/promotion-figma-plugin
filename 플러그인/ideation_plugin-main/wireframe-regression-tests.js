const fs = require("fs");
const vm = require("vm");

function createMockElement(id) {
  return {
    id,
    value: "",
    textContent: "",
    innerHTML: "",
    className: "",
    disabled: false,
    checked: false,
    files: [],
    style: {},
    classList: {
      add() {},
      remove() {},
      toggle() {}
    },
    addEventListener() {},
    appendChild() {},
    removeChild() {},
    select() {},
    querySelector() {
      return null;
    },
    querySelectorAll() {
      return [];
    }
  };
}

function loadUiScript() {
  const html = fs.readFileSync("ui.html", "utf8");
  const scripts = [...html.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/gi)].map((match) => match[1]);
  if (!scripts.length) throw new Error("ui.html 안에서 script를 찾지 못했습니다.");

  const elements = new Map();
  const document = {
    getElementById(id) {
      if (!elements.has(id)) elements.set(id, createMockElement(id));
      return elements.get(id);
    },
    createElement(tagName) {
      return createMockElement(tagName);
    },
    querySelector() {
      return null;
    },
    querySelectorAll() {
      return [];
    }
  };

  const window = {
    localStorage: {
      getItem() {
        return "";
      },
      setItem() {}
    }
  };
  const timers = [];

  const context = {
    console,
    document,
    window,
    parent: {
      postMessage() {}
    },
    setTimeout(callback) {
      if (typeof callback === "function") timers.push(callback);
      return timers.length;
    },
    clearTimeout() {},
    Uint8Array,
    ArrayBuffer,
    FileReader: function FileReader() {}
  };
  context.globalThis = context;
  window.window = window;
  window.document = document;
  window.parent = context.parent;

  vm.createContext(context);
  scripts.forEach((script) => vm.runInContext(script, context, { filename: "ui.html" }));
  timers.splice(0).forEach((callback) => callback());
  return context;
}

const context = loadUiScript();
const outputPolicy = JSON.parse(fs.readFileSync("src/output/wireframe-output-policy.json", "utf8"));
const uiPolicyVersion = vm.runInContext("WIRE_OUTPUT_POLICY.version", context);
if (uiPolicyVersion !== outputPolicy.version) {
  throw new Error(`ui.html output policy가 동기화되지 않았습니다. ui=${uiPolicyVersion} source=${outputPolicy.version}`);
}

const codeSource = fs.readFileSync("code.js", "utf8");
if (!codeSource.includes(`"version": "${outputPolicy.version}"`)) {
  throw new Error("code.js output policy가 src/output/wireframe-output-policy.json과 동기화되지 않았습니다.");
}

const runChecks = context.window.runWireframeRegressionChecks;
if (typeof runChecks !== "function") {
  throw new Error("window.runWireframeRegressionChecks 함수가 등록되지 않았습니다.");
}

const result = runChecks();
if (!result || !result.ok) {
  console.error("[wireframe-regression] FAIL", result && result.failures);
  process.exit(1);
}

console.log("[wireframe-regression] PASS");
