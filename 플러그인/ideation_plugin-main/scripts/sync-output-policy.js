const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const policyPath = path.join(root, "src", "output", "wireframe-output-policy.json");
const uiPath = path.join(root, "ui.html");
const codePath = path.join(root, "code.js");

function readPolicy() {
  const raw = fs.readFileSync(policyPath, "utf8");
  const policy = JSON.parse(raw);
  if (!policy || !policy.version) throw new Error("wireframe-output-policy.json에 version이 필요합니다.");
  if (!Number.isFinite(Number(policy.targetWidthPx)) || Number(policy.targetWidthPx) <= 0) {
    throw new Error("targetWidthPx는 0보다 큰 숫자여야 합니다.");
  }
  return policy;
}

function replaceBlock(filePath, declaration, policy) {
  const start = "// OUTPUT_POLICY_START";
  const end = "// OUTPUT_POLICY_END";
  const source = fs.readFileSync(filePath, "utf8");
  const startIndex = source.indexOf(start);
  const endIndex = source.indexOf(end);
  if (startIndex < 0 || endIndex < 0 || endIndex < startIndex) {
    throw new Error(`${path.basename(filePath)} 안에서 OUTPUT_POLICY marker를 찾지 못했습니다.`);
  }
  const linePrefix = declaration.startsWith("const") ? "    " : "";
  const json = JSON.stringify(policy, null, 2).replace(/\n/g, `\n${linePrefix}`);
  const block = `${start}\n${linePrefix}${declaration} ${json};\n${linePrefix}${end}`;
  const next = source.slice(0, startIndex) + block + source.slice(endIndex + end.length);
  fs.writeFileSync(filePath, next);
}

const policy = readPolicy();
replaceBlock(uiPath, "const WIRE_OUTPUT_POLICY =", policy);
replaceBlock(codePath, "var WIRE_OUTPUT_POLICY =", policy);

console.log("[sync-output-policy] synced", policy.version);
