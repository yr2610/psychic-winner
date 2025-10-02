const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");

const sourcePath = path.resolve(__dirname, "../src/TXT2JSON.js");
const source = fs.readFileSync(sourcePath, "utf8");

function extractFunction(name) {
  const marker = "function " + name + "(";
  const start = source.indexOf(marker);
  if (start === -1) {
    throw new Error("Could not find function " + name);
  }
  const braceIndex = source.indexOf("{", start);
  if (braceIndex === -1) {
    throw new Error("Could not find opening brace for function " + name);
  }

  let depth = 0;
  for (let i = braceIndex; i < source.length; i++) {
    const char = source.charAt(i);
    if (char === "{") {
      depth++;
    } else if (char === "}") {
      depth--;
      if (depth === 0) {
        return source.slice(start, i + 1);
      }
    }
  }

  throw new Error("Failed to extract function body for " + name);
}

const context = {
  previousPlaceholderWarningsBySheet: {},
  shouldReuseParsedSheetNode: null
};

vm.createContext(context);
vm.runInContext(
  "shouldReuseParsedSheetNode = " + extractFunction("shouldReuseParsedSheetNode"),
  context
);

assert.strictEqual(
  context.shouldReuseParsedSheetNode({ id: "node" }, "Sheet A"),
  true,
  "Sheets without cached warnings should be eligible for reuse"
);

context.previousPlaceholderWarningsBySheet = {
  "Sheet A": [
    { kind: "undefinedPlaceholder", message: "stale" }
  ]
};

assert.strictEqual(
  context.shouldReuseParsedSheetNode({ id: "node" }, "Sheet A"),
  false,
  "Sheets with cached placeholder warnings should be forced to reprocess"
);

context.previousPlaceholderWarningsBySheet = {
  "Sheet A": []
};

assert.strictEqual(
  context.shouldReuseParsedSheetNode({ id: "node" }, "Sheet A"),
  true,
  "Empty cached warning lists should not block reuse"
);

assert.strictEqual(
  context.shouldReuseParsedSheetNode(null, "Sheet A"),
  false,
  "Missing parsed sheet nodes should not be reused"
);

console.log("placeholderWarningsCacheReuse tests passed.");
