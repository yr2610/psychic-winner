const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");

const sourcePath = path.resolve(__dirname, "../src/TXT2JSON.js");
const source = fs.readFileSync(sourcePath, "utf8");

function extractDeclaration(pattern) {
  const match = pattern.exec(source);
  if (!match) {
    throw new Error("Failed to extract declaration for pattern: " + pattern);
  }
  return match[0];
}

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

function TemplateError(message, node) {
  this.errorMessage = message;
  this.node = node;
}

const context = {
  _: require("../src/lib/lodash.js"),
  TemplateError,
  templateError(message, node) {
    throw new TemplateError(message, node);
  },
  toRepeatList() { return null; },
  runAnchorDeclarations() {},
  runInitDirectives() {},
  replacePlaceholdersInNode() { return true; },
  evalTemplateParameters() {
    throw new Error("Nested template expansion is not expected in this test");
  },
  MyError(message) {
    throw new Error(message);
  },
  globalScope: {}
};

vm.createContext(context);

vm.runInContext(extractDeclaration(/var kindH\s*=\s*"H";/), context);
vm.runInContext(extractDeclaration(/var kindUL\s*=\s*"UL";/), context);
vm.runInContext(extractFunction("extendScope"), context);
vm.runInContext("extractOwnScopeLayer = " + extractFunction("extractOwnScopeLayer"), context);
vm.runInContext("getInheritedScopeLayer = " + extractFunction("getInheritedScopeLayer"), context);
vm.runInContext(extractFunction("attachArgAliases"), context);
vm.runInContext(extractFunction("expandInlineParamArray"), context);
vm.runInContext(extractFunction("forAllNodes_Recurse"), context);
vm.runInContext(extractFunction("cloneTemplateTree"), context);
vm.runInContext(extractFunction("shrinkChildrenArray"), context);
vm.runInContext(extractFunction("addTemplate"), context);

const templateRoot = {
  text: "&Dummy()",
  kind: context.kindUL,
  params: {
    arr: [
      { $value: "first", $id: "first" },
      { $value: "second", $id: "second" }
    ]
  },
  children: []
};

const inlineArrayNode = {
  text: "*arr",
  kind: context.kindUL,
  id: "node",
  children: [],
  parent: templateRoot
};

templateRoot.children.push(inlineArrayNode);

context.findTemplate_Recurse = function(name) {
  if (name === "Dummy") {
    return templateRoot;
  }
  return null;
};

const parentNode = { children: [] };
const targetNode = {
  text: "*Dummy()",
  id: "call",
  parent: parentNode,
  children: []
};
parentNode.children.push(targetNode);

context.addTemplate(targetNode, 0, "Dummy", {}, {});

const expandedChildren = parentNode.children.slice(1);

assert.strictEqual(parentNode.children[0], null, "Original call site should be replaced with null");
assert.strictEqual(expandedChildren.length, 2, "Two entries should be produced from the inline array");
assert.deepStrictEqual(
  expandedChildren.map(child => child.text),
  ["first", "second"],
  "Expanded nodes should contain the inline array values"
);
assert.deepStrictEqual(
  expandedChildren.map(child => child.id),
  ["call_node_first", "call_node_second"],
  "Expanded nodes should inherit IDs with the inline suffix"
);
assert.ok(expandedChildren.every(child => child.parent === parentNode), "Expanded nodes should belong to the call parent");

console.log("Inline arrays inside templates expand using template parameters.");
