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

const context = {
  _: require("../src/lib/lodash.js"),
  templateError: function(message) {
    throw new Error(message);
  },
  expandInlineParamArray: function() {
    throw new Error("Inline array expansion should not be used in this test");
  },
  shrinkChildrenArray: function() {},
  globalScope: {}
};

vm.createContext(context);

vm.runInContext(extractDeclaration(/var kindUL\s*=\s*"UL";/), context);
vm.runInContext(extractDeclaration(/var templateParamFnCache\s*=\s*Object\.create\(null\);/), context);
vm.runInContext(extractFunction("attachArgAliases"), context);
vm.runInContext(extractFunction("extendScope"), context);
vm.runInContext("getInheritedScopeLayer = " + extractFunction("getInheritedScopeLayer"), context);
vm.runInContext("evalTemplateParameters = " + extractFunction("evalTemplateParameters"), context);
vm.runInContext(extractFunction("forAllNodes_Recurse"), context);
vm.runInContext("expandAllTemplateCalls = " + extractFunction("expandAllTemplateCalls"), context);

const root = {
  kind: context.kindUL,
  text: "- root",
  children: []
};

const callNode = {
  kind: context.kindUL,
  text: "*Dummy(count)",
  children: [],
  parent: root,
  _initScopeLayer: { count: 3 }
};

root.children.push(callNode);
context.root = root;

let captured = null;
context.addTemplate = function(node, index, templateName, parameters, localScope) {
  captured = { node, index, templateName, parameters, localScope };
};

context.expandAllTemplateCalls();

assert.ok(captured, "Template call should invoke addTemplate");
assert.strictEqual(captured.templateName, "Dummy", "Template name should match the invocation");
assert.strictEqual(captured.parameters, 3, "Parameters should resolve using the inherited init scope");
assert.strictEqual(captured.localScope.count, 3, "Local scope passed to addTemplate should expose the init value");

console.log("Template call scope propagation test passed.");
