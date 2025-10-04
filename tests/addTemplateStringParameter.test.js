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

function TemplateError(message, node) {
  this.errorMessage = message;
  this.node = node;
}

const context = {
  _: require("../src/lib/lodash.js"),
  TemplateError,
  toRepeatList: function() { return null; },
  findTemplate_Recurse: function() {
    return { id: "template", parent: null, children: [] };
  },
  cloneTemplateTree: function(node) { return node; },
  extendScope: function() {
    throw new Error("extendScope should not be called for invalid parameters");
  },
  attachArgAliases: function() {
    throw new Error("attachArgAliases should not be called for invalid parameters");
  },
  runAnchorDeclarations: function() {
    throw new Error("runAnchorDeclarations should not be called for invalid parameters");
  },
  runInitDirectives: function() {
    throw new Error("runInitDirectives should not be called for invalid parameters");
  }
};

vm.createContext(context);
vm.runInContext(extractFunction("addTemplate"), context);

const targetNode = {
  id: "call",
  parent: { children: [] }
};

assert.throws(
  function() {
    context.addTemplate(targetNode, 0, "SampleTemplate", "invalid-parameter", {});
  },
  function(error) {
    assert.ok(error instanceof TemplateError, "Error should be an instance of TemplateError");
    assert.strictEqual(
      error.errorMessage,
      "テンプレート'SampleTemplate'では文字列引数は使用できません。"
    );
    assert.strictEqual(error.node, targetNode);
    return true;
  }
);

console.log("addTemplate rejects string parameters.");
