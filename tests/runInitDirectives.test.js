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
  execInScope: function(code, scope) {
    /* eslint no-with: 0 */
    with (scope) { eval(code); }
  }
};

vm.createContext(context);

vm.runInContext(extractDeclaration(/var kindUL\s*=\s*"UL";/), context);
vm.runInContext(extractFunction("installInitHelpers"), context);
vm.runInContext(extractFunction("cloneTemplateTree"), context);
vm.runInContext("extendScope = " + extractFunction("extendScope"), context);
vm.runInContext("runInitDirectives = " + extractFunction("runInitDirectives"), context);

const kindUL = context.kindUL;

function makeNode(text, parent) {
  return {
    kind: kindUL,
    text: text,
    children: [],
    parent: parent
  };
}

// Deep template with @init nested inside an extra node to exercise recursion.
const deepTemplate = {
  kind: kindUL,
  text: "&Deep()",
  children: [],
  templates: {}
};
const wrapperNode = makeNode("- wrapper", deepTemplate);
const initNode = makeNode("@init: $set('runs', ($get('runs') || 0) + 1)", wrapperNode);
wrapperNode.children.push(initNode);
deepTemplate.children.push(wrapperNode);

// Register deep template inside another template to emulate nested definitions.
const innerTemplate = {
  kind: kindUL,
  text: "&Inner()",
  children: [],
  templates: { Deep: deepTemplate }
};
deepTemplate.parent = innerTemplate;

function expandDeep(scope) {
  const clone = context.cloneTemplateTree(deepTemplate);
  context.runInitDirectives(clone, scope);
}

const scope = {};
expandDeep(scope);
assert.strictEqual(scope.runs, 1, "First expansion should run @init once");
expandDeep(scope);
assert.strictEqual(scope.runs, 2, "Second expansion should run @init again");

// Ensure params attached to a node are visible to nested @init directives.
const paramAwareTemplate = {
  kind: kindUL,
  text: "&ParamAware()",
  children: [],
  templates: {}
};

const paramParent = makeNode("- holder", paramAwareTemplate);
paramParent.params = { foo: "bar" };

const paramInit = makeNode("@init: spy(foo)", paramParent);
paramParent.children.push(paramInit);
paramAwareTemplate.children.push(paramParent);

const paramScope = {};
paramScope.spy = function(value) {
  paramScope.captured = value;
};
context.runInitDirectives(paramAwareTemplate, paramScope);
assert.strictEqual(paramScope.captured, "bar", "@init should access params inherited via scope chaining");

console.log("runInitDirectives nested template test passed.");
