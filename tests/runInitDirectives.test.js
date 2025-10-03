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
vm.runInContext("extractOwnScopeLayer = " + extractFunction("extractOwnScopeLayer"), context);
vm.runInContext("extendScope = " + extractFunction("extendScope"), context);
vm.runInContext("getInheritedScopeLayer = " + extractFunction("getInheritedScopeLayer"), context);
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
  return clone;
}

const scope = {};
const firstRun = expandDeep(scope);
assert.strictEqual(
  Object.prototype.hasOwnProperty.call(scope, "runs"),
  false,
  "Brand-new keys assigned in @init should remain scoped to the branch"
);
assert.strictEqual(
  firstRun.children[0]._initScopeLayer.runs,
  1,
  "First expansion should record run count within the branch overlay"
);

const secondRun = expandDeep(scope);
assert.strictEqual(
  secondRun.children[0]._initScopeLayer.runs,
  1,
  "Repeated expansions should still execute @init without leaking counters"
);

const seededScope = { runs: 0 };
expandDeep(seededScope);
assert.strictEqual(
  seededScope.runs,
  1,
  "Pre-existing keys should continue to receive propagated updates"
);

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

// Ensure child.params defined on @init can share values with subsequent directives.
const sharedTemplate = {
  kind: kindUL,
  text: "&Shared()",
  children: [],
  templates: {}
};

const sharedParent = makeNode("- holder", sharedTemplate);
const initWithParams = makeNode("@init: $set('shared', ($get('shared') || 0) + extra)", sharedParent);
initWithParams.params = { extra: 1 };
const initFollower = makeNode("@init: $set('shared', ($get('shared') || 0) + 1)", sharedParent);
sharedParent.children.push(initWithParams, initFollower);
sharedTemplate.children.push(sharedParent);

const sharedScope = {};
context.runInitDirectives(sharedTemplate, sharedScope);
assert.strictEqual(
  Object.prototype.hasOwnProperty.call(sharedScope, "shared"),
  false,
  "Branch-local accumulators should not leak onto the root scope"
);
assert.strictEqual(
  sharedParent._initScopeLayer.shared,
  2,
  "@init with params should propagate scope changes to later siblings via the branch overlay"
);

// Ensure sibling branches do not leak conflicting params into each other.
const siblingTemplate = {
  kind: kindUL,
  text: "&SiblingIsolation()",
  children: [],
  templates: {}
};

function makeBranch(name, flavor, parent) {
  const branch = makeNode(name, parent);
  branch.params = { flavor: flavor };
  const init = makeNode("@init: $set('dessert', flavor); outputs.push($get('dessert'))", branch);
  branch.children.push(init);
  return branch;
}

const siblingParent = makeNode("- siblings", siblingTemplate);
siblingTemplate.children.push(siblingParent);

const firstBranch = makeBranch("- first", "strawberry", siblingParent);
const secondBranch = makeBranch("- second", "vanilla", siblingParent);
siblingParent.children.push(firstBranch, secondBranch);

const siblingScope = { outputs: [] };
context.runInitDirectives(siblingTemplate, siblingScope);
assert.deepStrictEqual(
  siblingScope.outputs,
  ["strawberry", "vanilla"],
  "Each sibling @init should observe its own params without interference"
);

assert.strictEqual(
  Object.prototype.hasOwnProperty.call(siblingScope, "dessert"),
  false,
  "Root scope should not be polluted by branch-local @init variables"
);

assert.strictEqual(
  firstBranch._initScopeLayer.dessert,
  "strawberry",
  "First branch should capture its own dessert override"
);
assert.strictEqual(
  secondBranch._initScopeLayer.dessert,
  "vanilla",
  "Second branch should capture its own dessert override"
);

const templateLayer = context.getInheritedScopeLayer(siblingTemplate) || {};
assert.strictEqual(
  Object.prototype.hasOwnProperty.call(templateLayer, "dessert"),
  false,
  "Template-level overlays should not inherit branch-local keys"
);

const evaluationRootScope = context.extendScope(
  siblingScope,
  templateLayer
);
assert.strictEqual(
  Object.prototype.hasOwnProperty.call(evaluationRootScope, "dessert"),
  false,
  "Evaluated parent scope should remain free of branch-local keys"
);

const firstBranchScope = context.extendScope(
  evaluationRootScope,
  firstBranch._initScopeLayer || {}
);
const secondBranchScope = context.extendScope(
  evaluationRootScope,
  secondBranch._initScopeLayer || {}
);

assert.strictEqual(
  firstBranchScope.dessert,
  "strawberry",
  "Branch scopes should prefer their own @init values when rendering"
);
assert.strictEqual(
  secondBranchScope.dessert,
  "vanilla",
  "Sibling scopes should not leak @init values between branches"
);

console.log("runInitDirectives nested template test passed.");
