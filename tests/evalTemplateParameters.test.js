const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");

const sourcePath = path.resolve(__dirname, "../src/TXT2JSON.js");
const source = fs.readFileSync(sourcePath, "utf8");

function extractDeclaration(pattern) {
  var match = pattern.exec(source);
  if (!match) {
    throw new Error("Failed to extract declaration for pattern: " + pattern);
  }
  return match[0];
}

function extractFunction(name) {
  var marker = "function " + name;
  var start = source.indexOf(marker);
  if (start === -1) {
    throw new Error("Could not find function " + name);
  }
  var braceIndex = source.indexOf("{", start);
  if (braceIndex === -1) {
    throw new Error("Could not find opening brace for function " + name);
  }

  var depth = 0;
  for (var i = braceIndex; i < source.length; i++) {
    var char = source.charAt(i);
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

var context = {
  _: require("../src/lib/lodash.js"),
  globalScope: {}
};

vm.createContext(context);

var cacheDeclaration = extractDeclaration(/var templateParamFnCache\s*=\s*Object\.create\(null\);/);
vm.runInContext(cacheDeclaration, context);

var attachArgAliasesSrc = extractFunction("attachArgAliases");
vm.runInContext(attachArgAliasesSrc, context);

var evalTemplateParametersSrc = extractFunction("evalTemplateParameters");
vm.runInContext(evalTemplateParametersSrc, context);

function resetCache() {
  vm.runInContext("templateParamFnCache = Object.create(null);", context);
}

function callEval(paramsStr, currentParameters) {
  var node = { parent: undefined };
  return context.evalTemplateParameters(paramsStr, node, currentParameters);
}

function getCacheEntries() {
  return Object.keys(context.templateParamFnCache);
}

function getCachedFunction(key) {
  return context.templateParamFnCache[key];
}

// --- Tests ---

resetCache();
var result1 = callEval("foo", { foo: { value: 1 }, bar: { value: 2 } });
assert.strictEqual(result1.value, 1);
assert.deepStrictEqual(Object.keys(result1), ["value"]);

var keysAfterFirst = getCacheEntries();
assert.strictEqual(keysAfterFirst.length, 1);
var cachedFn = getCachedFunction(keysAfterFirst[0]);

var result2 = callEval("foo", { foo: { value: 10 }, bar: { value: 20 } });
assert.strictEqual(result2.value, 10);
assert.deepStrictEqual(Object.keys(result2), ["value"]);
assert.strictEqual(getCacheEntries().length, 1);
assert.strictEqual(getCachedFunction(keysAfterFirst[0]), cachedFn);

resetCache();
var aliasResult1 = callEval("$args", { alpha: { n: 1 }, beta: { n: 2 } });
assert.deepStrictEqual(Object.keys(aliasResult1), ["alpha", "beta"]);
assert.deepStrictEqual(aliasResult1.alpha, { n: 1 });
assert.deepStrictEqual(aliasResult1.beta, { n: 2 });

var aliasKeys = getCacheEntries();
assert.strictEqual(aliasKeys.length, 1);
var aliasCachedFn = getCachedFunction(aliasKeys[0]);

var aliasResult2 = callEval("$args", { alpha: { n: 3 }, beta: { n: 4 } });
assert.deepStrictEqual(Object.keys(aliasResult2), ["alpha", "beta"]);
assert.deepStrictEqual(aliasResult2.alpha, { n: 3 });
assert.deepStrictEqual(aliasResult2.beta, { n: 4 });
assert.strictEqual(getCacheEntries().length, 1);
assert.strictEqual(getCachedFunction(aliasKeys[0]), aliasCachedFn);

console.log("All evalTemplateParameters tests passed.");
