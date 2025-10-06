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
  _: require("../src/lib/lodash.js")
};

vm.createContext(context);
vm.runInContext(extractFunction("getInheritedScopeLayer"), context);

const paramsLayer = { arr: ["e1", "e2"], foo: "bar" };
const initLayer = { counter: 10 };

const combined = context.getInheritedScopeLayer({
  params: paramsLayer,
  _initScopeLayer: initLayer
});

assert.ok(combined);
assert.notStrictEqual(combined, paramsLayer, "Combined layer should not reuse params reference");
assert.deepStrictEqual(combined.arr, paramsLayer.arr, "Params values should be preserved");
assert.strictEqual(combined.foo, "bar", "Params keys should be kept");
assert.strictEqual(combined.counter, 10, "Init layer keys should be merged");

const onlyInit = context.getInheritedScopeLayer({ _initScopeLayer: initLayer });
assert.strictEqual(onlyInit, initLayer, "Init layer alone should be returned directly");

const onlyParams = context.getInheritedScopeLayer({ params: paramsLayer });
assert.strictEqual(onlyParams, paramsLayer, "Params alone should be returned directly");

const none = context.getInheritedScopeLayer(null);
assert.strictEqual(none, null, "Null node should yield null layer");

console.log("getInheritedScopeLayer merges params with init scope.");
