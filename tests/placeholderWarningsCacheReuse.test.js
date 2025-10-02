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

function extractSection(startMarker, endMarker) {
  const start = source.indexOf(startMarker);
  if (start === -1) {
    throw new Error("Could not find start marker: " + startMarker);
  }
  const end = source.indexOf(endMarker, start);
  if (end === -1) {
    throw new Error("Could not find end marker: " + endMarker);
  }
  return source.slice(start, end);
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

const placeholderWarningsSection = extractSection(
  "var PLACEHOLDER_WARN_ON_UNDEFINED",
  "function mergeCachedPlaceholderWarnings"
);

const recordedPaths = {};

const placeholderWarningsContext = {
  outfilePath: path.join("C:", "project", "output", "result.json"),
  fso: {
    GetParentFolderName(filePath) {
      recordedPaths.parent = filePath;
      return path.dirname(filePath);
    },
    BuildPath(directory, fileName) {
      recordedPaths.build = { directory, fileName };
      return path.join(directory, fileName);
    },
    FileExists(filePath) {
      recordedPaths.fileExists = filePath;
      return false;
    }
  },
  CL: {
    readTextFileUTF8() {
      throw new Error("Cache file should not be read when it does not exist");
    }
  },
  _: {
    assign: Object.assign,
    some: (array, predicate) => array.some(predicate),
    isArray: Array.isArray,
    map: (array, iteratee) => array.map(iteratee),
    pick: (object, keys) => {
      const result = {};
      for (const key of keys) {
        if (Object.prototype.hasOwnProperty.call(object, key)) {
          result[key] = object[key];
        }
      }
      return result;
    },
    isString: value => typeof value === "string"
  },
  JSON
};

vm.createContext(placeholderWarningsContext);
vm.runInContext(placeholderWarningsSection, placeholderWarningsContext);

assert.deepStrictEqual(
  recordedPaths,
  {
    parent: placeholderWarningsContext.outfilePath,
    build: {
      directory: path.dirname(placeholderWarningsContext.outfilePath),
      fileName: "placeholder_warnings.json"
    },
    fileExists: path.join(
      path.dirname(placeholderWarningsContext.outfilePath),
      "placeholder_warnings.json"
    )
  },
  "Placeholder warning cache should be resolved relative to the outfile using the expected filename"
);

assert.ok(
  Object.prototype.hasOwnProperty.call(
    placeholderWarningsContext,
    "previousPlaceholderWarningsBySheet"
  ),
  "Placeholder warning cache map should be defined on the global context"
);

assert.deepStrictEqual(
  Object.keys(placeholderWarningsContext.previousPlaceholderWarningsBySheet),
  [],
  "Missing cache files should result in an empty placeholder warning map"
);

console.log("placeholderWarningsCacheReuse tests passed.");
