#!/usr/bin/env node

const commander = require("commander");
const pkg = require("./package.json");
const { xlsxToJson } = require("./xlsx-to-json");
const { jsonToXlsx } = require("./json-to-xlsx");

commander.program
  .version(pkg.version)
  .option("-p, --path [path]", "path of xlsx file")
  .option("-k, --keyCode [keyCode]", "keyCode")
  .option("-v, --valueCode [valueCode]", "valueCode")
  .option("-t, --type [type]", "output file")
  .parse(process.argv);

const path = commander.program.getOptionValue("path");
const keyCode = commander.program.getOptionValue("keyCode") || "code";
const valueCode = commander.program.getOptionValue("valueCode") || "value";
const type = commander.program.getOptionValue("type") || "xtj";

if (!path) {
  console.error("path is required");
  process.exit(1);
}

if (type === "xtj") {
  xlsxToJson({
    filePath: path,
    keyCode,
    valueCode,
  });
}

if (type === "jtx") {
  jsonToXlsx({
    dirPath: path,
    keyCode,
    valueCode,
  });
}
