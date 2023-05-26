// @ts-check

const commander = require("commander");
const pkg = require("./package.json");
const { xlsxToJson } = require("./xlsx-to-json");

commander.program
  .version(pkg.version)
  .option("-p, --path [path]", "path of xlsx file")
  .parse(process.argv);

const path = commander.program.getOptionValue("path");
console.log(path);

xlsxToJson({
  filePath: path,
  keyCode: "Code",
  valueCode: "final",
});
