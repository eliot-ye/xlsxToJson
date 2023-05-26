// @ts-check

const fs = require("fs");
const XLSX = require("xlsx");

const filePath = "/Users/eliot/Downloads/";
const fileName = "language_mapping_to_mtel_v0.2";
const dirPath = `${filePath}${fileName}`;

const keyCode = "Code";
const valueCode = "final";

// 读取Excel文件
const workbook = XLSX.readFile(`${dirPath}.xlsx`);

try {
  fs.readdirSync(`${dirPath}`);
  fs.rmSync(`${dirPath}`, { recursive: true });
  fs.mkdirSync(`${dirPath}`);
} catch (error) {
  fs.mkdirSync(`${dirPath}`);
}

workbook.SheetNames.forEach((sheetName) => {
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(worksheet);

  if (data.length == 0) {
    return;
  }
  if (!data[0][keyCode]) {
    return;
  }

  const result = {};
  data.forEach((row) => {
    result[row[keyCode].replace(/"/g, "").replace("    ", "")] = row[valueCode]
      .replace('",', "")
      .replace(/^(?:")/, "");
  });

  const resultStr = JSON.stringify(result);

  fs.writeFile(`${dirPath}/${sheetName}.json`, resultStr, {}, () => {});
});
