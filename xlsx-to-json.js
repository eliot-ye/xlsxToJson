// @ts-check

const fs = require("fs");
const XLSX = require("xlsx");

/**
 *
 * @typedef {Object} Option
 * @property {string} filePath
 * @property {string} [keyCode]
 * @property {string} [valueCode]
 */

/**
 *
 * @param {Option} option
 * @returns
 */
exports.xlsxToJson = (option) => {
  const keyCode = option.keyCode || "Code";
  const valueCode = option.valueCode || "Value";

  if (!option.filePath) {
    return;
  }

  const filePathList = option.filePath.split("/");
  const fileFullName = filePathList.pop();

  if (!fileFullName) {
    return;
  }

  const fileFullNameList = fileFullName.split(".");
  fileFullNameList.pop();
  const fileName = fileFullNameList.join(".");

  const dirPath = `${filePathList.join("/")}/${fileName}`;
  try {
    fs.readdirSync(`${dirPath}`);
    fs.rmSync(`${dirPath}`, { recursive: true });
    fs.mkdirSync(`${dirPath}`);
  } catch (error) {
    fs.mkdirSync(`${dirPath}`);
  }

  // 读取Excel文件
  const workbook = XLSX.readFile(option.filePath);
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
      result[row[keyCode].replace(/"/g, "").replace("    ", "")] = row[
        valueCode
      ]
        .replace('",', "")
        .replace(/^(?:")/, "");
    });

    const resultStr = JSON.stringify(result);

    fs.writeFile(`${dirPath}/${sheetName}.json`, resultStr, {}, () => {});
  });
};
