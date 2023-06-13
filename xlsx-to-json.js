// @ts-check

const fs = require("fs");
const XLSX = require("xlsx");

/**
 * 将 Excel 文件 (.xlsx) 转换为一个或多个 JSON 文件。Excel 文件中的每个工作表将转换为一个同名的 JSON 文件，这个 JSON 文件将被创建在与源 Excel 文件相同的文件夹中。如果工作表中没有使用指定键代码的列，那么该工作表将被跳过。如果工作表具有使用指定键代码的列，但整个工作表为空，则该工作表也将被跳过。
 *
 * @param {XlsxToJsonOption} option - 包含以下属性的对象：
 *     - filePath：源 Excel 文件的路径。
 *     - keyCode：Excel 表中包含键的列的名称。
 *     - valueCode：Excel 表中包含值的列的名称。
 * @return {void} 该函数不返回任何值。
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
      const jsonKey = row[keyCode].replace(/"/g, "").replace("    ", "");
      let jsonValue = row[valueCode];

      try {
        const _jsonValue = jsonValue.replace(/",/g, '"').replace(/\\\\"/g, "\\\"")
        // console.log("jsonValue1", jsonKey, _jsonValue);
        jsonValue = JSON.parse(_jsonValue);
        // console.log("jsonValue2", jsonValue);
      } catch (error) {
        // console.log(error);
        jsonValue = jsonValue.replace('",', "").replace(/^(?:")/, "").replace(/\\\"/g, "\"");
      }

      result[jsonKey] = jsonValue;
    });

    const resultStr = JSON.stringify(result);

    fs.writeFile(`${dirPath}/${sheetName}.json`, resultStr, {}, () => {});
  });
};
