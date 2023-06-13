// @ts-check

const fs = require("fs");
const XLSX = require("xlsx");

/**
 *
 * @param {JsonToXlsxOption} option - 包含以下属性的对象：
 *     - dirPath: 源 Excel 文件的路径。
 *     - keyCode: Excel 表中包含键的列的名称。
 *     - valueCode: Excel 表中包含值的列的名称。
 * @return {void} 该函数不返回任何值。
 */
exports.jsonToXlsx = (option) => {
  /** @type {XLSX.WorkBook} */
  const workBook = {
    SheetNames: [],
    Sheets: {},
  };

  fs.readdirSync(option.dirPath).forEach((fileName) => {
    if (!fileName.endsWith(".json")) {
      console.log("error fileName:", fileName);
      return;
    }
    console.log("fileName:", fileName);

    try {
      const filePath = `${option.dirPath}/${fileName}`;
      const data1 = fs.readFileSync(filePath, { encoding: "utf8" });

      /** @type {Record<string, string>} */
      const data2 = JSON.parse(data1);

      let data3 = [];
      for (const key in data2) {
        data3.push({
          [option.keyCode]: key,
          [option.valueCode]: data2[key],
        });
      }

      const sheetName = fileName.replace(".json", "");
      workBook.SheetNames.push(sheetName);
      workBook.Sheets[sheetName] = XLSX.utils.json_to_sheet(data3);
    } catch (error) {
      console.log("error fileName:", fileName);
      console.error(error);
    }
  });

  XLSX.writeFile(workBook, `${option.dirPath}.xlsx`);
};
