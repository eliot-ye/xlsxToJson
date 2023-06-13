// @ts-check

const fs = require("fs");
const XLSX = require("xlsx");

const path = "/Users/eliot/Desktop/work/HKHA-iHousing/src/assets/i18n";

const workBook = {
  SheetNames: [],
  Sheets: {},
};

fs.readdirSync(path).forEach((fileName) => {
  if (!fileName.endsWith(".json")) {
    console.log("error fileName:", fileName);
    return;
  }
  console.log("fileName:", fileName);

  try {
    const filePath = `${path}/${fileName}`;
    const data1 = fs.readFileSync(filePath, { encoding: "utf8" });

    /** @type {{[key: string]: string}} */
    const data2 = JSON.parse(data1);

    let data3 = [];
    for (const key in data2) {
      data3.push({ code: key, value: data2[key] });
    }

    const sheetName = fileName.replace(".json", "");
    workBook.SheetNames.push(sheetName);
    workBook.Sheets[sheetName] = XLSX.utils.json_to_sheet(data3);
  } catch (error) {
    console.log("error fileName:", fileName);
    console.error(error);
  }
});

XLSX.writeFile(workBook, `${path}.xlsx`);
