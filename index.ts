import * as Excel from "exceljs";
import * as fs from "fs";

const getXlsxSheet = async (filePath) => {
  const workbook = new Excel.Workbook();
  const worksheet = (await workbook.xlsx.readFile(filePath)).worksheets[0];
  return worksheet;
  // return worksheet.getColumn(column).values;
};

const writeXlsxSheet = async (filePath: string, data: any[]) => {
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet();
  worksheet.addRows(data);
  return await workbook.xlsx.writeFile(filePath);
};

const config: { path: string; getSheet: Function }[] = [
  {
    path: "db/OSS1.0.xlsx",
    getSheet: getXlsxSheet,
  },
  {
    path: "db/OSS2.0.xlsx",
    getSheet: getXlsxSheet,
  },
];

const formatTitle = (title: string) => title.trim().toLocaleLowerCase();

(async () => {
  const result: Object[] = [];
  const map = {};
  const sheet1 = await config[0].getSheet(config[0].path);
  const sheet2 = await config[1].getSheet(config[1].path);

  const titles1 = sheet1.getColumn(2).values.slice(2);

  titles1.forEach((title, i) => {
    // if (i < 5) console.log(title);
    map[formatTitle(title)] = true;
  });

  sheet2.eachRow((row, i) => {
    if (i === 1) {
      result.push(row.values);
      return;
    }
    const title = row.values[2];

    // if (i < 5) console.log(title);

    if (!map[formatTitle(title)]) {
      result.push(row.values);
    }
  });

  console.log(sheet1.rowCount - 1, sheet2.rowCount - 1, result.length - 1);

  await writeXlsxSheet("output.xlsx", result);
})();
