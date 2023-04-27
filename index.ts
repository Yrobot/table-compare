import * as Excel from "exceljs";
import * as csv from "csv-parser";
import * as fs from "fs";

const getXlsxTitles = async (filePath) => {
  const workbook = new Excel.Workbook();
  const worksheet = (await workbook.xlsx.readFile(filePath)).getWorksheet(1);
  return worksheet.getColumn(5).values;
};

const getCsvTitles = async (filePath) =>
  new Promise((resolve, reject) => {
    let result: string[] = [];
    fs.createReadStream(filePath)
      .pipe(csv({ separator: "\t" }))
      .on("data", (d) => {
        result.push(d.Title);
      })
      .on("end", () => {
        resolve(result);
      });
  });

const config: { path: string; getTitles: Function }[] = [
  {
    path: "db/OSS1.0.xlsx",
    getTitles: async (path) => getXlsxTitles(path).then((arr) => arr.slice(2)),
  },
  {
    path: "db/OSS2.0.csv",
    getTitles: getCsvTitles,
  },
];

(async () => {
  const result: string[] = [];
  const map = {};
  const titles1 = await config[0].getTitles(config[0].path);
  const titles2 = await config[1].getTitles(config[1].path);

  titles1.forEach((title, i) => {
    map[title] = true;
  });
  titles2.forEach((title, i) => {
    if (!map[title]) {
      result.push(title);
    } else {
    }
  });
  await fs.writeFileSync("result.txt", result.join("\n"));

  console.log(titles1.length, titles2.length, result.length);
})();
