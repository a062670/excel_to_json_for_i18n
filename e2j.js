const fs = require("fs");
const XLSX = require("xlsx");

let inputFile = "input.xlsx";
let outputDir = "outputJson";

const workbook = XLSX.readFile(inputFile);
// 取得第一個sheet
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

let rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }).filter(row => {
  return row.length > 0;
});

let json = {};
let cols = [];
rows.forEach((row, rowIdx) => {
  if (rowIdx === 0) {
    row.forEach((item, colIdx) => {
      cols.push(item);
      if (colIdx > 0) {
        json[cols[colIdx]] = {};
      }
    });
  } else {
    let key = row[0];
    row.forEach((item, colIdx) => {
      if (colIdx > 0) {
        json[cols[colIdx]][key] = item;
      }
    });
  }
});

// 如果 output 資料夾存在，砍掉
if (fs.existsSync(outputDir)) {
  deleteDir(outputDir);
}
fs.mkdirSync(outputDir);

// {name}.json
for (let name in json) {
  fs.writeFileSync(
    `${outputDir}/${name}.json`,
    JSON.stringify(json[name]),
    "utf8"
  );
}

function deleteDir(rootFile) {
  var emptyDir = function(fileUrl) {
    var files = fs.readdirSync(fileUrl);
    files.forEach(function(file) {
      var stats = fs.statSync(fileUrl + "/" + file);
      if (stats.isDirectory()) {
        emptyDir(fileUrl + "/" + file);
      } else {
        fs.unlinkSync(fileUrl + "/" + file);
      }
    });
  };
  var rmEmptyDir = function(fileUrl) {
    var files = fs.readdirSync(fileUrl);
    if (files.length > 0) {
      var tempFile = 0;
      files.forEach(function(fileName) {
        tempFile++;
        rmEmptyDir(fileUrl + "/" + fileName);
      });
      if (tempFile == files.length) {
        fs.rmdirSync(fileUrl);
      }
    } else {
      fs.rmdirSync(fileUrl);
    }
  };
  emptyDir(rootFile);
  rmEmptyDir(rootFile);
}
