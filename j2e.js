const fs = require("fs");
const XLSX = require("xlsx");

let inputDir = "inputJson";
let outputFile = "output.xlsx";

// 取得檔案們
let files = fs.readdirSync(inputDir);

// 檔案內容處理成 {
//   key1:{lang1:'xx',lang2:'xx'}
//   key2:{lang1:'yy',lang2:'yy'}...
// }
let langs = [];
let i18n = {};
for(let file of files) {
  if(getFileExtension(file) === 'json') {
    let data = JSON.parse(fs.readFileSync(`${inputDir}/${file}`, 'utf8'));
    let lang = getFileName(file);
    langs.push(lang);
    for(let key in data){
      if(!i18n[key]){
        i18n[key] = {};
      }
      i18n[key][lang] = data[key];
    }
  }
}

// 轉成 Sheet
let ws_data = [];
// header
let ws_data_header = ['key'].concat(langs);
ws_data.push(ws_data_header);
// row
for(let key in i18n){
  let ws_data_row = [key];
  for(let lang of langs){
    ws_data_row.push(i18n[key][lang]);
  }
  ws_data.push(ws_data_row);
}
let ws = XLSX.utils.aoa_to_sheet(ws_data);


// 如果 output 檔案存在，砍掉
if (fs.existsSync(outputFile)) {
  fs.unlinkSync(outputFile);
}

// 建立 Workbooks
// 將 Sheet 加入
// 存檔
let wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
XLSX.writeFile(wb, outputFile);

function getFileExtension(filename) {
  return filename.slice((filename.lastIndexOf(".") - 1 >>> 0) + 2).toLowerCase();
}
function getFileName(filename) {
  let idx = filename.lastIndexOf(".");
  return idx > 0 ? filename.slice(0, idx) : filename;
}