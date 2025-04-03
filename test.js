const XLSX = require("xlsx");
const path = require("path");
const { log } = require("console");

// 读取 Excel 文件
function readExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet);
}

const morePath = `./system_mengdameng.xlsx`;

const lesspath = `./mengdameng_missing_keys.xlsx`;

// 读取 Excel 文件
const moreData = readExcel(morePath);
const lessData = readExcel(lesspath);

// 将数据转换为以 Key 为键的对象，方便查找
const moreMap = new Map(moreData.map(item => [item.Key, item]));
const lessMap = new Map(lessData.map(item => [item.Key, item]));


const missingKeys = [];
for (const key of moreMap.keys()) {
    if (!lessMap.has(key)) {
        missingKeys.push(moreMap.get(key));
    }
}

console.log(missingKeys,'missingKeys');


