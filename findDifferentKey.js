const xlsx = require("xlsx");
const path = require("path");

// 读取 Excel 文件
function readExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet);
}

const Ops266path = `D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本运营配置表@OpsEvenTranslationConfiguration.xlsx`;
const OpsXYpath = `./opsXY.xlsx`;

// 读取Excel 文件
const Ops266Data = readExcel(Ops266path, "Ops266path");
const OpsXYData = readExcel(OpsXYpath, "OpsXYpath");