const XLSX = require("xlsx");
const path = require("path");

// 读取 Excel 文件
function readExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet);
}

const map266path = `D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本关卡配置表@MapTranslationConfiguration.xlsx`;
const total266path = `D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本配置表@TotalTranslationConfiguration.xlsx`;
const system266path = `D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本系统配置表@SystemTranslationConfiguration.xlsx`;
const ops266path = `D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本运营配置表@OpsEvenTranslationConfiguration.xlsx`;
const battle266path = `D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本战斗配置表@BattleTranslationConfiguration.xlsx`;
const OpsXYpath = `./batterXS.xlsx`;


// 读取 Excel 文件
const map266Data = readExcel(map266path);
const total266Data = readExcel(total266path);
const system266Data = readExcel(system266path);
const ops266Data = readExcel(ops266path);
const battle266Data = readExcel(battle266path);

const OpsXYData = readExcel(OpsXYpath);
const merge266Data = [...map266Data,...total266Data,...system266Data,...ops266Data,...battle266Data]


// 将数据转换为以 Key 为键的对象，方便查找
const Ops266Map = new Map(merge266Data.map(item => [item.Key, item]));
const OpsXYMap = new Map(OpsXYData.map(item => [item.Key, item]));

// 任务1：找到 OpsXYData 有但 Ops266Data 没有的 Key
const missingKeys = [];
for (const key of OpsXYMap.keys()) {
    if (!Ops266Map.has(key)) {
        missingKeys.push(OpsXYMap.get(key));
    }
}

// 生成 Excel 表格
function writeExcel(data, outputPath, sheetName = "Sheet1") {
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    XLSX.writeFile(workbook, outputPath);
}

// 输出结果到文件
writeExcel(missingKeys, "./missing_keys.xlsx", "Missing Keys");

console.log("处理完成！");