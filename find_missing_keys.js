const XLSX = require("xlsx");
const path = require("path");

// 读取 Excel 文件
function readExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet);
}

const Ops266path = `D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本运营配置表@OpsEvenTranslationConfiguration.xlsx`;
const OpsXYpath = `./opsXY2.xlsx`;

// 读取 Excel 文件
const Ops266Data = readExcel(Ops266path);
const OpsXYData = readExcel(OpsXYpath);

// 将数据转换为以 Key 为键的对象，方便查找
const Ops266Map = new Map(Ops266Data.map(item => [item.Key, item]));
const OpsXYMap = new Map(OpsXYData.map(item => [item.Key, item]));

// 任务1：找到 OpsXYData 有但 Ops266Data 没有的 Key
const missingKeys = [];
for (const key of OpsXYMap.keys()) {
    if (!Ops266Map.has(key)) {
        missingKeys.push(OpsXYMap.get(key));
    }
}

// 任务2：找到 Key 相同但 Translate 不一致的数据
const inconsistentTranslates = [];
for (const [key, value] of OpsXYMap.entries()) {
    if (Ops266Map.has(key)) {
        const ops266Item = Ops266Map.get(key);
        if (value.Translate !== ops266Item.Translate) {
            inconsistentTranslates.push({
                Key: key,
                OpsXY_Translate: value.Translate,
                Ops266_Translate: ops266Item.Translate,
                ToolRemark: value.ToolRemark || ops266Item.ToolRemark
            });
        }
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
writeExcel(inconsistentTranslates, "./inconsistent_translates.xlsx", "Inconsistent Translates");

console.log("处理完成！");