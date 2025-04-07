const xlsx = require('xlsx');

// 输入和输出文件路径
const inputFilePath = 'opsXY2.xlsx'; // 原始表格文件
const outputFilePath = 'opsXY2_with_remarks.xlsx'; // 更新后的表格文件

// 读取 Excel 文件
const workbook = xlsx.readFile(inputFilePath);
const sheetName = workbook.SheetNames[0]; // 默认处理第一个工作表
const worksheet = workbook.Sheets[sheetName];

// 将工作表转换为 JSON 数据
const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

// 获取表头
const headers = rows[0];
headers.push('备注'); // 新增“备注”列

// 处理每一行数据
for (let i = 1; i < rows.length; i++) {
  const row = rows[i];
  const toolRemark = row[headers.indexOf('ToolRemark')] || '';
  const version = row[headers.indexOf('Version')] || '';
  const po = row[headers.indexOf('PO')] || '';

  // 根据逻辑生成“备注”列，并添加换行符 \n
  const remark = `场景：${toolRemark}\n使用版本：${version}\n负责人：${po}`;
  row.push(remark); // 添加新列
}

// 将更新后的数据转换回工作表
const updatedWorksheet = xlsx.utils.aoa_to_sheet(rows);

// 设置单元格样式以支持换行
updatedWorksheet['!cols'] = headers.map(() => ({ width: 20 })); // 设置列宽
for (let i = 1; i < rows.length; i++) {
  const cellAddress = xlsx.utils.encode_cell({ c: headers.length - 1, r: i }); // 备注列的单元格地址
  if (updatedWorksheet[cellAddress]) {
    updatedWorksheet[cellAddress].s = { alignment: { wrapText: true } }; // 启用自动换行
  }
}

// 创建一个新的工作簿并将更新后的工作表添加到其中
const newWorkbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(newWorkbook, updatedWorksheet, sheetName);

// 写入新的 Excel 文件
xlsx.writeFile(newWorkbook, outputFilePath);

console.log(`新文件已生成：${outputFilePath}`);