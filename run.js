const Excel = require("exceljs");
const fs = require("node:fs");
const date = require("date-and-time");

async function copyExcel() {
  //NOTE: get the date of yesterday
  var now = new Date();
  now = date.addDays(now, -2);
  var yesterday = date.format(now, "YYYY-MM-DD");

  var targetWorkbook = new Excel.Workbook();
  targetWorkbook = await targetWorkbook.xlsx.readFile("./target.xlsx");
  //SESSION:领星产品表现
  var input1 = new Excel.Workbook();
  var lxProduct = targetWorkbook.getWorksheet("领星产品表现");
  input1 = await input1.xlsx.readFile("./input1.xlsx");
  //worksheets:use array instaed of name as if different workbooks have the same worksheet named "sheet1",it won't work
  var sheet1 = input1.worksheets[0];
  //NOTE: add date to source worksheet
  sheet1.spliceColumns(1, 0, ["时间"]); //have to specify some words like['时间‘]
  const dateColumn = sheet1.getColumn("A");
  dateColumn.eachCell({ includeEmpty: false }, (cell, rowNumber) => {
    if (rowNumber > 1) {
      cell.value = yesterday;
      cell.numFmt = "yyyy-mm-dd";
    }
  });
  await input1.xlsx.writeFile("./input1.xlsx");

  //NOTE: find the empty row in target worksheet
  var emptyRowNumber = [];
  for (let i = 1; i >= 0; i++) {
    if (lxProduct.getRow(i).values[1] === undefined) {
      emptyRowNumber.push(i);
      break;
    }
  }
  emptyRowNumber = emptyRowNumber[0];

  //NOTE: copy source worksheet to the target worksheet begin from an empty row.
  var diff = emptyRowNumber - 2;
  for (let i = emptyRowNumber; i > 0; i++) {
    if (sheet1.getRow(i - diff).values[1] !== undefined) {
      lxProduct.getRow(i).values = sheet1.getRow(i - diff).values;
    } else {
      break;
    }
  }

  //SESSION:亿数通产品表现
  var input2 = new Excel.Workbook();
  var ystProduct = targetWorkbook.getWorksheet("亿数通产品表现");
  input2 = await input2.xlsx.readFile("./input2.xlsx");
  var sheet2 = input2.worksheets[0];

  //NOTE: add date to source worksheet
  sheet2.spliceColumns(1, 0, ["时间"]); //have to specify some words like['时间‘]
  const dateColumn2 = sheet2.getColumn("A");
  dateColumn2.eachCell({ includeEmpty: false }, (cell, rowNumber) => {
    if (rowNumber > 1) {
      cell.value = yesterday;
      cell.numFmt = "yyyy-mm-dd";
    }
  });
  await input2.xlsx.writeFile("./input2.xlsx");

  //NOTE: find the empty row in target worksheet
  var emptyRowNumber2 = [];
  for (let i = 1; i >= 0; i++) {
    if (ystProduct.getRow(i).values[1] === undefined) {
      emptyRowNumber2.push(i);
      break;
    }
  }
  emptyRowNumber2 = emptyRowNumber2[0];

  //NOTE: copy source worksheet to the target worksheet begin from an empty row.
  var diff2 = emptyRowNumber2 - 2;
  for (let i = emptyRowNumber2; i > 0; i++) {
    if (sheet2.getRow(i - diff2).values[1] !== undefined) {
      ystProduct.getRow(i).values = sheet2.getRow(i - diff2).values;
    } else {
      break;
    }
  }

  //SESSION:亿数通业务报告
  var ystReport = targetWorkbook.getWorksheet("亿数通业务报告");
  var input3 = new Excel.Workbook();
  input3 = await input3.xlsx.readFile("./input3.xlsx");
  var sheet3 = input3.worksheets[0];

  //NOTE: find the empty row in target worksheet
  var emptyRowNumber3 = [];
  for (let i = 1; i >= 0; i++) {
    if (ystReport.getRow(i).values[1] === undefined) {
      emptyRowNumber3.push(i);
      break;
    }
  }
  emptyRowNumber3 = emptyRowNumber3[0];

  //NOTE: copy source worksheet to the target worksheet begin from an empty row.
  var diff3 = emptyRowNumber3 - 2;
  for (let i = emptyRowNumber3; i > 0; i++) {
    if (sheet3.getRow(i - diff3).values[1] !== undefined) {
      ystReport.getRow(i).values = sheet3.getRow(i - diff3).values;
    } else {
      break;
    }
  }

  // //亿数通广告日报
  // var ystLog = targetWorkbook.getWorksheet("亿数通广告日报");
  // var input4 = new Excel.Workbook();
  // input4 = await input4.xlsx.readFile("./input4.xlsx");
  // var sheet4 = input4.worksheets[0];
  // sheet4.eachRow({ includeEmpty: false }, (row, rowNumber) => {
  //   var targetRow = ystLog.getRow(rowNumber);
  //   row.eachCell({ includeEmpty: false }, (cell, cellNumber) => {
  //     if (targetRow.getCell(cellNumber).value === null) {
  //       targetRow.getCell(cellNumber).value = cell.value;
  //     }
  //   });
  // });
  await targetWorkbook.xlsx.writeFile("./target.xlsx");
}
copyExcel();
