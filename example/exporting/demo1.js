/*
 * @Description: 未描述
 * @Author: danielmlc
 * @Date: 2019-10-05 23:25:33
 * @LastEditTime: 2019-10-06 20:04:55
 */
const Excel = require("exceljs");
const fs = require("fs");
// 创建工作薄
var workbook = new Excel.Workbook();

// 设置工作薄属性

workbook.creator = "danielmlc";
workbook.lastModifiedBy = "danielmlc";
workbook.created = new Date();
workbook.modified = new Date();
workbook.lastPrinted = new Date();

// 工作薄视图设置
workbook.views = [
  {
    x: 0,
    y: 0,
    width: 10,
    height: 10,
    firstSheet: 0, // 指定第一个sheet
    activeTab: 0, // 打开的第一个工作表
    showRuler: false,
    visibility: "visible"
  }
];

let sheet = workbook.addWorksheet("demo1", {
  properties: {
    // 工作表属性
    tabColor: { argb: "FFC0000" },
    defaultRowHeight: 15
  },
  pageSetup: {
    //工作表打印的属性
    pageSetup: {
      showGridLines: false,
      paperSize: 9,
      orientation: "landscape",
      fitToPage: true,
      fitToHeight: 5,
      fitToWidth: 7,
      margins: {
        left: 0.7,
        right: 0.7,
        top: 0.75,
        bottom: 0.75,
        header: 0.3,
        footer: 0.3
      },
      printArea: "A1:G20",
      printTitlesRow: "1:3",
      printTitlesColumn: "A:C"
    }
  },
  headerFooter: {
    differentFirst: true,
    firstHeader: "Hello Exceljs",
    firstFooter: "Hello World"
  }
});

sheet.columns = [
  { header: "创建日期", key: "create_time", width: 15, 
    style: { 
      font: { name: 'Arial Black', bold: true, size: 16, color: { argb: 'FF00FF00' } },
      alignment:{ vertical: 'middle', horizontal: 'center'},
    } 
  },
  { header: "单号", key: "id", width: 15 },
  { header: "电话号码", key: "phone", width: 15 },
  { header: "地址", key: "address", width: 40 }
];
sheet.getCell('A1').border = {
    top: {style:'thin'},
    left: {style:'thin'},
    bottom: {style:'thin'},
    right: {style:'thin'}
  };
sheet.getCell('A1').fill = {
    type: 'pattern',
    pattern:'darkGray',
    fgColor:{argb:'CCCCCC00'},
    bgColor:{argb:'CCCCCC00'}
  };
// sheet.mergeCells('A1', 'A2');
// sheet.mergeCells('C1', 'C2');
// sheet.mergeCells('D1', 'G2');
// sheet.getCell('A1').dataValidation = {
//     type: 'list',
//     allowBlank: true,
//     formulae: ['"One,Two,Three,Four"']
//   };

const data = [
  {
    create_time: "2018-10-01",
    id: "787818992109210",
    phone: "11111111111",
    address: "深圳市"
  },
  {
    create_time: "2018-10-01",
    id: "787818992109210",
    phone: "11111111111",
    address: "深圳市"
  },
  {
    create_time: "2018-10-01",
    id: "787818992109210",
    phone: "11111111111",
    address: "深圳市"
  },
  {
    create_time: "2018-10-01",
    id: "787818992109210",
    phone: "11111111111",
    address: "深圳市"
  }
];

sheet.addRows(data);

// 保存表格
workbook.xlsx.writeFile("../files/demo11.xlsx").then(function() {
  // done
  console.log("done");
});
