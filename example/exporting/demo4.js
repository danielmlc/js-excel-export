const fs = require('fs')
const Excel = require('exceljs')



let options = {
    filename: '../files/demo4.xlsx',
    useStyles: true,
    useSharedStrings: true
  };
let workbook = new Excel.stream.xlsx.WorkbookWriter(options);

workbook.creator = 'test'
workbook.lastModifiedBy = 'test'
workbook.created = new Date()
workbook.modified = new Date()

let sheet = workbook.addWorksheet('2018-10报表')

// # Add column headers and define column keys and widths
// 添加表头
sheet.getRow(10).values = ['种类', '销量', , , , '店铺'];
sheet.getRow(11).values = ['', '2018-05', '2018-06', '2018-07', '2018-08', ''];
// 添加数据项定义，与之前不同的是，此时去除header字段

//columns不能用push直接添加数据需要先动态创建好数据然后sheet.columns=arr;格式如下
sheet.columns = [
    {key: 'category', width: 30},
    {key: '2018-08', width: 30},
    {key: '2018-05', width: 30},
    {key: '2018-06', width: 30},
    {key: '2018-07', width: 30},
    {key: 'store', width: 30},
]

// 合并单元格
//
sheet.mergeCells(10, 2, 10, 5);//第1行  第2列  合并到第1行的第5列
sheet.mergeCells(10, 1, 11, 1);
sheet.mergeCells(10, 6, 11, 6);



const data = [{
    category: '衣服',
    '2018-05': 300,
    '2018-06': 230,
    '2018-07': 730,
    '2018-08': 630,
    '2018-066': 782,
    'store': '王小二旗舰店'
}, {
    category: '零食',
    '2018-05': 672,
    '2018-06': 826,
    '2018-07': 302,
    '2018-08': 389,
    'store': '吃吃货'
}]
sheet.addRow(data[0])

// 设置每一列样式
const row = sheet.getRow(10)
row.eachCell((cell, rowNumber) => {
    cell.font = {
        name: '微软雅黑',
        color: { argb: 'FFFFCCFF' },
        family: 4,
        size: 12,
        underline: false,
        bold: true
    };
    cell.alignment = {
        vertical: 'middle', 
        horizontal: 'center'
    };
})
const row1 = sheet.getRow(11)
row1.eachCell((cell, rowNumber) => {
    cell.font = {
        name: '微软雅黑',
        color: { argb: 'FFFFCCFF' },
        family: 4,
        size: 12,
        underline: false,
        bold: true
    };
    cell.alignment = {
        vertical: 'middle', 
        horizontal: 'center'
    };
})


sheet.getRow(20).values = ['c1', 'c2'];
sheet.columns = [
    {key: 'a', width: 30},
    {key: 'b', width: 30},
]

sheet.addRow({
    a: '衣服',
    b: 300,
})

sheet.commit()
workbook.commit() 
// workbook.xlsx.writeFile("../files/demo4.xlsx").then(function() {
//     // done
//     console.log("done");
//   });