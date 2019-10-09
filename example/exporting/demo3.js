/*
 * @Description: 未描述
 * @Author: danielmlc
 * @Date: 2019-10-08 16:59:20
 * @LastEditTime: 2019-10-09 11:08:59
 */
const excel = require('exceljs')
let options = {
    filename: '../files/demo3.xlsx',
    useStyles: true,
    useSharedStrings: true
  };
let workbook = new excel.stream.xlsx.WorkbookWriter(options);

let worksheet = workbook.addWorksheet('ceshi');
let headers = [
    { header: 'csv.content', width: 15 },
    { header: 'csv.ancestor', width: 10 },
    { header: 'csv.note',  width: 20 },
    { header: 'csv.priority', width: 10 },
    { header: 'csv.executor', width: 10 },
    { header: 'csv.startDate', width: 20 },
    { header: 'csv.dueDate', width: 20 },
    { header: 'csv.creator', width: 20 },
    { header: 'csv.created', width: 20 },
    { header: 'csv.isDone', width: 10 },
    { header: 'csv.accomplished', width: 20 },
    { header: 'csv.tasklist', width: 10 },
    { header: 'csv.stage', width: 10 },
    { header: 'csv.delayDays', width: 10 },
    { header: 'csv.delayed', width: 10 },
    { header: 'csv.totaltime', width: 10 },
    { header: 'csv.usedtime', width: 10 },
    { header: 'csv.tag', width: 10 }
  ]

worksheet.columns = headers

let headRow = worksheet.getRow(1);

headRow.eachCell(function(cell,colNum){
    cell.border ={
        top: {style:'thin'},
        left: {style:'thin'},
        bottom: {style:'thin'},
        right: {style:'thin'}
      };
    cell.fill = {
        type: 'pattern',
        pattern:'darkGray',
        fgColor:{argb:'CCCCCC00'},
        bgColor:{argb:'CCCCCC00'}
    }
})

worksheet.addRow([
    1111,
    2222,
    334,
    444,
    3333
]);

worksheet.commit()
workbook.commit()  
console.log('done!')