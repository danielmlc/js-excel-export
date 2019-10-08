/*
 * @Description: 未描述
 * @Author: danielmlc
 * @Date: 2019-10-08 16:59:20
 * @LastEditTime: 2019-10-08 17:24:03
 */
const ExcelCls = require('../../lib/xlsx.js')


let _excel = new  ExcelCls(
    {
        workbook:{
            creator: 'xxxxx',
        }
    }  
)

let workbook = _excel.getWorkbook();


// 保存表格
workbook.xlsx.writeFile("../files/demo2.xlsx").then(function() {
    // done
    console.log("done");
});
  
console.log('done!')