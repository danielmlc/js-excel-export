/*
 * @Description: 未描述
 * @Author: danielmlc
 * @Date: 2019-10-08 16:59:20
 * @LastEditTime: 2019-10-10 19:10:34
 */
const ExcelCls = require('../../lib/xlsx.js')
const configJson = require('../data/data1.structure.json')
const testData = require('../data/simple.data.json')
let _excel = new  ExcelCls(
    {
        workbook:{
            creator: '租户管理员',
        }
    }  
)
_excel.getWorkbook(null,function(wb){
    let sheetConf = configJson.sheetConf;
     _excel.getSheet(wb, sheetConf.title,function(ws){
        _excel.addSimpleTable(ws, sheetConf.header,testData);
     });
     _excel.getSheet(wb, 'hahahaha',function(ws){
        _excel.addSimpleTable(ws, sheetConf.header,testData);
     });
});

console.log('done!')