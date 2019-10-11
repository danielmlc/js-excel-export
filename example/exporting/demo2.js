/*
 * @Description: 未描述
 * @Author: danielmlc
 * @Date: 2019-10-08 16:59:20
 * @LastEditTime: 2019-10-11 09:51:56
 */
const ExcelCls = require('../../lib/xlsx.js')
const configJson = require('../data/data1.structure.json')
const testData = require('../data/simple.data.json')

const configJson1 = require('../data/complex.structure.json')
const testData1 = require('../data/data1.json')

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
     let sheetConf1 = configJson1.sheetConf;
     _excel.getSheet(wb, '复杂表格写入',function(ws){
        _excel.addComplexTable(ws, sheetConf1.header,testData1);
     });
});

console.log('done!')