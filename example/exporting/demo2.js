/*
 * @Description: 未描述
 * @Author: danielmlc
 * @Date: 2019-10-08 16:59:20
 * @LastEditTime: 2019-10-09 12:32:21
 */
const ExcelCls = require('../../lib/xlsx.js')
const configJson = require('../data/data1.structure.json')
const testData = require('../data/data1.json')
let _excel = new  ExcelCls(
    {
        workbook:{
            creator: 'xxxxx',
        }
    }  
)
let workbook = _excel.getWorkbook();
let worksheet = _excel.getSheet(workbook,configJson.sheetConf,testData);
worksheet.commit()
workbook.commit() 
console.log('done!')