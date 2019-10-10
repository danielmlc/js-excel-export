/*
 * @Description: 未描述
 * @Author: danielmlc
 * @Date: 2019-10-06 21:46:16
 * @LastEditTime: 2019-10-10 19:09:10
 */
const excel = require("exceljs");
const { defaultConf } = require('./conf.js')
const _ = require('lodash')

module.exports = class XLSX {
    constructor(config){
        // 初始化变量
        this._config =  _.defaultsDeep(config, defaultConf)
    }
    // 初始化工作薄
    getWorkbook(option, goFn){
        let _option = {
            filename: '../files/demo3.xlsx',
            useStyles: true,
            useSharedStrings: true
        };
        if(option){
            _option = _.defaultsDeep(option, _option)
        }
        let workbook =  new excel.stream.xlsx.WorkbookWriter(_option);
        // 设置工作薄属性
        workbook.creator = this._config.workbook.creator;
        workbook.lastModifiedBy = this._config.workbook.creator;
        workbook.created = this._config.workbook.created;
        workbook.modified = this._config.workbook.modified;
        workbook.lastPrinted = this._config.workbook.lastPrinted;
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
        goFn(workbook);
        workbook.commit();
        return workbook;
    }
    // 创建sheet
    getSheet(wb,title,goFn) {
        let worksheet = wb.addWorksheet(title, this._config.worksheet);
        goFn(worksheet);
        worksheet.commit();
        return worksheet;
    }
    // 添加单表 
    addSimpleTable(worksheet,tableConf,data,lineNum = 1){
        if(tableConf && tableConf.columns){
            let columns = []
            tableConf.columns.map(
                (data,index)=>{
                    columns.push({
                        header: data.title||'默认', 
                        key: data.name||index, 
                        width: data.width||12
                    })
            })
            worksheet.columns = columns;
            //格式化
            let _this = this
            let headRow = worksheet.getRow(lineNum);
            headRow.eachCell(function(cell){
                const { font,alignment,border,fill } = _this._config.worksheet.selfOptions.headConf.cell
                cell.font = font;
                cell.alignment = alignment;
                cell.border = border;   
                cell.fill = fill;
            })
            if(data){
                data.map((item,index)=>{
                    // 处理值
                    if(item.format){
                        fn = item.format;
                        fn(item)
                    }
                    worksheet.addRow(item);    
                    //添加样式
                    let contentRow = worksheet.getRow(index + lineNum + 1);
                    contentRow.eachCell(function(cell){
                        const { font,alignment,border,fill } = _this._config.worksheet.selfOptions.bodyConf.cell
                        cell.font = font;
                        cell.alignment = alignment;
                        cell.border = border;   
                        if(index%2){
                            cell.fill = fill;
                        }
                    })
                })
            }
        }
    }
    //添加复杂表格
    addcomplexTable(){
        
    }
}