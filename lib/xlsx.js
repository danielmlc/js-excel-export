/*
 * @Description: 未描述
 * @Author: danielmlc
 * @Date: 2019-10-06 21:46:16
 * @LastEditTime: 2019-10-09 13:08:14
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
    getWorkbook(option){
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
        return workbook;
    }
    // 创建sheet
    getSheet(wb,structure,data) {
        let worksheet = null
        if(structure){
            worksheet = wb.addWorksheet(structure.title, this._config.worksheet);
            //添加列头
            this._addColumnsHeader(worksheet,structure.header);
            //构造数据
            if(data){
            }
        }
        return worksheet;
    }


    _addColumnsHeader(worksheet,hc){
        if(hc && hc.columns){
            let columns = []
            hc.columns.map(
                (data,index)=>{
                    columns.push({
                        header: data.title||'默认', 
                        key: data.name||index, 
                        width: data.width||12
                    })
            })
            worksheet.columns = columns;

            //格式化
            let headRow = worksheet.getRow(1);
            headRow.eachCell(function(cell){
                console.log(this._config)
                const { font,alignment,border,fill } = this._config.worksheet.selfOptions.headConf.cell
                cell.font = font;
                cell.alignment = alignment;
                cell.border = border;   
                cell.fill = fill;
            })
        }
    }

}