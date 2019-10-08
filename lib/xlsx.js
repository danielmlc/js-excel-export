/*
 * @Description: 未描述
 * @Author: danielmlc
 * @Date: 2019-10-06 21:46:16
 * @LastEditTime: 2019-10-08 18:32:45
 */
const Excel = require("exceljs");
const { defaultConf } = require('./conf.js')
const fs = require("fs");
const _ = require('lodash')

module.exports = class XLSX {
    constructor(config){
        // 初始化变量
        this._config =  _.defaultsDeep(config, defaultConf)
    }

    // 初始化工作薄
    getWorkbook(){
        let workbook = new Excel.Workbook();
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
    getSheet(wb,title,structure,data) {
        let worksheet = wb.addWorksheet(title, this._config.worksheet);
        if(data){
            //构造数据
            if(structure){
                
            }
        }
        return worksheet;
    }
    // 

}