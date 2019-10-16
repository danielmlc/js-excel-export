/*
 * @Description: 未描述
 * @Author: danielmlc
 * @Date: 2019-10-06 21:46:16
 * @LastEditTime: 2019-10-16 18:53:01
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
    getWorkbook(filesPath, goFn){
        let _option = {
            filename: filesPath,
            useStyles: true,
            useSharedStrings: true
        };
        let workbook =  new excel.stream.xlsx.WorkbookWriter(_option);
        // 设置工作薄属性
        workbook.creator = this._config.workbook.creator;
        workbook.lastModifiedBy = this._config.workbook.creator;
        workbook.created = this._config.workbook.created;
        workbook.modified = this._config.workbook.modified;
        workbook.lastPrinted = this._config.workbook.lastPrinted;
        workbook.views = this._config.workbook.views;
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
    // 添加单表 (只能从第一行添加)
    addSimpleTable(worksheet,tableConf,data){
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
            let headRow = worksheet.getRow(1);
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
                    let contentRow = worksheet.getRow(index + 2);
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
    addComplexTable(worksheet,tableConf,data,lineNum = 1){
        let _this = this
        let _columnsInfo = {
            level:1,
            _level:1,
            columns:[],
            headerInfo:[],
            mergeArr:[]
        }
        if(tableConf && tableConf.columns){
            this._getParamsForObj(tableConf.columns, _columnsInfo)
            // 处理层级 根据层级生成二维数组
            _columnsInfo.headerArr = []
            for(let i = 0; i < _columnsInfo._level; i++){
                _columnsInfo.headerArr[i]=[]
            }
            this._generateArr(_columnsInfo,lineNum)

            // 布局列头结构
            _columnsInfo.headerArr.map((item, index)=>{
                worksheet.getRow(lineNum + index).values = item
                //格式化
                let headRow = worksheet.getRow(lineNum + index);
                headRow.eachCell(function(cell){
                    const { font,alignment,border,fill } = _this._config.worksheet.selfOptions.headConf.cell
                    cell.font = font;
                    cell.alignment = alignment;
                    cell.border = border;   
                    cell.fill = fill;
                })
            })
            // 定义列属性
            worksheet.columns = _columnsInfo.columns

            // 合并单元格
            _columnsInfo.mergeArr.map(data =>{
               worksheet.mergeCells(data.sy,data.sx,data.gy,data.gx)
            })
            
        }
        // 绘制数据项
        if(data){
            data.map((item,index)=>{
                // 处理值
                if(item.format){
                    fn = item.format;
                    fn(item)
                }
                worksheet.addRow(item);    
                //添加样式
                let contentRow = worksheet.getRow(index + lineNum + _columnsInfo._level);
                contentRow.eachCell(function(cell){
                    const { font,alignment,border,fill } = _this._config.worksheet.selfOptions.bodyConf.cell
                    cell.font = font;
                    cell.alignment = alignment;
                    cell.border = border;   
                    if(index % 2){
                        cell.fill = fill;
                    }
                })
            })
        }
    }
    
    // 获取列层深度和长度
    _getParamsForObj(conf, columnsInfo){
        const nowLevel = columnsInfo.level
        conf.map((i,index) =>{
              // 计算当前节点宽度
              let lengthArr = []
              if(i.item){
                this._getTreeWith(i.item, lengthArr)
              }else{
                lengthArr = ['']
              }
            columnsInfo.headerInfo.push({
                title:i.title,
                nowLevel:nowLevel,
                leaf:!i.item,
                width:lengthArr.length
            })
            if(i.item){
                columnsInfo.level += 1
                this._getParamsForObj(i.item, columnsInfo)

            }else{
                columnsInfo.columns.push({
                    key: i.name||index,
                    width: i.width||12
                })
            }
        })
        columnsInfo._level = Math.max(columnsInfo._level,columnsInfo.level);
        columnsInfo.level = 1
    }

    // 根据列配置生成列数组
    _generateArr(columnsInfo,lineNum) {
        columnsInfo.headerInfo.map((i) =>{
           

           if(i.leaf){
                //生成纵向合并信息
                if(columnsInfo._level - i.nowLevel > 0){
                    let obj = {
                        sy : lineNum + i.nowLevel-1,
                        sx : columnsInfo.headerArr[i.nowLevel-1].length + 1,
                        gy : lineNum + columnsInfo._level - 1,
                        gx : columnsInfo.headerArr[i.nowLevel-1].length + 1,
                    }
                    columnsInfo.mergeArr.push(obj)
                }
                for(let ii=i.nowLevel;ii<=columnsInfo._level;ii++){
                    if(ii === i.nowLevel){
                        columnsInfo.headerArr[ii-1].push(i.title)
                    }else{
                        columnsInfo.headerArr[ii-1].push('')
                    }
                }
           }else{
               columnsInfo.headerArr[i.nowLevel-1].push(i.title)
               if(i.width>1){
                    // 生成横向合并信息
                    let obj = {
                        sy : lineNum + i.nowLevel-1,
                        sx : columnsInfo.headerArr[i.nowLevel-1].length,
                        gy : lineNum + i.nowLevel-1,
                        gx : columnsInfo.headerArr[i.nowLevel-1].length +i.width-1,
                    }
                    columnsInfo.mergeArr.push(obj)
                    for(let ii=1;ii<i.width;ii++){
                        columnsInfo.headerArr[i.nowLevel-1].push('')
                    }
               }
              

           }
        })
    }
    
    _getTreeWith(conf,expandObj) {
        conf.map((data)=>{
            if(data.item){
                this._getTreeWith(data.item,expandObj)
            }else{
                expandObj.push(data.title)
            }
        })
    }

    addRichText(worksheet,tableConf,data,lineNum = 1){
        worksheet.getCell('A1').value = {
            'richText': [
              {'font': {'size': 12,'color': {'theme': 0},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'This is '},
              {'font': {'italic': true,'size': 12,'color': {'theme': 0},'name': 'Calibri','scheme': 'minor'},'text': 'a'},
              {'font': {'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': ' '},
              {'font': {'size': 12,'color': {'argb': 'FFFF6600'},'name': 'Calibri','scheme': 'minor'},'text': 'colorful'},
              {'font': {'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': ' text '},
              {'font': {'size': 12,'color': {'argb': 'FFCCFFCC'},'name': 'Calibri','scheme': 'minor'},'text': 'with'},
              {'font': {'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': ' in-cell '},
              {'font': {'bold': true,'size': 12,'color': {'theme': 1},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'format'}
            ]
          };
          
        //   expect(worksheet.getCell('A1').text).to.equal('This is a colorful text with in-cell format');
        //   expect(worksheet.getCell('A1').type).to.equal(Excel.ValueType.RichText);
    }
}