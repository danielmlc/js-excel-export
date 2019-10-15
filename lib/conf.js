/*
 * @Description: 未描述
 * @Author: danielmlc
 * @Date: 2019-10-06 23:33:09
 * @LastEditTime: 2019-10-15 15:09:55
 */
const defaultConf = {
    workbook:{
        creator: 'yearrow',
        lastModifiedBy: 'yearrow',
        created: new Date(),
        modified: new Date(),
        lastPrinted: new Date(),
    },
    worksheet:{
        properties: {
            // 工作表属性
        },
        pageSetup: {
            //工作表打印的属性
            showGridLines: false,
            paperSize: 9,
            orientation: "landscape",
            fitToPage: true,
            fitToHeight: 5,
            fitToWidth: 7,
            margins: {
                left: 0.7,
                right: 0.7,
                top: 0.75,
                bottom: 0.75,
                header: 0.3,
                footer: 0.3
            },
            printArea: "A1:G20",
            printTitlesRow: "1:3",
            printTitlesColumn: "A:C"
        },
        selfOptions:{
            headConf:{
                cell:{
                    font:{
                        name: '微软雅黑',
                        color: { argb: 'FFFFFFFF' },
                        family: 4,
                        size: 10,
                        underline: false,
                        bold: true
                    },
                    alignment:{
                        vertical: 'middle', 
                        horizontal: 'center'
                    },
                    border:{
                        top: {style:'thin'},
                        left: {style:'thin'},
                        bottom: {style:'thin'},
                        right: {style:'thin'}
                    },
                    fill:{
                        type: 'pattern',
                        pattern:'solid',
                        fgColor:{argb:'FF2E4053'},
                        bgColor:{argb:'FF2E4053'}
                    }
                }
            },
            bodyConf:{
                cell:{
                    font:{
                        name: '微软雅黑',
                        color: { argb: 'FF000000' },
                        family: 4,
                        size: 10,
                        underline: false,
                        bold: false
                    },
                    alignment:{
                        vertical: 'middle', 
                        horizontal: 'left'
                    },
                    border:{
                        top: {style:'thin'},
                        left: {style:'thin'},
                        bottom: {style:'thin'},
                        right: {style:'thin'}
                    },
                    fill:{
                        type: 'pattern',
                        pattern:'solid',
                        fgColor:{argb:'F9F9F9'},
                        bgColor:{argb:'F9F9F9'}
                    }
                }
            }
        }
    }
}
module.exports={
    defaultConf
}