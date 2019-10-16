/*
 * @Description: 未描述
 * @Author: danielmlc
 * @Date: 2019-10-06 23:33:09
 * @LastEditTime: 2019-10-16 10:39:53
 */
const defaultConf = {
    workbook:{
        creator: 'yearrow',
        lastModifiedBy: 'yearrow',
        created: new Date(),
        modified: new Date(),
        lastPrinted: new Date(),
        views:[
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
        ]
    },
    worksheet:{
        properties: {
            // 工作表属性
        },
        pageSetup: {
            //工作表打印的属性
            showGridLines: false,
            paperSize: 9,
            orientation: "portrait",
            fitToPage: true,
            fitToHeight:1,
            fitToWidth: 1,
            margins: {
                left: 0.7,
                right: 0.7,
                top: 0.75,
                bottom: 0.75,
                header: 0.3,
                footer: 0.3
            }
        },
        headerFooter:{
            oddFooter:"第 &P 页，共 &N 页"
        },
         // 导出默认样式
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