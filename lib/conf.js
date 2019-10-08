/*
 * @Description: 未描述
 * @Author: danielmlc
 * @Date: 2019-10-06 23:33:09
 * @LastEditTime: 2019-10-08 18:17:28
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
            tabColor: { argb: "FFC0000" },
            defaultRowHeight: 15
        },
        pageSetup: {
            //工作表打印的属性
            pageSetup: {
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
            }
        },
        headerFooter: {
            differentFirst: true,
            firstHeader: "",
            firstFooter: ""
        }
    }
}

module.exports={
    defaultConf
}