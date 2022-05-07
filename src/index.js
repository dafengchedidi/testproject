import React from 'react'; // eslint-disable-line no-unused-vars
import ReactDOM from 'react-dom'; // eslint-disable-line no-unused-vars
import ReminderBoxFrame from "./ReminderBox"; // eslint-disable-line no-unused-vars
const ExcelJS = require('exceljs'); // eslint-disable-line no-unused-vars

class MainOne extends React.Component {
    isDisabled = false;

    // 指定单元格内容样式：四个方向的黑边框
    static contentCellStyle = {
        border: {
            top: {
                style: "medium", color: "#000"
            },
            bottom: {
                style: "medium", color: "#000"
            },
            left: {
                style: "medium", color: "#000"
            },
            right: {
                style: "medium", color: "#000"
            },
        }
    };
    // 指定标题单元格样式：加粗居中
    static headerStyle = {
        font: {
            bold: true
        },
        alignment: {
            horizontal: "center"
        }
    }



    cusPMROnClick = (dataInfo1, event) => {
        console.log('点击对应得到的数据！' + JSON.stringify(dataInfo1));
        event.currentTarget.style = {disable: false};
        console.log('点击事件信息：' + event.currentTarget.style);
    }

    cusOnBlur = (e) => {
        console.log('触发客户号！');

        var objList = [];

        var obj1 = Object();
        obj1.name = '客户名称1';
        obj1.otherInfo = '其他数据1';

        var obj2 = Object();
        obj2.name = '客户名称2';
        obj2.otherInfo = '其他数据2';

        var obj3 = Object();
        obj3.name = '客户名称3';
        obj3.otherInfo = '其他数据3';

        objList.push(obj1);
        objList.push(obj2);
        objList.push(obj3);

        // 需要循环对应营销提醒数据
        ReminderBoxFrame.openFrame({
            alertTip:
                <div>
                    {
                        objList.map((itemInfo, index) =>
                            (
                                <button key={index} disable={ this.isDisabled } onClick={this.cusPMROnClick.bind(this, itemInfo)}>对公营销提醒数据</button>
                            )
                        )
                    }
                </div>,
            closeAlert: function () {
                console.log("框体关闭了...");
            }
        });
    }

    cusAcOnBlur = (e) => {
        console.log('触发客户账号！');
    }


    /**
     * 生成带固定格式的头部信息
     * @param sheet 页签对象
     * @param tableOtherData 表格外数据
     */
    createHeadSheet = (sheet, tableOtherData) => {
        console.log('createHeadSheet总行数：' + sheet.rowCount);
        // 合并单元格
        sheet.mergeCells('A1:K1');
        // 获取指定行数对象
        let row1 = sheet.getRow(1);
        // 设置单元格数据
        row1.getCell(1).value = tableOtherData.headOrgName + '银行大连商品交易所对账表';
        // 设置文本居中
        row1.getCell(1).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        // 设置字体样式
        row1.getCell(1).font = { name: '黑体', size: 16, bold: true };
        // 设置行高
        row1.height = 27;

        // 合并单元格
        sheet.mergeCells('A2:C2');
        // 获取指定行数对象
        let row2 = sheet.getRow(2);
        // 设置单元格数据
        row2.getCell(1).value = '户名：' + tableOtherData.accountName;
        // 设置文本靠左
        row2.getCell(1).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        // 设置字体样式
        row2.getCell(1).font = { name: '黑体', size: 14, bold: true };
        // 设置行高（每行只能设置一个有效行高）
        row2.height = 20;

        // 合并单元格
        sheet.mergeCells('F2:G2');
        // 设置单元格数据
        row2.getCell(6).value = '方向：' + tableOtherData.direction;
        // 设置文本靠左
        row2.getCell(6).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        // 设置字体样式
        row2.getCell(6).font = { name: '黑体', size: 14, bold: true };

        // 合并单元格
        sheet.mergeCells('I2:J2');
        // 设置单元格数据
        row2.getCell(9).value = '币种：' + tableOtherData.ccy;
        // 设置文本靠左
        row2.getCell(9).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        // 设置字体样式
        row2.getCell(9).font = { name: '黑体', size: 14, bold: true };

        // 设置单元格数据
        row2.getCell(11).value = '单位：' + tableOtherData.unit;
        // 设置文本靠左
        row2.getCell(11).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        // 设置字体样式
        row2.getCell(11).font = { name: '黑体', size: 14, bold: true };

        // 合并单元格
        sheet.mergeCells('A3:C3');
        // 获取指定行数对象
        let row3 = sheet.getRow(3);
        // 设置单元格数据
        row3.getCell(1).value = '账号：' + tableOtherData.accountNo;
        // 设置文本靠左
        row3.getCell(1).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        // 设置字体样式
        row3.getCell(1).font = { name: '黑体', size: 14, bold: true };
        // 设置行高
        row3.height = 20;

        // 合并单元格
        sheet.mergeCells('A4:C4');
        // 获取指定行数对象
        let row4 = sheet.getRow(4);
        // 设置单元格数据
        row4.getCell(1).value = '开户行：' + tableOtherData.depositBank;
        // 设置文本靠左
        row4.getCell(1).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
        // 设置字体样式
        row4.getCell(1).font = { name: '黑体', size: 14, bold: true };
        // 设置行高
        row4.height = 20;
    }


    /**
     * 生成带固定格式的尾部信息
     * @param sheet 页签对象
     * @param tableOtherData 表格外数据
     * @param endRowNumber 尾部行数
     */
    createBottomSheet = (sheet, tableOtherData, endRowNumber) => {
        console.log('createBottomSheet总行数：' + sheet.rowCount + '\t尾部行数：' + endRowNumber);
        // 合并单元格
        sheet.mergeCells('A' + endRowNumber + ':' + 'F' + endRowNumber);
        // 获取指定行数对象
        let endRow = sheet.getRow(endRowNumber);
        // 设置单元格数据
        endRow.getCell(1).value = '合计项';
        // 设置文本居中
        endRow.getCell(1).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        // 设置字体样式
        endRow.getCell(1).font = { name: '黑体', size: 11 };
        // 设置背景样式
        endRow.getCell(1).fill = {
            type: 'pattern',
            pattern: 'darkTrellis',
            fgColor: { argb: 'D9D9D9' },
            bgColor: { argb: 'D9D9D9' }
        };
        // 设置行高（每行只能设置一个有效行高）
        endRow.height = 18;
        // 设置边框
        endRow.getCell(1).border = {
            top: {style: 'thin', color: {argb: 'black'}},
            left: {style: 'thin', color: {argb: 'black'}},
            bottom: {style: 'medium', color: {argb: 'black'}},
            right: {style: 'thin', color: {argb: 'black'}},
        };

        // 设置单元格数据
        endRow.getCell(7).value = tableOtherData.countAmt;
        // 设置文本居中
        endRow.getCell(7).alignment = { vertical: 'middle', horizontal: 'right', wrapText: true };
        // 设置字体样式
        endRow.getCell(7).font = { name: '黑体', size: 11 };
        // 设置边框
        endRow.getCell(7).border = {
            top: {style: 'thin', color: {argb: 'black'}},
            left: {style: 'thin', color: {argb: 'black'}},
            bottom: {style: 'medium', color: {argb: 'black'}},
            right: {style: 'thin', color: {argb: 'black'}},
        };

        // 合并单元格
        sheet.mergeCells('H' + endRowNumber + ':' + 'K' + endRowNumber);
        // 设置单元格数据
        endRow.getCell(8).value = '';
        // 设置文本居中
        endRow.getCell(8).alignment = { vertical: 'middle', horizontal: 'right', wrapText: true };
        // 设置字体样式
        endRow.getCell(8).font = { name: '黑体', size: 11 };
        // 设置边框
        endRow.getCell(8).border = {
            top: {style: 'thin', color: {argb: 'black'}},
            left: {style: 'thin', color: {argb: 'black'}},
            bottom: {style: 'medium', color: {argb: 'black'}},
            right: {style: 'medium', color: {argb: 'black'}},
        };

        // 尾行加1
        endRowNumber++;
        // 合并单元格
        sheet.mergeCells('J' + endRowNumber + ':' + 'K' + endRowNumber);
        // 获取指定行数对象
        let endRow1 = sheet.getRow(endRowNumber);
        // 设置单元格数据
        endRow1.getCell(10).value = '打印日期：' + tableOtherData.printDate;
        // 设置字体样式
        endRow1.getCell(10).font = { name: '黑体', size: 10 };
        // 设置文本样式
        endRow1.getCell(10).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

        // 尾行加1
        endRowNumber++;
        // 合并单元格
        sheet.mergeCells('J' + endRowNumber + ':' + 'K' + endRowNumber);
        // 获取指定行数对象
        let endRow2 = sheet.getRow(endRowNumber);
        // 设置单元格数据
        endRow2.getCell(10).value = '柜员号：' + tableOtherData.teller;
        // 设置字体样式
        endRow2.getCell(10).font = { name: '黑体', size: 10 };
        // 设置文本样式
        endRow2.getCell(10).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };

        // 尾行加1
        endRowNumber++;
        // 合并单元格
        sheet.mergeCells('J' + endRowNumber + ':' + 'K' + endRowNumber);
        // 获取指定行数对象
        let endRow3 = sheet.getRow(endRowNumber);
        // 设置单元格数据
        endRow3.getCell(10).value = '银行盖章：暂无';
        // 设置字体样式
        endRow3.getCell(10).font = { name: '黑体', size: 10 };
        // 设置文本样式
        endRow3.getCell(10).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    }


    /**
     * 填充表格标题数据以及单元格样式
     * @param sheet 页签对象
     * @param startRow 从第几行开始写入
     * @param startCell 从第几个单元格开始写入
     * @param columns 表格头部标题数据
     */
    createTableSheet = (sheet, startRow, startCell, columns) => {
        console.log('createTableSheet总行数：' + sheet.rowCount);
        // 获取页签的所有行数
        let row = sheet.getRow(startRow);
        // 表格标题循环下标，从0开始
        let titleIndex = 0;
        // 取得表格头部标题集合的长度，+1是因为页签表格是从1开始，但是数组下标是从0开始的原因。
        let cellLength = columns.length + 1;
        for (let i = startCell; i < cellLength; i++) {
            // 设置表标题
            row.getCell(i).value = columns[titleIndex];
            // 表格头部标题下标+1
            titleIndex++;
            // 设置样式
            row.getCell(i).fill = {
                type: 'pattern',
                pattern: 'darkTrellis',
                fgColor: { argb: 'D9D9D9' },
                bgColor: { argb: 'D9D9D9' }
            };
            // 设置文本居中
            row.getCell(i).alignment = {
                vertical: 'middle',
                horizontal: 'center',
                wrapText: true
            };
            // 设置单元格样式
            row.getCell(i).numFmt = '@';
            // 设置边框（最后一个单元格“时间”需要设置右边框加粗）
            if ((i + 1) == cellLength) {
                row.getCell(i).border = {
                    top: {style: 'medium', color: {argb: 'black'}},
                    left: {style: 'thin', color: {argb: 'black'}},
                    bottom: {style: 'thin', color: {argb: 'black'}},
                    right: {style: 'medium', color: {argb: 'black'}},
                };
            } else {
                row.getCell(i).border = {
                    top: {style: 'medium', color: {argb: 'black'}},
                    left: {style: 'thin', color: {argb: 'black'}},
                    bottom: {style: 'thin', color: {argb: 'black'}},
                    right: {style: 'thin', color: {argb: 'black'}},
                };
            }
        }
    }


    /**
     * 设置数据对应的每个单元格样式
     * @param sheet 页签对象
     * @param startCellIndex 需要设置的一行开始单元格位置
     * @param endCellIndex 需要设置的一行结束单元格位置
     */
    setSheetStyle(sheet, startCellIndex, endCellIndex) {
        console.log('setSheetStyle总行数：' + sheet.rowCount);
        //sheet1样式
        for (let i = 1; i <= sheet.rowCount; i++) {
            // 跳过前5行数据（表头）
            if(i <= 5) {
                continue;
            }
            // 获取当前行
            let row = sheet.getRow(i);
            // 循环给当前行设置样式
            for (let j = startCellIndex; j <= endCellIndex; j++) {
                // 设置文本居中（转账金额第7个单元格默认靠右）
                if (j == 7) {
                    row.getCell(j).alignment = {
                        vertical: 'middle',
                        horizontal: 'right',
                        wrapText: true
                    };
                } else {
                    row.getCell(j).alignment = {
                        vertical: 'middle',
                        horizontal: 'center',
                        wrapText: true
                    };
                }
                // 设置单元格样式
                row.getCell(j).numFmt = '@';
                // 设置边框（最后一个单元格“时间”需要设置右边框加粗）
                if (j == endCellIndex) {
                    row.getCell(j).border = {
                        top: {style: 'thin', color: {argb: 'black'}},
                        left: {style: 'thin', color: {argb: 'black'}},
                        bottom: {style: 'thin', color: {argb: 'black'}},
                        right: {style: 'medium', color: {argb: 'black'}},
                    };
                } else {
                    row.getCell(j).border = {
                        top: {style: 'thin', color: {argb: 'black'}},
                        left: {style: 'thin', color: {argb: 'black'}},
                        bottom: {style: 'thin', color: {argb: 'black'}},
                        right: {style: 'thin', color: {argb: 'black'}},
                    };
                }

            }
        }
    }


    /**
     * 设置列宽
     * @param column 列的信息
     * @param width 需要设置的长度
     */
    setColumnWidth = (column, width) => {
        column.width = width;
    }


    /**
     * 保存Excel文件
     * @param workBookInfo 工作簿对象
     * @param fileName 文件名称
     */
    saveExcel = (workBookInfo, fileName) => {
        // 将workbook对象转成流类型的文件，并返回流信息
        this.sheetBlob(workBookInfo, function (blob) {
            // 此为绑定标签自动下载，也可以用其他方式下载
            let tempa = document.createElement('a');
            tempa.download = fileName ? fileName + '.xlsx' : new Date().getTime() + '.xlsx';
            // 文件数据源绑定a标签
            tempa.href = URL.createObjectURL(blob);
            // 创建事件流
            let event;
            // 判断当前窗口类型
            if (window.MouseEvent) {
                event = new MouseEvent('click');
            } else {
                event = document.createEvent('MouseEvent');
                event.initMouseEvent('click',true,false,window,0,0,0,0,0,false,false,false,false,0,null);
            }
            // 派发事件模拟点击下载
            tempa.dispatchEvent(event);

            // 延时释放文件流(0.1秒)
            setTimeout(function () {
                //用URL.revokeObjectURL来释放文件的流
                URL.revokeObjectURL(blob);
            }, 100);
        });
    }


    /**
     * 将workbook对象转成流类型的文件，并返回流信息
     * @param workBookInfo workbook对象
     * @param callback 回调函数
     */
    sheetBlob = (workBookInfo, callback) => {
        workBookInfo.xlsx.writeBuffer().then(function (buffer) {
            let blob = new Blob([buffer],{type: 'application/octet-stream'});
            callback(blob);
        });
    }


    /**
     * 创建带格式的Xlsx文件
     * @param columns 表格头部标题数据
     * @param tableData 表格数据
     * @param tableOtherData 表格外数据
     * @param sheetName 页签名称
     * @param fileName 文件名称（带后缀）
     */
    createFormatXlsxFile = (columns, tableData, tableOtherData, sheetName, fileName) => {
        // 获取并创建一个工作簿对象
        const workBookInfo = new ExcelJS.Workbook();
        // 设置工作簿创建时间
        workBookInfo.created = new Date();
        // 设置工作簿最后一次修改时间
        workBookInfo.modified = new Date();
        // 新建一个指定名称的页签
        workBookInfo.addWorksheet(sheetName); // eslint-disable-line no-unused-vars
        // 获取指定名称的页签对象信息
        let sheet1 = workBookInfo.getWorksheet(sheetName);

        // 生成带固定格式的头部信息
        this.createHeadSheet(sheet1, tableOtherData);

        // 填充表格标题数据以及单元格样式（从第5行开始，单元格从第一个开始）
        this.createTableSheet(sheet1, 5, 1, columns);

        // 设置第一页页签中的表格数据列宽，如不设置，则按照表格头部数据长度的文字默认长度设置
        this.setColumnWidth(sheet1.getColumn(1), 5); // 序号
        this.setColumnWidth(sheet1.getColumn(2), 10); // 会员号
        this.setColumnWidth(sheet1.getColumn(3), 25); // 会员名称
        this.setColumnWidth(sheet1.getColumn(4), 25); // 账号
        this.setColumnWidth(sheet1.getColumn(5), 40); // 开户行
        this.setColumnWidth(sheet1.getColumn(6), 30); // 资金用途
        this.setColumnWidth(sheet1.getColumn(7), 20); // 转账金额
        this.setColumnWidth(sheet1.getColumn(8), 20); // 银行流水号
        this.setColumnWidth(sheet1.getColumn(9), 18); // 交易所流水号
        this.setColumnWidth(sheet1.getColumn(10), 12); // 日期
        this.setColumnWidth(sheet1.getColumn(11), 12); // 时间

        // 按照模板从第6行开始添加表格中数据
        let currentRow = 6;
        // 循环填充表格数据
        for (let i = 0; i < tableData.length; i++) {
            // 获取从第几行开始填充数据
            let sheetRow = sheet1.getRow(currentRow);
            // 开始填充数据
            sheetRow.getCell(1).value = tableData[i].xuhao;
            sheetRow.getCell(2).value = tableData[i].huiyuanhao;
            sheetRow.getCell(3).value = tableData[i].huiyuanmingcheng;
            sheetRow.getCell(4).value = tableData[i].zhanghao;
            sheetRow.getCell(5).value = tableData[i].kaihuhang;
            sheetRow.getCell(6).value = tableData[i].zijinyongtu;
            sheetRow.getCell(7).value = tableData[i].zhuanzhangjine;
            sheetRow.getCell(8).value = tableData[i].yinhangliushuihao;
            sheetRow.getCell(9).value = tableData[i].jiaoyisuoliushuihao;
            sheetRow.getCell(10).value = tableData[i].riqi;
            sheetRow.getCell(11).value = tableData[i].shijian;
            // 下一行
            currentRow++;
        }
        // 设置数据对应的每个单元格样式（从单元格1开始，到11结束）
        this.setSheetStyle(sheet1, 1, 11);

        // 生成带固定格式的尾部信息
        this.createBottomSheet(sheet1, tableOtherData, currentRow);

        // 保存并下载Excel文件
        this.saveExcel(workBookInfo, fileName);
    }


    /**
     * 导出事件
     */
    exportXlsx = (e) => {

        console.log('~开始导出~');
        //sheet1表格头部数据
        let columns = ['序号', '会员号', '会员名称', '账号', '开户行', '资金用途', '转账金额', '银行流水号', '交易所流水号', '日期', '时间'];

        // 表格外部数据
        let tableOtherData = Object();
        tableOtherData.headOrgName = '湖南'; // 头部银行名称
        tableOtherData.accountName = '湖南商品交易所'; // 户名
        tableOtherData.direction = '场内入金'; // 方向
        tableOtherData.ccy = '人民币'; // 币种
        tableOtherData.unit = '元'; // 单位
        tableOtherData.accountNo = '310069079013006460347'; // 账号
        tableOtherData.depositBank = '交通银行'; // 开户行
        tableOtherData.countAmt = '13.06'; // 合计金额
        tableOtherData.printDate = '2022年04月29日'; // 打印日期
        tableOtherData.teller = '3100271'; // 柜员号

        //表格数据
        let tableDatas = [];
        let tableData1 = Object();
        tableData1.xuhao = '1';
        tableData1.huiyuanhao = '0050';
        tableData1.huiyuanmingcheng = '周专用测试有限公司';
        tableData1.zhanghao = '310069079013006460347';
        tableData1.kaihuhang = '交通银行上海嘉定支行';
        tableData1.zijinyongtu = '场内入金';
        tableData1.zhuanzhangjine = '1.00';
        tableData1.yinhangliushuihao = 'EBD0000000512254';
        tableData1.jiaoyisuoliushuihao = '1414472810';
        tableData1.riqi = '20220424';
        tableData1.shijian = '14:52:35';
        tableDatas.push(tableData1);

        let tableData2 = Object();
        tableData2.xuhao = '2';
        tableData2.huiyuanhao = '0050';
        tableData2.huiyuanmingcheng = '周专用测试有限公司';
        tableData2.zhanghao = '310069079013006460347';
        tableData2.kaihuhang = '交通银行上海嘉定支行';
        tableData2.zijinyongtu = '入金';
        tableData2.zhuanzhangjine = '2.06';
        tableData2.yinhangliushuihao = 'EBD0000000512254';
        tableData2.jiaoyisuoliushuihao = '1518174646';
        tableData2.riqi = '20220424';
        tableData2.shijian = '15:17:55';
        tableDatas.push(tableData2);

        let tableData3 = Object();
        tableData3.xuhao = '3';
        tableData3.huiyuanhao = '0109';
        tableData3.huiyuanmingcheng = '永安期货股份有限公司';
        tableData3.zhanghao = '212060120013002597735';
        tableData3.kaihuhang = '交通银行大连商品交易所支行';
        tableData3.zijinyongtu = '强制换汇';
        tableData3.zhuanzhangjine = '10.00';
        tableData3.yinhangliushuihao = 'EBD0000000520040';
        tableData3.jiaoyisuoliushuihao = 'A030147216';
        tableData3.riqi = '20220424';
        tableData3.shijian = '15:18:36';
        tableDatas.push(tableData3);

        this.createFormatXlsxFile(columns, tableDatas, tableOtherData, '大连商品交易所支行对账表', '对账表');
        console.log('~导出完成~');
    }



    render() {
        return (
            <div>
                <h1>Hello Word</h1>
                <p>客户号：<input type='text' onBlur={this.cusOnBlur}/></p>
                <div>
                    <button onClick={this.exportXlsx}>导出模拟数据</button>
                </div>
            </div>
        );
    }

}

ReactDOM.render(<MainOne/>, document.getElementById('root'));