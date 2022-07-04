# SpreadJS_EchartsPDF
在纯前端在线表格中实现集成 Echarts 并导出 PDF功能
# SpreadJS_EchartsPDF

#### 介绍
在纯前端在线表格中实现集成 Echarts 并导出 PDF功能


### SpreadJS 示例，SpreadJS集成 Echarts 并导出 PDF
该示例包括使用 SpreadJS API 的演示脚本，可用于实现SpreadJS集成 Echarts 并导出 PDF。
有关 SpreadJS API 的更多信息，请参阅[SpreadJS API指南]( https://demo.grapecity.com.cn/spreadjs/help/api/) 和[帮助手册]( https://help.grapecity.com.cn/pages/viewpage.action?pageId=5963808)。



### 运行步骤
1、在开始之前，请确保您已满足以下先决条件：
要运行 SpreadJS，浏览器必须支持 HTML5，客户端导入和导出 Excel 需要 IE10及以上。
请先了解 [SpreadJS 的产品使用环境]( https://www.grapecity.com.cn/developer/spreadjs/selection-guide/product-use-environment)，并申请临时部署授权激活
安装并更新NodeJS和NPM
2、克隆或下载此代码库
3、初始化控件，并运行示例脚本
#### 控件初始化
首先，创建一个新页面，并在页面上输入以下代码：
```
<!DOCTYPE html>
    <html>
    <head>
        <title>SpreadJS HTML Test Page</title>
```
2、在页面中添加对 SpreadJS 的引用。代码如下。需要注意的是，SpreadJS 提供压缩过
```
//（minified）的 JavaScript 文件和和用于调试的文件：
<script src="[Your_Scripts_Path]/gc.spread.sheets.all.xxxx.min.js" type="text/javascript"></script>
```
3、添加 CSS 文件以改变Spread.JS 的外观。默认的CSS文件名为： 
gc.spread.sheets.xxxx.css，里面包含了所有的默认样式。该 CSS 文件将会影响滚动条，筛选框及其子元素，单元格和下方标签栏的样式。引入 CSS 的代码如下：
```
//<link href="[Your_CSS_Path]/gc.spread.sheets.xxxx.css" rel="stylesheet" type="text/css"/>
```
4、添加产品授权，代码为（本地测试可以不添加）：
```
GC.Spread.Sheets.LicenseKey = "xxx";
```
5. 添加控件初始化代码。本例会在一个 id 为 “ss” 的 DOM 元素上初始化 SpreadJS：
```
<script type="text/javascript">
// Add your license
// If run this in local for testing, remove or comment below code
 GC.Spread.Sheets.LicenseKey = "xxx";

// Add your code
 window.onload = function(){
var spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"),{sheetCount:3});
var sheet = spread.getActiveSheet();
 }
</script>
</head>
<body>
```
6、 创建一个 id 为 “ss” 的元素，SpreadJS 将在该 DOM 中初始化：
```
<div id="ss" style="height: 500px; width: 800px"></div>
</body>
</html>
```
#### 示例代码
```
HTML：
<p>在SpreadJS中集成Ecahrts</p>
<input id="exportPDF" value="导出PDF" type="button" />
<div id='ss'></div>


CSS：
#ss {
    height: 600px;
    width: 100%
}
input{
    margin-bottom: 10px;
}

p{
    color: #336699;
    text-align: center;
    margin: 12px 0;
}

JavaScript：
/**
* 需要引入三方库： https://lib.baomitu.com/echarts/5.1.0/echarts.simple.js
* 点击外部资源，在展开的输入框中输入三方库地址，之后点击+号添加链接
* 点击运行，看看效果吧
**/

/**
* 导出PDF是因为字体缺失的问题，会出现乱码问题，我们提供了相关的解决方案：
* https://www.grapecity.com.cn/blogs/spreadjs-solve-font-garbled-problem
**/



GC.Spread.Common.CultureManager.culture('zh-cn');

// chart数据
var charts = {
    "bar": {
        id: "barChart",
        tableName: "bar",
        source: [{
            衬衫: 5,
            羊毛衫: 20,
            雪纺衫: 36,
            裤子: 10,
            高跟鞋: 10,
            袜子: 20
        }],
        startRow: 1,
        endRow: 15,
        startColumn: 8,
        endColumn: 16,
        echart: false
    },
    "pie": {
        id: "pieChart",
        tableName: "pie",
        source: [{
            直接访问: 335,
            邮件营销: 310,
            联盟广告: 234,
            视频广告: 135,
            搜索引擎: 1548
        }],
        startRow: 16,
        endRow: 30,
        startColumn: 8,
        endColumn: 16,
        echart: false
    },
    "line": {
        id: "lineChart",
        tableName: "line",
        source: [{
            Mon: 820,
            Tue: 932,
            Wed: 901,
            Thu: 934,
            Fri: 1290,
            Sat: 1330,
            Sun: 1320
        }],
        startRow: 31,
        endRow: 45,
        startColumn: 8,
        endColumn: 16,
        echart: false
    }

};

$(document).ready(function() {

    /*
			初始化Spread
        */
    var spread = new GC.Spread.Sheets.Workbook(document.getElementById("ss"));
    var tempSpread = new GC.Spread.Sheets.Workbook();
    var sheet = spread.getSheet(0);
    sheet.suspendPaint();

    var defaultShowRows = 600 / 20;

    for (var chart in charts) {
        // 初始化数据表格
        initDataTable(sheet, charts[chart]);
        // 初始化浮动对象
        initFloatingObject(sheet, charts[chart]);

    }

    sheet.resumePaint();

    for (var chart in charts) {
        // 初始化图表
        initCharts(charts[chart]);
    }

    spread.bind(GC.Spread.Sheets.Events.ValueChanged, function(s, e) {

        var row = e.row;
        var col = e.col;

        for (var chart in charts) {
            var range = new GC.Spread.Sheets.Range(charts[chart].table.row, charts[chart].table.col, charts[chart].table.rowCount, charts[chart].table.colCount);
            if (range.contains(row, col, 1, 1)) {
                refreshCharts(charts[chart].id, getChartDataFromTables(charts[chart].source));
                break;
            }
        }

    });

    // TopRowChanged 随事件加载
    /*
     *   SpreadJS 出于性能需要，采用了lazyload机制，因此
     *   需要在事件中进行判断，当滚动条滚动到floating object 所在位置时
     *   再加载ECharts图像。
     * */

    // 判断：当存在未加载的charts时，注册事件
    spread.bind(GC.Spread.Sheets.Events.TopRowChanged, function(s, e) {

        console.log(e);
        var newTopRow = e.newTopRow;

        if ((charts["bar"].startRow - defaultShowRows < newTopRow) && (!charts["bar"].echart)) {
            initCharts(charts["bar"]);
        }

        if ((charts["pie"].startRow - defaultShowRows < newTopRow) && (!charts["pie"].echart)) {
            initCharts(charts["pie"]);
        }

        if ((charts["line"].startRow - defaultShowRows < newTopRow) && (!charts["line"].echart)) {
            initCharts(charts["line"]);
        }

    });



    $("#exportPDF").click(function() {

        //深拷贝
        tempSpread.fromJSON(JSON.parse(JSON.stringify(spread.toJSON({
            includeBindingSource: true
        }))));
        tempSpread.suspendPaint();
        tempSheet = tempSpread.getSheet(0);

        for (var chart in charts) {
            //遍历sheet删除多余floatObject
            tempSheet.floatingObjects.remove(charts[chart].id);
            //删除floatObject对应位置替换picture
            if (!charts[chart].echart) {
                sheet.showCell(charts[chart].startRow, charts[chart].startColumn, GC.Spread.Sheets.VerticalPosition.top, GC.Spread.Sheets.HorizontalPosition.left);
                initCharts(charts[chart]);
            }
            var img = charts[chart].echart.getDataURL();
            var picture = tempSheet.pictures.add(charts[chart].id, img, 0, 0, 100, 100);
            picture.startRow(charts[chart].startRow);
            picture.startColumn(charts[chart].startColumn);
            picture.endColumn(charts[chart].endColumn);
            picture.endRow(charts[chart].endRow);
        }
        tempSpread.resumePaint();

        //注册导出PDF fallback字体为宋体
        /*GC.Spread.Sheets.PDF.PDFFontsManager.fallbackFont = function (font) {
                     return fontsObj["simkai.ttf"];
                 }*/

        tempSpread.savePDF(
            function(blob) {
                saveAs(blob, 'download.pdf');
            },
            function(error) {
                console.log(error);
            }, {
                title: 'Test Title',
                author: 'Test Author',
                subject: 'Test Subject',
                keywords: 'Test Keywords',
                creator: 'test Creator'
            }
        );

    });



});

function initCharts(chart) {

    switch (chart.id) {
        case "barChart":
            chart.echart = initBarECharts(chart);
            break;
        case "pieChart":
            chart.echart = initPieECharts(chart);
            break;
        case "lineChart":
            chart.echart = initLineECharts(chart);
            break;
    }

}

function initDataTable(sheet, chart) {
    //id, startRow, source
    sheet.setColumnWidth(1, 90);
    sheet.setColumnWidth(2, 90);
    sheet.setColumnWidth(3, 90);
    sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 90);
    sheet.setColumnWidth(6, 90);
    sheet.setColumnWidth(7, 90);

    var chartTable = sheet.tables.addFromDataSource(chart.tableName, chart.startRow + 1, 1, chart.source, GC.Spread.Sheets.Tables.TableThemes.medium2);

    var chartTableRange = chartTable.dataRange();

    var table = {}
    table.row = chartTableRange.row;
    table.rowCount = chartTableRange.rowCount;
    table.col = chartTableRange.col;
    table.colCount = chartTableRange.colCount;
    chart.table = table;



}

function getChartDataFromTables(tableSource) {
    var categoriesArr = [];
    var dataArr = [];
    for (var prop in tableSource[0]) {
        categoriesArr.push(prop);
        dataArr.push(tableSource[0][prop]);
    }
    var barData = {
        categories: categoriesArr,
        data: dataArr
    };
    return barData;
}

function refreshCharts(id, data) {
    var myChart = echarts.getInstanceByDom(document.getElementById(id));
    if (myChart) {
        switch (id) {
            case "barChart":
                myChart.setOption({
                    xAxis: {
                        data: data.categories
                    },
                    series: [{
                        data: data.data
                    }]
                });
                break;
            case "pieChart":
                var dataArr = [];
                for (var i = 0; i < data.categories.length; i++) {
                    dataArr.push({
                        value: data.data[i],
                        name: data.categories[i]
                    });
                }
                myChart.setOption({
                    legend: {
                        data: data.categories
                    },
                    series: [{
                        data: dataArr
                    }]
                });
                break;
            case "lineChart":
                myChart.setOption({
                    xAxis: {
                        data: data.categories
                    },
                    series: [{
                        data: data.data
                    }]
                });
                break;
        }
    }
}


function initFloatingObject(sheet, chart) {

    // 初始化浮动对象
    var customFloatingObject = new GC.Spread.Sheets.FloatingObjects.FloatingObject(chart.id);
    customFloatingObject.startRow(chart.startRow);
    customFloatingObject.startColumn(chart.startColumn);
    customFloatingObject.endColumn(chart.endColumn);
    customFloatingObject.endRow(chart.endRow);

    // 创建ECharts容器
    var div = document.createElement('div');
    div.innerHTML = '<div id="' + chart.id + '" style="width: 500px;height:300px; "></div>';
    $(div).css({
        background: "#FFFFFF"
    });
    // 将ECharts添加到浮动层中
    customFloatingObject.content(div);
    sheet.floatingObjects.add(customFloatingObject);

}

function initBarECharts(chart) {

    var div = document.getElementById(chart.id);

    if (!div) {
        return;
    }

    var dataObj = getChartDataFromTables(chart.source);

    // 基于准备好的dom，初始化echarts实例
    var myChart = echarts.init(div);

    // 指定图表的配置项和数据
    var option = {
        title: {
            text: 'ECharts 入门示例'
        },
        tooltip: {},
        legend: {
            data: ['销量']
        },
        xAxis: {
            data: dataObj.categories
        },
        yAxis: {},
        series: [{
            name: '销量',
            type: 'bar',
            data: dataObj.data,
            animation: false
        }]
    };

    // 使用刚指定的配置项和数据显示图表。
    myChart.setOption(option);

    //EChartsArr.bar.chart = myChart;

    return myChart;
}

function initPieECharts(chart) {

    var div = document.getElementById(chart.id);

    if (!div) {
        return;
    }

    var dataObj = getChartDataFromTables(chart.source);
    var dataArr = [];
    for (var i = 0; i < dataObj.categories.length; i++) {
        dataArr.push({
            value: dataObj.data[i],
            name: dataObj.categories[i]
        });
    }

    var myChart = echarts.init(div);

    var option = {
        tooltip: {
            trigger: 'item',
            formatter: "{a} <br/>{b}: {c} ({d}%)"
        },
        legend: {
            orient: 'vertical',
            x: 'left',
            data: dataObj.categories
        },
        series: [{
            name: '访问来源',
            type: 'pie',
            radius: ['50%', '70%'],
            avoidLabelOverlap: false,
            label: {
                normal: {
                    show: false,
                    position: 'center'
                },
                emphasis: {
                    show: true,
                    textStyle: {
                        fontSize: '30',
                        fontWeight: 'bold'
                    }
                }
            },
            labelLine: {
                normal: {
                    show: false
                }
            },
            data: dataArr,
            animation: false
        }]
    };

    myChart.setOption(option);

    //EChartsArr.pie.chart = myChart;

    return myChart;
}

function initLineECharts(chart) {

    var div = document.getElementById(chart.id);

    if (!div) {
        return;
    }

    var dataObj = getChartDataFromTables(chart.source);

    var myChart = echarts.init(div);

    var option = {
        xAxis: {
            type: 'category',
            data: dataObj.categories
        },
        yAxis: {
            type: 'value'
        },
        series: [{
            data: dataObj.data,
            type: 'line',
            animation: false
        }]
    };

    myChart.setOption(option);

    //EChartsArr.line.chart = myChart;

    return myChart;
}
```
#### 关于 SpreadJS
[SpreadJS]( https://www.grapecity.com.cn/developer/spreadjs) 是一款基于 HTML5 的纯前端表格控件，兼容 450 多种 Excel 公式，具备“高性能、跨平台、与 Excel 高度兼容”的产品特性。使用 SpreadJS，可直接在 Angular、 React、 Vue 等前端框架中实现高效的模板设计、在线编辑和数据绑定等功能，为最终用户提供高度类似 Excel 的使用体验。




