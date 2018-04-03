
window.onload = function () {

    var file = $('#file')[0];
    var $drop = $('#drop');
    var view = $('#view')[0];
    // var $export = $('#export');
    var $keySelect = $('#filter-key');
    var $valueSelect = $('#filter-value');
    var $nameSelect = $('#filter-name');
    var $handleChart = $('#handleChart');
    var $chartType = $('#filter-charttype');
    var $fileList = $('#file-list');

    var excelView = $('#excel-view')[0];
    var chart = echarts.init(excelView);
    var chooseKey;
    var chooseValue;
    var gdata = [];
    var chartType = 'pie';
    var products = [];



    /**
     * 读取
     */
    function readExcelFile(filedata, cb) {
        // https://developer.mozilla.org/zh-CN/docs/Web/API/FileReader
        var reader = new FileReader();

        var types = [
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        ];


        if (types.indexOf(filedata.type) === -1) {
            return alert('文件类型不是 Excel 格式');
        }

        reader.readAsBinaryString(filedata);
        reader.onload = function (e) {
            // 解析数据
            var bstr = e.target.result;
            var wb = XLSX.read(bstr, { type: 'binary', cellStyles: true });
            console.log(wb);

            // 只取第一个 sheet
            // var wsname = wb.SheetNames[0];
            // var ws = wb.Sheets[wsname];
            // 渲染
            typeof cb === 'function' && cb(wb);
            $fileList.html(`<div class="file-item"><i class="excel"></i><span>${filedata.name}</span></div>`);
        };
    }

    /**
     * 渲染数据
     */

    var transfer = {
        toJson: function (ws) {
            return XLSX.utils.sheet_to_json(ws);
        },

        toHTML(ws) {
            return XLSX.utils.sheet_to_html(ws);
        }
    };

    var render = function(filedata, type, cb) {
        var type = type || 'toJson';
        readExcelFile(filedata, function (wb) {
            // 只取第一个 sheet
            var wsname = wb.SheetNames[0];
            var ws = wb.Sheets[wsname];
            // 渲染
            cb && cb(transfer[type](ws));
        })
    };

    var handleChart = function(data) {
        console.log(data);
        var result = [];
        data.forEach(function(i) {
            var o = {};
            o.value = i[chooseValue];
            o.name = i[chooseKey];
            if (products.indexOf(o.name) > -1) {
                result.push(o);
            }
        });
        var option = {
            series : [
                {
                    name: '访问来源',
                    type: chartType,
                    radius: '55%',
                    data: result,
                    // 高亮样式。
                    emphasis: {
                        itemStyle: {
                        },
                        textStyle: {
                           fontSize: 20 // 用 legend.textStyle.fontSize 更改示例大小
                        },
                        label: {
                            show: true,
                            fontSize: 20,
                            formatter: '{b}\n{c}\n ({d}%)'
                        }
                    }
                }

            ]
        };

        // 使用刚指定的配置项和数据显示图表。
        chart.setOption(option);
    };

    var initFilter = function(data) {
        var str = '<option></option>';
        for (var i in data[0]) {
            str += `<option value="${i}">${i}</option>`;
        }
        $keySelect.html(str);
        $valueSelect.html(str);
    };

    var initNames = function() {
        var str = ''
        gdata[0].forEach(function(i) {
            str += `<option value="${i[chooseKey]}">${i[chooseKey]}</option>`
        })
        $nameSelect.html(str);
        $nameSelect.select2({
            placeholder:'请选择',
            placeholderOption: "first",
            allowClear: true,
        });
    };

    /**
     * 上传的文件
     */
    file.onchange = function (event) {
        var files = event.target.files;

        if (files && files[0]) {
            render(files[0], 'toJson', function(data) {
                gdata.unshift(data);
                initFilter(gdata[0]);
            });
        }
    }

    /**
     * 拖拽上传
     * https://developer.mozilla.org/zh-CN/docs/Web/Events/drop
     */
    $drop.on('drop', function (event) {
        event.stopPropagation();
        event.preventDefault();
        var files = event.dataTransfer.files;

        if (files && files[0]) {
            render(files[0], 'toJson', function(data) {
                gdata.unshift(data);
                initFilter(gdata[0]);
            });
            $drop.removeClass('active');
            $drop.text('把 Excel 文件拖动到这个区域！');
        }
    }, false);

    function dragover(event) {
        event.stopPropagation();
        event.preventDefault();
        event.dataTransfer.dropEffect = 'copy';

        if (!$drop.hasClass('active')) {
            $drop.addClass('active');
            $drop.text('松开吧！');
        }
    }

    $drop.on('dragenter', dragover, false);
    $drop.on('dragover', dragover, false);
    $drop.on('dragleave', function (event) {
        $drop.removeClass('active');
        $drop.text('把 Excel 文件拖拽到这个区域里！');
    }, false);

    $keySelect.on('change', function() {
        chooseKey = $keySelect.val();
        initNames();
    });
    $valueSelect.on('change', function() {
        chooseValue = $valueSelect.val();
    });

    $handleChart.on('click', function() {
        handleChart(gdata[0]);
    });

    $nameSelect.on('change', function() {
        products = $nameSelect.val();
    });

    $chartType.on('change', function() {
        chartType = $chartType.val();
        handleChart(gdata[0]);
    });

}

