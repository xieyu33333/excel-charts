
window.onload = function () {

    var file = $('#file')[0];
    var $drop = $('#drop');
    var view = $('#view')[0];
    var $export = $('#export');
    var $fileList = $('#file-list');

    var excelView = $('#excel-view')[0];
    var gdata = [];
    var fileList = [];
    var products = [];
    var cources = {
        '数学': /数学/,
        '英语': /英语/,
        '政治': /政治/,
        '翻译': /翻译/,
        '经济': /经济/,
        '金融': /金融/,
        '计算机': /数学/,
        '汉语': /汉语/,
        '西医': /西医/,
        '法律': /法律|法学/,
        '法律': /法律/,
        '心理': /心理/,
        '艺术类': /音乐|美术|艺术/,
        '管理': /管理/,
        '会计': /会计/,
        '日语': /日语/,
    }


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
            fileList.unshift(filedata.name);
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

    /**
     * 上传的文件
     */
    file.onchange = function (event) {
        var files = event.target.files;

        if (files && files[0]) {
            render(files[0], 'toJson', function(data) {
                gdata.unshift(data);
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

    // --------- export ---------//

    // 把 string 转为 ArrayBuffer
    function s2ab(str) {
        var buf = new ArrayBuffer(str.length);
        var _view = new Uint8Array(buf);
        for (var i = 0, len = str.length; i < len; i++) {
            _view[i] = str.charCodeAt(i) & 0xFF;
        }
        return buf;
    }

    // 根据表格内容，生成 Excel 文件
    $export.on('click', function (event) {
        var edata = JSON.parse(JSON.stringify(gdata[0]));
        edata.forEach(function(i) {
            var name = i['产品名称']
            if (name.match(/直通车/)) {
                i['产品类型'] = '直通车';
            }
            else if (name.match(/全程班/)) {
                i['产品类型'] = '全程班';
            }
            else {
                i['产品类型'] = '其他';
            }

            for (item in cources) {
                if (name.match(cources[item])) {
                    i['学科'] = item;
                    break;
                }
                else {
                    i['学科'] = '';
                }
            }
        })
        console.log(edata);
        var sheet = XLSX.utils.json_to_sheet(edata);

        sheet['A1'] = Object.assign(sheet['A1'], {
            // 样式？
            s: {
                fill: {
                    fgColor: { rgb: "FFFF0000" }
                }
            },
        });

        var wb = XLSX.utils.book_new({ cellStyles: true });
        XLSX.utils.book_append_sheet(wb, sheet, "SheetJS");
        // 渲染
        var wbout = XLSX.write(wb, { type: "binary", bookType: "xlsx" });
        // 保存 - https://github.com/eligrey/FileSaver.js
        saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }),  fileList[0].replace(/\.\w+/, '') + "-new.xlsx");

    });

}

