/**
 * TODO:
 * - 引入 mochajs 增加 Unit test
 * - 增加 tag: attribute(^), 讓使用者可以在 JSON 的 top-level 增加 attribute
 * - 增加 CLI 方法, 傳入檔案列表, 依列表依序 parse 各個資料
 * - 將 tags 改為使用 Objec 傳入 { tagTitle: '!', tagIgnore: '#', tagAttr: '^' }
 * - 改用 Asynchronous 的方式去寫, 能的話引入 async
 * - 錯誤的地方使用明確的 exception, 而非回傳 null
 */

var excel2jsontemplate = (function() {
    function InitException(message) {
        this.message = message;
        this.name = 'InitException';
    }

    if (typeof require === 'undefined') {
        throw new InitException('require is undefined');
    }

    var path = require('path');
    var fs = require('fs');
    var util = require('util');
    var XLSX = require('xlsx');
    var jsonfile = require('jsonfile');
    var _u = require('underscore');
    var _l = require('lodash');

    /**
     * Load the sheet by file path and sheet name.
     *
     * @method loadSheet
     *
     * @param  {string}  filePath
     * @param  {string}  sheetName
     *
     * @return Object
     */
    function loadSheet(filePath, sheetName) {
        if (_u.isEmpty(filePath) || _u.isEmpty(sheetName)) {
            return;
        }

        try {
            var workbook = XLSX.readFile(filePath);
            if (!workbook) {
                return;
            }
            if (_l.indexOf(workbook.SheetNames, sheetName) === -1) {
                return;
            }

            return workbook.Sheets[sheetName];
        } catch (e) {
            return;
        }
    }

    /**
     * The cell factory.
     *
     * @method factoryCell
     *
     * @param  {Object} cell_address [{ c: column, r: row }]
     * @param  {string} z_value [number format string associated with the cell]
     * @param  {Object} rawCell [Raw cell object]
     *
     * @return {Object} The cell Object
     */
    function factoryCell(cell_address, z_value, rawCell) {
        var cell = {
            o: rawCell,
            v: rawCell ? rawCell.v : null, // raw value
            w: rawCell ? rawCell.w : null, // formatted text (if applicable)
            c: cell_address.c,
            r: cell_address.r,
            z: z_value,
        };

        return cell;
    }

    /**
     * Check the cell is need ignore or not.
     * <pre>
     * Rules:
     * 1. Same column with ignore cell.
     * 2. If the column of ignore cell is zero, the same row with ignore cell will be ignored.
     * </pre>
     *
     * @method isNeedIgnore
     *
     * @param  {Object} cell
     * @param  {Array}  ignore cells
     *
     * @return {Boolean}
     */
    function isNeedIgnore(cell, ignores) {
        for (var k in ignores) {
            if (!ignores.hasOwnProperty(k)) {
                continue;
            }

            var fc = ignores[k];
            if (fc.c == cell.c) {
                return true;
            }

            if (fc.c === 0 && fc.r == cell.r) {
                return true;
            }
        }

        return false;
    }

    /**
     * Check is tag or not.
     *
     * @method isTag
     *
     * @param  {Object} cell
     * @param  {string} tagChar
     *
     * @return {Boolean}
     */
    function isTag(cell, tagChar) {
        if (!cell || !cell.v) {
            return false;
        }

        if (typeof cell.v !== 'string') {
            return false;
        }

        if (cell.v != tagChar) {
            return false;
        }

        return true;
    }

    /**
     * Find the cell object of title tag by title char.
     *
     * @method findTitleTagCell
     *
     * @param  {Object} worksheet
     * @param  {Object} range
     * @param  {string} tagTitle
     *
     * @return {Object} The cell of title.
     */
    function findTitleTagCell(worksheet, range, tagTitle) {
        for (var ir = range.s.r; ir <= range.e.r; ++ir) {
            for (var ic = range.s.c; ic <= range.e.c; ++ic) {

                var cell_address = {
                    c: ic,
                    r: ir
                };
                z = XLSX.utils.encode_cell(cell_address);
                if (!isTag(worksheet[z], tagTitle)) {
                    continue;
                }

                var cell = factoryCell(cell_address, z, worksheet[z]);
                return cell;
            }
        }

        return;
    }

    /**
     * Find the cells of ignore by ignore char.
     *
     * @method findIgnoreTagCells
     *
     * @param  {Object} worksheet
     * @param  {Object} range
     * @param  {string} tagIgnore
     *
     * @return {Array}  The array of cell of ignores.
     */
    function findIgnoreTagCells(worksheet, range, tagIgnore) {
        var cells = [];
        for (var ir = range.s.r; ir <= range.e.r; ++ir) {
            for (var ic = range.s.c; ic <= range.e.c; ++ic) {

                var cell_address = {
                    c: ic,
                    r: ir
                };
                z = XLSX.utils.encode_cell(cell_address);
                if (!isTag(worksheet[z], tagIgnore)) {
                    continue;
                }

                var cell = factoryCell(cell_address, z, worksheet[z]);
                cells.push(cell);
            }
        }

        return cells;
    }

    /**
     * Find the cells of title.
     *
     * @method findTitleCells
     *
     * @param  {Object} worksheet
     * @param  {Object} range
     * @param  {Object} titleCell
     * @param  {Object} ignoreCells
     *
     * @return {Array}  The cells of title
     */
    function findTitleCells(worksheet, range, titleCell, ignoreCells) {
        var cells = [];
        var ir = titleCell.r;
        for (var ic = range.s.c; ic <= range.e.c; ++ic) {
            var cell_address = {
                c: ic,
                r: ir
            };
            z = XLSX.utils.encode_cell(cell_address);
            if (z == titleCell.z) {
                continue;
            }

            var cell = factoryCell(cell_address, z, worksheet[z]);

            // check the cell is need to ignore or not
            if (isNeedIgnore(cell, ignoreCells)) {
                continue;
            }

            if (titleCell && cell.v && (cell.r == titleCell.r)) {
                cells[ic] = cell;
            }
        }

        return cells;
    }

    /**
     * Parse the excel file and export the data with JSON format.
     *
     * @method parseSheet
     *
     * @param  {string} filePath
     * @param  {string} sheetName
     * @param  {string} tagTitle = '#'
     * @param  {string} tagIgnore = '!'
     *
     * @return {Array}  The raw data of array of objects.
     */
    function parseSheet(filePath, sheetName,
        tagTitle = '#', tagIgnore = '!') {
        var ws = loadSheet(filePath, sheetName);
        if (!ws) {
            return;
        }

        var ref = ws['!ref'];
        if (!ref) {
            return;
        }

        var range = XLSX.utils.decode_range(ref);

        // check the title cell is exists or not
        var titleCell = findTitleTagCell(ws, range, tagTitle);
        if (!titleCell) {
            return;
        }

        var ignoreCells = findIgnoreTagCells(ws, range, tagIgnore);
        var titles = findTitleCells(ws, range, titleCell, ignoreCells);
        if (titles.length === 0) {
            return;
        }

        var cells = [];
        var rawDatas = [];
        for (var ir = range.s.r; ir <= range.e.r; ++ir) {
            var data = {};
            for (var ic = range.s.c; ic <= range.e.c; ++ic) {
                var cell_address = {
                    c: ic,
                    r: ir
                };
                z = XLSX.utils.encode_cell(cell_address);
                var cell = factoryCell(cell_address, z, ws[z]);

                // title row
                if (cell.r == titleCell.r) {
                    continue;
                }

                // check the cell is need to ignore or not
                if (isNeedIgnore(cell, ignoreCells)) {
                    continue;
                }

                var t = titles[ic];
                if (!t || !t.w) {
                    continue;
                }

                // save the data
                data[t.w] = cell.w;
                cells.push(cell);
            }

            if (_u.isEmpty(data)) {
                continue;
            }

            rawDatas.push(data);
        }

        return rawDatas;
    }

    /**
     * Map the data to tempalte.
     *
     * @method map
     *
     * @param  {Object} data
     * @param  {JSON}   template
     *
     * @return {JSON} Mapped data.
     */
    function map(data, template) {
        var obj = {};
        for (var key in template) {
            if (!template.hasOwnProperty(key)) {
                continue;
            }

            var val = template[key];
            switch (val.constructor) {
                case Array:
                    obj[key] = [];
                    for (var index in val) {
                        if (val.hasOwnProperty(index)) {
                            obj[key].push(data[val[index]]);
                        }
                    }
                    break;

                case Object:
                    obj[key] = map(data, val);
                    break;

                default:
                    obj[key] = data[val];
                    break;
            }
        }

        return obj;
    }

    /**
     * Transform the raw data to the template JSON object.
     *
     * @method transform
     *
     * @param  {Array}  rawData
     * @param  {JSON}   template
     *
     * @return {Array}  Array of JSON objects with template.
     */
    function transform(rawData, template) {
        // init result object
        var result = [];

        // foreach data to mapping to template
        for (var index in rawData) {
            if (!rawData.hasOwnProperty(index)) {
                continue;
            }

            // mapping
            var obj = map(rawData[index], template);
            result.push(obj);
        }

        return result;
    }

    /**
     * Parse the excel file and export the data with JSON format.
     *
     * @method parse
     *
     * @param  {string} filePath
     * @param  {string} sheetName
     * @param  {JSON} template
     * @param  {string} tagTitle = '#'
     * @param  {string} tagIgnore = '!'
     *
     * @return {Object} The JSON Object of data
     */
    function parse(filePath, sheetName, template,
        tagTitle = '#', tagIgnore = '!') {
        // check file is exists or not
        if (!fs.existsSync(filePath)) {
            throw new Error(util.format("The file %s is not exists.", filePath));
        }

        // check sheet name is empty or not
        if (!sheetName) {
            throw new Error("The sheet name is null or empty.");
        }

        // check template
        template = template || {};
        if (!_u.isObject(template) || _u.isArray(template)) {
            throw new Error("The JSON template is invaild.");
        }

        // get raw data
        var rawDatas = parseSheet(filePath, sheetName,
            tagTitle, tagIgnore) || [];

        // transform the data if necessary
        var datas = _u.isEmpty(template) ?
            rawDatas :
            transform(rawDatas, template);

        return {
            name: sheetName,
            datas: datas,
        };
    }

    /**
     * Save the JSON to the file.
     *
     * @method save
     *
     * @param  {string} filePath
     * @param  {string} jsonObj
     * @param  {Object} [options] Set replacer for a JSON replacer.
     * @param  {Function}   callback
     *
     * @return
     */
    function save(filePath, jsonObj, options, callback) {
        if (!callback) {
            callback = options;
            options = {};
        }

        // check the output directory is exists or not, if not create it.
        var dir = path.dirname(filePath);
        fs.stat(dir, function(err, data) {
            if (err) {
                fs.mkdirSync(dir);
            }

            jsonfile.writeFile(filePath, jsonObj, options, function(err) {
                if (callback) {
                    return callback(err);
                }
            });
        });
    }

    /**
     * Load json file.
     *
     * @method load
     *
     * @param  {string} filePath
     * @param  {Function}   callback
     *
     * @return
     */
    function loadTemplate(filePath, callback) {
        jsonfile.readFile(filePath, function(err, obj) {
            if (callback) {
                return callback(err, obj);
            }
        });
    }

    return {
        save: save,
        loadTemplate: loadTemplate,
        parse: parse,

        loadSheet: loadSheet,
        factoryCell: factoryCell,
        isNeedIgnore: isNeedIgnore,
        isTag: isTag,
        findTitleTagCell: findTitleTagCell,
        findIgnoreTagCells: findIgnoreTagCells,
        findTitleCells: findTitleCells,
        parseSheet: parseSheet,
        map: map,
        transform: transform,
    };
}());

module.exports = excel2jsontemplate;
