var excel2jsontemplate = (function() {
    function InitException(message) {
        this.message = message;
        this.name = 'InitException';
    }

    if (typeof require === 'undefined') {
        throw new InitException('require is undefined');
    }

    const fs = require('fs');
    const util = require('util');
    var XLSX = require('xlsx');
    var _ = require('underscore');
    var jsonfile = require('jsonfile');

    /**
     * Load the sheet by file path and sheet name.
     *
     * @method loadSheet
     *
     * @param  {string}  filePath
     * @param  {string}  sheet
     *
     * @return Object
     */
    function loadSheet(filePath, sheet) {
        var workbook = XLSX.readFile(filePath);
        if (!workbook) {
            return;
        }

        return workbook.Sheets[sheet];
    }

    /**
     * Load the sheet in excel and output the data with JSON format.
     *
     * @method toJson
     *
     * @param  {string} filePath
     * @param  {string} sheet
     *
     * @return {JSON}
     */
    function toJson(filePath, sheet) {
        var ws = loadSheet(filePath, sheet);
        return XLSX.utils.sheet_to_json(ws);
    }

    /**
     * The cell factory.
     *
     * @method factoryCell
     *
     * @param  {number} col
     * @param  {number} row
     * @param  {string} zVal
     * @param  {string} rawVal
     *
     * @return {Object} The cell Object
     */
    function factoryCell(cell_address, z_value, obj) {
        var cell = {
            o: obj,
            v: obj ? obj.v : null, // raw value
            w: obj ? obj.w : null, // formatted text (if applicable)
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
     * @param  {Array}  ignores
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
     * @param  {object} worksheet
     * @param  {object} range
     * @param  {string} titleChar
     *
     * @return {object} The cell of title.
     */
    function findTitleTagCell(worksheet, range, titleChar) {
        for (var ir = range.s.r; ir <= range.e.r; ++ir) {
            for (var ic = range.s.c; ic <= range.e.c; ++ic) {

                var cell_address = {
                    c: ic,
                    r: ir
                };
                z = XLSX.utils.encode_cell(cell_address);
                if (!isTag(worksheet[z], titleChar)) {
                    continue;
                }

                var cell = factoryCell(cell_address, z, worksheet[z]);
                return cell;
            }
        }

        return null;
    }

    /**
     * Find the cells of ignore by ignore char.
     *
     * @method findIgnoreTagCells
     *
     * @param  {object} worksheet
     * @param  {object} range
     * @param  {string} ignoreChar
     *
     * @return {Array}  The array of cell of ignores.
     */
    function findIgnoreTagCells(worksheet, range, ignoreChar) {
        var cells = [];
        for (var ir = range.s.r; ir <= range.e.r; ++ir) {
            for (var ic = range.s.c; ic <= range.e.c; ++ic) {

                var cell_address = {
                    c: ic,
                    r: ir
                };
                z = XLSX.utils.encode_cell(cell_address);
                if (!isTag(worksheet[z], ignoreChar)) {
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
     * @param  {object} worksheet
     * @param  {object} range
     * @param  {object} titleCell
     * @param  {object} ignoreCells
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
     * @param  {string} titleChar = '#'
     * @param  {string} ignoreChar = '!'
     *
     * @return {Array}  The raw data of array of objects.
     */
    function parseSheet(filePath, sheetName,
        titleChar = '#', ignoreChar = '!') {
        var ws = loadSheet(filePath, sheetName);
        if (!ws) {
            return null;
        }

        var ref = ws['!ref'];
        if (!ref) {
            return null;
        }

        var range = XLSX.utils.decode_range(ref);

        // check the title cell is exists or not
        var titleCell = findTitleTagCell(ws, range, titleChar);
        if (titleCell === null) {
            return null;
        }

        var ignoreCells = findIgnoreTagCells(ws, range, ignoreChar);
        var titles = findTitleCells(ws, range, titleCell, ignoreCells);
        if (titles.length === 0) {
            return null;
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

            if (_.isEmpty(data)) {
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
     * @param  {object} data
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
     * @param  {string} titleChar = '#'
     * @param  {string} ignoreChar = '!'
     *
     * @return {Object} The JSON Object of data
     */
    function parse(filePath, sheetName, template,
        titleChar = '#', ignoreChar = '!') {
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
        if (!_.isObject(template) || _.isArray(template)) {
            throw new Error("The JSON template is invaild.");
        }

        var rawDatas = parseSheet(filePath, sheetName, titleChar, ignoreChar);
        rawDatas = rawDatas || [];

        var datas = _.isEmpty(template) ?
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
     * @param  {object} [options] Set replacer for a JSON replacer.
     * @param  {Function}   callback
     *
     * @return
     */
    function save(filePath, jsonObj, options, callback) {
        jsonfile.writeFile(filePath, jsonObj, options, function(err) {
            if (callback) {
                return callback(err, null);
            }
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
    function load(filePath, callback) {
        jsonfile.readFile(filePath, function(err, obj) {
            if (callback) {
                return callback(err, obj);
            }
        });
    }

    return {
        save: save,
        load: load,
        parse: parse,
    };
}());

module.exports = excel2jsontemplate;
