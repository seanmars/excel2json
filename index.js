/**
 * TODO:
 * - Check the value of title cell is valid or not (only alphabet)
 * - 增加 CLI 方法, 傳入檔案列表, 依列表依序 parse 各個資料
 * - 將 tags 改為使用 Objec 傳入 { tagTitle: '!', tagIgnore: '#', tagAttr: '^' }
 * - 改用 Asynchronous 的方式去寫, 能的話引入 async
 * - 錯誤的地方使用明確的 exception, 而非回傳 null
 * Refactoring:
 * 	- 修改 function A(x,y) call function B(y), 但卻在 A, B 裡都詳細檢查判斷了 y;
 * 	修正成為只需要在 function B(y) 內檢查就好, A 只需要判斷回傳出來的值
 */

var excel2jsontemplate = (function () {
    const DEFAULT_ATTR_SHEETNAME = 'sheetname';
    const DEFAULT_ATTR_DATAS = 'datas';

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
     * @param {string} filePath
     * @param {string} sheetName
     *
     * @return {Object} sheet Object of sheet.
     */
    function loadSheet(filePath, sheetName) {
        if (_u.isEmpty(filePath) || _u.isEmpty(sheetName)) {
            return;
        }

        try {
            // check file is exists or not
            if (!fs.existsSync(filePath)) {
                throw new Error('The file is not exists.');
            }

            var workbook = XLSX.readFile(filePath);
            if (!workbook) {
                return;
            }
            if (_l.indexOf(workbook.SheetNames, sheetName) === -1) {
                return;
            }

            return workbook.Sheets[sheetName];
        } catch (e) {
            throw e;
        }
    }

    /**
     * The cell factory.
     *
     * @method factoryCell
     *
     * @param {Object} cell_address { c: column, r: row }
     * @param {string} z_value number format string associated with the cell
     * @param {Object} rawCell Raw cell object
     *
     * @return {Object} cell The cell Object.
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
     * @param {Object} cell
     * @param {Array} ignore cells
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
     * @param {Object} cell
     * @param {string} tagChar
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
     * @param {Object} worksheet
     * @param {Object} range
     * @param {string} tagTitle
     *
     * @return {Object} cell The cell of title.
     */
    function findTitleTagCell(worksheet, range, tagTitle) {
        if (!tagTitle) {
            return;
        }

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
     * Find the cells by tag char.
     *
     * @method findTagCells
     *
     * @param {Object} worksheet
     * @param {Object} range
     * @param {Object} endCell
     * @param {string} tagChar
     *
     * @return {Array} cells The array of cell(s).
     */
    function findTagCells(worksheet, range, endCell, tagChar) {
        var cells = [];
        if (!tagChar) {
            return cells;
        }

        for (var ir = range.s.r; ir <= endCell.r; ++ir) {
            for (var ic = range.s.c; ic <= range.e.c; ++ic) {

                var cell_address = {
                    c: ic,
                    r: ir
                };
                z = XLSX.utils.encode_cell(cell_address);
                if (!isTag(worksheet[z], tagChar)) {
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
     * @param {Object} worksheet
     * @param {Object} range
     * @param {Object} titleCell
     * @param {Object} ignoreCells
     *
     * @return {Array} cells The cells of title.
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

            // TODO: check the value is valid or not (only alphabet)

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
     * fetch all the attribute(s)
     *
     * @method fetchAttrJson
     *
     * @param {Object} ws
     * @param {Array} attrCells
     *
     * @return {JSON} cells The JSON Object of attribute.
     */
    function fetchAttrJson(ws, attrCells) {
        var cells = {};
        for (var cell of attrCells) {
            var keyAddr = {
                    c: cell.c + 1,
                    r: cell.r
                },
                valAddr = {
                    c: cell.c + 2,
                    r: cell.r
                };

            var zKey = XLSX.utils.encode_cell(keyAddr),
                zVal = XLSX.utils.encode_cell(valAddr);

            cells[ws[zKey].v] = ws[zVal].v;
        }

        return cells;
    }

    /**
     * fetch name of key
     *
     * @method fetchNameOfKeyJson
     *
     * @param {Object} ws
     * @param {Array} cells
     *
     * @return {string} name of key.
     */
    function fetchNameOfKeyJson(ws, cells) {
        for (var cell of cells) {
            var valAddr = {
                c: cell.c + 1,
                r: cell.r
            };

            var zVal = XLSX.utils.encode_cell(valAddr);

            return ws[zVal].v;
        }
    }

    /**
     * fetch raw data(s)
     *
     * @method fetchRawData
     *
     * @param {Object} ws Object of worksheet
     * @param {Object} range The range from XLSX.utils.decode_range
     * @param {Object} titleCell
     * @param {Array} titles
     * @param {Array} ignoreCells
     *
     * @return {Array} rawDatas Array of Cells.
     */
    function fetchRawData(ws, range, titleCell, titles, ignoreCells) {
        var cells = [];
        var rawDatas = [];
        var cell_address = {};
        for (var ir = titleCell.r + 1; ir <= range.e.r; ++ir) {
            var data = {};
            for (var ic = range.s.c; ic <= range.e.c; ++ic) {
                cell_address = {
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

                // check title cell content
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
     * Parse the excel file and export the data with JSON format.
     *
     * @method parseSheet
     *
     * @param {string} filePath
     * @param {string} sheetName
     * @param {string} tagTitle
     * @param {string} tagAttribute
     * @param {string} tagNameOfKey
     * @param {string} tagIgnore
     *
     * @return {JSON} result The data of JSON.
     */
    function parseSheet(filePath, sheetName, tagTitle, tagAttribute, tagNameOfKey, tagIgnore) {
        var ws = loadSheet(filePath, sheetName);
        if (!ws) {
            return;
        }

        var ref = ws['!ref'];
        if (!ref) {
            return;
        }

        var range = XLSX.utils.decode_range(ref);

        // find all tag of title cells
        var titleCell = findTitleTagCell(ws, range, tagTitle);
        if (!titleCell) {
            return;
        }
        // find all tag of attribute cells
        var attrCells = findTagCells(ws, range, titleCell, tagAttribute);
        // find all tag of ignore cells
        var ignoreCells = findTagCells(ws, range, titleCell, tagIgnore);
        // find all tag of name of key cells
        var nameOfKeyCells = findTagCells(ws, range, titleCell, tagNameOfKey);

        // find all titles with titleCell
        var titles = findTitleCells(ws, range, titleCell, ignoreCells);
        if (titles.length === 0) {
            return;
        }

        // fetch all attribute cell(s)
        var attrs = fetchAttrJson(ws, attrCells);
        // fetch name of key
        var nameOfKey = fetchNameOfKeyJson(ws, nameOfKeyCells);

        // fetch all raw datas
        var rawDatas = fetchRawData(ws, range, titleCell, titles, ignoreCells);

        return {
            nameOfKey: nameOfKey,
            attrs: attrs,
            rawDatas: rawDatas
        };
    }

    /**
     * Map the data to tempalte.
     *
     * @method map
     *
     * @param {Object} data
     * @param {JSON} template
     * @param {Boolean} isKeyVal
     *
     * @return {JSON} result Mapped data.
     */
    function map(data, template, isKeyVal) {
        var type = template.constructor;

        var obj;
        switch (type) {
        case Array:
            obj = [];
            break;
        case Object:
            obj = {};
            break;
        default:
            return data[template];
        }

        expr = /^\$/;
        for (var key in template) {
            if (!template.hasOwnProperty(key)) {
                continue;
            }

            var isValKey = expr.test(key);
            var keyVal = isValKey ? map(data, key.replace('$', '')) : key;
            var val = template[key];
            switch (val.constructor) {
            case Array:
                obj[keyVal] = [];
                for (var index in val) {
                    var arrVal;
                    if (val.hasOwnProperty(index)) {
                        switch (val[index].constructor) {
                        case Array:
                            arrVal = map(data, val[index]);
                            break;
                        case Object:
                            arrVal = map(data, val[index]);
                            break;
                        default:
                            arrVal = data[val[index]];
                            break;
                        }
                        obj[keyVal].push(arrVal);
                    }
                }
                break;
            case Object:
                switch (type) {
                case Array:
                    obj.push(map(data, val));
                    break;
                case Object:
                    obj[keyVal] = map(data, val);
                    break;
                }
                break;
            default:
                switch (type) {
                case Array:
                    obj.push(data[val]);
                    break;
                case Object:
                    obj[keyVal] = data[val];
                    break;
                }
                break;
            }

            if (isKeyVal) {
                obj._key = keyVal;
                obj._val = obj[keyVal];
            }
        }

        return obj;
    }

    /**
     * Transform the raw data to the template JSON object.
     *
     * @method transform
     *
     * @param {Array} rawData
     * @param {JSON} template
     * @param {Boolean} isKeyVal
     *
     * @return {Array} result Array of JSON objects with template.
     */
    function transform(rawData, template, isKeyVal) {
        // init result object
        var result = isKeyVal ? {} : [];

        // foreach data to mapping to template
        for (var index in rawData) {
            if (!rawData.hasOwnProperty(index)) {
                continue;
            }

            // mapping
            var obj = map(rawData[index], template, isKeyVal);
            if (isKeyVal) {
                result[obj._key] = obj._val;
            } else {
                result.push(obj);
            }
        }

        return result;
    }

    /**
     * Parse the excel file and export the data with JSON format.
     *
     * @method parse
     *
     * @param {string} filePath
     * @param {string} sheetName
     * @param {JSON} template
     * @param {Object} options
     * @param {Boolean} options.useSheetname = false
     * @param {Boolean} options.isKeyVal = false
     * @param {string} options.tagTitle = '#'
     * @param {string} options.tagAttribute = '^'
     * @param {string} options.tagNameOfKey = '~'
     * @param {string} options.tagIgnore = '!'
     *
     * @return {Object} The JSON Object of data
     */
    function parse(filePath, sheetName, template, options) {
        options = options || {};

        nameOfKey = options.nameOfKey;
        isAddSheetName = options.isAddSheetName || false;
        isKeyVal = options.isKeyVal || false;
        tagTitle = options.tagTitle || '#';
        tagAttribute = options.tagAttribute || '^';
        tagNameOfKey = options.tagNameOfKey || '~';
        tagIgnore = options.tagIgnore || '!';

        // check sheet name is empty or not
        if (!sheetName) {
            throw new Error("The sheet name is null or empty.");
        }

        // check template
        template = template || {};
        if (!_u.isObject(template) || _u.isArray(template)) {
            throw new Error("The JSON template is invaild.");
        }

        // get data
        var objJson = parseSheet(filePath, sheetName,
            tagTitle, tagAttribute, tagNameOfKey, tagIgnore);
        objJson = objJson || {
            attrs: {},
            rawDatas: []
        };

        // copy top-level attribute
        var result = objJson.attrs;
        if (isAddSheetName) {
            result[DEFAULT_ATTR_SHEETNAME] = sheetName;
        }
        // transform the data if necessary
        datas = _u.isEmpty(template) ? objJson.rawDatas : transform(objJson.rawDatas, template, isKeyVal);
        if (isKeyVal) {
            result = datas;
        } else {
            nameOfKey = nameOfKey ? nameOfKey : (objJson.nameOfKey || DEFAULT_ATTR_DATAS);
            result[nameOfKey] = datas;
        }

        return result;
    }

    /**
     * Save the JSON to the file.
     *
     * @method save
     *
     * @param {string} filePath
     * @param {string} jsonObj
     * @param {Object} options Set replacer for a JSON replacer[Options].
     * @param {Function} callback
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
        fs.stat(dir, function (err, data) {
            if (err) {
                fs.mkdirSync(dir);
            }

            jsonfile.writeFile(filePath, jsonObj, options, function (err) {
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
     * @param {string} filePath
     * @param {Function} callback
     *
     * @return
     */
    function loadTemplate(filePath, callback) {
        jsonfile.readFile(filePath, function (err, obj) {
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
        findTagCells: findTagCells,
        findTitleCells: findTitleCells,
        parseSheet: parseSheet,
        map: map,
        transform: transform,
    };
}());

module.exports = excel2jsontemplate;
