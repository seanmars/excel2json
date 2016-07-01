var excel2json = (function() {
    if (typeof require !== 'undefined') {
        XLSX = require('xlsx');
        _ = require('underscore');
    }

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
        console.log('Load file: ' + filePath);
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
     * Get custom excel cell Object
     *
     * @method getCell
     *
     * @param  {string} z The key of cell in Excel.
     * @param  {string} v The cell's value.
     *
     * @return {Object} { c: column, r: row, z: key of cell }
     */
    function getCell(z, v) {
        var cell = XLSX.utils.decode_cell(z);
        cell.z = z;
        cell.v = v;

        return cell;
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
            v: obj ? obj.v : null,
            w: obj ? obj.w : null,
            c: cell_address.c,
            r: cell_address.r,
            z: z_value,
        };

        return cell;
    }

    /**
     * Check the cell is need ignore or not.
     *
     * @method needIgnore
     *
     * @param  {Object} cell
     * @param  {Array}  filters
     *
     * @return {Boolean}
     */
    function needIgnore(cell, filters) {
        for (var k in filters) {
            if (!filters.hasOwnProperty(k)) {
                continue;
            }

            var fc = filters[k];
            if (fc.c === 0 && fc.r == cell.r) {
                return true;
            }

            if (fc.c == cell.c && fc.r < cell.r) {
                return true;
            }
        }

        return false;
    }

    /**
     * Parse the excel file and export the data with JSON format.
     *
     * @method parse
     *
     * @param  {string} filePath
     * @param  {string} sheetName
     * @param  {string} jsonTemplate
     * @param  {string} titleChar    = '#'
     * @param  {string} ignoreChar   = '!'
     *
     * @return {Array} [description]
     */
    function parse(filePath, sheetName, jsonTemplate,
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
        console.log(ref, range);

        // TODO: check the title cell is exists or not

        var filterCells = [];
        var titleCell;
        var titles = [];
        var cells = [];
        var result = [];

        for (var ir = range.s.r; ir <= range.e.r; ++ir) {
            var data = {};
            for (var ic = range.s.c; ic <= range.e.c; ++ic) {
                var cell_address = {
                    c: ic,
                    r: ir
                };
                z = XLSX.utils.encode_cell(cell_address);
                var cell = factoryCell(cell_address, z, ws[z]);

                // check type of cell
                if (typeof cell.v === 'string') {
                    if (cell.v == ignoreChar) {
                        filterCells.push(cell);
                        continue;
                    }

                    if (cell.v == titleChar) {
                        titleCell = cell;
                        continue;
                    }
                }

                // check the cell is title or not
                if (titleCell && cell.v && (cell.r == titleCell.r)) {
                    titles[ic] = cell;
                    continue;
                }

                // check the cell is need to ignore or not
                if (needIgnore(cell, filterCells)) {
                    // console.log(JSON.stringify(cell));
                    continue;
                }

                // check the titles
                if (titles.length === 0) {
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

            result.push(data);
        }

        console.log("\nValid count: " + cells.length);
        console.log("\nFilter cells:");
        console.log(filterCells);
        console.log("\nTitle cell:");
        console.log(titleCell);
        console.log("\nTitles:");
        console.log(titles);
        console.log("\nCells:");
        console.log(cells);

        return result;
    }

    return {
        toJson: toJson,
        parse: parse,
    };
}());

module.exports = excel2json;
