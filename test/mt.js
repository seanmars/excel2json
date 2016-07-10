var should = require('should');
var e2jt = require('../index.js');

// load file
describe('Load file', function() {
    // #loadSheet()
    describe('#loadSheet()', function() {
        describe('when type if xlsx', function() {
            describe('when file is NOT exist', function() {
                it('should return undefined', function() {
                    var sheet = e2jt.loadSheet('./test/data/none.xlsx', 'test');
                    should(sheet).be.not.ok();
                });
            });

            describe('when sheet is exist', function() {
                it('should return undefined', function() {
                    var sheet = e2jt.loadSheet('./test/data/test.xlsx', 'none');
                    should(sheet).be.not.ok();
                });
            });

            describe('when sheet is NOT exist', function() {
                it('should return sheet object', function() {
                    var sheet = e2jt.loadSheet('./test/data/test.xlsx', 'test');
                    should(sheet).be.ok();
                });
            });

            describe('without parameter [filePath] and [sheetName]', function() {
                it('should return undefined', function() {
                    var sheet = e2jt.loadSheet();
                    should(sheet).be.not.ok();
                });
            });

            describe('without parameter [sheetName]', function() {
                it('should return undefined', function() {
                    var sheet = e2jt.loadSheet('./test/data/test.xlsx');
                    should(sheet).be.not.ok();
                });
            });
        });
    });

    // #loadTemplate()
});

// #factoryCell()
// #isNeedIgnore()
// #isTag()
// #findTitleTagCell()
// #findIgnoreTagCells()
// #findTitleCells()
// #parseSheet()
// #map()
// #transform()
// #save()
// #parse()
