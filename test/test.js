var should = require('should');
var chai = require('chai');
var util = require('util');
var e2jt = require('../index.js');

// load file
describe('Load file', function () {
    // #loadSheet()
    describe('#loadSheet()', function () {
        describe('when type is xlsx', function () {
            describe('when file is NOT exist', function () {
                it('should throw expect', function (done) {
                    var filePath = './test/data/none.xlsx';
                    should.throws(() => e2jt.loadSheet(filePath, 'test'));

                    return done();
                });
            });

            describe('when sheet is exist', function () {
                it('should return undefined', function (done) {
                    var sheet = e2jt.loadSheet('./test/data/test.xlsx', 'none');
                    should(sheet).not.ok();

                    return done();
                });
            });

            describe('when sheet is NOT exist', function () {
                it('should return sheet object', function (done) {
                    var sheet = e2jt.loadSheet('./test/data/test.xlsx', 'test');
                    should(sheet).ok();

                    return done();
                });
            });

            describe('without parameter [filePath] and [sheetName]', function () {
                it('should return undefined', function (done) {
                    var sheet = e2jt.loadSheet();
                    should(sheet).not.ok();

                    return done();
                });
            });

            describe('without parameter [sheetName]', function () {
                it('should return undefined', function (done) {
                    var sheet = e2jt.loadSheet('./test/data/test.xlsx');
                    should(sheet).not.ok();

                    return done();
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
