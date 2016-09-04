var should = require('should');
var chai = require('chai');
var util = require('util');
var e2jt = require('../index.js');

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
describe('#loadTemplate', function () {
    it('should return right JSON Object', function (done) {
        e2jt.loadTemplate('./test/data/template.json', function (err, jsonObj) {
            should(err).be.a.null();
            should(jsonObj).not.be.a.null();

            // console.log(jsonObj);
            jsonObj.should.have.properties('key');
            jsonObj.should.have.properties('attributes');
            jsonObj.should.have.properties('array');
            jsonObj.should.have.properties('price');

            jsonObj.key.should.be.eql('id');

            jsonObj.attributes.should.be.a.Object();
            jsonObj.attributes.value.should.be.eql('value');
            jsonObj.attributes.name.should.be.eql('name');

            jsonObj.array.should.be.a.Array();
            jsonObj.array[0].should.be.eql('value');
            jsonObj.array[1].should.be.eql('name');
            jsonObj.array[2].should.be.a.Object();
            jsonObj.array[2].value.should.be.eql('value');
            jsonObj.array[2].name.should.be.eql('name');

            jsonObj.price.should.be.eql('value');

            return done();
        });
    });
});

// #parse()
describe('#parse()', function () {
    it('should return right JSON Object', function (done) {
        throw new exception('not completed.');
    });
});


// #save()
describe('#save()', function () {
    it('should save right JSON Object', function (done) {
        throw new exception('not completed.');
    });
});
