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

                done();
            });
        });

        describe('when sheet is exist', function () {
            it('should return undefined', function (done) {
                var sheet = e2jt.loadSheet('./test/data/test.xlsx', 'none');
                should(sheet).not.ok();

                done();
            });
        });

        describe('when sheet is NOT exist', function () {
            it('should return sheet object', function (done) {
                var sheet = e2jt.loadSheet('./test/data/test.xlsx', 'test');
                should(sheet).ok();

                done();
            });
        });

        describe('without parameter [filePath] and [sheetName]', function () {
            it('should return undefined', function (done) {
                var sheet = e2jt.loadSheet();
                should(sheet).not.ok();

                done();
            });
        });

        describe('without parameter [sheetName]', function () {
            it('should return undefined', function (done) {
                var sheet = e2jt.loadSheet('./test/data/test.xlsx');
                should(sheet).not.ok();

                done();
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

            should(jsonObj).have.properties('key');
            should(jsonObj).have.properties('attributes');
            should(jsonObj).have.properties('array');
            should(jsonObj).have.properties('price');

            should(jsonObj.key).be.eql('id');

            should(jsonObj.attributes).be.Object();
            should(jsonObj.attributes.value).be.eql('value');
            should(jsonObj.attributes.name).be.eql('name');
            should(jsonObj.attributes.array).be.Array();
            should(jsonObj.attributes.array[0]).be.eql('value');
            should(jsonObj.attributes.array[1]).be.eql('name');

            should(jsonObj.array).be.Array();
            should(jsonObj.array[0]).be.eql('value');
            should(jsonObj.array[1]).be.eql('name');

            should(jsonObj.array[2]).be.Object();
            should(jsonObj.array[2].value).be.eql('value');
            should(jsonObj.array[2].name).be.eql('name');

            should(jsonObj.array[3]).be.Array();
            should(jsonObj.array[3][0]).be.eql('value');
            should(jsonObj.array[3][1]).be.eql('name');

            should(jsonObj.price).be.eql('value');

            done();
        });
    });
});

// #parse()
describe('#parse()', function () {
    it('should return right JSON Object', function (done) {
        e2jt.loadTemplate('./test/data/template.json', function (err, jsonObj) {
            if (err) {
                done(err);
            } else {
                var data = e2jt.parse('./test/data/test.xlsx', 'test', jsonObj);
                // console.log(JSON.stringify(data, '', 4));
                should(data).be.Object();
                should(data.version).be.eql('1.0.0');
                should(data.name).be.eql('newname');

                should(data.datas).be.Array();
                should(data.datas.length).be.eql(7);
                for (var i = 0; i < data.datas.length; i++) {
                    var v = i + 1;
                    should(data.datas[i].key).be.eql('Id_' + v);

                    should(data.datas[i].attributes).be.Object();
                    should(data.datas[i].attributes.value).be.eql('Value_' + v);

                    should(data.datas[i].attributes.array).be.Array();
                    should(data.datas[i].attributes.array[0]).be.eql(data.datas[i].attributes.value);
                    should(data.datas[i].attributes.array[1]).be.eql(data.datas[i].attributes.name);

                    should(data.datas[i].array).be.Array();
                    should(data.datas[i].array[0]).be.eql(data.datas[i].attributes.value);
                    should(data.datas[i].array[1]).be.eql(data.datas[i].attributes.name);

                    should(data.datas[i].array[2]).be.Object();
                    should(data.datas[i].array[2].value).be.eql(data.datas[i].attributes.value);
                    should(data.datas[i].array[2].name).be.eql(data.datas[i].attributes.name);

                    should(data.datas[i].array[3]).be.Array();
                    should(data.datas[i].array[3][0]).be.eql(data.datas[i].attributes.value);
                    should(data.datas[i].array[3][1]).be.eql(data.datas[i].attributes.name);
                }

                done();
            }
        });
    });
});


// #save()
// describe('#save()', function () {
//     it('should save right JSON Object', function (done) {
//         done();
//     });
// });
