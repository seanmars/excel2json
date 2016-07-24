var fs = require('fs');
var path = require('path');
var e2jt = require('../index.js');

// e2jt.loadTemplate('./test/data/template.json', function(err, jsonObj) {
//     var outputPath = './test/output';
//     var stats = fs.statSync(outputPath);
//     if (!stats.isDirectory()) {
//         fs.mkdirSync('./test/output/');
//     }
//
//     console.log(jsonObj);
//     var fp = path.resolve('./test/data/test.xlsx');
//
//     var data = e2jt.parse(fp, 'demo', jsonObj);
//     console.log(JSON.stringify(data, null, 4));
//     e2jt.save('./test/output/test_with_template.json', data);
// });
//
// e2jt.loadTemplate('./test/data/template.json', function(err, jsonObj) {
//     var outputPath = './test/output';
//     var stats = fs.statSync(outputPath);
//     if (!stats.isDirectory()) {
//         fs.mkdirSync('./test/output/');
//     }
//
//     console.log(jsonObj);
//     var fp = path.resolve('./test/data/test.xlsx');
//
//     var data = e2jt.parse(fp, 'demo');
//     console.log(JSON.stringify(data, null, 4));
//     e2jt.save('./test/output/test_without_template.json', data);
// });
//
// e2jt.loadTemplate('./test/data/template.json', function(err, jsonObj) {
//     var outputPath = './test/output';
//     var stats = fs.statSync(outputPath);
//     if (!stats.isDirectory()) {
//         fs.mkdirSync('./test/output/');
//     }
//
//     console.log(jsonObj);
//     var fp = path.resolve('./test/data/test.xlsx');
//
//     var data = e2jt.parse(fp, 'nosheet');
//     console.log(JSON.stringify(data, null, 4));
//     e2jt.save('./test/output/test_nosheet.json', data);
// });
//
e2jt.loadTemplate('./test/data/template.json', function(err, jsonObj) {
    var outputPath = './test/output';
    var stats = fs.statSync(outputPath);
    if (!stats.isDirectory()) {
        fs.mkdirSync('./test/output/');
    }

    console.log(jsonObj);
    var fp = path.resolve('./test/data/test.xlsx');

    var data = e2jt.parse(fp, 'test2');
    e2jt.save('./test/output/test2_without_template.json', data);
});
