var path = require('path');
var jsonfile = require('jsonfile');
var e2j = require('./src/excel2json');

var fp = path.resolve('./data/test.xlsx');
var d = e2j.parse(fp, 'test');
// var d = e2j.parse(fp, 'demo');
console.log(JSON.stringify(d, null, 4));

// console.log(XLSX.utils.encode_row(0));
// console.log(XLSX.utils.encode_col(0));
