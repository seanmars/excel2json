var path = require('path');
var e2j = require('../excel2json.js');

var fp = path.resolve('./test/data/test.xlsx');
var dd = e2j.parse(fp, 'demo');
console.log(JSON.stringify(dd, null, 4));

// TODO check the dir is exists.
e2j.save('./test/output/test.json', dd);
