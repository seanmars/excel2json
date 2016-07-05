const fs = require('fs');
var path = require('path');
var e2j = require('../excel2json.js');

e2j.load('./test/data/template.json', function(err, jsonObj) {
    console.log(jsonObj);
    var fp = path.resolve('./test/data/test.xlsx');
    var dd = e2j.parse(fp, 'demo', jsonObj);
    console.log(JSON.stringify(dd, null, 4));

    fs.mkdirSync('./test/output/', 755);
    e2j.save('./test/output/test.json', dd);
});
