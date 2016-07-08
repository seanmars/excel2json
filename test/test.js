const fs = require('fs');
var path = require('path');
var e2j = require('../index.js');

e2j.load('./test/data/template.json', function(err, jsonObj) {
    const outputPath = './test/output';
    var stats = fs.statSync(outputPath);
    if (!stats.isDirectory()) {
        fs.mkdirSync('./test/output/', 755);
    }

    console.log(jsonObj);
    var fp = path.resolve('./test/data/test.xlsx');

    var dd = e2j.parse(fp, 'demo', jsonObj);
    console.log(JSON.stringify(dd, null, 4));
    e2j.save('./test/output/test_with_template.json', dd);

    dd = e2j.parse(fp, 'demo');
    console.log(JSON.stringify(dd, null, 4));
    e2j.save('./test/output/test_without_template.json', dd);
});
