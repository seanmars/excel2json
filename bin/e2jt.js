#!/usr/bin/env node

var path = require('path');
var fs = require('fs');
var util = require('util');
var cmder = require('commander');
var clc = require('cli-color');
var pkg = require('../package.json');
var e2jt = require('../index.js');

var inputFile;
var inputSheet;
var outputFile;

// setting commander
cmder.version(pkg.version)
    .usage('<parse file> [sheet name] [output name]')
    .arguments('<file> [sheet] [output]')
    .option('-s, --space <space>', 'The white space to insert into the output JSON string for readability purposes')
    .option('-t, --template <template>', 'The file path of JSON template')
    .option('-o, --outputdir <outputdir>', 'The path of output folder')
    .option('--key <name of key>', 'The name of key for data list')
    .option('--sheetname [true|false]', 'Auto add sheet name.[true|false]', /^(true|false)$/i, false)
    .option('--dict [true|false]', 'Is key-value JSON?[true|false]', /^(true|false)$/i, false)
    .action(function (file, sheet, output) {
        inputFile = file;
        inputSheet = sheet ? sheet : path.basename(file, path.extname(file));
        outputFile = (output ? output : inputSheet) + '.json';
    })
    .parse(process.argv);

// analysis parameter
if (!inputFile) {
    console.error(clc.red("ERROR:\n") + 'No file given!');
    process.exit(1);
}

if (cmder.template) {
    e2jt.loadTemplate(cmder.template, parseCallback);
} else {
    parseCallback();
}

function parseCallback(err, templateJson) {
    if (err) {
        throw err;
    }

    var fp = path.resolve(inputFile);
    var data = e2jt.parse(fp, inputSheet, templateJson, {
        nameOfKey: cmder.key,
        isKeyVal: cmder.dict,
        isAddSheetName: cmder.sheetname
    });
    var opdir = cmder.outputdir || path.dirname(inputFile);
    var space = {
        spaces: Number(cmder.space) || 0
    };
    outputFile = path.resolve(path.join(opdir, outputFile));
    e2jt.save(outputFile, data, space, function (err) {
        if (err) {
            console.log(err);
            return;
        }

        console.log(clc.greenBright('Parse "%s" successed.'), fp);
    });
}
