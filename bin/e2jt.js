#!/usr/bin/env node

const path = require('path');
const fs = require('fs');
const util = require('util');
const cmder = require('commander');
const clc = require('cli-color');
const pkg = require('../package.json');
const e2jt = require('../index.js');

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
    .action(function(file, sheet, output) {
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

function parseCallback(err, jsonObj) {
    if (err) {
        throw err;
    }

    var fp = path.resolve(inputFile);
    var data = e2jt.parse(fp, inputSheet, jsonObj);
    var opdir = cmder.outputdir || path.dirname(inputFile);
    outputFile = path.resolve(path.join(opdir, outputFile));
    e2jt.save(outputFile, data, function(err) {
        if (err) {
            console.log(err);
            return;
        }

        console.log(clc.greenBright('Parse "%s" successed.'), fp);
    });
}
