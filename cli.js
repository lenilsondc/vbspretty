#!/usr/bin/env node
var vbspretty = require('./src/vbspretty');

var argv = process.argv.slice(2);

var options = {};
var inFile = argv.shift();
var outFile = inFile;

for (var i = 0; i < argv.length; i++) {
  var arg = argv[i];
  var option = arg.substring(2); //remove -- prefix

  switch (option) {
    case 'level':
      var value = argv[++i];
      if (!isNaN(value)) {
        options[option] = Number(value);
      } else {
        console.error("Option '" + option + "' is expected to be a number, got '" + value + "' instead");
        process.exit(1);
      }
      break;
    case 'indentChar':
      var value = argv[++i];
      if (/(\ +|\\t)/.test(value)) {
        options[option] = value.replace('\\t', '\t');
      } else {
        console.error("Option '" + option + "' accepts tabs or spaces, got '" + value + "' instead");
        process.exit(2);
      }
      break;
    case 'breakLineChar':
      var value = argv[++i];
      if (/(\\n|\\r\\n)/) {
        options[option] = value.replace('\\n', '\n').replace('\\r', '\r');
      } else {
        console.error("Option '" + option + "' accepts \n or \r\n only, got '" + value + "' instead");
        process.exit(3);
      }
      break;
    case 'breakOnSeperator':
      options[option] = true
      break;
    case 'removeComments':
      options[option] = true
      break;
    case 'output':
      var value = argv[++i]
      if (value) {
        outFile = value;
      } else {
        console.error("Option '" + option + "' expects a value.");
        process.exit(4);
      }
      break;
    default:
      console.warn("Option '" + i + "' is not valid, will be ignored.");
  }
}

(function (fs) {
  console.info('Prettifying vbs file:', inFile);
  fs.readFile(inFile, 'utf8', function (err, data) {
    if (err) {
      console.error("Failed to read file", inFile)
      console.error(err);
      return;
    }

    var bsource = vbspretty({
      level: options.level,
      indentChar: options.indentChar,
      breakLineChar: options.breakLineChar,
      breakOnSeperator: options.breakOnSeperator,
      removeComments: options.removeComments,
      source: data
    });

    console.info('Writing to vbs file:', outFile);
    fs.writeFileSync(outFile, bsource);
    console.info('Done!');
  });
})(require('fs'));

//--level 1 --indentChar '\t' --breakLineChar '\r\n' --breakOnSeperator false --removeComments false 