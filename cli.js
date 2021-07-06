const vbspretty = require('./src/vbspretty');

// tests
let inFile = process.argv[2];
let outFile = inFile.replace('.vbs', '-pretty.vbs');
(function (fs) {
  console.log('Prettifying vbs file:', inFile);
  fs.readFile(inFile, 'utf8', function (err,data) {
    if (err) {
      return console.log(err);
    }

    var bsource = vbspretty({
      level: 1,
      indentChar: '\t',
      breakLineChar: '\r\n',
      breakOnSeperator: false,
      removeComments: false,
      source: data
    });

    fs.writeFileSync(outFile, bsource);
  });
})(require('fs'));
