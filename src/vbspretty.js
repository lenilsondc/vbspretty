var vbspretty = function vbspretty_(options){


  var
    parse = require('./vbsparser.js'),
    beautify = require('./vbsbeautifier.js');

  options = options || {};
  (function vbspretty_beautify(){
    var tparsedTemp = parse(options);

    // Combine any consecutive "STRING" items into one String
    var tparsed = {};
    tparsed.tokens = [];
    tparsed.tokenTypes = [];
    for(i = 0; i<tparsedTemp.tokens.length; i++) {
      let tkn = tparsedTemp.tokens[i];
      let tknType = tparsedTemp.tokenTypes[i];
      if (tknType === "STRING" && tparsed.tokenTypes[tparsed.tokenTypes.length-1] === "STRING") {
        // console.log('tparsed.tokenTypes.length:', tparsed.tokenTypes.length, i)
        // console.log("tparsed.tokenTypes[tparsed.tokenTypes.length-1]:", tparsed.tokenTypes[tparsed.tokenTypes.length-1])
        tparsed.tokens[tparsed.tokenTypes.length-1] += tkn;
      } else {
        tparsed.tokens.push(tkn);
        tparsed.tokenTypes.push(tknType);
      }
    }
    
    // require('fs').writeFileSync('./bundle-parsed.json', JSON.stringify(tparsed, null, 2));

    (function vbspretty_beautify_options(){
      options.tokens = tparsed.tokens || [];
      options.tokenTypes = tparsed.tokenTypes || [];
      options.level = options.level && /^\d+$/.test(options.level) ? parseInt(options.level) : 0;
      options.indentChar = options.indentChar ? options.indentChar.toString() : '  ';
      options.breakLineChar = options.breakLineChar ? options.breakLineChar.toString() : '\n';
      options.breakOnSeperator = options.breakOnSeperator === true || false;
      options.removeComments = options.removeComments === true || false;
    })();
  })();

  return beautify(options);
};

module.exports = vbspretty;

// tests

/*(function (fs) {

  fs.readFile('./input.vbs', 'utf8', function (err,data) {
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

    console.log(bsource)
  });
})(require('fs'));*/
