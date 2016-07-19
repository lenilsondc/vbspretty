var vbspretty = function vbspretty_(options){


  var
    parse = require('./vbsparser.js'),
    beautify = require('./vbsbeautifier.js');

  options = options || {};
  (function vbspretty_beautify(){
    var tparsed = parse(options);

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
