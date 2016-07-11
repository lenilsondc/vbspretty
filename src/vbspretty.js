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
      options.level = options.level && /^d+$/.test(options.level) ? options.level : 0;
    })();

    return beautify(options);
  })();
};


//test
var bsource = vbspretty({ source: 'Dim i\n i = 0\nIf i = 0 Then\n i = i+1\n End If'});
console.log(vbspretty);
