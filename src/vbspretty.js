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
      options.indentChar = options.indentChar ? options.indentChar.ToString() : '  ';
      options.breakLineChar = options.breakLineChar ? options.breakLineChar.ToString() : '\n';
    })();
  })();
console.log(options.tokenTypes)
  return beautify(options);
};


//test
var bsource = vbspretty({
  level: 1,
  source: `If i = 0 Then j = 0
  Dim i, j
  i = 0
  If i = 0 Then
  i = i+1
  End If



  FOR i = 0 to 10
    SELECT CASE j
    case 10:
      case 11:
j = 10
      CASE ELSE:
      j = 11
    End Select
  Next
`});
console.log(bsource)
