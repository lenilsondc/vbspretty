var vbsparser = function vbsparser_(options) {
  var
    tokens = [],
    tokenTypes = [],
    tokenTable = {
      "Step": "STEP",
      "Class": "CLASS",
      "Const": "CONST",
      "Function": "FUNCTION",
      "Property": "PROPERTY",
      "Sub": "SUB",
      "Goto": "GOTO",
      "Xor": "BINARY_OPERATOR",
      "Or": "BINARY_OPERATOR",
      "And": "BINARY_OPERATOR",
      "Not": "BINARY_OPERATOR",
      "Eqv": "BINARY_OPERATOR",
      "Imp": "BINARY_OPERATOR",
      "=": "COMPARISON_OPERATOR",
      "<=": "COMPARISON_OPERATOR",
      ">=": "COMPARISON_OPERATOR",
      "<>": "COMPARISON_OPERATOR",
      "Is": "COMPARISON_OPERATOR",
      "<": "COMPARISON_OPERATOR",
      ">": "COMPARISON_OPERATOR",
      "Mod": "ARTHMETIC_OPERATOR",
      "Dim": "DIM",
      "ReDim": "REDIM",
      "Preserve": "PRESERVE",
      "Public": "PUBLIC",
      "Private": "PRIVATE",
      "Default": "DEFAULT",
      "Next": "FOR_LOOP_NEXT",
      "Nothing": "OBJECT_NOTHING",
      "Null": "VALUE_NULL",
      "True": "VALUE_TRUE",
      "False": "VALUE_FALSE",
      "Empty": "VALUE_EMPTY",
      "ByVal": "BYVAL",
      "ByRef": "BYREF",
      "Select": "SELECT",
      "Case": "CASE",
      "If": "IF",
      "Else": "ELSE",
      "ElseIf": "ELSE_IF",
      "Exit": "EXIT",
      "End": "END",
      "Then": "THEN",
      "Err": "ERR",
      "RegExp": "REGEXP",
      "Call": "CALL",
      "Erase": "ERASE",
      "With": "WITH",
      "Stop": "STOP",
      "On": "ON",
      "Error": "ERROR",
      "Resume": "RESUME",
      "Option": "OPTION",
      "Explicit": "EXPLICIT",
      "Do": "DO_LOOP",
      "While": "WHILE_LOOP",
      "Wend": "WHILE_LOOP_WEND",
      "Until": "DO_LOOP_UNTIL",
      "Loop": "DO_LOOP_END",
      "For": "FOR_LOOP",
      //"Each": "eac",
      "To": "FOR_LOOP_TO",
      "In": "FOR_LOOP_IN",
      "Set": "SET_OPERATOR",
      "New": "NEW_OPERATOR",
      "Abs": "VBSCRIPT_FUNCTION",
      "Array": "VBSCRIPT_FUNCTION",
      "Asc": "VBSCRIPT_FUNCTION",
      "Atn": "VBSCRIPT_FUNCTION",
      "CBool": "VBSCRIPT_FUNCTION",
      "CByte": "VBSCRIPT_FUNCTION",
      "CCur": "VBSCRIPT_FUNCTION",
      "CDate": "VBSCRIPT_FUNCTION",
      "CDbl": "VBSCRIPT_FUNCTION",
      "Chr": "VBSCRIPT_FUNCTION",
      "CInt": "VBSCRIPT_FUNCTION",
      "CLng": "VBSCRIPT_FUNCTION",
      "Conversions": "VBSCRIPT_FUNCTION",
      "Cos": "VBSCRIPT_FUNCTION",
      "CreateObject": "VBSCRIPT_FUNCTION",
      "CSng": "VBSCRIPT_FUNCTION",
      "CStr": "VBSCRIPT_FUNCTION",
      "Date": "VBSCRIPT_FUNCTION",
      "DateAdd": "VBSCRIPT_FUNCTION",
      "DateDiff": "VBSCRIPT_FUNCTION",
      "DatePart": "VBSCRIPT_FUNCTION",
      "DateSerial": "VBSCRIPT_FUNCTION",
      "DateValue": "VBSCRIPT_FUNCTION",
      "Day": "VBSCRIPT_FUNCTION",
      "Derived Math": "VBSCRIPT_FUNCTION",
      "Escape": "VBSCRIPT_FUNCTION",
      "Eval": "VBSCRIPT_FUNCTION",
      "Exp": "VBSCRIPT_FUNCTION",
      "Filter": "VBSCRIPT_FUNCTION",
      "FormatCurrency": "VBSCRIPT_FUNCTION",
      "FormatDateTime": "VBSCRIPT_FUNCTION",
      "FormatNumber": "VBSCRIPT_FUNCTION",
      "FormatPercent": "VBSCRIPT_FUNCTION",
      "GetLocale": "VBSCRIPT_FUNCTION",
      "GetObject": "VBSCRIPT_FUNCTION",
      "GetRef": "VBSCRIPT_FUNCTION",
      "Hex": "VBSCRIPT_FUNCTION",
      "Hour": "VBSCRIPT_FUNCTION",
      "InputBox": "VBSCRIPT_FUNCTION",
      "InStr": "VBSCRIPT_FUNCTION",
      "InStrRev": "VBSCRIPT_FUNCTION",
      "Int, Fix": "VBSCRIPT_FUNCTION",
      "IsArray": "VBSCRIPT_FUNCTION",
      "IsDate": "VBSCRIPT_FUNCTION",
      "IsEmpty": "VBSCRIPT_FUNCTION",
      "IsNull": "VBSCRIPT_FUNCTION",
      "IsNumeric": "VBSCRIPT_FUNCTION",
      "IsObject": "VBSCRIPT_FUNCTION",
      "Join": "VBSCRIPT_FUNCTION",
      "LBound": "VBSCRIPT_FUNCTION",
      "LCase": "VBSCRIPT_FUNCTION",
      "Left": "VBSCRIPT_FUNCTION",
      "Len": "VBSCRIPT_FUNCTION",
      "LoadPicture": "VBSCRIPT_FUNCTION",
      "Log": "VBSCRIPT_FUNCTION",
      "LTrim": "VBSCRIPT_FUNCTION",
      "Maths": "VBSCRIPT_FUNCTION",
      "Mid": "VBSCRIPT_FUNCTION",
      "Minute": "VBSCRIPT_FUNCTION",
      "Month": "VBSCRIPT_FUNCTION",
      "MonthName": "VBSCRIPT_FUNCTION",
      "MsgBox": "VBSCRIPT_FUNCTION",
      "Now": "VBSCRIPT_FUNCTION",
      "Oct": "VBSCRIPT_FUNCTION",
      "Replace": "VBSCRIPT_FUNCTION",
      "RGB": "VBSCRIPT_FUNCTION",
      "Right": "VBSCRIPT_FUNCTION",
      "Rnd": "VBSCRIPT_FUNCTION",
      "Round": "VBSCRIPT_FUNCTION",
      "RTrim": "VBSCRIPT_FUNCTION",
      "ScriptEngine": "VBSCRIPT_FUNCTION",
      "ScriptEngineBuildVersion": "VBSCRIPT_FUNCTION",
      "ScriptEngineMajorVersion": "VBSCRIPT_FUNCTION",
      "ScriptEngineMinorVersion": "VBSCRIPT_FUNCTION",
      "Second": "VBSCRIPT_FUNCTION",
      "SetLocale": "VBSCRIPT_FUNCTION",
      "Sgn": "VBSCRIPT_FUNCTION",
      "Sin": "VBSCRIPT_FUNCTION",
      "Space": "VBSCRIPT_FUNCTION",
      "Split": "VBSCRIPT_FUNCTION",
      "Sqr": "VBSCRIPT_FUNCTION",
      "StrComp": "VBSCRIPT_FUNCTION",
      "String": "VBSCRIPT_FUNCTION",
      "StrReverse": "VBSCRIPT_FUNCTION",
      "Tan": "VBSCRIPT_FUNCTION",
      "Time": "VBSCRIPT_FUNCTION",
      "Timer": "VBSCRIPT_FUNCTION",
      "TimeSerial": "VBSCRIPT_FUNCTION",
      "TimeValue": "VBSCRIPT_FUNCTION",
      "Trim": "VBSCRIPT_FUNCTION",
      "TypeName": "VBSCRIPT_FUNCTION",
      "UBound": "VBSCRIPT_FUNCTION",
      "UCase": "VBSCRIPT_FUNCTION",
      "Unescape": "VBSCRIPT_FUNCTION",
      "VarType": "VBSCRIPT_FUNCTION",
      "Weekday": "VBSCRIPT_FUNCTION",
      "WeekdayName": "VBSCRIPT_FUNCTION",
      "Year": "VBSCRIPT_FUNCTION",
      "vbCr": "VBSCRIPT_CONSTANT_STRING",
      "VbCrLf": "VBSCRIPT_CONSTANT_STRING",
      "vbFormFeed": "VBSCRIPT_CONSTANT_STRING",
      "vbLf": "VBSCRIPT_CONSTANT_STRING",
      "vbNewLine": "VBSCRIPT_CONSTANT_STRING",
      "vbNullChar": "VBSCRIPT_CONSTANT_STRING",
      "vbNullString": "VBSCRIPT_CONSTANT_STRING",
      "vbTab": "VBSCRIPT_CONSTANT_STRING",
      "vbVerticalTab": "VBSCRIPT_CONSTANT_STRING",
      "vbEmpty": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbNull": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbInteger": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbLong": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbSingle": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbDouble": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbCurrency": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbDate": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbString": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbObject": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbError": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbBoolean": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbVariant": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbDataObject": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbDecimal": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbByte": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbArray": "VBSCRIPT_CONSTANT_VARTYPE",
      "vbUseDefault": "VBSCRIPT_CONSTANT_TRISTATE",
      "vbTrue": "VBSCRIPT_CONSTANT_TRISTATE",
      "vbFalse": "VBSCRIPT_CONSTANT_TRISTATE",
      "vbOKOnly": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbOKCancel": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbAbortRetryIgnore": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbYesNoCancel": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbYesNo": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbRetryCancel": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbCritical": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbQuestion": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbExclamation": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbInformation": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbDefaultButton1": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbDefaultButton2": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbDefaultButton3": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbDefaultButton4": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbApplicationModal": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbSystemModal": "VBSCRIPT_CONSTANT_MSGBOX",
      "vbOK": "VBSCRIPT_CONSTANT_MSGBOX_RETVAL",
      "vbCancel": "VBSCRIPT_CONSTANT_MSGBOX_RETVAL",
      "vbAbort": "VBSCRIPT_CONSTANT_MSGBOX_RETVAL",
      "vbRetry": "VBSCRIPT_CONSTANT_MSGBOX_RETVAL",
      "vbIgnore": "VBSCRIPT_CONSTANT_MSGBOX_RETVAL",
      "vbYes": "VBSCRIPT_CONSTANT_MSGBOX_RETVAL",
      "vbNo": "VBSCRIPT_CONSTANT_MSGBOX_RETVAL",
      "vbObjectError": "VBSCRIPT_CONSTANT_ERROR",
      "vbGeneralDate": "VBSCRIPT_CONSTANT_DATEFORMAT",
      "vbLongDate": "VBSCRIPT_CONSTANT_DATEFORMAT",
      "vbShortDate": "VBSCRIPT_CONSTANT_DATEFORMAT",
      "vbLongTime": "VBSCRIPT_CONSTANT_DATEFORMAT",
      "vbShortTime": "VBSCRIPT_CONSTANT_DATEFORMAT",
      "vbBinaryCompare": "VBSCRIPT_CONSTANT_COMPARE",
      "vbTextCompare": "VBSCRIPT_CONSTANT_COMPARE",
      "vbBlack": "VBSCRIPT_CONSTANT_COLOR",
      "vbRed": "VBSCRIPT_CONSTANT_COLOR",
      "vbGreen": "VBSCRIPT_CONSTANT_COLOR",
      "vbYellow": "VBSCRIPT_CONSTANT_COLOR",
      "vbBlue": "VBSCRIPT_CONSTANT_COLOR",
      "vbMagenta": "VBSCRIPT_CONSTANT_COLOR",
      "vbCyan": "VBSCRIPT_CONSTANT_COLOR",
      "vbWhite": "VBSCRIPT_CONSTANT_COLOR",
    },
    source = options.source || '',
    lastParsedToken = '',
    lastNonWSParsedToken = '',
    pushToken = function vbsparser_pushToken(token, tokenType) {
      tokens.push(token);
      tokenTypes.push(tokenType);
      lastParsedToken = tokenType;
      if (tokenType === 'WHITESPACE' || tokenType === 'NEWLINE')
        lastNonWSParsedToken = tokenType;
    };

  //Lexical Analysis
  (function vbsparser_tokenizer() {
    var
      index = 0,
      bLength = source.length,
      buffer = options.source.split(''),
      n = 0,
      isSpace = function vbsparser_tokenizer_isSpace(char) {
        return (char === ' ' || char === '\t' || char === '\f' || char ===
          '\v');
      },
      isEOLorEOF = function vbsparser_tokenizer_isEOLorEOF(char) {
        char = char || buffer[index];
        if (char === -1) return true;
        if (char === '\r' && nextChar() === '\n') {
          return true;
        }
        if (char === '\n') return true;
        return false;
      },
      isAlphaNumeric = function vbsparser_tokenizer_isAlphaNumeric(char) {
        return char !== -1 && /[a-zA-Z0-9]/.test(char);
      },
      isDigit = function vbsparser_tokenizer_isDigit(char) {
        return /[0-9]/.test(char);
      },
      read = function vbsparser_tokenizer_read(length) {
        var str = '';
        length = length || 1;
        if (index + length > bLength) {
          return -1;
        }
        str = buffer.slice(index, index + length).join('');
        index += length;
        return str;
      },
      currentChar = function vbsparser_tokenizer_currentChar() {
        return (index >= 0 && index < bLength) ? buffer[index] : -1;
      },
      charAt = function vbsparser_tokenizer_charAt(chrIndex) {
        if (chrIndex >= 0 && chrIndex < bLength) {
          return buffer[chrIndex]
        }
        return -1;
      },
      prevChar = function vbsparser_tokenizer_prevChar() {
        if (index - 1 >= 0 && bLength > 0) {
          return -1;
        }
        return buffer[index - 1]
      },
      nextChar = function vbsparser_tokenizer_nextChar() {
        if (index + 1 >= bLength) {
          return -1;
        }
        return buffer[index + 1]
      },
      readTill = function vbsparser_tokenizer_readTill(fn) {
        var n = 0,
          str = '',
          peeked;
        while ((peeked = charAt(index + n)) !== -1 && !isEOLorEOF(peeked) &&
          fn(peeked, buffer, index)) {
          n++;
        }
        if (n === 0) return '';
        str = read(n);
        return str;
      },
      readSpace = function vbsparser_tokenizer_readSpace() {
        return readTill(function(char) {
          return isSpace(char);
        });
      },
      nextWord = function vbsparser_tokenizer_nextWord() {
        var ch,
          str = '';
        //n = 0;
        do {
          ch = charAt(index + n++);
        } while (isSpace(ch) || ch === '_' || ch === '\r' || ch === '\n');
        n--;
        if (n === 0) {
          return '';
        }
        while (isAlphaNumeric(ch = charAt(index + n++))) {
          str += ch;
        }
        n--;
        return str;
      },
      readString = function vbsparser_tokenizer_readString() {
        return readTill(function(char) {
          return char === "\"" && prevChar() !== "\""
        });
      },
      readLine = function vbsparser_tokenizer_readLine() {
        return readTill(function(char) {
          return isEOLorEOF(char);
        });
      },
      readAlphaNumeric = function vbsparser_tokenizer_readAlphaNumeric() {
        return readTill(function(char) {
          return isAlphaNumeric(char);
        });
      },
      readNumber = function vbsparser_tokenizer_readNumber(char) {
        var str = '';
        str += readTill(function(chr) {
          return isDigit(chr);
        });
        if (nextChar() === '.') str += read();
        if (str.length === 0) return '';
        str += readTill(function(chr) {
          return isDigit(chr);
        });
        if (nextChar() === 'e' || nextChar() === 'E') {
          str += read();
          if (nextChar() === '+' || nextChar() === '-') str += read();
          str += readTill(function(chr) {
            return isDigit(chr);
          });
        }
        return str;
      };

    var
      curChar,
      nextChar,
      ch,
      word;

    while ((ch = currentChar()) !== -1) {
      word = '';

      switch (ch) {
        /*case '~':
        case ';':
        case '?':
        case '|':
        case '`':
        case '!':
        case '{':
        case '}':
            pushToken(ch, UNKNOWN);
            break;*/
        case '\t':
        case '\v':
        case ' ':
        case '\f':
          pushToken(readSpace(), 'WHITESPACE');
          break;
        case '(':
          pushToken(read(), 'OPEN_BRACKET');
          break;
        case ')':
          pushToken(read(), 'CLOSE_BRACKET');
          break;
        case '\"':
          pushToken(readString(), 'STRING');
          break;
        case '\'':
          pushToken(readLine, 'COMMENT');
          break;
        case '#':
          read();
          word = '#' + readTill(function(char) {
            return char !== '#';
          }) + '#';
          read();
          pushToken(word, 'DATE');
          break;
        case '[':
          word = readTill(function(char) {
            return char !== ']';
          }) + "]";
          read();
          pushToken(word, 'IDENTIFIER');
          break;
        case '_':
          pushToken(read(), 'STATEMENT_CONTINUATION');
          break;
        case ':':
          pushToken(read(), 'STATEMENT_SEPARATOR');
          break;
        case '@':
          if (nextChar() === '@') pushToken(readLine(), 'COMMENT');
          else pushToken(read(), 'UNKNOWN');
          break;
        case '+':
        case '^':
        case '%':
        case '-':
        case '*':
        case '/':
        case '\\':
        case '=':
          pushToken(read(), 'ARTHMETIC_OPERATOR');
          break;
        case '<':
          nextChar = nextChar();

          if (nextChar === '>') {
            pushToken(read(2), 'COMPARISON_OPERATOR');
          } else if (nextChar === '=') {
            pushToken(read(2), 'COMPARISON_OPERATOR');
          } else {
            pushToken(read(), 'COMPARISON_OPERATOR');
          }

          break;
        case '>':
          nextChar = nextChar();

          if (nextChar === '=') {
            pushToken(read(2), 'COMPARISON_OPERATOR');
          } else {
            pushToken(read(), 'COMPARISON_OPERATOR');
          }

          break;
        case '&':
          nextChar = nextChar();

          if (nextChar === 'H' || nextChar === 'h') {
            //this is a hexa decimal value
            pushToken(read() + readAlphaNumeric(), 'HEXNUMBER');
          } else {
            pushToken(read(), 'ARTHMETIC_OPERATOR');
          }

          break;
        case '.':
          nextChar = nextChar();

          if (isDigit(nextChar)) {
            pushToken(readNumber(), 'NUMBER');
          } else {
            //record dot operator
            pushToken(read(), 'DOT_OPERATOR');
          }
          break;
        case '\r':
          if (nextChar() === '\n') {
            pushToken(read(2), 'NEWLINE');
          } else {
            pushToken(read(), 'NEWLINE');
          }

          break;
        case '\n':
          pushToken(read(), 'NEWLINE');
          break;
        case '0':
        case '1':
        case '2':
        case '3':
        case '4':
        case '5':
        case '6':
        case '7':
        case '8':
        case '9':
          pushToken(readNumber(), 'NUMBER');
          break;
        case ',':
          pushToken(read(), 'COMMA');
          break;
        default:

          if (!isAlphaNumeric(ch)) {
            pushToken(read(), 'INVALID');
            continue;
          }

          word = readAlphaNumeric();
          n = 0;

          if (lastNonWSParsedToken !== 'DOT_OPERATOR' && tokenTable[word] !== undefined) {

            switch (word.toLowerCase()) {
              case 'do':
                nextWord = nextWord().toLowerCase();

                if (nextWord === 'while') {
                  read(n);
                  pushToken('Do While', 'DO_LOOP_START_WHILE');
                } else if (nextWord === 'until') {
                  read(n);
                  pushToken('Do Until', 'DO_LOOP_START_UNTIL');
                } else {
                  pushToken(tokenTable[word], word);
                }

                break;
              case 'loop':
                nextWord = nextWord().toLowerCase();

                if (nextWord === 'while') {
                  read(n);
                  pushToken('Loop While', 'DO_LOOP_END_WHILE');
                } else if (nextWord === 'until') {
                  read(n);
                  pushToken('Loop Until', 'DO_LOOP_END_UNTIL');
                }

                break;
              case 'for':
                nextWord = nextWord();

                if (nextWord.toLowerCase() === 'each') {
                  read(n);
                  pushToken('For Each', 'FOR_EACHLOOP');
                } else {
                  pushToken(tokenTable[word], word);
                }

                break;
              case 'on':
                nextWord = nextWord();

                if (nextWord.toLowerCase() === 'error') {
                  nextWord1 = nextWord().toLowerCase();
                  nextWord2 = nextWord().toLowerCase();

                  if (nextWord1 === 'resume' && nextWord2 === 'next') {
                    read(n);
                    pushToken('On Error Resume Next',
                      'ON_ERROR_RESUME_NEXT');
                  } else if (nextWord1 === 'goto' && nextWord2 === '0') {
                    read(n);
                    pushToken('On Error GoTo 0', 'ON_ERROR_GOTO_0');
                  } else {
                    pushToken(tokenTable[word], word);
                  }
                }

                break;
              case 'case':
                nextWord = nextWord();

                if (nextWord.toLowerCase() === 'else') {
                  read(n);
                  pushToken('Case Else', 'CASE_ELSE');
                } else {
                  pushToken(tokenTable[word], word);
                }

                break;
              case 'select':
                nextWord = nextWord();

                if (nextWord.toLowerCase() === 'case') {
                  read(n);
                  pushToken('Select Case', 'SELECT_CASE');
                } else {
                  pushToken(word, tokenTable[word]);
                }

                break;
              case 'end':
                nextWord = nextWord();

                switch (nextWord.toLowerCase()) {
                  case 'function':
                    read(n);
                    pushToken('End Function', 'END_FUNCTION');

                    break;
                  case 'class':
                    read(n);
                    pushToken('End Class', 'END_CLASS');
                    break;
                  case 'sub':
                    read(n);
                    pushToken('End Sub', 'END_SUB');

                    break;
                  case 'property':
                    read(n);
                    pushToken('End Property', 'END_PROPERTY');

                    break;
                  case 'if':
                    read(n);
                    pushToken('End If', 'END_IF');

                    break;
                  case 'with':
                    read(n);
                    pushToken('End With', 'END_WITH');

                    break;
                  case 'select':
                    read(n);
                    pushToken('End Select', 'END_SELECT');

                    break;
                  default:
                    //RecordToken(tokenTable[word], oKeywordsCaseDict[word]);
                    pushToken(tokenTable[word], word);

                    break;
                }
                break;
              case 'exit':
                nextWord = nextWord();

                switch (nextWord.toLowerCase()) {
                  case 'function':
                    read(n);
                    pushToken('Exit Function', 'EXIT_FUNCTION');

                    break;
                  case 'for':
                    read(n);
                    pushToken('Exit For', 'EXIT_FOR');

                    break;
                  case 'do':
                    read(n);
                    pushToken('Exit Do', 'EXIT_DO');

                    break;
                  case 'property':
                    read(n);
                    pushToken('Exit Property', 'EXIT_PROPERTY');

                    break;
                  case 'sub':
                    read(n);
                    pushToken('Exit Sub', 'EXIT_SUB');

                    break;
                  default:
                    pushToken(tokenTable[word], word);
                    break;
                }

                break;
              default:
                //RecordToken(tokenTable[word], oKeywordsCaseDict[word]);
                pushToken(word, tokenTable[word]);
                break;
            }
          } else {
            /*switch (word.toUpperCase()) {
                            case 'REM':
                                pushToken(word + readLine(), 'COMMENT');
                                break;
                            default:*/
            pushToken(word, 'UNKNOWN');
            /*break;
                        }*/
          }

          break;
      }
    }

    pushToken(null, 'EOF');
  })();
  //Syntatic Analysis
  (function vbsparser_analizer() {
    var
      i = 0,
      lTokens = tokens.length,
      n = 0,
      curToken = null,
      curTokenType = null,
      nextToken = null,
      skipToToken = function vbsparser_analizer_skipToToken(tokenType) {
        while (n < lTokens && tokenType[n++] !== tokenType);
      },
      skipToEndOfStatement = function vbsparser_analizer_skipToEndOfStatement(fn) {
        var
          bIgnoreNewLine = false,
          tokenType = null;

        while (n < lTokens) {
          tokenType = tokenTypes[++n];
          switch (tokenType) {
            case 'STATEMENT_SEPARATOR':
              if (options.breakOnSeperator) return;
              break;
            case 'NEWLINE':
              if (bIgnoreNewLine) {
                bIgnoreNewLine = false;
                break;
              }
              return;
            case 'STATEMENT_CONTINUATION':
              bIgnoreNewLine = true;
              break;
            default:
              break;
          }
          if (fn) {
            fn(tokenType);
          }
        }
      };

    for (i = 0; i < lTokens - 1; i++) {

      curToken = tokens[i];
      curTokenType = tokenTypes[i];

      n = i + 1;

      while (n < lTokens && tokenTypes[n] === 'WHITESPACE') {
        n++;
      }

      if (n < lTokens) {
        nextToken = tokens[n];
      }

      n = i;

      switch (curTokenType) {
        case 'IF':
          skipToToken('THEN');
          while (tokenTypes[++n] !== 'NEWLINE') {
            if (!(tokenTypes[n] === 'WHITESPACE' || tokenTypes[n] ===
                'COMMENT')) {
              tokenTypes[n] = 'IF_ELSE_ONE_LINE';
              break;
            }
          }
          break;
        case 'DIM':
          skipToEndOfStatement(function(t) {
            if (t === 'UNKNOWN') {
              tokenTypes[n] = 'VARIABLE_NAME';
            }
          });

          break;
        case 'CLASS':
          skipToToken('UNKNOWN');

          tokens[n] = 'CLASS_NAME';

          break;
        case 'REDIM':
          skipToEndOfStatement(function(t) {
            if (t === 'UNKNOWN') {
              tokenTypes[n] = 'VARIABLE_NAME';
            }
          });

          break;
        case 'CONST':
          skipToEndOfStatement(function(t) {
            if (t === 'UNKNOWN') {
              tokenTypes[n] = 'CONST_VARIABLE_NAME';
            }
          });

          break;
        default:
          break;
      }

      lastTokenType = tokenTypes[i];

      if (lastTokenType != 'WHITESPACE' && lastTokenType != 'NEWLINE') {
        lastNonWSParsedToken = lastTokenType;
      }

      if (lastNonWSParsedToken != 'WHITESPACE') {
        lastNonWSLNParsedToken = lastTokenType;
      }
    }
  })();

  return {
    tokens: tokens,
    tokenTypes: tokenTypes
  };
};
module.exports = vbsparser;
