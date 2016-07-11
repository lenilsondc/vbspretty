var vbsbeautifier = function vbsbeautifier_(options) {
  var
    tokens = options.tokens || [],
    tokenTypes = options.tokenTypes || [],
    source = options.source || '',
    output = '',
    lastParsedToken = '',
    lastNonWSParsedToken = '';

  //beautify
  (function vbsbeautifier_beautify() {
    var
      i = 0,
      lTokens = tokens.length,
      n = 0,
      curToken = null,
      curTokenType = null,
      nextToken = null,
      nextTokenType = null,
      writeToken = function vbsbeautifier_beautify_writeToken() {
        var curTokenWSAfter, curTokenWSBefore, nextTokenWSBefore, nextTokenWSAfter;
            var curIndentDelta, curLineIndent, nextIndentDelta, nextLineIndent;

            GetTokenImpact(token, out curIndentDelta, out curLineIndent, out curTokenWSBefore, out curTokenWSAfter);

            //if next token is new line then lets not add a space after the current token
            curTokenWSAfter = nextTokenType == 'NEWLINE' ? "" : curTokenWSAfter;

            if (curTokenType != 'NEWLINE' /*&& !ignoreWSToken.Contains(curTokenType)*/)
            {
                GetTokenImpact(nextToken, out nextIndentDelta, out nextLineIndent, out nextTokenWSBefore, out nextTokenWSAfter);

                //Let's make sure we don't merge two different tokens because of spacing
                if (curTokenWSAfter == "" && nextTokenWSBefore == "")
                {
                    switch (nextcurTokenType)
                    {
                        case 'NEWLINE':
                        case 'DOT_OPERATOR':
                        case 'OPEN_BRACKET':
                        case 'CLOSE_BRACKET':
                        case 'COMMA':
                        case 'EOF':
                            break;
                        default:
                            curTokenWSAfter = " ";
                            break;
                    }

                }
            }


            switch (curTokenType)
            {
                 case 'STATEMENT_CONTINUATION':
                    bNextLineInContinuation = true;
                    break;
                case 'NEWLINE':
                    if (lastToken != null && lastNonWSLNParsedToken == 'NEWLINE')
                    {
                        bUseIndent = true;
                        AppendCode(GetIndent(curLineIndent) + "\r\n");
                    }
                    else
                        AppendCode("\r\n");
                    bUseIndent = true;
                    if (bNextLineInContinuation)
                    {
                        //we need to add some smart identation to make sure the continued line is extra indented
                        //int length = lastLine.Trim().Length;
                        //get tabe size
                        //length = (int) (length / 4 * 0.4);
                        dynamicIndent = "\t\t";
                        //indentTabs[length];
                        bNextLineInContinuation = false;
                    }
                    else
                        dynamicIndent = "";
                    break;
            }

            ChangeIndent(curIndentDelta);

            if (curTokenType == 'ELSE' && nextcurTokenType == 'IF')
                curTokenWSAfter = "\r\n";

            if (operatorTokens.Contains(curTokenType) &&
                operatorTokens.Contains(lastNonWSParsedToken))
            {
                curTokenWSAfter = "";
                curTokenWSBefore = "";

                if (curTokenType == 'BINARY_OPERATOR' && curToken === "Not")
                    curTokenWSAfter = " ";
            }

            if (options.removeComments && curTokenType == 'COMMENT')
            {
                //lets not do anything and remove this comment
                if (nextcurTokenType == 'NEWLINE')
                {
                    nextTokenType = 'UNKNOWN';
                    nextToken = "";
                }
            }
            else if ((curTokenType != 'WHITESPACE' && curTokenType != 'NEWLINE'))
            {
                AppendCode(GetIndent(curLineIndent) + curTokenWSBefore + token.value + curTokenWSAfter);
            }

            if (curTokenType == 'ELSE' && nextcurTokenType == 'IF')
            {
                //if have a mistaken else if and not ElseIf. Need to make sure
                //we add a new line to it
                bUseIndent = true;
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
        nextTokenType = tokenTypes[n];
      }

      n = i;

      if (curTokenType !== 'WHITESPACE'){
        writeToken(c);
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

  return output;
};
module.exports = vbsbeautifier;
