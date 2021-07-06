# vbspretty
Sophisticated VBScript beautifier 
## Usage

```shell
npm i @vbsnext/vbs-pretty
```
As command-line

```shell
npx vbs-pretty .\src\Excel.vbs
```

As nodejs module

```js
const vbspretty = require('@vbsnext/vbs-pretty')
var bsource = vbspretty({
    level: 1,
    indentChar: '\t',
    breakLineChar: '\r\n',
    breakOnSeperator: false,
    removeComments: false,
    source: require('fs').readFileSync('./src/Excel-unpretty.vbs').toString()
  });

  require('fs').writeFileSync('./src/Excel-pretty.vbs', bsource)
```