# vbspretty [![GitHub license](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/lenilsondc/vbspretty/blob/master/LICENSE) [![npm version](https://img.shields.io/npm/v/vbspretty.svg?style=flat)](https://www.npmjs.com/package/vbspretty) 

A sophisticated VBScript parser and beautifier powered by nodejs.

## Usage

```shell
npm i vbspretty
```
As command-line (See command line options at [CLI](#cli-usage))

```shell
npx vbspretty ./MyApp.vbs
```

### Nodejs usage

```js
const fs = require('fs');
const vbspretty = require('vbspretty');

const source = fs.readFileSync('./MyApp.vbs').toString();

var sourcePretty = vbspretty({
  level: 0,
  indentChar: '\t',
  breakLineChar: '\r\n',
  breakOnSeperator: false,
  removeComments: false,
  source: source,
});

fs.writeFileSync('./MyAppPretty.vbs', sourcePretty);
```

### CLI usage

Cli accepts all options from the [api](#api) plus an `--output` option to provide a different file to output formatted version, if `--output` is omitted, the input file will be overwritten.

First param should always be the input file and it's mandatory, other params are optionals to configure vbspretty options. See full example bellow.

```shell
vbspretty MyApp.vbs --level 0 --indentChar "\t" --breakLineChar "\r\n" --breakOnSeperator --removeComments --output ./MyAppPretty.vbs
```

## API

|Options|Type|Default|Description|
|---|---|---|---|
|level|`number`|`0`|Indent level to start off|
|indentChar|`String`| "<kbd>space</kbd><kbd>space</kbd>"|Indent character (e.g., `\t`, <kbd>space</kbd><kbd>space</kbd>)|
|breakLineChar|`String`|`"\n"`| Break line character (e.g., `\n`, `\r\n`)|
|breakOnSeperator|`boolean`|`false`| Whether it breaks the line on occurrences of the `":"` statement separator.|
|removeComments|`boolean`|`false`|Whether it removes comments from the input.|
