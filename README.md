# docxmarks   [![npm version](https://badge.fury.io/js/docxmarks.svg)](http://badge.fury.io/js/docxmarks)   [![js-standard-style](https://img.shields.io/badge/code%20style-standard-brightgreen.svg?style=flat)](https://github.com/feross/standard)   [![Dependency Status](https://dependencyci.com/github/doesdev/docxmarks/badge)](https://dependencyci.com/github/doesdev/docxmarks)

Replace text in native bookmarks in Open XML Document Files (.docx).

Formatting will follow the initial style inside the bookmark. That is to say,
if you have content inside the bookmark the first character defines the styling.
This is the same behavior supplied in Microsoft Word's libraries.

This library will also only handle text, nothing else.

I say all of that to say if you need more advanced formatting and content
options then you should be using
[docxtemplater](https://github.com/open-xml-templating/docxtemplater).

## Install
`npm i docxmarks --save`

## Usage

```javascript
const docxmarks = require('docxmarks')
const fs = require('fs')
const docx = fs.readFileSync('path/to/document.docx')

const replaceLast = (val) => val === 'Smith' ? 'Carpenter' : val
docxmarks(docx, {first: 'Andrew', last: replaceLast}).then((data) => {
  fs.writeFileSync('path/to/newDocument.docx', data)
})
```

## API

#### Takes docx data, replaces bookmarks, returns `Promise` resolving with new docx data in the same encoding as provided in input.

#### `docxmarks(*docxData, *replacements)`

- **docxData** *[base64 | Buffer | ArrayBuffer | Uint8Array - required]*
- **replacements** *[object - required]*
  - **key** - Name of bookmark to replace, is case sensitive
  - **value** *[string | function]* Value to replace bookmark with. If a function is supplied it will receive the text value currently in the bookmark and will use the function's return value as the replacement string.

## Upgrading

Version 2.0.0 is a complete re-write with 100% different API. Use new API if
upgrading from an old version, as there is no transitional API.


## License

MIT © [Andrew Carpenter](https://github.com/doesdev)
