# docxmarks

Replace text in native bookmarks in Open XML Document Files (.docx).

Formatting will follow the initial style inside the bookmark. That is to say,
if you have content inside the bookmark the first character defines the styling.
This is the same behavior supplied in Microsoft Word's libraries.

This library will also only handle text, nothing else.

I say all of that to say if you need more advanced formatting and content
options then you should be using
[docxtemplater](https://github.com/open-xml-templating/docxtemplater).

## usage

Basic Example (overwrites the source document)
```javascript
var DocxMarks = require('docxmarks')
var docx = 'path/to/document.docx'

var docxMarks = new DocxMarks(testDocxIn)

docxMarks
  .on('error', function (error) { return console.log(error.stack) })
  .on('ready', function () {
    // docxMarks.update() - where the magic happens
    // Param 1 (bookmark_name OR object)
    // Param 2 (bookmark_text OR output_path if using object)
    // Param 3 (output path if not using object)
    docxMarks.update('BOOKMARK_NAME', 'Text to set bookmark to')
  })
  .on('updated', function () { return console.log('successfully updated') })

```

Advanced Example (save to new file, update multiple bookmarks, use callback)
```javascript
var DocxMarks = require('docxmarks')
var docx = 'path/to/document.docx'
var docxOut = 'path/to/document_out.docx'

var docxMarks = new DocxMarks(testDocxIn)

docxMarks
  .on('error', function (error) { return console.log(error.stack) })
  .on('ready', function () {
    // Use object to set {bookmark_name: value} for multiple bookmarks
    docxMarks.update({
      // Set value with callback which gets the original content as param
      'FIRST_NAME': function (v) { return v + ' is his first name' },
      'LAST_NAME': 'Carpenter'
    // Pass output path as last param to output to that file
    }, docxOut)
  })
  .on('updated', function () { return console.log('successfully updated') })

```
