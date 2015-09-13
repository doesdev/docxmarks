var DocxMarks = require('./../index')
var path = require('path')
var testDocxIn = path.join(__dirname, 'assets', 'docx', 'test_in.docx')
var testDocxOut = path.join(__dirname, 'assets', 'docx', 'test_out.docx')

var docxMarks = new DocxMarks(testDocxIn)

docxMarks
  .on('error', function (error) { return console.log(error.stack) })
  .on('ready', function () {
    docxMarks.update({
      'DATE': '01/01/2015',
      'FIRST': function (v) { return v + ' is his name' },
      'FIRST_AGAIN': 'JERRY',
      'MIDDLE': 'S.',
      'LAST': 'Jones',
      'LAST_AGAIN': 'JONES'
    }, testDocxOut)
  })
  .on('updated', function () { return console.log('saved succesfully') })
