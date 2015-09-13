var DocxMarks = require('./../index')
var path = require('path')
var testDocxIn = path.join(__dirname, 'assets', 'docx', 'test_in.docx')
var testDocxOut = path.join(__dirname, 'assets', 'docx', 'test_out.docx')

var docxMarks = new DocxMarks(testDocxIn)

docxMarks.on('error', function (error) {
  return console.log(error.stack)
})

docxMarks.on('ready', function () {
  docxMarks.update({
    'DATE': '01/01/2015',
    'FIRST': 'Jerry',
    'FIRST_AGAIN': 'JERRY',
    'MIDDLE': 'S.',
    'LAST': 'Jones',
    'LAST_AGAIN': 'JONES'
  }, testDocxOut)
})

docxMarks.on('saved', function () { return console.log('saved succesfully') })
