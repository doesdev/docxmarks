// Setup
const Zip = require('jszip')
const fs = require('fs')
const util = require('util')
const Event = require('events').EventEmitter
const Bookmark = require('./bookmark')

module.exports = DocxMarks

function DocxMarks (path) {
  var self = this
  self.initialize()
  function noPath () { self.emit('error', 'Initilized document without path') }
  if (!path) {
    setTimeout(noPath, 50)
    return self
  }
  self.path = path
  self.startProcessing()
  return self
}

util.inherits(DocxMarks, Event)

// INSTANCE METHODS
// Initialize listeners
DocxMarks.prototype.initialize = function () {
  var self = this
  self.data = {}
  // Emit data when ready (i.e. checklist is of specified length)
  self.on('item-ready', function (item) {
    switch (item) {
      case 'extract':
        self.getBody()
        break
      case 'body':
        self.getHeaders()
        break
      case 'headers':
        self.getBookmarks()
        break
      case 'bookmarks':
        self.emit('data', self.data)
        self.emit('ready')
        break
      default: return
    }
  })
}

// Start initial processing
DocxMarks.prototype.startProcessing = function () {
  var self = this
  self.extract()
}

// Extract docx contents as Zip
DocxMarks.prototype.extract = function () {
  var self = this
  fs.readFile(self.path, function (err, data) {
    if (err) return self.emit('error', err)
    try {
      self.data.structure = new Zip(data)
      self.emit('item-ready', 'extract')
    } catch (e) {
      return self.emit('error', e)
    }
  })
}

// Write docx file to disk
DocxMarks.prototype.save = function (path) {
  var self = this
  self.updateBody()
  var newDocx = self.data.structure.generate({type: 'nodebuffer'})
  fs.writeFile(path || self.path, newDocx, function (err, data) {
    if (err) return self.emit('error', err)
    return self.emit('saved')
  })
}

// Get document XML as text
DocxMarks.prototype.getBody = function () {
  var self = this
  self.data.body = self.data.structure.file('word/document.xml').asText()
  self.emit('item-ready', 'body')
}

// Get document XML as text
DocxMarks.prototype.getHeaders = function () {
  var self = this
  self.data.headers = []
  var i = 1
  var xml
  while ((xml = self.data.structure.files['word/header' + i + '.xml'])) {
    xml = self.data.structure.file('word/header' + i + '.xml')
    self.data.headers.push({
      file: 'word/header' + i + '.xml',
      content: xml.asText()
    })
    i = i + 1
  }
  self.emit('item-ready', 'headers')
}

// Set document.xml content from self.data.body
DocxMarks.prototype.updateBody = function () {
  var self = this
  var files = {}
  self.data.bookmarks.forEach(function (b) {
    if (!b.newXml) return
    files[b.file] = files[b.file] || self.data.structure.file(b.file).asText()
    files[b.file] = files[b.file].replace(b.xml, b.newXml)
    self.data.structure.file(b.file, files[b.file])
  })
}

// Get array of bookmarkObjects from documentObject
DocxMarks.prototype.getBookmarks = function () {
  var self = this
  self.data.bookmarks = self.data.bookmarks || []
  var bkAryEventBody = Bookmark.getArrayOfBookmarks(
    self.data.body, 'word/document.xml'
  )
  var length = self.data.headers.length + 1
  var numberReady = 0
  bkAryEventBody.on('data', function (bkAry) {
    self.data.bookmarks = self.data.bookmarks.concat(bkAry)
    numberReady = numberReady + 1
    if (numberReady === length) self.emit('item-ready', 'bookmarks')
  })
  var bkAryEventsHeader = {}
  self.data.headers.forEach(function (h, i) {
    bkAryEventsHeader[i] = Bookmark.getArrayOfBookmarks(h.content, h.file)
    bkAryEventsHeader[i].on('data', function (bkAry) {
      self.data.bookmarks = self.data.bookmarks.concat(bkAry)
      numberReady = numberReady + 1
      if (numberReady === length) self.emit('item-ready', 'bookmarks')
    })
  })
}

// Convenience method for bookmark replacement and save
DocxMarks.prototype.update = function (name, text, path) {
  var self = this
  if (typeof name === 'object' && text && !path) path = text
  self.on('saved', function () { return self.emit('updated') })
  self.on('bookmark-text-set', function () {
    self.save(path)
  })
  self.setBookmarkText(name, text)
}

// Get bookmarkObject by name
DocxMarks.prototype.getBookmark = function (name) {
  var self = this
  return (self.data.bookmarks.filter(function (b) {
    return b.name === name
  }) || [null])[0]
}

// Update bookmark text by name
DocxMarks.prototype.setBookmarkText = function (name, text) {
  var self = this
  var bk
  if (typeof name === 'object') {
    for (var bkName in name) {
      bk = self.getBookmark(bkName)
      if (bk) {
        if (typeof name[bkName] === 'function') {
          name[bkName] = name[bkName](bk.initValue)
        }
        bk.setText(name[bkName])
      }
    }
  } else {
    bk = self.getBookmark(name)
    if (bk) {
      if (typeof name[bkName] === 'function') {
        name[bkName] = name[bkName](bk.initValue)
      }
      bk.setText(text)
    }
  }
  self.emit('bookmark-text-set')
}
