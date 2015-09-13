// Setup
const util = require('util')
const Event = require('events').EventEmitter

module.exports = Bookmark

function Bookmark (xml, file) {
  var self = this
  function noXml () {
    self.emit('error', 'Initilized bookmark without xml content')
  }
  if (!xml) {
    setTimeout(noXml, 50)
    return self
  }
  self.xml = xml
  self.file = file
  setTimeout(self.extractProperties.bind(self), 0)
  return self
}

util.inherits(Bookmark, Event)

// CLASS METHODS
Bookmark.getArrayOfBookmarks = function (xml, file) {
  var event = new Event()
  var bookmarksOut = []
  var bookmarks = xml.match(/<w:bookmarkStart.+?<w:bookmarkEnd.*?\/>/g)
  bookmarks.forEach(function (b) {
    while (b.match(/bookmarkStart/g).length > b.match(/bookmarkEnd/g).length) {
      var rgx = new RegExp(b.replace(/\//g, '\\/') + '.*?<w:bookmarkEnd.*?\\/>')
      var match = xml.match(rgx)
      b = match ? match[0] : 'bookmarkStartbookmarkEnd'
    }
    if (b === 'bookmarkStartbookmarkEnd') return
    var bookmark = new Bookmark(b, file)
    bookmark.on('ready', function () {
      bookmarksOut.push(bookmark)
      if (bookmarksOut.length === bookmarks.length) {
        event.emit('data', bookmarksOut)
      }
    })
  })
  return event
}

// INSTANCE METHODS
// Extract bookmark properties
Bookmark.prototype.extractProperties = function () {
  var self = this
  self.id = self.xml.match(/^<w:bookmarkStart.*?w:id="(\d+)"/)[1]
  self.name = self.xml.match(/^<w:bookmarkStart.*?w:name="([^"]+)"/)[1]
  var oXmlRgx = /^<w:bookmarkStart.*?\/>(.*)<w:bookmarkEnd.+?\/>/
  self.origXmlValue = self.xml.match(oXmlRgx)[1]
  self.initValue = (self.xml.match(/<w:r.*?<w:t.*?>(.*?)<\/w:t>/) || [])[1]
  self.initValue = self.initValue || ''
  self.initTags = (self.xml.match(/<w:r.*?<\/w:r>/) || [])[0]
  self.initTags = self.initTags || '<w:r><w:t></w:t></w:r>'
  self.emit('ready')
}

// Replace bookmark value
Bookmark.prototype.setText = function (text) {
  var self = this
  var startTag = '' +
    '<w:bookmarkStart w:id="' + self.id + '" w:name="' + self.name + '"/>'
  var endTag = '<w:bookmarkEnd w:id="' + self.id + '"/>'
  var value = self.initTags.replace(
    /<w:t.*?<\/w:t>/, '<w:t xml:space="preserve">' + text + '</w:t>'
  )
  self.newXml = startTag + value + endTag
}
