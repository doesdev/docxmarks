'use strict'

// setup
const jszip = require('jszip')
const matcher = /<w:bookmarkStart.+?(?:(?:(?:w:id="(.+?)").+?(?:w:name="(.+?)"))|(?:(?:w:name="(.+?)").+?(?:w:id="(.+?)"))).+?<w:bookmarkEnd.+?(?:w:id="(\1|\4)").+?>/g
const defTags = '<w:r><w:t></w:t></w:r>'
const primitive = (v) => {
  return ['string', 'number', 'boolean', 'undefined'].includes(typeof v)
}
const coal = (v, alt) => (v === 0 || v === false) ? v : (v || alt)
const charConv = {
  '<': `&lt;`,
  '>': `&gt;`,
  '"': `&quot;`,
  '\'': `&apos;`,
  '&': `&amp;`
}
const convChar = {
  '&lt;': '<',
  '&gt;': '>',
  '&quot;': '"',
  '&apos;': `'`,
  '&amp;': '&'
}
const xmlCleanRgx = /(?!&lt;|&gt;|&quot;|&apos;|&amp;)[<>"'&]/g
const xmlClean = (v) => v.replace(xmlCleanRgx, (m) => charConv[m])
const xmlDirtyRgx = /&lt;|&gt;|&quot;|&apos;|&amp;/g
const xmlDirty = (v) => v.replace(xmlDirtyRgx, (m) => convChar[m])

// helpers
const getType = (o) => {
  if (typeof o === 'string') return 'base64'
  if (typeof Buffer === 'function' && o instanceof Buffer) return 'nodebuffer'
  if (typeof Uint8Array === 'function' && o instanceof Uint8Array) {
    return 'uint8array'
  }
  if (typeof ArrayBuffer === 'function' && o instanceof ArrayBuffer) {
    return 'arraybuffer'
  }
  return 'blob'
}

const getReplacer = (marks, found, ids) => {
  return (match, id, name, name2, id2) => {
    id = id || id2
    name = name || name2
    found[name] = true
    ids.push(id)
    if (!marks._dxmGetter && !marks[name]) return match
    let realValue = match.replace(/(.*<w:t(?:>|\s.+?>))(.*?)(<\/w:t>.*)/, `$1${marks[name].setter()}$3`)
    return realValue;
  }
}

// main
module.exports = (docx, marks) => {
  let append
  let bookmarks = {}
  if (!marks) marks = { _dxmGetter: (k, v) => { bookmarks[k] = v } }
  marks = Object.assign({}, marks)
  Object.entries(marks).forEach(([k, v]) => {
    if (k === '_dxmGetter') return
    if (primitive(v)) marks[k] = { setter: () => xmlClean(`${coal(v, '')}`) }
    else if (typeof v === 'function') marks[k] = { setter: v }
    else if (v.hasOwnProperty('setter') && typeof v.setter === 'string') {
      marks[k] = Object.assign({}, v, { setter: () => v.setter })
    } else delete marks[k]
    append = append || marks[k].append
  })
  let type = getType(docx)
  let found = {}
  let ids = []
  let replacer = getReplacer(marks, found, ids)
  return jszip.loadAsync(docx, { base64: type === 'base64' }).then((zip) => {
    let files = Object.keys(zip.files).filter((k) => k.match(/^word\/.+\.xml$/))
    let replace = (f) => {
      return zip.file(f).async('string').then((text) => {
        let ogTmp
        let tmp = (text.match(matcher) || [])[0]
        while (tmp && (tmp = (tmp.substr(1).match(matcher) || [])[0])) {
          ogTmp = tmp
          tmp = tmp.replace(matcher, replacer)
          text = text.replace(ogTmp, tmp)
        }
        text = text.replace(matcher, replacer)
        zip.file(f, text)
        return Promise.resolve()
      })
    }
    let finish = () => {
      if (marks._dxmGetter) return Promise.resolve(bookmarks)
      if (!append) return zip.generateAsync({ type })
      let doc = 'word/document.xml'
      return zip.file(doc).async('string').then((text) => {
        ids = ids.map((i) => parseInt(i, 10)).sort((a, b) => b - a)
        let lastId = ids[0] || 0
        Object.entries(marks).forEach(([k, v]) => {
          if (found[k] || !v.append) return
          let id = lastId = lastId + 1
          let val = v.setter('')
          let start = `<w:p><w:bookmarkStart w:id="${id}" w:name="${k}"/>`
          let content = `<w:r><w:t xml:space="preserve">${val}</w:t></w:r>`
          let end = `<w:bookmarkEnd w:id="${id}"/></w:p>/</w:body>`
          text = text.replace(/<\/w:body>/, `${start}${content}${end}`)
        })
        zip.file(doc, text)
        return zip.generateAsync({ type })
      })
    }
    return Promise.all(files.map(replace)).then(finish)
  })
}
