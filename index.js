'use strict'

// setup
const jszip = require('jszip')
const matcher = /<w:bookmarkStart.+?(?:(?:(?:w:id="(.+?)").+?(?:w:name="(.+?)"))|(?:(?:w:name="(.+?)").+?(?:w:id="(.+?)"))).+?<w:bookmarkEnd.+?\/>/g
const defTags = '<w:r><w:t></w:t></w:r>'

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
    if (!marks[name]) return match
    let wrap = (match.match(/(<w:r[> ].+<\/w:r>)/) || [])[0] || defTags
    let [raw, val] = (wrap.match(/<w:t(?:>|\s.+?>)(.*)<\/w:t>/) || []) || []
    raw = raw || '<w:t></w:t>'
    val = !val ? '' : val.replace(/<.+?>/g, '')
    val = marks[name].setter(val)
    let start = `<w:bookmarkStart w:id="${id}" w:name="${name}"/>`
    let content = `<w:t xml:space="preserve">${val}</w:t>`
    let end = `<w:bookmarkEnd w:id="${id}"/>`
    return `${start}${wrap.replace(raw, content)}${end}`
  }
}

// main
module.exports = (docx, marks) => {
  let append
  Object.entries(marks).forEach(([k, v]) => {
    if (typeof v === 'string') marks[k] = {setter: () => v}
    else if (typeof v === 'function') marks[k] = {setter: v}
    else if (v.hasOwnProperty('setter') && typeof v.setter === 'string') {
      let oldVal = v.setter
      marks[k].setter = () => oldVal
    }
    append = append || marks[k].append
  })
  let type = getType(docx)
  let found = {}
  let ids = []
  let replacer = getReplacer(marks, found, ids)
  return jszip.loadAsync(docx, {base64: type === 'base64'}).then((zip) => {
    let files = Object.keys(zip.files).filter((k) => k.match(/^word\/.+\.xml$/))
    let replace = (f) => {
      return zip.file(f).async('string').then((text) => {
        text = text.replace(matcher, replacer)
        zip.file(f, text)
        return Promise.resolve()
      })
    }
    let finish = () => {
      if (!append) return zip.generateAsync({type})
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
        return zip.generateAsync({type})
      })
    }
    return Promise.all(files.map(replace)).then(finish)
  })
}
