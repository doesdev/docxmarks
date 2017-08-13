'use strict'

// setup
const jszip = require('jszip')
const matcher = /<w:bookmarkStart.+?(?:(?:(?:w:id="(.+?)").+?(?:w:name="(.+?)"))|(?:(?:w:name="(.+?)").+?(?:w:id="(.+?)"))).+?<w:bookmarkEnd.+?\/>/g
const defTags = '<w:r><w:t></w:t></w:r>'
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

// main
module.exports = async (docx, marks) => {
  let type = getType(docx)
  let zip = await jszip.loadAsync(docx, {base64: type === 'base64'})
  let files = Object.keys(zip.files).filter((k) => k.match(/^word\/.+\.xml$/))
  for (let f of files) {
    let text = await zip.file(f).async('string')
    text = text.replace(matcher, (match, id, name, name2, id2) => {
      id = id || id2
      name = name || name2
      if (!marks[name]) return match
      let wrap = (match.match(/(<w:r[> ].+<\/w:r>)/) || [])[0] || defTags
      let [raw, val] = (wrap.match(/<w:t(?:>|\s.+?>)(.*)<\/w:t>/) || []) || []
      val = !val ? '' : val.replace(/<.+?>/g, '')
      val = (typeof marks[name] === 'function') ? marks[name](val) : marks[name]
      let start = `<w:bookmarkStart w:id="${id}" w:name="${name}"/>`
      let content = `<w:t xml:space="preserve">${val}</w:t>`
      let end = `<w:bookmarkEnd w:id="${id}"/>`
      return `${start}${wrap.replace(raw, content)}${end}`
    })
    zip.file(f, text)
  }
  return await zip.generateAsync({type})
}
