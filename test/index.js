'use strict'

import test from 'ava'
import fs from 'fs'
import path from 'path'
import docxmarks from './../index'
const docx = fs.readFileSync(path.resolve(__dirname, 'assets', 'test.docx'))

test(`bookmark replace works with Buffer`, async (assert) => {
  let oldLast
  let newLast = 'Jones'
  let setOld = (last) => {
    oldLast = last
    return newLast
  }
  let newDx = await docxmarks(docx, {LAST: setOld})
  assert.is(oldLast, 'Doe')
  let finalDx = await docxmarks(newDx, {LAST: setOld})
  assert.is(oldLast, newLast)
  assert.true(finalDx instanceof Buffer)
})

test(`bookmark replace works with ArrayBuffer`, async (assert) => {
  let oldLast
  let newLast = 'Jones'
  let setOld = (last) => {
    oldLast = last
    return newLast
  }
  let newDx = await docxmarks(docx.buffer, {LAST: setOld})
  assert.is(oldLast, 'Doe')
  let finalDx = await docxmarks(newDx, {LAST: setOld})
  assert.is(oldLast, newLast)
  assert.true(finalDx instanceof ArrayBuffer)
})

test(`bookmark replace works with Uint8Array`, async (assert) => {
  let oldLast
  let newLast = 'Jones'
  let setOld = (last) => {
    oldLast = last
    return newLast
  }
  let newDx = await docxmarks(new Uint8Array(docx.buffer), {LAST: setOld})
  assert.is(oldLast, 'Doe')
  let finalDx = await docxmarks(newDx, {LAST: setOld})
  assert.is(oldLast, newLast)
  assert.true(finalDx instanceof Uint8Array)
})

test(`bookmark replace works with base64`, async (assert) => {
  let oldLast
  let newLast = 'Jones'
  let setOld = (last) => {
    oldLast = last
    return newLast
  }
  let newDx = await docxmarks(docx.toString('base64'), {LAST: setOld})
  assert.is(oldLast, 'Doe')
  let finalDx = await docxmarks(newDx, {LAST: setOld})
  assert.is(oldLast, newLast)
  assert.true(typeof finalDx === 'string')
})
