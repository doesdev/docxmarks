'use strict'

import test from 'ava'
import fs from 'fs'
import path from 'path'
import docxmarks from './index'
const docx = fs.readFileSync(path.resolve(__dirname, 'test.docx'))

test(`test a bunch of bookmarks`, async (assert) => {
  let old = {
    FIRST: 'John',
    LAST: 'Doe',
    DATE: '12/01/2000',
    FIRST_AGAIN: 'John',
    MIDDLE: '',
    SPACES_3: '   ',
    LAST_AGAIN: 'Doe',
    MASSIVE_TEXT: 'This is a whole bunch of freaking text, even a new paragraph.And here’s that new paragraph. All inside of a bookmark.'
  }
  let replacements = {
    FIRST: 'Andrew',
    LAST: 'Carpenter',
    DATE: '12/01/2017',
    FIRST_AGAIN: 'Jerry',
    MIDDLE: 'R',
    SPACES_3: '   ',
    LAST_AGAIN: 'Smith',
    MASSIVE_TEXT: 'This is some textage, not so massive.'
  }
  let updated = {}
  let updaters = {}
  Object.keys(old).forEach((k) => {
    updaters[k] = (v) => {
      updated[k] = v
      return replacements[k]
    }
  })
  let newDx = await docxmarks(docx, updaters)
  Object.keys(old).forEach((k) => assert.is(updated[k], old[k]))
  await docxmarks(newDx, updaters)
  Object.keys(old).forEach((k) => assert.is(updated[k], replacements[k]))
})

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
