import assert from 'node:assert/strict'
import { randomUUID, createHash } from 'node:crypto'
import { writeFile } from 'node:fs/promises'
import { openPptx } from '@genoffice/pptx-engine'
import { PresentationService } from '../src/main/presentation-service'

const service = new PresentationService(),
  document = randomUUID(),
  browser = randomUUID()
let owner = document + ':' + browser,
  sessionId = '',
  revision = 0
const call = async (
  method: string,
  payload: Record<string, unknown> = {},
  requestId = randomUUID(),
) => {
  const value = await service.dispatch({
    owner,
    requestId,
    method,
    payload: { sessionId, expectedRevision: revision, ...payload },
  })
  if (value?.state) revision = value.state.revision
  else if (typeof value?.revision === 'number') revision = value.revision
  return value
}
const opened = await call('open', { revision: 0, sourceVersion: '', fitWidthPx: 1280 })
sessionId = opened.sessionId
assert.equal(opened.dirty, false)
assert.equal((await call('rpc', { channel: 'slides:redo', args: [] })).value, null)
assert.equal(revision, 0)
assert.equal(opened.defaultFont, 'Calibri')
assert.equal(opened.channels.length, 117)
const text = await call('rpc', {
  channel: 'slides:add-element',
  args: [
    {
      slideIndex: 0,
      kind: 'textbox',
      xPx: 80,
      yPx: 70,
      wPx: 700,
      hPx: 70,
      fitWidthPx: 1280,
      paragraphs: [
        {
          runs: [
            { text: 'Native font host and Office draft', fontFamily: 'Calibri', fontSize: 24 },
          ],
        },
      ],
    },
  ],
})
assert.ok(text.value.sourceId)
assert.equal(text.state.dirty, true)
assert.match(text.value.slide.nodes[0].text.lines[0].runs[0].fontFamily, /Carlito/)
const picture = await call('rpc', {
  channel: 'slides:insert-image',
  args: [0, 1280],
  asset: {
    name: 'sample.png',
    base64:
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aSioAAAAASUVORK5CYII=',
  },
})
assert.ok(picture.value.sourceId)
await call('rpc', {
  channel: 'slides:add-table',
  args: [
    { slideIndex: 0, rows: 2, cols: 2, xPx: 80, yPx: 200, wPx: 480, hPx: 120, fitWidthPx: 1280 },
  ],
})
await call('rpc', {
  channel: 'slides:add-chart',
  args: [
    {
      slideIndex: 0,
      kind: 'bar',
      title: 'Native chart',
      categories: ['A', 'B'],
      series: [{ name: 'Series', values: [10, 20] }],
      xPx: 640,
      yPx: 210,
      wPx: 400,
      hPx: 260,
      fitWidthPx: 1280,
    },
  ],
})
const first = await call('prepareSave')
const second = await call('prepareSave')
assert.equal(first.bytes, second.bytes)
const saved = Buffer.from(second.bytes, 'base64')
await call('acceptSaved', { token: second.token, sha256: second.sha256, sourceVersion: '1' })
assert.equal((await call('read')).dirty, false)
const undone = await call('rpc', { channel: 'slides:undo', args: [] })
assert.equal(undone.state.canRedo, true)
assert.equal((await call('read')).dirty, true)
const redone = await call('rpc', { channel: 'slides:redo', args: [] })
assert.equal(redone.state.canRedo, false)
assert.equal((await call('read')).dirty, false)
assert.equal((await call('rpc', { channel: 'slides:redo', args: [] })).value, null)
const receiptId = randomUUID(),
  operation = {
    expectedRevision: revision,
    ops: [
      {
        op: 'setNotes',
        target: { slide: 0 },
        text: 'Native actor, native history and native save.',
      },
    ],
  }
await call('apply', operation, receiptId)
const once = revision
await call('apply', operation, receiptId)
assert.equal((await call('read')).revision, once)
assert.equal((await call('query', { requestId: receiptId })).state, 'completed')
const prior = owner
owner = document + ':' + randomUUID()
const takeover = await call('open', { revision: once, sourceVersion: '1' })
assert.equal(takeover.sessionId, sessionId)
await assert.rejects(
  service.dispatch({
    owner: prior,
    requestId: randomUUID(),
    method: 'read',
    payload: { sessionId },
  }),
  /PPTX_SESSION_EXPIRED/,
)
const final = await call('prepareSave'),
  bytes = Buffer.from(final.bytes, 'base64'),
  deck = await openPptx(bytes)
assert.equal(deck.deck.slides[0]!.elements.length, 4)
const pristine = await service.dispatch({
  owner: randomUUID() + ':' + randomUUID(),
  requestId: randomUUID(),
  method: 'open',
  payload: { revision: 0, sourceVersion: '1' },
  bytes: final.bytes,
})
assert.equal(pristine.dirty, false)
await writeFile('pptx-service-sample.pptx', bytes)
const report = {
  at: new Date().toISOString(),
  nativeChannels: opened.channels.length,
  fontHost: 'Original SystemFontMetrics, Carlito and HarfBuzz',
  nativeElements: deck.deck.slides[0]!.elements.map((x) => x.type),
  revision,
  bytes: bytes.length,
  sha256: createHash('sha256').update(bytes).digest('hex'),
  checks: [
    'text-picture-table-chart',
    'same-revision-byte-cache',
    'undo-across-save',
    'redo-to-saved',
    'request-deduplication',
    'query-without-replay',
    'browser-takeover-fencing',
    'native-save-reopen-pristine',
  ],
}
await writeFile('pptx-service-result.json', JSON.stringify(report, null, 2) + '\n')
console.log(JSON.stringify(report))
