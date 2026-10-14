/** Real Node/Linux smoke test for the shared original PPTX engine, outside Electron/jsdom. */
import assert from 'node:assert/strict'
import { createHash, randomUUID } from 'node:crypto'
import { writeFile } from 'node:fs/promises'
import { createBlankPptx, openPptx, getSlideNotes } from '@genoffice/pptx-engine'
import { opNames } from '@genoffice/pptx-ops'
import { HeuristicMetrics } from '@genoffice/pptx-render'
import { PresentationRuntime } from '../src/main/presentation-runtime'

const host = { getFontMetrics: () => new HeuristicMetrics() }
const original = await createBlankPptx()
const baseline = await openPptx(original)
const runtime = await PresentationRuntime.open(original, host)
assert.equal(runtime.read().dirty, false)
const text = {
  op: 'addElement',
  target: { slide: 0 },
  kind: 'textbox',
  offset: { x: 914400, y: 914400, cx: 7315200, cy: 914400 },
  paragraphs: [{ runs: [{ text: 'GenOffice 原生 PPTX 接入', fontSize: 28, bold: true }] }],
}
const mutation = {
  kind: 'apply' as const,
  transaction: {
    ops: [
      text,
      {
        op: 'addElement',
        target: { slide: 0 },
        kind: 'roundRect',
        offset: { x: 914400, y: 2743200, cx: 4572000, cy: 914400 },
        fill: '#4472C4',
        paragraphs: [{ runs: [{ text: '共享事务、撤销与原格式保存', color: '#FFFFFF' }] }],
      },
      { op: 'setNotes', target: { slide: 0 }, text: '原生 Node 引擎验收；Web 编辑器接线待完成。' },
      { op: 'duplicateSlide', target: { slide: 0 } },
    ],
  },
}
const requestId = randomUUID()
const result = runtime.mutate(requestId, 0, mutation)
assert.equal(result.result.applied, true)
assert.equal(runtime.read().slides.length, 2)
assert.equal(runtime.read().dirty, true)
assert.deepEqual(runtime.mutate(requestId, 0, mutation), result)
assert.equal(runtime.read().revision, 1)
runtime.mutate(randomUUID(), 1, { kind: 'undo' })
assert.equal(runtime.read().slides.length, 1)
assert.equal(runtime.read().slides[0]!.nodes.length, 0)
assert.equal(runtime.read().dirty, false)
runtime.mutate(randomUUID(), 2, { kind: 'redo' })
assert.equal(runtime.read().slides.length, 2)
const saved = await runtime.prepareSave(3)
assert.equal(runtime.read().dirty, true)
assert.throws(() => runtime.acceptSaved(saved.token, 3, 'not-published'), /PPTX_SAVE_CONFLICT/)
runtime.acceptSaved(saved.token, 3, saved.sha256)
assert.equal(runtime.read().dirty, false)
const reopened = await openPptx(saved.bytes)
assert.equal(reopened.deck.slides.length, 2)
assert.equal(reopened.deck.slides[0]!.elements.length, 2)
assert.match(JSON.stringify(reopened.deck.slides[0]!.elements), /GenOffice 原生 PPTX 接入/)
assert.match(getSlideNotes(reopened.archive, reopened.deck.slides[0]!.path), /Web 编辑器接线待完成/)
const preserved: string[] = []
for (const [path, bytes] of baseline.archive.entries) {
  if (!/^ppt\/(theme|slideMasters|slideLayouts)\//.test(path)) continue
  assert.deepEqual(reopened.archive.entries.get(path), bytes, path)
  preserved.push(path)
}
assert.equal((await openPptx(original)).deck.slides[0]!.elements.length, 0)
await writeFile('pptx-native-sample.pptx', saved.bytes)
const report = {
  at: new Date().toISOString(),
  runtime: process.version,
  platform: process.platform,
  nativeOpNames: opNames(),
  nativeOpCount: opNames().length,
  checks: [
    'blank-clean',
    'native-transaction',
    'text-shape-notes',
    'two-slides',
    'request-deduplication',
    'undo-to-pristine',
    'redo',
    'publish-acknowledgement',
    'native-save-reopen',
    'immutable-input',
    'untouched-theme-master-layout',
  ],
  preserved,
  bytes: saved.bytes.length,
  sha256: createHash('sha256').update(saved.bytes).digest('hex'),
  fontMetrics:
    'upstream HeuristicMetrics for pure engine smoke; original SystemFontMetrics verified separately in service integration',
  browserAccepted: false,
  actualPowerPointAccepted: false,
}
await writeFile('pptx-runtime-result.json', JSON.stringify(report, null, 2) + '\n')
console.log(JSON.stringify(report))
