import { randomUUID } from 'node:crypto'
import { describe, expect, it, vi } from 'vitest'
import { createBlankPptx, openPptx } from '@genoffice/pptx-engine'
import { HeuristicMetrics } from '@genoffice/pptx-render'
import { PresentationRuntime } from '../src/main/presentation-runtime'

const host = { getFontMetrics: () => new HeuristicMetrics() }
const text = {
  op: 'addElement',
  target: { slide: 0 },
  kind: 'textbox',
  offset: { x: 914400, y: 914400, cx: 7315200, cy: 914400 },
  paragraphs: [{ runs: [{ text: '原生 PPTX 接入验收', fontSize: 24 }] }],
}
const apply = (runtime: PresentationRuntime, ops: unknown[], dryRun = false) =>
  runtime.mutate(randomUUID(), runtime.read().revision, {
    kind: 'apply',
    transaction: { ops, dryRun },
  })

describe('Node presentation host with original GenOffice document operations', () => {
  it('keeps serialized bytes stable at the same revision across ZIP timestamps', async () => {
    vi.useFakeTimers({ toFake: ['Date'] })
    try {
      vi.setSystemTime(new Date('2026-10-14T00:00:00Z'))
      const runtime = await PresentationRuntime.open(null, host)
      apply(runtime, [text])
      const first = await runtime.prepareSave(1)
      vi.setSystemTime(new Date('2026-10-14T00:00:10Z'))
      const second = await runtime.prepareSave(1)
      expect(second.bytes).toEqual(first.bytes)
      expect(second.sha256).toBe(first.sha256)
    } finally {
      vi.useRealTimers()
    }
  })
  it('compares undo and redo with the last saved baseline, including undo across save', async () => {
    const runtime = await PresentationRuntime.open(null, host)
    apply(runtime, [text])
    const saved = await runtime.prepareSave(1)
    runtime.acceptSaved(saved.token, 1, saved.sha256)
    runtime.mutate(randomUUID(), 1, { kind: 'undo' })
    expect(runtime.read()).toMatchObject({ revision: 2, dirty: true })
    runtime.mutate(randomUUID(), 2, { kind: 'redo' })
    expect(runtime.read()).toMatchObject({ revision: 3, dirty: false })
  })
  it('opens a blank deck clean without naming or publishing a file', async () => {
    const runtime = await PresentationRuntime.open(null, host)
    expect(runtime.read()).toMatchObject({ revision: 0, dirty: false, canUndo: false })
    expect(runtime.read().slides).toHaveLength(1)
    const prepared = await runtime.prepareSave(0)
    expect((await openPptx(prepared.bytes)).deck.slides).toHaveLength(1)
    expect(runtime.read().dirty).toBe(false)
  })

  it('uses native transactions, dry-run, undo and redo on the same draft', async () => {
    const runtime = await PresentationRuntime.open(null, host)
    expect(apply(runtime, [text], true).result).toMatchObject({ dryRun: true, applied: false })
    expect(runtime.read()).toMatchObject({ revision: 0, dirty: false, canUndo: false })
    const result = apply(runtime, [text])
    expect(result).toMatchObject({ revision: 1, result: { applied: true } })
    expect(result.result.records?.[0]?.created).toHaveLength(1)
    expect(runtime.read().slides[0]!.nodes).toHaveLength(1)
    runtime.mutate(randomUUID(), 1, { kind: 'undo' })
    expect(runtime.read()).toMatchObject({ revision: 2, dirty: false, canRedo: true })
    expect(runtime.read().slides[0]!.nodes).toHaveLength(0)
    runtime.mutate(randomUUID(), 2, { kind: 'redo' })
    expect(runtime.read()).toMatchObject({ revision: 3, dirty: true })
    expect(runtime.read().slides[0]!.nodes).toHaveLength(1)
  })

  it('reuses the native atomic planner; rejected edits do not change the draft', async () => {
    const runtime = await PresentationRuntime.open(null, host)
    const result = apply(runtime, [text, { op: 'setText', target: { slide: 0, el: 'missing' } }])
    expect(result.result.applied).toBe(false)
    expect(result.result.failures).toHaveLength(1)
    expect(runtime.read()).toMatchObject({ revision: 0, dirty: false, canUndo: false })
    expect(runtime.read().slides[0]!.nodes).toHaveLength(0)
  })

  it('returns duplicate outcomes without repeating edits and refuses reused IDs', async () => {
    const runtime = await PresentationRuntime.open(null, host)
    const id = randomUUID()
    const mutation = { kind: 'apply' as const, transaction: { ops: [text] } }
    const first = runtime.mutate(id, 0, mutation)
    runtime.mutate(randomUUID(), 1, { kind: 'undo' })
    expect(runtime.mutate(id, 0, mutation)).toEqual(first)
    expect(runtime.read().slides[0]!.nodes).toHaveLength(0)
    expect(() => runtime.mutate(id, 2, mutation)).toThrow('PPTX_REQUEST_ID_REUSED')
  })

  it('retains a conflict receipt instead of applying it after the document changes', async () => {
    const runtime = await PresentationRuntime.open(null, host)
    const id = randomUUID()
    const mutation = { kind: 'apply' as const, transaction: { ops: [text] } }
    expect(() => runtime.mutate(id, 1, mutation)).toThrow('PPTX_DRAFT_CONFLICT')
    apply(runtime, [text])
    expect(() => runtime.mutate(id, 1, mutation)).toThrow('PPTX_DRAFT_CONFLICT')
    expect(runtime.read().slides[0]!.nodes).toHaveLength(1)
  })

  it('keeps dirty until the exact serialized revision has been published', async () => {
    const runtime = await PresentationRuntime.open(null, host)
    apply(runtime, [text])
    const prepared = await runtime.prepareSave(1)
    expect(runtime.read().dirty).toBe(true)
    expect(() => runtime.acceptSaved(prepared.token, 1, 'wrong digest')).toThrow(
      'PPTX_SAVE_CONFLICT',
    )
    expect(runtime.read().dirty).toBe(true)
    runtime.acceptSaved(prepared.token, 1, prepared.sha256)
    expect(runtime.read().dirty).toBe(false)
    const reopened = await openPptx(prepared.bytes)
    const element = reopened.deck.slides[0]!.elements[0]!
    expect(element.type).toBe('text')
    expect(JSON.stringify(element)).toContain('原生 PPTX 接入验收')
  })

  it('rejects acknowledging an old saved snapshot after a newer manual edit', async () => {
    const runtime = await PresentationRuntime.open(null, host)
    apply(runtime, [text])
    const prepared = await runtime.prepareSave(1)
    apply(runtime, [{ op: 'duplicateSlide', target: { slide: 0 } }])
    expect(() => runtime.acceptSaved(prepared.token, 1, prepared.sha256)).toThrow(
      'PPTX_SAVE_CONFLICT',
    )
    expect(runtime.read()).toMatchObject({ revision: 2, dirty: true })
    expect(runtime.read().slides).toHaveLength(2)
  })

  it('preserves untouched theme/master/layout bytes through original-format serialization', async () => {
    const input = await createBlankPptx()
    const baseline = await openPptx(input)
    const runtime = await PresentationRuntime.open(input, host)
    apply(runtime, [text])
    const saved = await runtime.prepareSave(1)
    const reopened = await openPptx(saved.bytes)
    for (const [path, bytes] of baseline.archive.entries) {
      if (/^ppt\/(theme|slideMasters|slideLayouts)\//.test(path))
        expect(reopened.archive.entries.get(path), path).toEqual(bytes)
    }
    expect((await openPptx(input)).deck.slides[0]!.elements).toHaveLength(0)
  })
})
