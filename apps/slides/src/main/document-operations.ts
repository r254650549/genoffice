/** Original native transaction funnel and text autofit post-pass, shared by both hosts. */
import { runTxn, type OpRecord, type TxnRequest, type TxnResult } from '@genoffice/pptx-ops'
import { matchesElementRef } from '@genoffice/pptx-engine/identity'
import {
  materializeSlide,
  slideDurableId,
  patchBodyPrAutofit,
  type Slide,
  type TextElement,
} from '@genoffice/pptx-engine'
import {
  layoutText,
  makeViewport,
  EMU_PER_PX_96,
  type RenderSlide,
  type RenderTextLayout,
} from '@genoffice/pptx-render'
import type { ApplyTxnOp, ApplyTxnResult } from '../shared/ipc'
import type { Session, OpLogEntry, SlidesDocumentRuntime } from './session-document'

export function createSlidesDocumentOperations(runtime: SlidesDocumentRuntime) {
  const {
    pushHistory,
    journalOps,
    scheduleDeckBroadcast,
    markMetaDirty,
    rebuildSlide,
    rebuildSlideWithReparse,
    buildAllRenderSlides,
    getFontMetrics,
  } = runtime
  function findEl(slide: Slide, sourceId: string): TextElement | undefined {
    const el = slide.elements.find((e) => matchesElementRef(e, sourceId))
    if (el && (el.type === 'text' || el.type === 'shape')) return el as TextElement
    return undefined
  }

  // The single funnel for non-dry transactions in this module: every applied
  // batch lands in the session's op journal (collab groundwork).
  function journaledTxn(
    session: Session,
    source: Exclude<OpLogEntry['source'], 'reset'>,
    req: TxnRequest,
  ): TxnResult {
    const r = runTxn(session.opened, req)
    if (r.applied) {
      journalOps(session, source, r.records ?? [])
      scheduleDeckBroadcast(session)
    }
    return r
  }

  /**
   * AI batch surface core, shared by the `slides:apply-txn` IPC handler and the
   * shell's MCP slides bridge (one implementation so both stay behaviorally
   * identical): raw ops arrive as one transaction. The registry validates (guided
   * errors), the executor owns atomicity/rollback/journal; dry-run rehearses the
   * plan without touching the deck or its history.
   */
  function applySessionTxn(session: Session, req: ApplyTxnOp): ApplyTxnResult | null {
    const ops = Array.isArray(req?.ops) ? (req.ops as Parameters<typeof runTxn>[1]['ops']) : []
    if (ops.length === 0 || ops.length > 50) {
      return {
        applied: false,
        failures: [
          { index: 0, error: 'ops must be a non-empty array (at most 50 per transaction).' },
        ],
      }
    }
    const isolation = req.isolation === 'per_op' ? ('per_op' as const) : ('atomic' as const)
    const compact = (fails?: Array<{ index: number; error: string }>) =>
      fails?.map((f) => ({ index: f.index, error: f.error }))
    if (req.dryRun) {
      const r = runTxn(session.opened, { ops, isolation, dryRun: true })
      return {
        applied: false,
        dryRun: true,
        plan: r.plan ?? [],
        ...(r.failures?.length ? { failures: compact(r.failures) } : {}),
      }
    }
    // Plan before pushing history (a no-op request must not clear the redo stack)
    const plan = runTxn(session.opened, { ops, isolation, dryRun: true })
    const invalid = plan.failures?.length ?? 0
    if (isolation === 'atomic' ? invalid > 0 : invalid >= ops.length) {
      return { applied: false, failures: compact(plan.failures) }
    }
    pushHistory(session)
    const r = journaledTxn(session, 'batch', { ops, isolation })
    if (!r.applied) {
      session.undoStack.pop()
      return { applied: false, failures: compact(r.failures) }
    }
    // Some ops change only package state (setNotes, a theme commit) and leave no
    // element dirty, so without this the session would still look clean and a
    // close could discard the edit. Element-level ops set their own flags; this
    // covers the archive-only ones.
    markMetaDirty(session)
    // Post-pass mirroring the dedicated shims (autofit/reparse are render concerns and live
    // outside the executor): text ops get autofit resize + fontScale write-back, level changes
    // materialize, and XML-patching ops reparse the page so the final render reflects them.
    // Slides are re-found by the executor-stamped durable id: a numeric target.slide drifts
    // when a later structural op (deleteSlide/moveSlide/duplicateSlide) shifts pages.
    const slideIdxOf = (rec: OpRecord): number => {
      if (rec.slideId)
        return session.opened.deck.slides.findIndex((s) => slideDurableId(s) === rec.slideId)
      return -1
    }
    const renderedByIdx = new Map<number, ReturnType<typeof rebuildSlide>>()
    for (const rec of r.records ?? []) {
      const o = rec.op
      const idx = slideIdxOf(rec)
      if (idx < 0) continue
      if (o.op === 'setTableStyle' || o.op === 'setChart') {
        rebuildSlideWithReparse(session, idx)
        renderedByIdx.delete(idx)
        continue
      }
      const id = o.target?.el
      if (!id || o.group) continue
      if (o.op !== 'setText' && o.op !== 'setFont' && o.op !== 'setParagraphFormat') continue
      if (o.op === 'setText' && (rec.after as { levelDirty?: boolean } | undefined)?.levelDirty)
        continue
      if (
        o.op === 'setParagraphFormat' &&
        (o.format as { indentDelta?: number } | undefined)?.indentDelta
      ) {
        materializeSlide(session.opened, idx)
        renderedByIdx.delete(idx)
        continue
      }
      let rendered = renderedByIdx.has(idx) ? renderedByIdx.get(idx)! : rebuildSlide(session, idx)
      rendered = applyAutofitResize(session, idx, id, rendered)
      rendered = syncAutofitScale(session, idx, id, rendered)
      renderedByIdx.set(idx, rendered)
    }
    return {
      applied: true,
      records: (r.records ?? []).map((rec) => ({
        op: rec.op.op,
        ...(rec.op.target
          ? { target: `${rec.op.target.slide}${rec.op.target.el ? `/${rec.op.target.el}` : ''}` }
          : {}),
        ...(rec.created ? { created: rec.created } : {}),
      })),
      ...(r.failures?.length ? { failures: compact(r.failures) } : {}),
      slides: buildAllRenderSlides(session.opened, session.fitWidthPx),
    }
  }

  /**
   * spAutoFit (autofit='resize', "resize shape to fit text"): after a text change, the box height
   * grows/shrinks with the content and is written back to cy; a wrap="none" body (PowerPoint's
   * click-to-type text box) also follows its widest line in cx. rendered = the rebuilt result
   * after this change; when the size changed, update the transform and rebuild once more.
   * Top-level elements only (group children use a different coordinate system, skip).
   */
  function applyAutofitResize(
    session: Session,
    slideIndex: number,
    sourceId: string,
    rendered: RenderSlide | null,
  ): RenderSlide | null {
    if (!rendered) return rendered
    const slide = session.opened.deck.slides[slideIndex]
    const el = slide ? findEl(slide, sourceId) : undefined
    if (!el?.text || el.text.autofit !== 'resize') return rendered
    const node = rendered.nodes.find((n) => n.sourceId === el.id)
    if (!node || (node.type !== 'shape' && node.type !== 'text') || !node.text) return rendered
    const text = node.text
    const needH = text.contentHeight + text.insets.t + text.insets.b
    const needW = !text.wrap && !text.vert ? nowrapContentWidth(text) : node.box.w
    if (Math.abs(needH - node.box.h) < 1 && Math.abs(needW - node.box.w) < 1) return rendered
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = session.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.max(Math.round((px / scale) * EMU_PER_PX_96), 1)
    // A nowrap box grows around its alignment: right-aligned text keeps the right edge, centered keeps the center
    const align = text.lines[0]?.align
    const shiftX =
      align === 'right' ? node.box.w - needW : align === 'center' ? (node.box.w - needW) / 2 : 0
    el.transform = {
      ...el.transform,
      offset: {
        ...el.transform.offset,
        x: el.transform.offset.x + Math.round((shiftX / scale) * EMU_PER_PX_96),
        cx: toEmu(needW),
        cy: toEmu(needH),
      },
    }
    el.dirtyTransform = true
    return rebuildSlide(session, slideIndex)
  }

  /** Box width a nowrap body needs: widest laid-out line plus the insets (an empty body keeps the insets alone). */
  function nowrapContentWidth(text: RenderTextLayout): number {
    let left = Infinity
    let right = -Infinity
    for (const line of text.lines)
      for (const run of line.runs) {
        left = Math.min(left, run.x)
        right = Math.max(right, run.x + run.widthPx)
      }
    const content = right > left ? right - left : 0
    return content + text.insets.l + text.insets.r
  }

  /**
   * normAutofit fontScale write-back: when the shrink ratio the layout actually used after a text
   * edit (≤ the stored cap, shrink-only) differs from the stored model value, sync the model and
   * patch the bodyPr attribute — only then does PowerPoint show the same size on open.
   * Triggered only by text edits (resize gestures do not write: the layout cap locks the stored
   * value, and writing back during a gesture would ratchet one way); top-level elements only.
   */
  function syncAutofitScale(
    session: Session,
    slideIndex: number,
    sourceId: string,
    rendered: RenderSlide | null,
  ): RenderSlide | null {
    if (!rendered) return rendered
    const slide = session.opened.deck.slides[slideIndex]
    const el = slide ? findEl(slide, sourceId) : undefined
    if (!el?.text || el.text.autofit !== 'shrink') return rendered
    const node = rendered.nodes.find((n) => n.sourceId === el.id)
    if (!node || (node.type !== 'shape' && node.type !== 'text') || !node.text) return rendered
    // The plain render honors the stored fontScale as-is (PowerPoint-on-open semantics);
    // after a text edit, re-run the layout with the autofit ladder enabled to find the
    // ratio PowerPoint would now use (still capped at the stored value).
    const refit = layoutText({
      body: el.text,
      boxWidthPx: node.box.w,
      boxHeightPx: node.box.h,
      metrics: getFontMetrics(),
      vp: makeViewport(session.opened.deck.size, session.fitWidthPx),
      refitAutofit: true,
    })
    const effective = refit.fontScale
    const effectiveRed = refit.lnSpcReduction ?? 0
    if (
      Math.abs(effective - (el.text.fontScale ?? 1)) < 0.005 &&
      Math.abs(effectiveRed - (el.text.lnSpcReduction ?? 0)) < 0.005
    )
      return rendered
    el.text.fontScale = effective
    if (effectiveRed) el.text.lnSpcReduction = effectiveRed
    else delete el.text.lnSpcReduction
    el.anchor.originalXml = patchBodyPrAutofit(el.anchor.originalXml, effective, effectiveRed)
    slide!.structureDirty = true
    // The render above used the pre-edit stored scale; redo it at the written-back value
    return rebuildSlide(session, slideIndex) ?? rendered
  }

  return { findEl, journaledTxn, applySessionTxn, applyAutofitResize, syncAutofitScale }
}
