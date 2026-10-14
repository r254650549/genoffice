/** Original Slides document state, history and render helpers, shared by Electron and Node hosts. */
import {
  materializeSlide,
  parseClrMap,
  parseTheme,
  resolveSchemeColor,
  type OpenedPptx,
  type Slide,
} from '@genoffice/pptx-engine'
import {
  buildRenderSlide,
  type FontMetricsProvider,
  type RenderSlide,
} from '@genoffice/pptx-render'
import { tiffToPng } from './tiff-decode'
import { neutralizeJpegOrientation } from './jpeg-orientation'
import { displayMime } from './media-mime'

export function deckDefaultFont(opened: OpenedPptx): string | undefined {
  try {
    const slidePath = opened.archive.readPresentation().slidePaths[0]
    if (!slidePath) return undefined
    const themePath = opened.archive.resolveSlideChain(slidePath).themePath
    const xml = themePath ? opened.archive.readText(themePath) : undefined
    return xml ? parseTheme(xml).minorFont : undefined
  } catch {
    return undefined
  }
}

export interface Session {
  path: string
  opened: OpenedPptx
  fitWidthPx: number
  undoStack: HistorySnapshot[]
  redoStack: HistorySnapshot[]
  /** Nested history transaction used to collapse an AI tool/run into one undo step. */
  historyBatch?: {
    depth: number
    undoStart: number
    before: HistorySnapshot
  }
  /** Rollback points for the AI panel's Snapshots list, keyed by id (one per AI run that edited the deck). */
  aiSnapshots?: Map<number, HistorySnapshot>
  /** Edits that only touch archive entries (notes/comments; element-level dirty cannot detect them), reset after save */
  metaDirty?: boolean
  /** Monotonic count of metaDirty transitions, so a save can tell a notes/comments
      edit committed while its write was streaming from one the write carried. */
  metaRev?: number
  /** Transform preview gesture in progress (the first preview already pushed an undo snapshot; later previews/final commit do not) */
  transformPreview?: boolean
  /** Redo stack as it was before the preview gesture's pushHistory cleared it; a cancel puts it back */
  previewRedo?: HistorySnapshot[]
  /** The part currently edited in master view (exception to the fidelity rule: only that part is written back) */
  masterEdit?: { partPath: string; slide: Slide } | null
  /** A history-state notification is already queued for this session (coalesces per task) */
  historyNotifyScheduled?: boolean
  /** A deck-changed broadcast is already queued for this session (coalesces per task) */
  deckBroadcastScheduled?: boolean
  /** Monotonic sequence of the last journaled op entry (collab groundwork) */
  opSeq?: number
  /** Applied-op journal, capped ring — the attachment point for a future sync transport */
  opLog?: OpLogEntry[]
}

export interface OpLogEntry {
  seq: number
  source: 'edit' | 'batch' | 'script' | 'generate' | 'reset'
  ops: Array<{ op: { op: string; [k: string]: unknown }; slideId?: string; created?: string[] }>
}

export interface HistorySnapshot {
  slides: Slide[]
  entries: Map<string, Uint8Array>
  size: { cx: number; cy: number }
  /** archive-only edits (notes, theme) flag the session, so undo must restore that too */
  metaDirty: boolean
}

export interface SlidesDocumentHost {
  getFontMetrics(): FontMetricsProvider
  historyChanged?(session: Session): void
  deckChanged?(session: Session): void
}

export function createSlidesDocumentRuntime(host: SlidesDocumentHost) {
  const getFontMetrics = () => host.getFontMetrics()
  function sessionDirty(session: Session): boolean {
    return (
      !!session.metaDirty ||
      session.opened.deck.slides.some(
        (s) => s.structureDirty || s.elements.some((el) => el.dirty || el.dirtyTransform),
      )
    )
  }
  /** Flag an archive-only edit (notes/comments) and advance the save-snapshot counter. */
  function markMetaDirty(session: Session): void {
    session.metaDirty = true
    session.metaRev = (session.metaRev ?? 0) + 1
  }

  // ── Op journal (collab groundwork) ──────────────────────────────────────
  // Every applied transaction appends its records here in order. Snapshot restores
  // (undo/redo/AI rollback) append a `reset` marker instead of inverse entries: a
  // consumer that cannot invert must full-resync past one. Payloads (e.g. picture
  // bytes) are kept verbatim; content-addressing them is the transport layer's job.

  const OP_LOG_MAX = 200

  function journalOps(
    session: Session,
    source: Exclude<OpLogEntry['source'], 'reset'>,
    records: Array<{
      op: { op: string; [k: string]: unknown }
      slideId?: string
      created?: string[]
    }>,
  ): void {
    if (records.length === 0) return
    const seq = (session.opSeq = (session.opSeq ?? 0) + 1)
    const log = (session.opLog ??= [])
    log.push({
      seq,
      source,
      ops: records.map((r) => ({
        op: r.op,
        ...(r.slideId ? { slideId: r.slideId } : {}),
        ...(r.created ? { created: r.created } : {}),
      })),
    })
    while (log.length > OP_LOG_MAX) log.shift()
  }

  // ── Undo/redo (snapshot-based) ─────────────────────────────────────────
  // The document's source of truth lives in the main process (deck.slides mutated in place +
  // archive.entries surgery), so history lives here too: snapshot both before every edit.
  // slides needs a deep copy (elements are mutated in place); entries only needs a shallow Map
  // copy (byte Buffers are never mutated in place, only replaced wholesale).

  const MAX_HISTORY = 50

  function trimHistory(stack: HistorySnapshot[]): void {
    while (stack.length > MAX_HISTORY) stack.shift()
  }

  function takeSnapshot(session: Session): HistorySnapshot {
    return {
      slides: structuredClone(session.opened.deck.slides),
      entries: new Map(session.opened.archive.entries),
      size: { ...session.opened.deck.size },
      metaDirty: !!session.metaDirty,
    }
  }

  // slides must be deep-copied: restoreSnapshot hands them to the live deck, which mutates in place
  function cloneSnapshot(snap: HistorySnapshot): HistorySnapshot {
    return {
      slides: structuredClone(snap.slides),
      entries: new Map(snap.entries),
      size: { ...snap.size },
      metaDirty: snap.metaDirty,
    }
  }

  /**
   * Tell the renderer whether undo/redo have anything to apply (drives the QAT
   * button gray states). Deferred with setImmediate so no-op handlers that push
   * a snapshot and then pop it back in the same turn report the settled state,
   * and multiple stack changes per turn coalesce into one message.
   */
  function scheduleHistoryNotify(session: Session): void {
    if (session.historyNotifyScheduled) return
    session.historyNotifyScheduled = true
    setImmediate(() => {
      session.historyNotifyScheduled = false
      host.historyChanged?.(session)
    })
  }

  /** Coalesce native deck notifications through the explicitly supplied host. */
  function scheduleDeckBroadcast(session: Session): void {
    if (session.deckBroadcastScheduled) return
    session.deckBroadcastScheduled = true
    setImmediate(() => {
      session.deckBroadcastScheduled = false
      host.deckChanged?.(session)
    })
  }

  /** Call before an edit operation: push onto the undo stack and clear the redo stack. */
  function pushHistory(session: Session): void {
    session.undoStack.push(takeSnapshot(session))
    trimHistory(session.undoStack)
    session.redoStack = []
    scheduleHistoryNotify(session)
  }

  /** Begin a nestable transaction. Individual edit handlers keep their normal rollback behavior. */
  function beginHistoryBatch(session: Session): void {
    if (session.historyBatch) {
      session.historyBatch.depth += 1
      return
    }
    session.historyBatch = {
      depth: 1,
      undoStart: session.undoStack.length,
      before: takeSnapshot(session),
    }
  }

  /**
   * End a transaction and collapse every successful edit since begin into the pre-transaction
   * snapshot. Failed/no-op handlers can continue popping their own snapshots safely.
   * Returns the pre-transaction snapshot when the outermost end collapsed real edits (null otherwise),
   * so the caller can register it as an AI-panel rollback point.
   */
  function endHistoryBatch(session: Session): HistorySnapshot | null {
    const batch = session.historyBatch
    if (!batch) return null
    batch.depth -= 1
    if (batch.depth > 0) return null
    session.historyBatch = undefined
    if (session.undoStack.length <= batch.undoStart) return null
    session.undoStack.splice(batch.undoStart)
    session.undoStack.push(batch.before)
    trimHistory(session.undoStack)
    scheduleHistoryNotify(session)
    return batch.before
  }

  /** Preserve the old deck and its history when AI replaces the entire presentation. */
  function carryHistoryForReplacement(previous: Session | undefined, replacement: Session): void {
    if (!previous) return
    pushHistory(previous)
    replacement.undoStack = previous.undoStack
    replacement.redoStack = previous.redoStack
    replacement.historyBatch = previous.historyBatch
    replacement.aiSnapshots = previous.aiSnapshots
    scheduleHistoryNotify(replacement)
  }

  const MAX_AI_SNAPSHOTS = 20
  let nextAiSnapshotId = 1

  /** Register a rollback point (stored as its own copy; `snap` typically also sits on the undo stack). */
  function registerAiSnapshot(session: Session, snap: HistorySnapshot): number {
    const map = (session.aiSnapshots ??= new Map())
    const id = nextAiSnapshotId++
    map.set(id, cloneSnapshot(snap))
    while (map.size > MAX_AI_SNAPSHOTS) map.delete(map.keys().next().value as number)
    return id
  }

  /** Roll the deck back to a registered AI snapshot; the pre-rollback state becomes one undo step. */
  function restoreAiSnapshot(session: Session, id: number): boolean {
    const snap = session.aiSnapshots?.get(id)
    if (!snap) return false
    pushHistory(session)
    restoreSnapshot(session, snap)
    session.aiSnapshots?.delete(id)
    return true
  }

  function restoreSnapshot(session: Session, snap: HistorySnapshot): void {
    // Clone: the live deck mutates elements in place, so handing a snapshot's own
    // arrays over would let later edits rewrite history still referenced by the
    // other stack (undo → edit → redo would replay mutated state).
    const fresh = cloneSnapshot(snap)
    session.opened.deck.slides = fresh.slides
    session.opened.deck.size = fresh.size
    session.metaDirty = fresh.metaDirty
    const entries = session.opened.archive.entries
    entries.clear()
    for (const [k, v] of fresh.entries) entries.set(k, v)
    // The journal cannot express a snapshot jump as ops; mark the divergence so a
    // future consumer knows to full-resync rather than replay across it.
    if (session.opLog?.length) {
      const seq = (session.opSeq = (session.opSeq ?? 0) + 1)
      session.opLog.push({ seq, source: 'reset', ops: [] })
      while (session.opLog.length > OP_LOG_MAX) session.opLog.shift()
    }
  }

  /**
   * Close a history batch that outlived its run: an AI tool path that
   * throws between begin and end would otherwise leave historyBatch set forever,
   * and undo/redo — which refuse to run mid-batch — would silently do nothing for
   * the rest of the session. Collapsing here keeps the run's edits as one step.
   */
  function settleStaleHistoryBatch(session: Session): void {
    while (session.historyBatch) {
      const collapsed = endHistoryBatch(session)
      if (collapsed) registerAiSnapshot(session, collapsed)
    }
  }

  function buildAllRenderSlides(opened: OpenedPptx, fitWidthPx: number): RenderSlide[] {
    return opened.deck.slides.map((s, i) =>
      buildRenderSlide(s, opened.deck.size, {
        fitWidthPx,
        media: makeMediaResolver(opened, s.path),
        metrics: getFontMetrics(),
        slideNo: i + 1,
      }),
    )
  }

  /** Office theme-class slot (MsftOfcThm_<slot>_Fill/_Stroke) → schemeClr name. */
  const SVG_THEME_SLOTS: Record<string, string> = {
    background1: 'bg1',
    text1: 'tx1',
    background2: 'bg2',
    text2: 'tx2',
    accent1: 'accent1',
    accent2: 'accent2',
    accent3: 'accent3',
    accent4: 'accent4',
    accent5: 'accent5',
    accent6: 'accent6',
    hyperlink: 'hlink',
    followedhyperlink: 'folHlink',
  }

  /**
   * Office-exported SVGs carry `.MsftOfcThm_<slot>_Fill/_Stroke` CSS classes whose baked
   * values PowerPoint rewrites to the CURRENT theme's colors at render time (the static
   * value is just the export-time snapshot — probe deck: a bg1-classed blob draws white
   * on a white theme, not its baked blue). Mirror that rewrite before serving the SVG.
   */
  function retintThemedSvg(svg: string, opened: OpenedPptx, slidePath?: string): string {
    const path = slidePath ?? opened.deck.slides[0]?.path
    if (!path) return svg
    let theme
    try {
      const chain = opened.archive.resolveSlideChain(path)
      const themeXml = chain.themePath ? opened.archive.readText(chain.themePath) : null
      if (!themeXml) return svg
      theme = parseTheme(themeXml)
      const masterXml = chain.masterPath ? opened.archive.readText(chain.masterPath) : undefined
      const layoutXml = chain.layoutPath ? opened.archive.readText(chain.layoutPath) : undefined
      theme.clrMap = parseClrMap(
        masterXml ?? undefined,
        layoutXml ?? undefined,
        opened.archive.readText(path) ?? undefined,
      )
    } catch {
      return svg
    }
    return svg.replace(
      /\.MsftOfcThm_(\w+?)_(Fill|Stroke)\w*\s*\{[^}]*\}/g,
      (rule, slot: string, kind: string) => {
        const scheme = SVG_THEME_SLOTS[slot.toLowerCase()]
        const color = scheme ? resolveSchemeColor(scheme, theme) : undefined
        if (!color) return rule
        const prop = kind === 'Stroke' ? 'stroke' : 'fill'
        return rule.replace(new RegExp(`${prop}\\s*:\\s*[^;}]+`, 'g'), `${prop}:${color}`)
      },
    )
  }

  /** Image mediaRef -> dataUrl (lazily decoded). TIFF is transcoded to PNG for display
    (Chromium can't decode it); the archive keeps the original bytes for save fidelity.
    The mime comes from magic-byte sniffing first (legacy decks mislabel media — a PNG
    stored as .emf must not enter the EMF parser), extension second. */
  function makeMediaResolver(opened: OpenedPptx, slidePath?: string) {
    const cache = new Map<string, string | undefined>()
    return (mediaRef: string): string | undefined => {
      if (cache.has(mediaRef)) return cache.get(mediaRef)
      const bytes = opened.archive.readBytes(mediaRef)
      let url: string | undefined
      if (bytes) {
        const mime = displayMime(mediaRef, bytes)
        if (mime === 'image/tiff') {
          const decoded = tiffToPng(bytes)
          if (decoded) url = `data:image/png;base64,${Buffer.from(decoded.png).toString('base64')}`
        } else if (mime === 'image/svg+xml') {
          let text = Buffer.from(bytes).toString('utf8')
          if (text.includes('MsftOfcThm_')) text = retintThemedSvg(text, opened, slidePath)
          url = `data:${mime};base64,${Buffer.from(text, 'utf8').toString('base64')}`
        } else {
          // PowerPoint ignores EXIF orientation; Chromium applies it on decode — neutralize
          // the flag so rotated-pixel JPEGs with a shape-level rot don't double-rotate
          const served = mime === 'image/jpeg' ? neutralizeJpegOrientation(bytes) : bytes
          url = `data:${mime};base64,${Buffer.from(served).toString('base64')}`
        }
      }
      cache.set(mediaRef, url)
      return url
    }
  }

  /** Rebuild a single page's RenderSlide (sent back to the renderer after an edit). */
  function rebuildSlide(session: Session, slideIndex: number): RenderSlide | null {
    const slide = session.opened.deck.slides[slideIndex]
    if (!slide) return null
    return buildRenderSlide(slide, session.opened.deck.size, {
      fitWidthPx: session.fitWidthPx,
      media: makeMediaResolver(session.opened, slide.path),
      metrics: getFontMetrics(),
      slideNo: slideIndex + 1,
    })
  }

  /**
   * Rebuild the RenderSlide after re-parsing the whole page (the model is stale after a chart
   * edit and needs a reparse). Equivalent to materializeSlide then rebuildSlide, but without the
   * structureDirty save logic.
   */
  function rebuildSlideWithReparse(session: Session, slideIndex: number): RenderSlide | null {
    const fresh = materializeSlide(session.opened, slideIndex)
    if (!fresh) return null
    return buildRenderSlide(fresh, session.opened.deck.size, {
      fitWidthPx: session.fitWidthPx,
      media: makeMediaResolver(session.opened, fresh.path),
      metrics: getFontMetrics(),
      slideNo: slideIndex + 1,
    })
  }

  return {
    sessionDirty,
    markMetaDirty,
    journalOps,
    takeSnapshot,
    pushHistory,
    beginHistoryBatch,
    endHistoryBatch,
    carryHistoryForReplacement,
    registerAiSnapshot,
    restoreAiSnapshot,
    restoreSnapshot,
    settleStaleHistoryBatch,
    buildAllRenderSlides,
    retintThemedSvg,
    makeMediaResolver,
    rebuildSlide,
    rebuildSlideWithReparse,
    scheduleHistoryNotify,
    scheduleDeckBroadcast,
    getFontMetrics,
  }
}

export type SlidesDocumentRuntime = ReturnType<typeof createSlidesDocumentRuntime>
