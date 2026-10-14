/** Original native Slides IPC document handlers. Both hosts register these same callbacks. */
import type { Dialog, Clipboard, NativeImage, BrowserWindow } from 'electron'
import type { readFile as readFileType } from 'node:fs/promises'
import type { showOpenDialogWithMemory as dialogType } from '@genoffice/electron-utils'
import type { SlidesDocumentRuntime } from './session-document'
import { createElementClipboard, type ClipboardImagePort } from './element-clipboard-core'
import { createSlidesDocumentOperations } from './document-operations'
import {
  EditTextOp,
  SetElementFontOp,
  SetElementParagraphFormatOp,
  EditTransformOp,
  EditTransformMultiOp,
  EditConnectorEndpointsOp,
  BatchEditTransformOp,
  ApplyTxnOp,
  ApplyTxnResult,
  ApplyEditScriptOp,
  AddElementOp,
  DeleteElementOp,
  DeleteElementsOp,
  EditStrokeOp,
  FlipElementOp,
  EditPictureSrcRectOp,
  GroupElementsOp,
  UngroupElementOp,
  EditBackgroundOp,
  EditFillImageOp,
  EditFillOp,
  AddSlideOp,
  CopySlidesOp,
  PasteSlideOp,
  RepasteSlideOp,
  AddBlankSlideOp,
  AddSlideWithLayoutOp,
  MasterEnterResult,
  MasterEditTextOp,
  MasterEditTransformOp,
  MasterEditFillOp,
  MasterEditStrokeOp,
  MasterDeleteElementOp,
  EditPictureOpacityOp,
  SetSlideSizeOp,
  SetSlideLayoutOp,
  FindReplaceOp,
  DeleteSlidesOp,
  DuplicateSlidesOp,
  SetSlidesHiddenOp,
  EditTableCellOp,
  TableMergeIpcOp,
  TableStructureIpcOp,
  SetTableRowHeightOp,
  SetTableCellAnchorOp,
  SetTableColWidthOp,
  EditTableStyleOp,
  EditChartOp,
  ReorderElementOp,
  SetShapeGeometryOp,
  SetEffectsPatch,
  CopyElementsOp,
  PasteElementsOp,
  DuplicateElementsOp,
  AddTableOp,
  AddInkOp,
  AddChartOp,
  AddSmartArtOp,
  AddImageBytesOp,
  ReplacePictureBytesOp,
  AddMediaBytesOp,
  SetLinkOp,
  HeaderFooterOp,
  ApplyThemeOp,
  SetTransitionOp,
  SetAdvanceTimesOp,
  AnimationItem,
  ShapeKey,
  SetAnimationsOp,
  SetSlideHiddenOp,
  AddSectionOp,
  RenameSectionOp,
  RemoveSectionOp,
  RemoveSectionSlidesOp,
  MoveSlideOp,
  MoveSlidesOp,
  MoveSectionOp,
  SetNotesOp,
  AddCommentOp,
  DeleteCommentOp,
  GradientFillSpec,
  PasteSlideResult,
  AddMediaResult,
} from '../shared/ipc'
import {
  materializeSlide,
  type Paragraph,
  EMU_PER_PT,
  type SlideBundle,
  copySlide,
  listSlideLayouts,
  shouldOfferBuiltinLayouts,
  builtinLayoutInfos,
  listMasterParts,
  parseMasterPart,
  editChartElement,
  getChartElementData,
  copyElementData,
  getElementLink,
  getSlideLinks,
  getRunLinks,
  readHeaderFooter,
  reparseDeck,
  getSlideTransition,
  elementSpid,
  getSlideAnimations,
  getSections,
  type SectionInfo,
  normalizeSections,
  getSlideNotes,
  getSlideComments,
  findGroupChild,
  BUILTIN_LAYOUT_PREFIX,
  ensureBuiltinLayout,
  type Slide,
  type OpenedPptx,
  type ElementClipboardItem,
  parseTheme,
} from '@genoffice/pptx-engine'
import {
  runTxn,
  mapScriptOps,
  type TxnResult,
  resolveGroupChildId,
  type TxnRequest,
  type Op,
} from '@genoffice/pptx-ops'
import {
  EMU_PER_PX_96,
  imageDpiFromBytes,
  buildRenderSlide,
  type RenderSlide,
} from '@genoffice/pptx-render'
import { tm } from './i18n-main'
import { tiffToPng } from './tiff-decode'
import { pictureFrame, DEFAULT_PICTURE_DPI } from './picture-frame'
import { planSlideDuplicates, planSlideMoves } from '../shared/slide-selection'
import { matchesElementRef } from '@genoffice/pptx-engine/identity'
import { randomUUID } from 'node:crypto'
import {
  newPasteCascade,
  pageKey,
  pasteShiftPx,
  recordPaste,
  type PasteCascade,
} from './paste-cascade'
import { ELEMENT_CLIPBOARD_FORMAT } from './element-clipboard-core'
import { VIDEO_EXTS, AUDIO_EXTS } from '../shared/media-kinds'
import { baseName } from '../shared/base-name'
import { safeExternalUrl } from '../../../../packages/electron-utils/src/safe-external-url'
import { DECK_LINK_PROTOCOLS } from '../shared/run-link'
import { type Session, type OpLogEntry } from './session-document'
import { unplayableAudioCodec } from './mp4-audio-sniff'
import { type Size, type Point, videoFrame, videoSize, audioFrame } from './video-size'
export interface SlidesDocumentEvent {
  sender: { id: number; cut(): void; copy(): void; paste(): void }
}
export interface SlidesDocumentRegistrar {
  handle(channel: string, handler: (event: SlidesDocumentEvent, ...args: any[]) => any): void
}
export interface SlidesDocumentPorts {
  dialog: Dialog
  dialogParent(): BrowserWindow | undefined
  clipboard: Clipboard
  nativeImage: {
    createFromBuffer(bytes: Buffer): NativeImage
    createFromPath(path: string): NativeImage
    createThumbnailFromPath(
      path: string,
      size: { width: number; height: number },
    ): Promise<NativeImage>
  }
  showOpenDialogWithMemory: typeof dialogType
  readFile: typeof readFileType
  commentAuthorName(): string
}
export function registerSlidesDocumentHandlers(
  ipcMain: SlidesDocumentRegistrar,
  sessions: Map<number, Session>,
  documentRuntime: SlidesDocumentRuntime,
  ports: SlidesDocumentPorts,
) {
  const {
    getFontMetrics,
    beginHistoryBatch,
    endHistoryBatch,
    buildAllRenderSlides,
    journalOps,
    makeMediaResolver,
    markMetaDirty,
    pushHistory,
    rebuildSlide,
    rebuildSlideWithReparse,
    registerAiSnapshot,
    restoreAiSnapshot,
    restoreSnapshot,
    settleStaleHistoryBatch,
    scheduleDeckBroadcast,
    scheduleHistoryNotify,
    takeSnapshot,
    sessionDirty,
  } = documentRuntime
  const { findEl, journaledTxn, applySessionTxn, applyAutofitResize, syncAutofitScale } =
    createSlidesDocumentOperations(documentRuntime)
  const {
    dialog,
    dialogParent,
    clipboard,
    nativeImage,
    showOpenDialogWithMemory,
    readFile,
    commentAuthorName,
  } = ports
  const {
    canWriteElementClipboardImage,
    elementClipboardMarkerMatches,
    isElementClipboardToken,
    writeElementClipboardImage,
  } = createElementClipboard({ clipboard, nativeImage } as ClipboardImagePort)

  /** One slide, copied from any deck open in this process, waiting to be pasted into another. */
  let slideClipboard: { bundles: SlideBundle[]; pngs?: string[] } | null = null

  /** The immediately preceding slide paste per webContents, so the paste-options floater can redo it with another mode. */
  const lastSlidePaste = new Map<number, { afterIndex: number; undoLen: number }>()

  /** IPC gradient → engine stop list (full stop list wins over the two-color from/to form). */
  const gradientStops = (g: GradientFillSpec['gradient']): Array<{ pos: number; color: string }> =>
    g.stops?.length
      ? g.stops
      : [
          { pos: 0, color: g.from },
          { pos: 1, color: g.to },
        ]

  /** IPC gradient → path kind (radial is a legacy alias for circle; undefined = linear). */
  const gradientPathKind = (
    g: GradientFillSpec['gradient'],
  ): 'circle' | 'rect' | 'shape' | undefined => g.path ?? (g.radial ? 'circle' : undefined)

  /** IPC gradient focus point → <a:fillToRect> insets (undefined when unspecified). */
  const gradientFillTo = (
    g: GradientFillSpec['gradient'],
  ): { l: number; t: number; r: number; b: number } | undefined =>
    g.center ? { l: g.center.x, t: g.center.y, r: 1 - g.center.x, b: 1 - g.center.y } : undefined

  // ── In-app element clipboard (app-wide, so elements copied in one deck paste into any other open deck; the cascade decides the paste offset) ─
  let elementClipboard: {
    items: ElementClipboardItem[]
    cascade: PasteCascade
    token: string
    senderId: number
  } | null = null

  /** Legacy fixed color schemes (AI tools/old files still pass these keys; kept as fallback). */
  const CHART_COLOR_SCHEMES: Record<string, string[]> = {
    default: [],
    blue: ['#2E75B6', '#4472C4', '#5B9BD5', '#70AD47', '#ED7D31'],
    warm: ['#ED7D31', '#FFC000', '#FF0000', '#C55A11', '#833C00'],
    cool: ['#0070C0', '#00B0F0', '#00B0A0', '#7030A0', '#2E75B6'],
    mono: ['#404040', '#666666', '#888888', '#AAAAAA', '#CCCCCC'],
  }

  /** PowerPoint default theme accent sequence (fallback when the deck has no theme colors). */
  const FALLBACK_ACCENTS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']

  /** Mix a hex color with black/white by ratio (for mono-gradient steps). */
  function mixHex(hex: string, target: number, ratio: number): string {
    const m = /^#?([0-9a-fA-F]{6})$/.exec(hex)
    if (!m) return hex
    const v = parseInt(m[1]!, 16)
    const ch = (x: number) => Math.round(x + (target - x) * ratio)
    const r = ch((v >> 16) & 255)
    const g = ch((v >> 8) & 255)
    const b = ch(v & 255)
    return `#${((r << 16) | (g << 8) | b).toString(16).padStart(6, '0').toUpperCase()}`
  }

  /** Current deck's theme accent1..6 (read from the theme part of the first slide's inheritance chain). */
  function deckAccents(opened: OpenedPptx): string[] {
    const slide = opened.deck.slides[0]
    if (!slide) return FALLBACK_ACCENTS
    try {
      const chain = opened.archive.resolveSlideChain(slide.path)
      const xml = chain.themePath ? opened.archive.readText(chain.themePath) : null
      const colors = xml ? parseTheme(xml).colors : undefined
      const acc = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
        .map((k) => colors?.[k])
        .filter((c): c is string => !!c)
      return acc.length >= 3 ? acc : FALLBACK_ACCENTS
    } catch {
      return FALLBACK_ACCENTS
    }
  }

  /** Theme-derived color schemes for the chart "Change Colors" gallery: two colorful sets + one mono gradient per accent. */
  function chartColorSchemes(
    opened: OpenedPptx,
  ): Array<{ key: string; label: string; colors: string[] }> {
    const acc = deckAccents(opened)
    const rot = [...acc.slice(3), ...acc.slice(0, 3)]
    const mono = (c: string) => [
      mixHex(c, 0, 0.25),
      c,
      mixHex(c, 255, 0.25),
      mixHex(c, 255, 0.45),
      mixHex(c, 255, 0.65),
    ]
    return [
      { key: 'default', label: tm('schemeThemeDefault'), colors: [] },
      { key: 'colorful', label: tm('schemeColorful'), colors: acc },
      { key: 'colorful2', label: tm('schemeColorful2'), colors: rot },
      ...acc.map((c, i) => ({
        key: `mono-accent${i + 1}`,
        label: tm('schemeMono', { n: i + 1 }),
        colors: mono(c),
      })),
    ]
  }

  // Shared shim core: one undo step wrapping a transaction; on failure the step is
  // erased (the executor already restored the model, so history stays consistent).
  // Handlers keep only surface translation and result shaping around this.
  // Plan first: pushHistory clears the redo stack (and can evict the oldest undo
  // entry at the cap), so an invalid request must not touch history at all —
  // legacy handlers validated existence before their history push.

  const sessionTxn = (
    session: Session,
    req: TxnRequest,
    source: Exclude<OpLogEntry['source'], 'reset'> = 'edit',
  ): TxnResult | null => {
    const plan = runTxn(session.opened, { ...req, dryRun: true })
    const invalid = plan.failures?.length ?? 0
    if ((req.isolation ?? 'atomic') === 'atomic' ? invalid > 0 : invalid >= req.ops.length)
      return null
    pushHistory(session)
    const r = journaledTxn(session, source, req)
    if (!r.applied) {
      session.undoStack.pop()
      return null
    }
    return r
  }

  // Shim over the canonical setText op (rich-text rebuild, link rels, resource cleanup,
  // level rematerialization live in the op); autofit is a render concern and stays here.
  ipcMain.handle('slides:edit-text', (e, op: EditTextOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const r = journaledTxn(session, 'edit', {
      ops: [
        {
          op: 'setText',
          target: { slide: op.slideIndex, el: op.sourceId },
          paragraphs: op.paragraphs,
          ...(op.groupId ? { group: op.groupId } : {}),
        },
      ],
    })
    if (!r.applied) {
      session.undoStack.pop()
      return null
    }
    const levelDirty = (r.records?.[0]?.after as { levelDirty?: boolean } | undefined)?.levelDirty
    if (op.groupId || levelDirty) return rebuildSlide(session, op.slideIndex)
    const rendered = applyAutofitResize(
      session,
      op.slideIndex,
      op.sourceId,
      rebuildSlide(session, op.slideIndex),
    )
    return syncAutofitScale(session, op.slideIndex, op.sourceId, rendered)
  })

  // Shim over the canonical setFont op: one per_op transaction covers the whole
  // selection (non-text elements fail their own op and are skipped, matching the
  // legacy "changed if any succeeded" semantics).
  ipcMain.handle('slides:set-element-font', (e, op: SetElementFontOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const font = {
      fontFamily: op.fontFamily,
      fontSizePt: op.fontSizePt,
      fontSizeStep: op.fontSizeStep,
      strike: op.strike,
      bold: op.bold,
      italic: op.italic,
      underline: op.underline,
      color: op.color,
    }
    pushHistory(session)
    const r = journaledTxn(session, 'edit', {
      isolation: 'per_op',
      ops: op.sourceIds.map((id) => ({
        op: 'setFont',
        target: { slide: op.slideIndex, el: id },
        font,
        ...(op.groupId ? { group: op.groupId } : {}),
      })),
    })
    if (!r.applied) {
      session.undoStack.pop() // All non-text elements (images etc.): nothing happened, pop the just-pushed history
      return null
    }
    let rendered = rebuildSlide(session, op.slideIndex)
    for (const id of op.sourceIds) {
      rendered = applyAutofitResize(session, op.slideIndex, id, rendered)
      rendered = syncAutofitScale(session, op.slideIndex, id, rendered)
    }
    return rendered
  })

  // Shim over the canonical setParagraphFormat op (same per_op selection semantics as setFont).
  ipcMain.handle('slides:set-element-paragraph-format', (e, op: SetElementParagraphFormatOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const format = {
      bullet: op.bullet,
      bulletChar: op.bulletChar,
      bulletFont: op.bulletFont,
      numType: op.numType,
      startAt: op.startAt,
      bulletImage: op.bulletImage,
      bulletHangEmu: op.bulletHangEmu,
      bulletSizePct: op.bulletSizePct,
      bulletColor: op.bulletColor,
      lineSpacingPct: op.lineSpacingPct,
      spaceBeforePt: op.spaceBeforePt,
      spaceAfterPt: op.spaceAfterPt,
      align: op.align,
      rtl: op.rtl,
      indentDelta: op.indentDelta,
    }
    pushHistory(session)
    const r = journaledTxn(session, 'edit', {
      isolation: 'per_op',
      ops: op.sourceIds.map((id) => ({
        op: 'setParagraphFormat',
        target: { slide: op.slideIndex, el: id },
        format,
        ...(op.groupId ? { group: op.groupId } : {}),
      })),
    })
    if (!r.applied) {
      session.undoStack.pop()
      return null
    }
    if (op.indentDelta) {
      // Level changes affect inherited defaults; bake into bytes then reparse
      materializeSlide(session.opened, op.slideIndex)
      return rebuildSlide(session, op.slideIndex)
    }
    let rendered = rebuildSlide(session, op.slideIndex)
    for (const id of op.sourceIds) {
      rendered = applyAutofitResize(session, op.slideIndex, id, rendered)
      rendered = syncAutofitScale(session, op.slideIndex, id, rendered)
    }
    return rendered
  })

  // px→EMU (inverting the viewport scale) for one setTransform. In-group editing: the
  // pixel box is in group-local coords (with ext/chExt scaling baked in); divide out
  // the group scale first, then convert back to the child EMU coordinate system.
  const transformPayload = (
    session: Session,
    slideIndex: number,
    fitWidthPx: number,
    item: Omit<EditTransformOp, 'slideIndex' | 'fitWidthPx' | 'preview'>,
  ) => {
    const slide = session.opened.deck.slides[slideIndex]
    if (!slide) return null
    const childId = item.groupId
      ? resolveGroupChildId(slide, item.groupId, item.sourceId)
      : item.sourceId
    const grpChild = item.groupId ? findGroupChild(slide, item.groupId, childId) : null
    if (item.groupId && !grpChild) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    let box: { x: number; y: number; cx: number; cy: number }
    if (grpChild) {
      const ch = grpChild.grp.childOffset
      const chX = ch?.x ?? grpChild.grp.transform.offset.x
      const chY = ch?.y ?? grpChild.grp.transform.offset.y
      const gExt = grpChild.grp.transform.offset
      const gsx = ch?.cx ? gExt.cx / ch.cx : 1
      const gsy = ch?.cy ? gExt.cy / ch.cy : 1
      box = {
        x: toEmu(item.xPx / gsx) + chX,
        y: toEmu(item.yPx / gsy) + chY,
        cx: toEmu(item.wPx / gsx),
        cy: toEmu(item.hPx / gsy),
      }
    } else {
      box = { x: toEmu(item.xPx), y: toEmu(item.yPx), cx: toEmu(item.wPx), cy: toEmu(item.hPx) }
    }
    return {
      op: 'setTransform' as const,
      target: { slide: slideIndex, el: item.sourceId },
      box,
      rotDeg: item.rotationDeg,
      // Tables redistribute gridCol widths / tr heights so the file matches the frame
      ...(item.groupId ? { group: item.groupId } : { resizeTableGrid: true }),
    }
  }

  // Shim over the canonical setTransform op. Preview-gesture undo bookkeeping is a
  // surface concern and stays here.
  ipcMain.handle('slides:edit-transform', (e, op: EditTransformOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const payload = transformPayload(session, op.slideIndex, op.fitWidthPx, op)
    if (!payload) return null
    // Validate BEFORE the preview bookkeeping: a failed first preview must not set
    // transformPreview (later frames would skip pushHistory and the eventual commit
    // would lose its undo step) and must not clear the redo stack.
    if (runTxn(session.opened, { dryRun: true, ops: [payload] }).failures?.length) return null
    // Undo semantics for preview gestures: one whole drag = one undo step.
    // The first preview pushes a pre-gesture snapshot; later previews and the final commit do not.
    let pushed = false
    if (op.preview) {
      if (!session.transformPreview) {
        session.previewRedo = session.redoStack
        pushHistory(session)
        pushed = true
        session.transformPreview = true
      }
    } else if (session.transformPreview) {
      session.transformPreview = false
      session.previewRedo = undefined
    } else {
      pushHistory(session)
      pushed = true
    }
    // Journal only the settled commit: preview frames would flood the ring with
    // intermediate boxes (undo coalesces the gesture the same way via transformPreview).
    const r = op.preview
      ? runTxn(session.opened, { ops: [payload] })
      : journaledTxn(session, 'edit', { ops: [payload] })
    if (!r.applied) {
      // Apply-time failure (e.g. a group-child slice that validation cannot see):
      // unwind everything this call did, including the preview flag it set — a stuck
      // flag would make later frames skip pushHistory and cost the commit its undo step.
      if (pushed) {
        session.undoStack.pop()
        if (op.preview) {
          session.transformPreview = false
          if (session.previewRedo) session.redoStack = session.previewRedo
          session.previewRedo = undefined
        }
      }
      return null
    }
    return rebuildSlide(session, op.slideIndex)
  })

  // Escape mid-resize: the preview frames live in the model and their undo snapshot is the
  // pre-gesture state, so restoring that snapshot and dropping it leaves no trace of the gesture
  ipcMain.handle('slides:cancel-transform-preview', (e, op: { slideIndex: number }) => {
    const session = sessions.get(e.sender.id)
    if (!session || !session.transformPreview) return null
    session.transformPreview = false
    const snap = session.undoStack.pop()
    if (snap) restoreSnapshot(session, snap)
    if (session.previewRedo) session.redoStack = session.previewRedo
    session.previewRedo = undefined
    scheduleHistoryNotify(session)
    scheduleDeckBroadcast(session)
    return rebuildSlide(session, op.slideIndex)
  })

  // A multi-selection nudge/rotate: every item commits exactly like edit-transform,
  // the whole batch is one undo step (PowerPoint undoes the action, not each shape).
  ipcMain.handle('slides:edit-transform-multi', (e, op: EditTransformMultiOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const ops: NonNullable<ReturnType<typeof transformPayload>>[] = []
    for (const item of op.items) {
      const payload = transformPayload(session, op.slideIndex, op.fitWidthPx, item)
      if (!payload) return null
      ops.push(payload)
    }
    if (!ops.length) return null
    const r = sessionTxn(session, { ops })
    return r ? rebuildSlide(session, op.slideIndex) : null
  })

  // Connector endpoint drag: box+flip re-derived from the two endpoints;
  // attach/detach writes a:stCxn/a:endCxn so the connector follows later shape moves
  ipcMain.handle('slides:edit-connector-endpoints', (e, op: EditConnectorEndpointsOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'setConnectorEndpoints',
          target: { slide: op.slideIndex, el: op.sourceId },
          p1: { x: toEmu(op.x1Px), y: toEmu(op.y1Px) },
          p2: { x: toEmu(op.x2Px), y: toEmu(op.y2Px) },
          start: op.start,
          end: op.end,
        },
      ],
    })
    return r ? rebuildSlide(session, op.slideIndex) : null
  })

  // Read-only: RenderSlide for every page of the current session (E2E driver/debug use, no state change)
  ipcMain.handle('slides:get-render-slides', (e) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    return session.opened.deck.slides.map((_, i) => rebuildSlide(session, i))
  })

  // Shim: the whole selection is one atomic transaction of setTransform ops — the
  // executor's plan step reproduces the legacy "every element must exist" gate.
  ipcMain.handle('slides:batch-edit-transform', (e, op: BatchEditTransformOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    const r = sessionTxn(session, {
      ops: op.items.map((item) => ({
        op: 'setTransform',
        target: { slide: op.slideIndex, el: item.sourceId },
        box: { x: toEmu(item.xPx), y: toEmu(item.yPx), cx: toEmu(item.wPx), cy: toEmu(item.hPx) },
        rotDeg: item.rotationDeg,
        // Legacy parity: the batch path never redistributed table grids
      })),
    })
    return r ? rebuildSlide(session, op.slideIndex) : null
  })

  // AI batch surface: raw ops arrive as one transaction — the shared core in
  // applySessionTxn (validation, atomicity/rollback/journal, autofit render pass)
  // is the same code the shell's MCP slides bridge drives.
  ipcMain.handle('slides:apply-txn', (e, req: ApplyTxnOp): ApplyTxnResult | null => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    return applySessionTxn(session, req)
  })

  // The whole edit script as ONE transaction: the collected primitives arrive in a
  // single IPC, compile to ops (script-map), and apply atomically — the executor
  // owns validation, rollback and the journal. Autofit is a render concern and
  // stays here, mirroring the per-op shims the script used to fan out to.
  ipcMain.handle('slides:apply-edit-script', (e, op: ApplyEditScriptOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const ops = mapScriptOps(session.opened, op)
    if (ops.length === 0) return null
    const plan = runTxn(session.opened, { ops, dryRun: true })
    if (plan.failures?.length) return { error: plan.failures[0]!.error }
    pushHistory(session)
    const r = journaledTxn(session, 'script', { ops })
    if (!r.applied) {
      session.undoStack.pop()
      return { error: r.failures?.[0]?.error ?? 'the transaction could not be applied' }
    }
    let rendered = rebuildSlide(session, op.slideIndex)
    for (const rec of r.records ?? []) {
      const o = rec.op
      const id = o.target?.el
      if (!id || o.group || !rendered) continue
      if (o.op === 'setText') {
        // Level changes rematerialized inside the op; autofit would fight the reparse
        if ((rec.after as { levelDirty?: boolean } | undefined)?.levelDirty) continue
        rendered = applyAutofitResize(session, op.slideIndex, id, rendered)
        rendered = syncAutofitScale(session, op.slideIndex, id, rendered)
      } else if (o.op === 'setFont' || o.op === 'setParagraphFormat') {
        rendered = applyAutofitResize(session, op.slideIndex, id, rendered)
        rendered = syncAutofitScale(session, op.slideIndex, id, rendered)
      }
    }
    return rendered ? { slide: rendered } : null
  })

  ipcMain.handle('slides:add-element', (e, op: AddElementOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    const paragraphs: Paragraph[] | undefined = op.paragraphs?.length
      ? (op.paragraphs as Paragraph[])
      : op.text
        ? op.text.split('\n').map((line) => ({ runs: [{ text: line }] }))
        : undefined
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'addElement',
          target: { slide: op.slideIndex },
          kind: op.kind,
          offset: { x: toEmu(op.xPx), y: toEmu(op.yPx), cx: toEmu(op.wPx), cy: toEmu(op.hPx) },
          ...(paragraphs ? { paragraphs } : {}),
          ...(op.fillColor ? { fill: op.fillColor } : {}),
          ...(op.stroke
            ? {
                stroke: {
                  color: op.stroke.color,
                  widthEmu: Math.round(op.stroke.widthPt * EMU_PER_PT),
                },
              }
            : {}),
          ...(op.bodyPr ? { bodyPr: op.bodyPr } : {}),
        },
      ],
    })
    if (!r) return null
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: r.records![0]!.created![0]! } : null
  })

  // Shim over the canonical op (see @genoffice/pptx-ops): the op owns validation/mutation/journal;
  // the shim keeps session lookup, undo bookkeeping, and RenderSlide rebuilding.
  ipcMain.handle('slides:delete-element', (e, op: DeleteElementOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const r = journaledTxn(session, 'edit', {
      ops: [{ op: 'deleteElement', target: { slide: op.slideIndex, el: op.sourceId } }],
    })
    if (!r.applied) {
      session.undoStack.pop()
      return null
    }
    return rebuildSlide(session, op.slideIndex)
  })

  // Whole-selection delete as one undo step. per_op mirrors the old per-element
  // loop: an id that already vanished is skipped instead of aborting the rest.
  ipcMain.handle('slides:delete-elements', (e, op: DeleteElementsOp) => {
    const session = sessions.get(e.sender.id)
    if (!session || !op.sourceIds.length) return null
    const r = sessionTxn(session, {
      isolation: 'per_op',
      ops: op.sourceIds.map((id) => ({
        op: 'deleteElement' as const,
        target: { slide: op.slideIndex, el: id },
      })),
    })
    return r ? rebuildSlide(session, op.slideIndex) : null
  })

  // Shim over the canonical setStroke op: pt→EMU/angle conversion is surface translation
  // and stays here; validation/mutation/journal live in the op.
  ipcMain.handle('slides:edit-stroke', (e, op: EditStrokeOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const patch = op.stroke
      ? {
          color: op.stroke.color,
          widthEmu: Math.round(op.stroke.widthPt * EMU_PER_PT),
          ...(op.stroke.dash ? { dash: op.stroke.dash } : {}),
          ...(op.stroke.cap ? { cap: op.stroke.cap } : {}),
          ...(op.stroke.join ? { join: op.stroke.join } : {}),
          ...(op.stroke.compound ? { compound: op.stroke.compound } : {}),
          ...(op.stroke.gradient
            ? {
                gradient: {
                  stops: op.stroke.gradient.stops,
                  angle: Math.round(op.stroke.gradient.angleDeg * 60000),
                },
              }
            : {}),
        }
      : null
    pushHistory(session)
    const r = journaledTxn(session, 'edit', {
      ops: [
        {
          op: 'setStroke',
          target: { slide: op.slideIndex, el: op.sourceId },
          stroke: patch,
          ...(op.groupId ? { group: op.groupId } : {}),
        },
      ],
    })
    if (!r.applied) {
      session.undoStack.pop()
      return null
    }
    return rebuildSlide(session, op.slideIndex)
  })

  // Mirror elements across their own axis: flipH/flipV is the only way to
  // point an arrow the other way — rotation cannot express a single-axis mirror
  ipcMain.handle('slides:flip-elements', (e, op: FlipElementOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'flipElements',
          target: { slide: op.slideIndex },
          els: op.sourceIds,
          axis: op.axis,
          ...(op.groupId ? { group: op.groupId } : {}),
        },
      ],
    })
    return r ? rebuildSlide(session, op.slideIndex) : null
  })

  ipcMain.handle('slides:edit-picture-src-rect', (e, op: EditPictureSrcRectOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    let box: { x: number; y: number; cx: number; cy: number } | undefined
    if (op.boxPx && op.fitWidthPx) {
      const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
      const scale = op.fitWidthPx / baseWidthPx
      const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
      box = {
        x: toEmu(op.boxPx.x),
        y: toEmu(op.boxPx.y),
        cx: toEmu(op.boxPx.w),
        cy: toEmu(op.boxPx.h),
      }
    }
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'setPictureSrcRect',
          target: { slide: op.slideIndex, el: op.sourceId },
          srcRect: op.srcRect,
          ...(box ? { box } : {}),
        },
      ],
    })
    return r ? rebuildSlide(session, op.slideIndex) : null
  })

  ipcMain.handle('slides:group-elements', (e, op: GroupElementsOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, {
      ops: [{ op: 'groupElements', target: { slide: op.slideIndex }, els: op.sourceIds }],
    })
    if (!r) return null
    const renderSlide = rebuildSlide(session, op.slideIndex)
    return renderSlide ? { slide: renderSlide, groupId: r.records![0]!.created![0]! } : null
  })

  ipcMain.handle('slides:ungroup-element', (e, op: UngroupElementOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, {
      ops: [{ op: 'ungroupElement', target: { slide: op.slideIndex, el: op.sourceId } }],
    })
    return r ? rebuildSlide(session, op.slideIndex) : null
  })

  // Shim over the canonical setBackground op (one slide per op; apply-to-all fans out;
  // the full-bleed backdrop repaint lives in the op). File dialogs stay here.
  ipcMain.handle('slides:edit-background', async (e, op: EditBackgroundOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slides = session.opened.deck.slides
    const targets = op.slideIndex === -1 ? slides : [slides[op.slideIndex]].filter(Boolean)
    if (targets.length === 0) return null

    if (op.kind === 'image') {
      let source: { bytes: Uint8Array; ext: string } | { mediaPath: string }
      if (op.pick !== false) {
        const r = await showOpenDialogWithMemory(dialog, dialogParent(), {
          title: tm('dlgInsertImage'),
          properties: ['openFile' as const],
          filters: [
            {
              name: tm('filterImages'),
              extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'tif', 'tiff'],
            },
          ],
        })
        if (r.canceled || !r.filePaths[0]) return null
        const bytes = await readFile(r.filePaths[0])
        source = {
          bytes: new Uint8Array(bytes),
          ext: r.filePaths[0].split('.').pop()!.toLowerCase(),
        }
      } else {
        // Reuse an already-landed background image (mode change / apply-to-all)
        const src = slides[op.sourceSlideIndex ?? op.slideIndex]
        if (src?.background?.type !== 'image') return null
        source = { mediaPath: src.background.mediaRef }
      }
      pushHistory(session)
      // The picked bytes land as one media part; further slides only add a rel to it
      // (each op reports the landed media path, threaded into the next op's source)
      let landed: string | null = null
      for (const [i, s] of slides.entries()) {
        if (!targets.includes(s)) continue
        const r = journaledTxn(session, 'edit', {
          ops: [
            {
              op: 'setBackground',
              target: { slide: i },
              kind: 'image',
              source: landed ? { mediaPath: landed } : source,
              tile: op.mode === 'tile',
            },
          ],
        })
        if (r.applied) {
          const used = (r.records![0]!.after as { mediaPath?: string } | undefined)?.mediaPath
          if (used) landed = used
        }
      }
      if (!landed) {
        session.undoStack.pop()
        return null
      }
    } else {
      const common: Record<string, unknown> =
        op.kind === 'solid'
          ? { kind: 'solid', color: op.color }
          : op.kind === 'gradient'
            ? {
                kind: 'gradient',
                from: op.from,
                to: op.to,
                ...(op.angleDeg !== undefined ? { angleDeg: op.angleDeg } : {}),
                ...(op.radial ? { radial: true } : {}),
              }
            : op.kind === 'reset'
              ? { kind: 'reset' }
              : { kind: 'graphics', hidden: op.hidden }
      pushHistory(session)
      const r = journaledTxn(session, 'edit', {
        ops: slides
          .map((s, i) => ({ s, i }))
          .filter(({ s }) => targets.includes(s))
          .map(({ i }) => ({ op: 'setBackground', target: { slide: i }, ...common })),
      })
      if (!r.applied) {
        session.undoStack.pop()
        return null
      }
    }
    session.fitWidthPx = op.fitWidthPx
    return buildAllRenderSlides(session.opened, op.fitWidthPx)
  })

  ipcMain.handle('slides:edit-image-fill', async (e, op: EditFillImageOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide || op.targets.length === 0) return null
    let bytes: Uint8Array
    let ext: string
    if (op.source) {
      // bundled texture preset: bytes shipped inline, no picker
      bytes = new Uint8Array(Buffer.from(op.source.base64, 'base64'))
      ext = op.source.ext.toLowerCase()
    } else {
      const r = await showOpenDialogWithMemory(dialog, dialogParent(), {
        title: tm('dlgInsertImage'),
        properties: ['openFile' as const],
        filters: [
          {
            name: tm('filterImages'),
            extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'tif', 'tiff'],
          },
        ],
      })
      if (r.canceled || !r.filePaths[0]) return null
      bytes = new Uint8Array(await readFile(r.filePaths[0]))
      ext = r.filePaths[0].split('.').pop()!.toLowerCase()
    }
    pushHistory(session)
    // The picked bytes land as one media part; further targets only add rels to it
    // (each op reports the landed media path, threaded into the next op's source)
    let landed: string | null = null
    let applied = 0
    for (const target of op.targets) {
      const r = journaledTxn(session, 'edit', {
        ops: [
          {
            op: 'setImageFill',
            target: { slide: op.slideIndex, el: target.sourceId },
            source: landed ? { mediaPath: landed } : { bytes, ext },
            tile: op.mode === 'tile',
            ...(target.groupId ? { group: target.groupId } : {}),
          },
        ],
      })
      if (r.applied) {
        applied += 1
        const used = (r.records![0]!.after as { mediaPath?: string } | undefined)?.mediaPath
        if (used) landed = used
      }
    }
    if (!applied) {
      session.undoStack.pop()
      return null
    }
    return rebuildSlide(session, op.slideIndex)
  })

  // Replace picture: the renderer swaps the bytes in place through replacePictureBytes
  ipcMain.handle('slides:pick-picture-file', async () => {
    const r = await showOpenDialogWithMemory(dialog, dialogParent(), {
      title: tm('dlgReplacePicture'),
      properties: ['openFile' as const],
      filters: [
        {
          name: tm('filterImages'),
          extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'tif', 'tiff'],
        },
      ],
    })
    if (r.canceled || !r.filePaths[0]) return null
    const filePath = r.filePaths[0]
    return {
      base64: (await readFile(filePath)).toString('base64'),
      ext: filePath.split('.').pop()!.toLowerCase(),
    }
  })

  ipcMain.handle('slides:insert-image', async (e, slideIndex: number, fitWidthPx: number) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[slideIndex]
    if (!slide) return null
    const parent = dialogParent()
    const options = {
      title: tm('dlgInsertImage'),
      properties: ['openFile' as const],
      filters: [
        {
          name: tm('filterImages'),
          extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'tif', 'tiff'],
        },
      ],
    }
    const r = await showOpenDialogWithMemory(dialog, parent, options)
    if (r.canceled || !r.filePaths[0]) return null
    const filePath = r.filePaths[0]
    const bytes = await readFile(filePath)
    const ext = filePath.split('.').pop()!.toLowerCase()

    const deckSize = session.opened.deck.size
    let natural: { width: number; height: number } | null = null
    if (ext === 'tif' || ext === 'tiff') {
      const decoded = tiffToPng(new Uint8Array(bytes))
      if (decoded) natural = { width: decoded.width, height: decoded.height }
    } else {
      const img = nativeImage.createFromPath(filePath)
      if (!img.isEmpty()) natural = img.getSize()
    }
    const offset = pictureFrame(
      { width: deckSize.cx, height: deckSize.cy },
      natural,
      imageDpiFromBytes(new Uint8Array(bytes)),
      DEFAULT_PICTURE_DPI,
    )

    const txn = sessionTxn(session, {
      ops: [
        {
          op: 'addPicture',
          target: { slide: slideIndex },
          bytes: new Uint8Array(bytes),
          ext,
          offset,
        },
      ],
    })
    if (!txn) return { error: 'unsupported' as const, ext }
    session.fitWidthPx = fitWidthPx
    const rebuilt = rebuildSlide(session, slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: txn.records![0]!.created![0]! } : null
  })

  // Shim over the canonical setFill op: gradient normalization is surface translation
  // and stays here; validation/mutation/journal live in the op.
  ipcMain.handle('slides:edit-fill', (e, op: EditFillOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const fill =
      typeof op.fill === 'string'
        ? op.fill
        : {
            stops: gradientStops(op.fill.gradient),
            ...((pk) =>
              pk
                ? {
                    path: pk,
                    ...((ft) => (ft ? { fillTo: ft } : {}))(gradientFillTo(op.fill.gradient)),
                  }
                : { angle: Math.round((op.fill.gradient.angleDeg ?? 0) * 60000) })(
              gradientPathKind(op.fill.gradient),
            ),
          }
    pushHistory(session)
    const r = journaledTxn(session, 'edit', {
      ops: [
        {
          op: 'setFill',
          target: { slide: op.slideIndex, el: op.sourceId },
          fill,
          ...(op.groupId ? { group: op.groupId } : {}),
        },
      ],
    })
    if (!r.applied) {
      session.undoStack.pop()
      return null
    }
    return rebuildSlide(session, op.slideIndex)
  })

  ipcMain.handle('slides:add-slide', (e, op: AddSlideOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'duplicateSlide',
          target: { slide: op.sourceIndex },
          ...(op.clearText ? { clearText: true } : {}),
        },
      ],
    })
    if (!r) return null
    session.fitWidthPx = op.fitWidthPx
    markMetaDirty(session)
    return {
      slides: buildAllRenderSlides(session.opened, op.fitWidthPx),
      index: op.sourceIndex + 1,
    }
  })

  // App-wide, so a slide copied in one tab can be pasted into another deck.
  ipcMain.handle('slides:copy-slides', (e, op: CopySlidesOp) => {
    const session = sessions.get(e.sender.id)
    if (!session || !op.slideIndexes.length) return false
    const bundles: SlideBundle[] = []
    for (const i of op.slideIndexes) {
      const bundle = copySlide(session.opened, i)
      if (!bundle) return false
      bundles.push(bundle)
    }
    const pngs = op.pngs?.length === bundles.length ? op.pngs : undefined
    slideClipboard = { bundles, ...(pngs ? { pngs } : {}) }
    // Marker so plain ⌘V knows the latest copy was a slide (element copies / external copies overwrite it)
    clipboard.writeBuffer('io.genoffice.slides.slide', Buffer.from('1'))
    return true
  })

  ipcMain.handle('slides:has-slide-clipboard', () => slideClipboard !== null)

  // Several copied slides paste in order; 'picture' mode stacks every bitmap on the anchor slide.
  const performSlidePaste = (session: Session, op: PasteSlideOp): PasteSlideResult | null => {
    if (!slideClipboard) return null
    const picture = op.mode === 'picture'
    const { bundles, pngs } = slideClipboard
    if (picture && !pngs) return null
    const r = journaledTxn(session, 'edit', {
      ops: bundles.map((bundle, j) => ({
        op: 'pasteSlide',
        afterIndex: picture ? op.afterIndex : op.afterIndex + j,
        mode: op.mode,
        ...(picture ? { png: pngs![j] } : { bundle }),
      })),
    })
    if (!r.applied) return null
    const rec = r.records![0]!
    session.fitWidthPx = op.fitWidthPx
    // Pasted slides are parsed fresh, so no element dirty flag is set: flag the
    // session here or the paste is invisible to the close guard and autosave.
    // repaste-slide re-runs this after restoring a snapshot that cleared the flag.
    markMetaDirty(session)
    const created = r.records!.flatMap((x) => x.created ?? [])
    return {
      slides: buildAllRenderSlides(session.opened, op.fitWidthPx),
      index: (rec.after as { index: number }).index,
      count: picture ? 0 : bundles.length,
      ...(created.length ? { sourceIds: created } : {}),
    }
  }

  ipcMain.handle('slides:paste-slide', (e, op: PasteSlideOp) => {
    const session = sessions.get(e.sender.id)
    if (!session || !slideClipboard) return null
    pushHistory(session)
    const r = performSlidePaste(session, op)
    if (!r) {
      session.undoStack.pop()
      return null
    }
    lastSlidePaste.set(e.sender.id, {
      afterIndex: op.afterIndex,
      undoLen: session.undoStack.length,
    })
    return r
  })

  // Paste-options floater: undo the just-completed paste and redo it with another
  // mode. Refused when anything (edits, ⌘Z) touched the deck in between.
  ipcMain.handle('slides:repaste-slide', (e, op: RepasteSlideOp) => {
    const session = sessions.get(e.sender.id)
    const rec = lastSlidePaste.get(e.sender.id)
    if (!session || !slideClipboard || !rec) return null
    if (session.undoStack.length !== rec.undoLen) return null
    const snap = session.undoStack.pop()
    if (!snap) return null
    restoreSnapshot(session, snap)
    pushHistory(session)
    const r = performSlidePaste(session, {
      afterIndex: rec.afterIndex,
      fitWidthPx: op.fitWidthPx,
      mode: op.mode,
    })
    if (!r) {
      session.undoStack.pop()
      return null
    }
    rec.undoLen = session.undoStack.length
    return r
  })

  ipcMain.handle('slides:add-blank-slide', (e, op: AddBlankSlideOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, {
      ops: [
        { op: 'addBlankSlide', target: { slide: op.sourceIndex } },
        // ops are validated against the pre-transaction deck, so move the existing
        // source slide down past the new one rather than targeting the new index
        ...(op.before
          ? [{ op: 'moveSlide', target: { slide: op.sourceIndex }, to: op.sourceIndex + 1 }]
          : []),
      ],
    })
    if (!r) return null
    session.fitWidthPx = op.fitWidthPx
    // Always: the blank slide is parsed fresh, so neither structureDirty nor an
    // element dirty flag is set and the insert would otherwise be unsaveable.
    markMetaDirty(session)
    return {
      slides: buildAllRenderSlides(session.opened, op.fitWidthPx),
      index: op.before ? op.sourceIndex : op.sourceIndex + 1,
    }
  })

  // 'builtin:<key>' virtual paths get injected into the package on first use
  const resolveLayoutPath = (session: Session, layoutPath?: string): string | undefined => {
    if (!layoutPath?.startsWith(BUILTIN_LAYOUT_PREFIX)) return layoutPath
    return (
      ensureBuiltinLayout(
        session.opened.archive,
        session.opened.deck.size,
        layoutPath.slice(BUILTIN_LAYOUT_PREFIX.length),
      ) ?? undefined
    )
  }

  ipcMain.handle('slides:add-slide-with-layout', (e, op: AddSlideWithLayoutOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    // 'builtin:' resolution may inject a layout part — do it before the undo snapshot
    // is taken so a failed insert still rolls back to a consistent package
    pushHistory(session)
    const layoutPath = resolveLayoutPath(session, op.layoutPath)
    const r = layoutPath
      ? journaledTxn(session, 'edit', {
          ops: [{ op: 'addSlideWithLayout', target: { slide: op.sourceIndex }, layoutPath }],
        })
      : null
    if (!r?.applied) {
      session.undoStack.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    markMetaDirty(session)
    return {
      slides: buildAllRenderSlides(session.opened, op.fitWidthPx),
      index: op.sourceIndex + 1,
    }
  })

  ipcMain.handle('slides:get-layouts', (e) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const layouts = listSlideLayouts(session.opened.archive)
    // Decks whose own layouts carry no placeholders (AI-generated single blank layout)
    // get the built-in standard set, injected into the package on first use
    if (shouldOfferBuiltinLayouts(layouts)) {
      layouts.push(
        ...builtinLayoutInfos(session.opened.deck.size, new Set(layouts.map((l) => l.name))),
      )
    }
    return { layouts, size: { ...session.opened.deck.size } }
  })

  // ── Master edit view ───────────────────────────────────────────────
  // Exception to the fidelity rule: only parts the user actively changed in master view are
  // written back, using the same byte surgery as slides. Every commit writes the entry + fully
  // reparses all slides — inheritance takes effect immediately, and each undo snapshot's
  // (slides model, entries) pair stays self-consistent (rendering and file don't diverge after
  // undo).
  const buildMasterRenderSlide = (session: Session): RenderSlide | null => {
    const me = session.masterEdit
    if (!me) return null
    return buildRenderSlide(me.slide, session.opened.deck.size, {
      fitWidthPx: session.fitWidthPx,
      media: makeMediaResolver(session.opened),
      metrics: getFontMetrics(),
    })
  }

  // Master edits run as part-addressed op transactions seeded with the live
  // master slide: ops mutate me.slide itself (parse-time ids stay stable for
  // the renderer's selection/editing state — a fresh parse would re-mint them)
  // and the executor's flush serializes that same object to the entry.
  const masterTxn = (session: Session, me: { partPath: string; slide: Slide }, op: Op) =>
    sessionTxn(session, { ops: [op], parts: new Map([[me.partPath, me.slide]]) })

  const masterEditDone = (session: Session): RenderSlide | null => {
    markMetaDirty(session)
    return buildMasterRenderSlide(session)
  }

  ipcMain.handle('slides:master-enter', (e, fitWidthPx: number): MasterEnterResult | null => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    session.fitWidthPx = fitWidthPx
    const items: MasterEnterResult['items'] = []
    for (const p of listMasterParts(session.opened.archive)) {
      const slide = parseMasterPart(session.opened.archive, p.partPath)
      if (!slide) continue
      const rendered = buildRenderSlide(slide, session.opened.deck.size, {
        fitWidthPx,
        media: makeMediaResolver(session.opened),
        metrics: getFontMetrics(),
      })
      items.push({ partPath: p.partPath, kind: p.kind, name: p.name, slide: rendered })
      if (!session.masterEdit) session.masterEdit = { partPath: p.partPath, slide }
    }
    return items.length ? { items } : null
  })

  ipcMain.handle('slides:master-open', (e, partPath: string) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = parseMasterPart(session.opened.archive, partPath)
    if (!slide) return null
    session.masterEdit = { partPath, slide }
    return buildMasterRenderSlide(session)
  })

  ipcMain.handle('slides:master-close', (e) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    session.masterEdit = null
    // Edits were materialized one by one; here we only fetch the full render tree
    return buildAllRenderSlides(session.opened, session.fitWidthPx)
  })

  ipcMain.handle('slides:master-edit-text', (e, op: MasterEditTextOp) => {
    const session = sessions.get(e.sender.id)
    const me = session?.masterEdit
    if (!session || !me) return null
    const r = masterTxn(session, me, {
      op: 'setText',
      target: { part: me.partPath, el: op.sourceId },
      paragraphs: op.paragraphs,
    })
    if (!r) return null
    return masterEditDone(session)
  })

  ipcMain.handle('slides:master-edit-transform', (e, op: MasterEditTransformOp) => {
    const session = sessions.get(e.sender.id)
    const me = session?.masterEdit
    if (!session || !me) return null
    const el = me.slide.elements.find((x) => x.id === op.sourceId)
    if (!el) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    if (op.preview) {
      // Previews are not persisted: mutate the live master slide only (the final
      // commit at drag end goes through the op executor against the entry bytes)
      if (!session.transformPreview) {
        pushHistory(session)
        session.transformPreview = true
      }
      el.transform = {
        ...el.transform,
        offset: { x: toEmu(op.xPx), y: toEmu(op.yPx), cx: toEmu(op.wPx), cy: toEmu(op.hPx) },
        rot: Math.round(op.rotationDeg * 60000),
      }
      return buildMasterRenderSlide(session)
    }
    const payload = {
      op: 'setTransform',
      target: { part: me.partPath, el: op.sourceId },
      box: { x: toEmu(op.xPx), y: toEmu(op.yPx), cx: toEmu(op.wPx), cy: toEmu(op.hPx) },
      rotDeg: op.rotationDeg,
    }
    const seed = new Map([[me.partPath, me.slide]])
    let r: TxnResult | null
    if (session.transformPreview) {
      // The gesture's first preview frame already pushed the undo step
      session.transformPreview = false
      const applied = journaledTxn(session, 'edit', { ops: [payload], parts: seed })
      r = applied.applied ? applied : null
    } else {
      r = sessionTxn(session, { ops: [payload], parts: seed })
    }
    if (!r) return null
    return masterEditDone(session)
  })

  ipcMain.handle('slides:master-edit-fill', (e, op: MasterEditFillOp) => {
    const session = sessions.get(e.sender.id)
    const me = session?.masterEdit
    if (!session || !me) return null
    const fill =
      typeof op.fill === 'string'
        ? op.fill
        : {
            stops: gradientStops(op.fill.gradient),
            ...((pk) =>
              pk
                ? {
                    path: pk,
                    ...((ft) => (ft ? { fillTo: ft } : {}))(gradientFillTo(op.fill.gradient)),
                  }
                : { angle: Math.round((op.fill.gradient.angleDeg ?? 0) * 60000) })(
              gradientPathKind(op.fill.gradient),
            ),
          }
    const r = masterTxn(session, me, {
      op: 'setFill',
      target: { part: me.partPath, el: op.sourceId },
      fill,
    })
    if (!r) return null
    return masterEditDone(session)
  })

  ipcMain.handle('slides:master-edit-stroke', (e, op: MasterEditStrokeOp) => {
    const session = sessions.get(e.sender.id)
    const me = session?.masterEdit
    if (!session || !me) return null
    const r = masterTxn(session, me, {
      op: 'setStroke',
      target: { part: me.partPath, el: op.sourceId },
      stroke: op.stroke
        ? { color: op.stroke.color, widthEmu: Math.round(op.stroke.widthPt * EMU_PER_PT) }
        : null,
    })
    if (!r) return null
    return masterEditDone(session)
  })

  ipcMain.handle('slides:master-delete-element', (e, op: MasterDeleteElementOp) => {
    const session = sessions.get(e.sender.id)
    const me = session?.masterEdit
    if (!session || !me) return null
    const r = masterTxn(session, me, {
      op: 'deleteElement',
      target: { part: me.partPath, el: op.sourceId },
    })
    if (!r) return null
    return masterEditDone(session)
  })

  ipcMain.handle('slides:edit-picture-opacity', (e, op: EditPictureOpacityOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'setPictureOpacity',
          target: { slide: op.slideIndex, el: op.sourceId },
          opacity: op.opacity,
        },
      ],
    })
    return r ? rebuildSlide(session, op.slideIndex) : null
  })

  ipcMain.handle('slides:set-slide-size', (e, op: SetSlideSizeOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, { ops: [{ op: 'setSlideSize', cx: op.cx, cy: op.cy }] })
    if (!r) return null
    markMetaDirty(session)
    return buildAllRenderSlides(session.opened, session.fitWidthPx)
  })

  ipcMain.handle('slides:get-slide-size', (e) => {
    const session = sessions.get(e.sender.id)
    return session ? { ...session.opened.deck.size } : null
  })

  ipcMain.handle('slides:set-slide-layout', (e, op: SetSlideLayoutOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    pushHistory(session)
    const layoutPath = resolveLayoutPath(session, op.layoutPath)
    // A named layout that fails to resolve is an error; an absent layoutPath means reset
    const r =
      op.layoutPath && !layoutPath
        ? null
        : journaledTxn(session, 'edit', {
            ops: [
              {
                op: 'setSlideLayout',
                target: { slide: op.slideIndex },
                ...(layoutPath ? { layoutPath } : {}),
              },
            ],
          })
    if (!r?.applied) {
      session.undoStack.pop()
      return null
    }
    return rebuildSlide(session, op.slideIndex)
  })

  ipcMain.handle('slides:find-replace', (e, op: FindReplaceOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'findReplace',
          find: op.find,
          replace: op.replace,
          matchCase: op.matchCase,
          firstOnly: op.firstOnly,
          slideIndex: op.slideIndex,
          elementId: op.elementId,
        },
      ],
    })
    if (!r) return { count: 0, slides: null }
    const count = (r.records![0]!.after as { count: number }).count
    return { count, slides: buildAllRenderSlides(session.opened, session.fitWidthPx) }
  })

  ipcMain.handle('slides:delete-slide', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, { ops: [{ op: 'deleteSlide', target: { slide: slideIndex } }] })
    if (!r) return null
    markMetaDirty(session)
    return buildAllRenderSlides(session.opened, session.fitWidthPx)
  })

  // Highest index first: every op is validated against the pre-transaction deck
  // and applied in sequence, so the remaining indexes stay valid.
  ipcMain.handle('slides:delete-slides', (e, op: DeleteSlidesOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const indexes = [...new Set(op.slideIndexes)].sort((a, b) => b - a)
    if (!indexes.length || indexes.length >= session.opened.deck.slides.length) return null
    const r = sessionTxn(session, {
      ops: indexes.map((i) => ({ op: 'deleteSlide' as const, target: { slide: i } })),
    })
    if (!r) return null
    markMetaDirty(session)
    return buildAllRenderSlides(session.opened, session.fitWidthPx)
  })

  ipcMain.handle('slides:duplicate-slides', (e, op: DuplicateSlidesOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const steps = planSlideDuplicates(op.slideIndexes)
    if (!steps.length) return null
    const r = sessionTxn(session, {
      ops: steps.map((st) =>
        'duplicate' in st
          ? { op: 'duplicateSlide' as const, target: { slide: st.duplicate } }
          : { op: 'moveSlide' as const, target: { slide: st.from }, to: st.to },
      ),
    })
    if (!r) return null
    session.fitWidthPx = op.fitWidthPx
    // Not just for multi-slide plans: a single duplicate is parsed fresh too and
    // would otherwise leave the deck changed but reported clean.
    markMetaDirty(session)
    return {
      slides: buildAllRenderSlides(session.opened, op.fitWidthPx),
      index: Math.max(...op.slideIndexes) + 1,
    }
  })

  ipcMain.handle('slides:set-slides-hidden', (e, op: SetSlidesHiddenOp) => {
    const session = sessions.get(e.sender.id)
    if (!session || !op.slideIndexes.length) return null
    const r = sessionTxn(session, {
      ops: op.slideIndexes.map((i) => ({
        op: 'setHidden' as const,
        target: { slide: i },
        hidden: op.hidden,
      })),
    })
    return r ? buildAllRenderSlides(session.opened, session.fitWidthPx) : null
  })

  ipcMain.handle('slides:edit-table-cell', (e, op: EditTableCellOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'setTableCell',
          target: { slide: op.slideIndex, el: op.sourceId },
          row: op.row,
          col: op.col,
          paragraphs: op.paragraphs,
        },
      ],
    })
    return r ? rebuildSlide(session, op.slideIndex) : null
  })

  ipcMain.handle('slides:table-merge', (e, op: TableMergeIpcOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'tableMerge',
          target: { slide: op.slideIndex, el: op.sourceId },
          kind: op.kind,
          row: op.row,
          col: op.col,
        },
      ],
    })
    if (!r) return null
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt
      ? { slide: rebuilt, sourceId: (r.records![0]!.after as { elementId: string }).elementId }
      : null
  })

  ipcMain.handle('slides:table-structure', (e, op: TableStructureIpcOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'tableStructure',
          target: { slide: op.slideIndex, el: op.sourceId },
          kind: op.kind,
          index: op.index,
          ...(op.before ? { before: true } : {}),
        },
      ],
    })
    if (!r) return null
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt
      ? { slide: rebuilt, sourceId: (r.records![0]!.after as { elementId: string }).elementId }
      : null
  })

  ipcMain.handle('slides:set-table-row-height', (e, op: SetTableRowHeightOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'setTableRowHeight',
          target: { slide: op.slideIndex, el: op.sourceId },
          row: op.row,
          hEmu: (op.hPx / scale) * EMU_PER_PX_96,
        },
      ],
    })
    return r ? rebuildSlide(session, op.slideIndex) : null
  })

  ipcMain.handle('slides:set-table-cell-anchor', (e, op: SetTableCellAnchorOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'setTableCellAnchor',
          target: { slide: op.slideIndex, el: op.sourceId },
          row: op.row,
          col: op.col,
          anchor: op.anchor,
        },
      ],
    })
    return r ? rebuildSlide(session, op.slideIndex) : null
  })

  ipcMain.handle('slides:set-table-col-width', (e, op: SetTableColWidthOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'setTableColWidth',
          target: { slide: op.slideIndex, el: op.sourceId },
          col: op.col,
          wEmu: (op.wPx / scale) * EMU_PER_PX_96,
        },
      ],
    })
    return r ? rebuildSlide(session, op.slideIndex) : null
  })

  ipcMain.handle('slides:edit-table-style', (e, op: EditTableStyleOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    // A reparse regenerates element ids: look up the new id by element index; the renderer uses it to keep the selection
    const elIdx = slide.elements.findIndex((el) => matchesElementRef(el, op.sourceId))
    pushHistory(session)
    const { slideIndex, sourceId, ...style } = op
    const r = journaledTxn(session, 'edit', {
      ops: [{ op: 'setTableStyle', target: { slide: slideIndex, el: sourceId }, ...style }],
    })
    if (!r.applied) {
      session.undoStack.pop()
      return null
    }
    // The patch is written on anchor.originalXml; a materialize reparse is needed before it shows in the render model
    const rebuilt = rebuildSlideWithReparse(session, op.slideIndex)
    if (!rebuilt) return null
    const newId = session.opened.deck.slides[op.slideIndex]?.elements[elIdx]?.id ?? null
    return { slide: rebuilt, sourceId: newId }
  })

  ipcMain.handle('slides:edit-chart', async (e, op: EditChartOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    // A reparse regenerates element ids: look up the new id by element index; the renderer uses it to keep the selection
    const elIdx = slide.elements.findIndex((el) => matchesElementRef(el, op.sourceId))
    // Confirm before the first edit of a chart from an imported file: editing rebuilds it from the template,
    // and unmodeled fine-grained formatting (number formats/trendlines/error bars/per-point styles) is lost
    const chartEl = slide.elements[elIdx] as { type?: string; descr?: string } | undefined
    if (chartEl?.type === 'chart' && chartEl.descr !== 'aislides-chart') {
      const parent = dialogParent()
      const options = {
        type: 'warning' as const,
        buttons: [tm('chartSimplifyOk'), tm('btnCancel')],
        defaultId: 0,
        cancelId: 1,
        message: tm('chartSimplifyTitle'),
        detail: tm('chartSimplifyBody'),
      }
      const r = parent
        ? await dialog.showMessageBox(parent, options)
        : await dialog.showMessageBox(options)
      if (r.response !== 0) return null
    }
    const patch: Parameters<typeof editChartElement>[3] = {
      ...(op.kind ? { kind: op.kind === 'barH' ? 'bar' : op.kind } : {}),
      ...(op.kind === 'barH' ? { barDir: 'bar' as const } : {}),
      ...(op.categories ? { categories: op.categories } : {}),
      ...(op.series ? { series: op.series } : {}),
      ...(op.title !== undefined ? { title: op.title } : {}),
      ...(op.colorScheme
        ? {
            colorScheme:
              chartColorSchemes(session.opened).find((s) => s.key === op.colorScheme)?.colors ??
              CHART_COLOR_SCHEMES[op.colorScheme],
          }
        : {}),
      ...(op.legendPos ? { legendPos: op.legendPos } : {}),
      ...(op.dataLabels !== undefined ? { dataLabels: op.dataLabels } : {}),
      ...(op.gridlines !== undefined ? { gridlines: op.gridlines } : {}),
      ...(op.catAxisTitle !== undefined ? { catAxisTitle: op.catAxisTitle } : {}),
      ...(op.valAxisTitle !== undefined ? { valAxisTitle: op.valAxisTitle } : {}),
      ...(op.gapWidthPct !== undefined ? { gapWidthPct: op.gapWidthPct } : {}),
      ...(op.switchRowCol ? { switchRowCol: true } : {}),
      ...(op.pointColors ? { pointColors: op.pointColors } : {}),
    }
    const r = sessionTxn(session, {
      ops: [{ op: 'setChart', target: { slide: op.slideIndex, el: op.sourceId }, patch }],
    })
    if (!r) return null
    // The chart part XML is updated; reparse the whole page to refresh the model
    const rebuilt = rebuildSlideWithReparse(session, op.slideIndex)
    if (!rebuilt) return null
    const newId = session.opened.deck.slides[op.slideIndex]?.elements[elIdx]?.id ?? null
    return { slide: rebuilt, sourceId: newId }
  })

  ipcMain.handle('slides:chart-color-schemes', (e) => {
    const session = sessions.get(e.sender.id)
    return session ? chartColorSchemes(session.opened) : null
  })

  ipcMain.handle('slides:get-chart-data', (e, slideIndex: number, sourceId: string) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[slideIndex]
    if (!slide) return null
    return getChartElementData(slide, sourceId)
  })

  ipcMain.handle('slides:reorder-element', (e, op: ReorderElementOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, {
      ops: [
        { op: 'reorderElement', target: { slide: op.slideIndex, el: op.sourceId }, dir: op.dir },
      ],
    })
    return r ? rebuildSlide(session, op.slideIndex) : null
  })

  ipcMain.handle(
    'slides:change-shape',
    (e, op: { slideIndex: number; sourceId: string; prst: string; groupId?: string }) => {
      const session = sessions.get(e.sender.id)
      if (!session) return null
      const r = sessionTxn(session, {
        ops: [
          {
            op: 'setShapeGeometry',
            target: { slide: op.slideIndex, el: op.sourceId },
            prst: op.prst,
            ...(op.groupId ? { group: op.groupId } : {}),
          },
        ],
      })
      return r ? rebuildSlide(session, op.slideIndex) : null
    },
  )

  ipcMain.handle(
    'slides:set-shape-adjust',
    (
      e,
      op: {
        slideIndex: number
        sourceId: string
        adjust: Record<string, number>
        groupId?: string
        preview?: boolean
      },
    ) => {
      const session = sessions.get(e.sender.id)
      if (!session) return null
      const payload = {
        op: 'setShapeAdjust',
        target: { slide: op.slideIndex, el: op.sourceId },
        adjust: op.adjust,
        ...(op.groupId ? { group: op.groupId } : {}),
      }
      // Validate before the preview bookkeeping (same contract as edit-transform):
      // a failed first preview must not set transformPreview or clear the redo stack.
      if (runTxn(session.opened, { dryRun: true, ops: [payload] }).failures?.length) return null
      // Gesture undo semantics: one whole drag = one undo step. The first preview
      // pushes a pre-gesture snapshot; later previews and the final commit do not.
      let pushed = false
      if (op.preview) {
        if (!session.transformPreview) {
          pushHistory(session)
          session.transformPreview = true
          pushed = true
        }
      } else if (session.transformPreview) {
        session.transformPreview = false
      } else {
        pushHistory(session)
        pushed = true
      }
      const r = journaledTxn(session, 'edit', { ops: [payload] })
      if (!r.applied) {
        if (pushed) {
          session.undoStack.pop()
          if (op.preview) session.transformPreview = false
        }
        return null
      }
      return rebuildSlide(session, op.slideIndex)
    },
  )

  // Edit Points: same gesture-undo contract as set-shape-adjust (one drag = one undo step)
  ipcMain.handle('slides:set-shape-geometry', (e, op: SetShapeGeometryOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const scale = op.fitWidthPx / (session.opened.deck.size.cx / EMU_PER_PX_96)
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    const payload = {
      op: 'setShapeCustomGeometry',
      target: { slide: op.slideIndex, el: op.sourceId },
      path: {
        w: toEmu(op.pathPx.w),
        h: toEmu(op.pathPx.h),
        cmds: op.pathPx.cmds.map((c) => ({ op: c.op, pts: c.pts.map(toEmu) })),
      },
      ...(op.groupId ? { group: op.groupId } : {}),
    }
    if (runTxn(session.opened, { dryRun: true, ops: [payload] }).failures?.length) return null
    let pushed = false
    if (op.preview) {
      if (!session.transformPreview) {
        pushHistory(session)
        session.transformPreview = true
        pushed = true
      }
    } else if (session.transformPreview) {
      session.transformPreview = false
    } else {
      pushHistory(session)
      pushed = true
    }
    const r = journaledTxn(session, 'edit', { ops: [payload] })
    if (!r.applied) {
      if (pushed) {
        session.undoStack.pop()
        if (op.preview) session.transformPreview = false
      }
      return null
    }
    return rebuildSlide(session, op.slideIndex)
  })

  ipcMain.handle(
    'slides:set-text-anchor',
    (e, op: { slideIndex: number; sourceId: string; anchor: 'top' | 'middle' | 'bottom' }) => {
      const session = sessions.get(e.sender.id)
      if (!session) return null
      const r = sessionTxn(session, {
        ops: [
          {
            op: 'setTextAnchor',
            target: { slide: op.slideIndex, el: op.sourceId },
            anchor: op.anchor,
          },
        ],
      })
      return r ? rebuildSlide(session, op.slideIndex) : null
    },
  )

  ipcMain.handle(
    'slides:set-text-body-props',
    (
      e,
      op: {
        slideIndex: number
        sourceId: string
        props: {
          vert?: 'horz' | 'eaVert' | 'vert' | 'vert270' | 'wordArtVert'
          autofit?: 'none' | 'shrink' | 'resize'
          insets?: Partial<{ l: number; t: number; r: number; b: number }>
          wrap?: boolean
        }
      },
    ) => {
      const session = sessions.get(e.sender.id)
      if (!session) return null
      const r = sessionTxn(session, {
        ops: [
          {
            op: 'setTextBodyProps',
            target: { slide: op.slideIndex, el: op.sourceId },
            props: op.props,
          },
        ],
      })
      if (!r) return null
      // Autofit must take effect immediately on toggle, not only on the next text edit:
      // 'resize' fits the shape height to the content now, 'shrink' runs the ladder and
      // writes the used fontScale back into bodyPr (a bare <a:normAutofit/> opens at
      // 100% in PowerPoint). Insets/wrap/vert changes re-fit under the same rules.
      const rendered = applyAutofitResize(
        session,
        op.slideIndex,
        op.sourceId,
        rebuildSlide(session, op.slideIndex),
      )
      return syncAutofitScale(session, op.slideIndex, op.sourceId, rendered)
    },
  )

  ipcMain.handle(
    'slides:set-effects',
    (e, op: { slideIndex: number; sourceId: string; effects: SetEffectsPatch }) => {
      const session = sessions.get(e.sender.id)
      if (!session) return null
      const r = sessionTxn(session, {
        ops: [
          {
            op: 'setEffects',
            target: { slide: op.slideIndex, el: op.sourceId },
            effects: op.effects,
          },
        ],
      })
      return r ? rebuildSlide(session, op.slideIndex) : null
    },
  )

  // Our marker still present = the last copy came from this app -> use internal element paste
  // (on macOS custom formats don't appear in availableFormats, so check via readBuffer)
  const clipboardMarker = (format: string) => {
    try {
      return clipboard.readBuffer(format).length > 0
    } catch {
      return false
    }
  }

  ipcMain.handle('slides:clipboard-external', () => {
    if (slideClipboard && clipboardMarker('io.genoffice.slides.slide')) return { kind: 'slide' }
    if (elementClipboard && elementClipboardMarkerMatches(elementClipboard.token))
      return { kind: 'internal' }
    const img = clipboard.readImage()
    if (!img.isEmpty()) return { kind: 'image', base64: img.toPNG().toString('base64'), ext: 'png' }
    const text = clipboard.readText()
    if (text.trim()) return { kind: 'text', text }
    return { kind: 'none' }
  })

  // Menu-enable probe: is there anything a paste would act on? (no image decode)
  ipcMain.handle('slides:clipboard-probe', () => {
    if (slideClipboard && clipboardMarker('io.genoffice.slides.slide')) return true
    if (elementClipboard && elementClipboardMarkerMatches(elementClipboard.token)) return true
    if (clipboard.availableFormats().some((f) => f.startsWith('image/'))) return true
    return clipboard.readText().trim().length > 0
  })

  ipcMain.handle('slides:copy-elements', (e, op: CopyElementsOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return 0
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return 0
    const items = op.sourceIds
      .map((id) => slide.elements.find((el) => el.id === id))
      .filter((el): el is NonNullable<typeof el> => !!el)
      .map((el) => copyElementData(session.opened, slide, el))
    if (items.length) {
      const token = isElementClipboardToken(op.clipboardToken) ? op.clipboardToken : randomUUID()
      elementClipboard = {
        items,
        cascade: newPasteCascade(op.cut ? null : pageKey(e.sender.id, op.slideIndex)),
        token,
        senderId: e.sender.id,
      }
      // Write our marker to the OS clipboard: an external copy overwrites it, so at paste time it tells whether internal or external is newer
      clipboard.writeBuffer(ELEMENT_CLIPBOARD_FORMAT, Buffer.from(token))
    }
    return items.length
  })

  ipcMain.handle('slides:copy-elements-image', (e, clipboardToken: string, pngBase64: string) => {
    if (!canWriteElementClipboardImage(elementClipboard, e.sender.id, clipboardToken)) return false
    return writeElementClipboardImage(clipboardToken, pngBase64)
  })

  ipcMain.handle('slides:paste-elements', (e, op: PasteElementsOp) => {
    const session = sessions.get(e.sender.id)
    const clip = elementClipboard
    if (!session || !clip?.items.length) return null
    if (!session.opened.deck.slides[op.slideIndex]) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    // Cascade only past occupied spots: the first paste onto another page lands
    // at the source coordinates exactly; the copy page and repeat
    // pastes keep shifting 16px per landing relative to the original.
    const target = pageKey(e.sender.id, op.slideIndex)
    const shiftPx = pasteShiftPx(clip.cascade, target)
    const shift = Math.round((shiftPx / scale) * EMU_PER_PX_96)
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'pasteElements',
          target: { slide: op.slideIndex },
          items: clip.items,
          dx: shift,
          dy: shift,
        },
      ],
    })
    if (!r) return null
    recordPaste(clip.cascade, target)
    session.fitWidthPx = op.fitWidthPx
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceIds: r.records![0]!.created! } : null
  })

  // Duplicate in place (⌘D / Option+drag copy): does not touch the app clipboard; the caller supplies the offset
  ipcMain.handle('slides:duplicate-elements', (e, op: DuplicateElementsOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const items = op.sourceIds
      .map((id) => slide.elements.find((el) => el.id === id))
      .filter((el): el is NonNullable<typeof el> => !!el)
      .map((el) => copyElementData(session.opened, slide, el))
    if (!items.length) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'pasteElements',
          target: { slide: op.slideIndex },
          items,
          dx: toEmu(op.dxPx),
          dy: toEmu(op.dyPx),
        },
      ],
    })
    if (!r) return null
    session.fitWidthPx = op.fitWidthPx
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceIds: r.records![0]!.created! } : null
  })

  ipcMain.handle('slides:add-table', (e, op: AddTableOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    if (!session.opened.deck.slides[op.slideIndex]) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    const rowH = op.rowHeightEmu && op.rowHeightEmu > 0 ? Math.round(op.rowHeightEmu) : null
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'addTable',
          target: { slide: op.slideIndex },
          rows: op.rows,
          cols: op.cols,
          offset: {
            x: toEmu(op.xPx),
            y: toEmu(op.yPx),
            cx: toEmu(op.wPx),
            cy: rowH ? rowH * op.rows : toEmu(op.hPx),
          },
          ...(rowH ? { rowHeightsEmu: Array.from({ length: op.rows }, () => rowH) } : {}),
        },
      ],
    })
    if (!r) return null
    session.fitWidthPx = op.fitWidthPx
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: r.records![0]!.created![0]! } : null
  })

  // Freehand ink stroke commit: one transparent PNG picture element per stroke (cNvPr name has
  // the aislides-ink prefix, descr stores the vector points as JSON); undo/save/thumbnails all
  // go through the existing picture-element pipeline.
  ipcMain.handle('slides:add-ink', (e, op: AddInkOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'addPicture',
          target: { slide: op.slideIndex },
          bytes: new Uint8Array(Buffer.from(op.base64, 'base64')),
          ext: 'png',
          offset: {
            x: toEmu(op.xPx),
            y: toEmu(op.yPx),
            cx: Math.max(1, toEmu(op.wPx)),
            cy: Math.max(1, toEmu(op.hPx)),
          },
          name: `aislides-ink ${Date.now().toString(36)}`,
          descr: op.payload,
        },
      ],
    })
    if (!r) return null
    session.fitWidthPx = op.fitWidthPx
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: r.records![0]!.created![0]! } : null
  })

  // ── New insert capabilities: charts / SmartArt / icon bitmaps / audio-video / 3D / links / header-footer ──

  ipcMain.handle('slides:add-chart', (e, op: AddChartOp) => {
    const session = sessions.get(e.sender.id)
    if (!session || !session.opened.deck.slides[op.slideIndex]) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'addChart',
          target: { slide: op.slideIndex },
          kind: op.kind === 'barH' ? 'bar' : op.kind,
          ...(op.kind === 'barH' ? { barDir: 'bar' as const } : {}),
          ...(op.title ? { title: op.title } : {}),
          categories: op.categories,
          series: op.series,
          offset: { x: toEmu(op.xPx), y: toEmu(op.yPx), cx: toEmu(op.wPx), cy: toEmu(op.hPx) },
        },
      ],
    })
    if (!r) return null
    session.fitWidthPx = op.fitWidthPx
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: r.records![0]!.created![0]! } : null
  })

  ipcMain.handle('slides:add-smartart', (e, op: AddSmartArtOp) => {
    const session = sessions.get(e.sender.id)
    if (!session || !session.opened.deck.slides[op.slideIndex]) return null
    const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'addSmartArt',
          target: { slide: op.slideIndex },
          layout: op.layout,
          items: op.items,
          offset: { x: toEmu(op.xPx), y: toEmu(op.yPx), cx: toEmu(op.wPx), cy: toEmu(op.hPx) },
        },
      ],
    })
    if (!r) return null
    session.fitWidthPx = op.fitWidthPx
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: r.records![0]!.created![0]! } : null
  })

  ipcMain.handle('slides:add-image-bytes', (e, op: AddImageBytesOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const deckSize = session.opened.deck.size
    const baseWidthPx = deckSize.cx / EMU_PER_PX_96
    const scale = op.fitWidthPx / baseWidthPx
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    const bytes = new Uint8Array(Buffer.from(op.base64, 'base64'))
    let offset: { x: number; y: number; cx: number; cy: number }
    if ('naturalPx' in op) {
      offset = pictureFrame(
        { width: deckSize.cx, height: deckSize.cy },
        op.naturalPx,
        imageDpiFromBytes(bytes),
        DEFAULT_PICTURE_DPI,
        {
          center: op.centerPx ? { x: toEmu(op.centerPx.x), y: toEmu(op.centerPx.y) } : undefined,
          minEmu: toEmu(24),
        },
      )
    } else {
      offset = {
        x: toEmu(op.xPx),
        y: toEmu(op.yPx),
        cx: Math.max(1, toEmu(op.wPx)),
        cy: Math.max(1, toEmu(op.hPx)),
      }
    }
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'addPicture',
          target: { slide: op.slideIndex },
          bytes,
          ext: op.ext,
          offset,
          ...(op.name ? { name: op.name } : {}),
        },
      ],
    })
    if (!r) return { error: 'unsupported' as const, ext: op.ext }
    session.fitWidthPx = op.fitWidthPx
    const rebuilt = rebuildSlide(session, op.slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: r.records![0]!.created![0]! } : null
  })

  ipcMain.handle('slides:replace-picture-bytes', (e, op: ReplacePictureBytesOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'replacePicture',
          target: { slide: op.slideIndex, el: op.sourceId },
          bytes: new Uint8Array(Buffer.from(op.base64, 'base64')),
          ext: op.ext,
          ...(op.keepSrcRect ? { keepSrcRect: true } : {}),
        },
      ],
    })
    if (!r) return { error: 'unsupported' as const, ext: op.ext }
    return rebuildSlide(session, op.slideIndex)
  })

  // In-app playback caveats: AVI has no Chromium demuxer at all; mp4/m4v/mov
  // with e.g. AC-3/DTS audio plays silent.
  const mediaPlaybackWarning = (kind: 'video' | 'audio', ext: string, bytes: Uint8Array) => {
    if (kind !== 'video') return null
    if (ext === 'avi') return tm('mediaAviBody')
    if (ext === 'mp4' || ext === 'm4v' || ext === 'mov') {
      const codec = unplayableAudioCodec(bytes)
      if (codec) return tm('mediaNoAudioBody', { codec })
    }
    return null
  }

  // Video poster frame prefers the system thumbnail (QuickLook), falling back to a solid color on failure.
  const videoPoster = async (filePath: string) => {
    try {
      const thumb = await nativeImage.createThumbnailFromPath(filePath, { width: 960, height: 540 })
      if (!thumb.isEmpty()) return { bytes: new Uint8Array(thumb.toPNG()), ext: 'png' }
    } catch {
      /* Solid-color fallback */
    }
    return undefined
  }

  const embedMedia = async (
    session: Session,
    slideIndex: number,
    kind: 'video' | 'audio',
    bytes: Uint8Array,
    ext: string,
    name: string | undefined,
    opts: { filePath?: string; natural?: Size | null; center?: Point; fitWidthPx: number },
  ): Promise<AddMediaResult | null> => {
    const poster = kind === 'video' && opts.filePath ? await videoPoster(opts.filePath) : undefined
    const deckSize = session.opened.deck.size
    const slideSize = { width: deckSize.cx, height: deckSize.cy }
    const offset =
      kind === 'video'
        ? videoFrame(slideSize, opts.natural ?? videoSize(bytes, ext), opts.center)
        : audioFrame(slideSize, opts.center)
    const txn = sessionTxn(session, {
      ops: [
        {
          op: 'addMedia',
          target: { slide: slideIndex },
          kind,
          bytes,
          ext,
          ...(poster ? { poster } : {}),
          offset,
          ...(name ? { name } : {}),
        },
      ],
    })
    if (!txn) return null
    session.fitWidthPx = opts.fitWidthPx
    const rebuilt = rebuildSlide(session, slideIndex)
    if (!rebuilt) return null
    const warning = mediaPlaybackWarning(kind, ext, bytes)
    return {
      slide: rebuilt,
      sourceId: txn.records![0]!.created![0]!,
      ...(warning ? { warning } : {}),
    }
  }

  // Show a dialog to pick video/audio and embed it.
  ipcMain.handle(
    'slides:insert-media',
    async (e, slideIndex: number, kind: 'video' | 'audio', fitWidthPx: number) => {
      const session = sessions.get(e.sender.id)
      if (!session || !session.opened.deck.slides[slideIndex]) return null
      const parent = dialogParent()
      const filters =
        kind === 'video'
          ? [{ name: tm('filterVideo'), extensions: [...VIDEO_EXTS] }]
          : [{ name: tm('filterAudio'), extensions: [...AUDIO_EXTS] }]
      const options = {
        title: kind === 'video' ? tm('dlgInsertVideo') : tm('dlgInsertAudio'),
        properties: ['openFile' as const],
        filters,
      }
      const r = await showOpenDialogWithMemory(dialog, parent, options)
      if (r.canceled || !r.filePaths[0]) return null
      const filePath = r.filePaths[0]
      const bytes = new Uint8Array(await readFile(filePath))
      const ext = filePath.split('.').pop()!.toLowerCase()
      const fileName = baseName(filePath)

      // Warn up front, before the file lands on the slide
      const detail = mediaPlaybackWarning(kind, ext, bytes)
      if (detail) {
        const warn = {
          type: 'warning' as const,
          buttons: [tm('legacyPptOk')],
          message: tm('mediaUnsupportedTitle'),
          detail,
        }
        if (parent) await dialog.showMessageBox(parent, warn)
        else await dialog.showMessageBox(warn)
      }

      const result = await embedMedia(session, slideIndex, kind, bytes, ext, fileName, {
        filePath,
        fitWidthPx,
      })
      return result ? { slide: result.slide, sourceId: result.sourceId } : null
    },
  )

  // Double-click playback: read the media bytes of an audio/video element (embedded converts to dataUrl, external links return as-is)
  const AV_MIME: Record<string, string> = {
    mp4: 'video/mp4',
    m4v: 'video/mp4',
    // Chromium refuses to even load video/quicktime, but demuxes QuickTime bytes
    // fine through the ISO-BMFF path when served as video/mp4
    mov: 'video/mp4',
    webm: 'video/webm',
    avi: 'video/x-msvideo',
    mp3: 'audio/mpeg',
    wav: 'audio/wav',
    m4a: 'audio/mp4',
    aac: 'audio/aac',
    ogg: 'audio/ogg',
  }

  ipcMain.handle('slides:media-data', (e, slideIndex: number, sourceId: string) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    if (!session || !slide) return null
    const el = slide.elements.find((x) => x.id === sourceId)
    if (!el || el.type !== 'picture') return null
    const media = (
      el as { media?: { kind: 'video' | 'audio'; target?: string; external?: boolean } }
    ).media
    if (!media?.target) return null
    if (media.external) return { kind: media.kind, dataUrl: media.target }
    const bytes = session.opened.archive.readBytes(media.target)
    if (!bytes) return null
    const ext = media.target.split('.').pop()?.toLowerCase() ?? ''
    const mime = AV_MIME[ext] ?? (media.kind === 'video' ? 'video/mp4' : 'audio/mpeg')
    return {
      kind: media.kind,
      dataUrl: `data:${mime};base64,${Buffer.from(bytes).toString('base64')}`,
    }
  })

  // Media recorded by the renderer (screen-recording webm) or dropped on the canvas: placed like an inserted file
  ipcMain.handle('slides:add-media-bytes', async (e, op: AddMediaBytesOp) => {
    const session = sessions.get(e.sender.id)
    if (!session || !session.opened.deck.slides[op.slideIndex]) return null
    let bytes: Uint8Array
    try {
      bytes =
        'path' in op
          ? new Uint8Array(await readFile(op.path))
          : new Uint8Array(Buffer.from(op.base64, 'base64'))
    } catch {
      return null
    }
    const deckSize = session.opened.deck.size
    const scale = op.fitWidthPx / (deckSize.cx / EMU_PER_PX_96)
    const toEmu = (px: number) => Math.round((px / scale) * EMU_PER_PX_96)
    return embedMedia(session, op.slideIndex, op.kind, bytes, op.ext, op.name, {
      filePath: 'path' in op ? op.path : undefined,
      natural: op.natural,
      center: op.centerPx ? { x: toEmu(op.centerPx.x), y: toEmu(op.centerPx.y) } : undefined,
      fitWidthPx: op.fitWidthPx,
    })
  })

  // 3D model (simplified): glb embed + poster placeholder image
  ipcMain.handle('slides:insert-model3d', async (e, slideIndex: number, fitWidthPx: number) => {
    const session = sessions.get(e.sender.id)
    if (!session || !session.opened.deck.slides[slideIndex]) return null
    const parent = dialogParent()
    const options = {
      title: tm('dlgInsert3d'),
      properties: ['openFile' as const],
      filters: [{ name: tm('filter3d'), extensions: ['glb', 'gltf'] }],
    }
    const r = await showOpenDialogWithMemory(dialog, parent, options)
    if (r.canceled || !r.filePaths[0]) return null
    const filePath = r.filePaths[0]
    const bytes = await readFile(filePath)
    const ext = filePath.split('.').pop()!.toLowerCase()

    let poster: { bytes: Uint8Array; ext: string } | undefined
    try {
      const thumb = await nativeImage.createThumbnailFromPath(filePath, { width: 640, height: 640 })
      if (!thumb.isEmpty()) poster = { bytes: new Uint8Array(thumb.toPNG()), ext: 'png' }
    } catch {
      /* Dark-gray fallback */
    }

    const deckSize = session.opened.deck.size
    const cy = Math.round(deckSize.cy * 0.5)
    const cx = cy
    const txn = sessionTxn(session, {
      ops: [
        {
          op: 'addModel3d',
          target: { slide: slideIndex },
          bytes: new Uint8Array(bytes),
          ext,
          ...(poster ? { poster } : {}),
          offset: {
            x: Math.round((deckSize.cx - cx) / 2),
            y: Math.round((deckSize.cy - cy) / 2),
            cx,
            cy,
          },
          name: baseName(filePath),
        },
      ],
    })
    if (!txn) return null
    session.fitWidthPx = fitWidthPx
    const rebuilt = rebuildSlide(session, slideIndex)
    return rebuilt ? { slide: rebuilt, sourceId: txn.records![0]!.created![0]! } : null
  })

  ipcMain.handle('slides:set-link', (e, op: SetLinkOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    // A javascript:/file: target must never be saved into the package.
    if (
      op.target?.kind === 'url' &&
      safeExternalUrl(op.target.url, { allowedProtocols: DECK_LINK_PROTOCOLS }) === null
    ) {
      return null
    }
    const r = sessionTxn(session, {
      ops: [{ op: 'setLink', target: { slide: op.slideIndex, el: op.sourceId }, link: op.target }],
    })
    return r ? rebuildSlide(session, op.slideIndex) : null
  })

  ipcMain.handle('slides:get-link', (e, slideIndex: number, sourceId: string) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    return getElementLink(session.opened, slideIndex, sourceId)
  })

  ipcMain.handle('slides:get-slide-links', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    if (!session) return []
    return getSlideLinks(session.opened, slideIndex).map(({ elementId, target }) => ({
      sourceId: elementId,
      target,
    }))
  })

  ipcMain.handle('slides:get-run-links', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    if (!session) return []
    return getRunLinks(session.opened, slideIndex).map(({ elementId, ...rest }) => ({
      sourceId: elementId,
      ...rest,
    }))
  })

  ipcMain.handle('slides:apply-header-footer', (e, op: HeaderFooterOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'applyHeaderFooter',
          settings: {
            footer: op.footer ?? null,
            slideNum: !!op.slideNum,
            date: op.date ?? null,
            ...(op.dateAuto ? { dateAuto: true } : {}),
          },
        },
      ],
    })
    if (!r) return null
    session.fitWidthPx = op.fitWidthPx
    return buildAllRenderSlides(session.opened, op.fitWidthPx)
  })

  ipcMain.handle('slides:get-header-footer', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    return slide ? readHeaderFooter(slide) : { footer: null, slideNum: false, date: null }
  })

  // Apply a theme (Design tab theme gallery): rewrite theme*.xml colors/fonts (scheme-referenced
  // colors follow), and remap the deck's explicit srgbClr wholesale to the new theme palette
  // (real-world decks have almost entirely explicit colors, so swapping only the theme changes
  // nothing visually). Element resolved colors come from the parse-time inheritance chain, so
  // after the surgery the deck reparses in memory; undo snapshots roll back as usual.
  ipcMain.handle('slides:apply-theme', (e, op: ApplyThemeOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const payload = {
      op: 'applyTheme',
      name: op.name,
      colors: op.colors,
      ...(op.majorFont ? { majorFont: op.majorFont } : {}),
      ...(op.minorFont ? { minorFont: op.minorFont } : {}),
    }
    // Plan first so an invalid request doesn't clear the redo stack
    const plan = runTxn(session.opened, { ops: [payload], dryRun: true })
    if (plan.failures?.length) return { error: plan.failures[0]!.error }
    pushHistory(session)
    const r = journaledTxn(session, 'edit', { ops: [payload] })
    if (!r.applied) {
      session.undoStack.pop()
      return { error: r.failures?.[0]?.error ?? 'applyTheme failed' }
    }
    const after = r.records![0]!.after as { patched: number; remapped: number }
    if (after.patched === 0 && after.remapped === 0) {
      // Nothing matched the spec; the op only baked pending edits — not an undo step
      session.undoStack.pop()
      return null
    }
    // Reparse so every element's resolved colors/fonts refresh. In-memory
    // (reparseDeck) instead of savePptx -> openPptx: the zip roundtrip's
    // contiguous buffer fails on large decks
    session.opened = reparseDeck(session.opened)
    // Reopening cleared element-level dirty; the session-level flag preserves the "unsaved" state (reset on save)
    markMetaDirty(session)
    session.fitWidthPx = op.fitWidthPx
    return buildAllRenderSlides(session.opened, op.fitWidthPx)
  })

  ipcMain.handle('slides:set-transition', (e, op: SetTransitionOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return false
    const slides = session.opened.deck.slides
    const idxs =
      op.slideIndex === -1 ? slides.map((_, i) => i) : slides[op.slideIndex] ? [op.slideIndex] : []
    if (idxs.length === 0) return false
    const r = sessionTxn(session, {
      ops: idxs.map((i) => ({ op: 'setTransition', target: { slide: i }, kind: op.kind })),
    })
    return r !== null
  })

  ipcMain.handle('slides:get-transition', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    return slide ? getSlideTransition(slide) : 'none'
  })

  // Rehearsal timing save: batch-write each page's auto-advance time (<p:transition advTm>, ms)
  ipcMain.handle('slides:set-advance-times', (e, op: SetAdvanceTimesOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return false
    const slides = session.opened.deck.slides
    const targets = op.times.filter((t) => slides[t.slideIndex])
    if (targets.length === 0) return false
    const r = sessionTxn(session, {
      ops: targets.map((t) => ({
        op: 'setAdvanceTime',
        target: { slide: t.slideIndex },
        ms: t.ms,
      })),
    })
    return r !== null
  })

  // ── Shape animations (<p:timing>; the spid <-> temporary element id mapping happens here) ──
  ipcMain.handle('slides:get-animations', (e, slideIndex: number): AnimationItem[] => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    if (!slide) return []
    const bySpid = new Map<number, (typeof slide.elements)[number]>()
    for (const el of slide.elements) {
      const spid = elementSpid(el)
      if (spid != null && !bySpid.has(spid)) bySpid.set(spid, el)
    }
    const typeLabel: Record<string, string> = {
      text: tm('labelTextBox'),
      shape: tm('labelShape'),
      picture: tm('labelPicture'),
      group: tm('labelGroup'),
      table: tm('labelTable'),
      chart: tm('labelChart'),
      passthrough: tm('labelObject'),
    }
    const out: AnimationItem[] = []
    for (const a of getSlideAnimations(slide)) {
      const el = bySpid.get(a.spid)
      if (!el) continue // Leftover animations whose target shape was deleted are not echoed back
      out.push({
        sourceId: el.id,
        targetName: el.name || typeLabel[el.type] || tm('labelObject'),
        effect: a.effect,
        trigger: a.trigger,
        durationMs: a.durationMs,
        delayMs: a.delayMs,
        // a modelled directional effect carries its own direction; a top wipe must not
        // come back to the player as a bare 'wipe' and play bottom-up
        ...(a.direction != null ? { direction: a.direction } : {}),
        ...(a.motionPath != null ? { motionPath: a.motionPath } : {}),
        ...(a.paragraph != null ? { paragraph: a.paragraph } : {}),
      })
    }
    return out
  })

  // Pairing keys for Morph transitions: sourceId changes on every reparse, so match across pages by cNvPr id/name
  ipcMain.handle('slides:get-shape-keys', (e, slideIndex: number): ShapeKey[] => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    if (!slide) return []
    return slide.elements.map((el) => ({
      sourceId: el.id,
      spid: elementSpid(el),
      name: el.name ?? '',
    }))
  })

  ipcMain.handle('slides:set-animations', (e, op: SetAnimationsOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return false
    const r = sessionTxn(session, {
      ops: [{ op: 'setAnimations', target: { slide: op.slideIndex }, items: op.items }],
    })
    return r !== null
  })

  ipcMain.handle('slides:set-hidden', (e, op: SetSlideHiddenOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, {
      ops: [{ op: 'setHidden', target: { slide: op.slideIndex }, hidden: op.hidden }],
    })
    return r ? rebuildSlide(session, op.slideIndex) : null
  })

  // ── Section management: presentation.xml surgery, riding on snapshot undo and savePptx ──
  ipcMain.handle('slides:get-sections', (e) => {
    const session = sessions.get(e.sender.id)
    return session ? getSections(session.opened) : []
  })

  // Section shims share one shape: single op, metaDirty, echo the op's section payload back.
  const sectionShim = (e: SlidesDocumentEvent, op: Record<string, unknown>) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, { ops: [op as Parameters<typeof runTxn>[1]['ops'][0]] })
    if (!r) return null
    markMetaDirty(session)
    return r.records![0]!.after
  }

  ipcMain.handle('slides:set-sections', (e, sections: SectionInfo[]) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    if (!sectionShim(e, { op: 'setSections', sections })) return null
    return getSections(session.opened)
  })

  ipcMain.handle('slides:add-section', (e, op: AddSectionOp) =>
    sectionShim(e, { op: 'addSection', atSlideIndex: op.atSlideIndex, name: op.name }),
  )

  ipcMain.handle('slides:rename-section', (e, op: RenameSectionOp) =>
    sectionShim(e, { op: 'renameSection', id: op.id, name: op.name }),
  )

  ipcMain.handle('slides:remove-section', (e, op: RemoveSectionOp) =>
    sectionShim(e, { op: 'removeSection', id: op.id }),
  )

  // Section header + its slides in one undo step. Slides go highest index first so the
  // numeric targets stay valid as the deck shrinks; the deck keeps at least one slide.
  ipcMain.handle('slides:remove-section-slides', (e, op: RemoveSectionSlidesOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const total = session.opened.deck.slides.length
    const { lead, sections } = normalizeSections(getSections(session.opened), total)
    const indices = op.id == null ? lead : sections.find((s) => s.id === op.id)?.slideIndices
    if (!indices || indices.length >= total) return null
    const ops: Parameters<typeof runTxn>[1]['ops'] = [...indices]
      .reverse()
      .map((i) => ({ op: 'deleteSlide', target: { slide: i } }))
    if (op.id != null) ops.push({ op: 'removeSection', id: op.id })
    if (!ops.length) return null
    const r = sessionTxn(session, { ops })
    if (!r) return null
    markMetaDirty(session)
    return {
      slides: buildAllRenderSlides(session.opened, session.fitWidthPx),
      sections: getSections(session.opened),
    }
  })

  // Drag to reorder slides (sldIdLst + deck.slides + section membership); must send back the full RenderSlide set
  ipcMain.handle('slides:move-slide', (e, op: MoveSlideOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const r = sessionTxn(session, {
      ops: [{ op: 'moveSlide', target: { slide: op.fromIndex }, to: op.toIndex }],
    })
    if (!r) return null
    markMetaDirty(session)
    return {
      slides: buildAllRenderSlides(session.opened, session.fitWidthPx),
      sections: getSections(session.opened),
    }
  })

  ipcMain.handle('slides:move-slides', (e, op: MoveSlidesOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const steps = planSlideMoves(session.opened.deck.slides.length, op.slideIndexes, op.insertAt)
    if (!steps.length) return null
    const r = sessionTxn(session, {
      ops: steps.map((st) => ({ op: 'moveSlide' as const, target: { slide: st.from }, to: st.to })),
    })
    if (!r) return null
    markMetaDirty(session)
    return {
      slides: buildAllRenderSlides(session.opened, session.fitWidthPx),
      sections: getSections(session.opened),
    }
  })

  // Moving a whole section changes slide order (sldIdLst + deck.slides); must send back the full RenderSlide set
  ipcMain.handle('slides:move-section', (e, op: MoveSectionOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const sections = sectionShim(e, { op: 'moveSection', id: op.id, dir: op.dir })
    if (!sections) return null
    return {
      slides: buildAllRenderSlides(session.opened, session.fitWidthPx),
      sections,
    }
  })

  // ── Speaker notes / comments (archive surgery, riding on snapshot undo and savePptx) ────
  ipcMain.handle('slides:get-notes', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    return session && slide ? getSlideNotes(session.opened.archive, slide.path) : ''
  })

  ipcMain.handle('slides:set-notes', (e, op: SetNotesOp) => {
    const session = sessions.get(e.sender.id)
    if (!session) return false
    const r = sessionTxn(session, {
      ops: [{ op: 'setNotes', target: { slide: op.slideIndex }, text: op.text }],
    })
    if (r) markMetaDirty(session)
    return r !== null
  })

  ipcMain.handle('slides:get-comments', (e, slideIndex: number) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[slideIndex]
    return session && slide ? getSlideComments(session.opened.archive, slide.path) : []
  })

  ipcMain.handle('slides:add-comment', (e, op: AddCommentOp) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[op.slideIndex]
    if (!session || !slide) return null
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'addComment',
          target: { slide: op.slideIndex },
          author: commentAuthorName(),
          text: op.text,
        },
      ],
    })
    if (!r) return null
    markMetaDirty(session)
    return getSlideComments(session.opened.archive, slide.path)
  })

  ipcMain.handle('slides:delete-comment', (e, op: DeleteCommentOp) => {
    const session = sessions.get(e.sender.id)
    const slide = session?.opened.deck.slides[op.slideIndex]
    if (!session || !slide) return null
    const r = sessionTxn(session, {
      ops: [
        {
          op: 'deleteComment',
          target: { slide: op.slideIndex },
          authorId: op.authorId,
          idx: op.idx,
        },
      ],
    })
    if (!r) return null
    markMetaDirty(session)
    return getSlideComments(session.opened.archive, slide.path)
  })

  // System clipboard while text-editing (menu commands are echoed back by the renderer per context)
  ipcMain.handle('slides:native-clipboard', (e, op: 'cut' | 'copy' | 'paste') => {
    if (op === 'cut') e.sender.cut()
    else if (op === 'copy') e.sender.copy()
    else e.sender.paste()
  })

  ipcMain.handle('slides:history-batch-begin', (e) => {
    const session = sessions.get(e.sender.id)
    if (!session) return false
    beginHistoryBatch(session)
    return true
  })

  ipcMain.handle('slides:history-batch-end', (e) => {
    const session = sessions.get(e.sender.id)
    if (!session) return null
    const before = endHistoryBatch(session)
    return before ? registerAiSnapshot(session, before) : null
  })

  ipcMain.handle('slides:ai-snapshot-restore', (e, id: number) => {
    const session = sessions.get(e.sender.id)
    if (!session || session.masterEdit || session.historyBatch) return null
    if (!restoreAiSnapshot(session, id)) return null
    scheduleDeckBroadcast(session)
    return buildAllRenderSlides(session.opened, session.fitWidthPx)
  })

  ipcMain.handle('slides:undo', (e) => {
    const session = sessions.get(e.sender.id)
    // Undo disabled in master view: the masterEdit.slide model cannot roll back with snapshots (v1 trade-off; undoable after exiting)
    if (!session || session.masterEdit) return null
    settleStaleHistoryBatch(session)
    if (session.undoStack.length === 0) return null
    session.redoStack.push(takeSnapshot(session))
    restoreSnapshot(session, session.undoStack.pop()!)
    scheduleHistoryNotify(session)
    scheduleDeckBroadcast(session)
    return buildAllRenderSlides(session.opened, session.fitWidthPx)
  })

  ipcMain.handle('slides:redo', (e) => {
    const session = sessions.get(e.sender.id)
    if (!session || session.masterEdit) return null
    settleStaleHistoryBatch(session)
    if (session.redoStack.length === 0) return null
    session.undoStack.push(takeSnapshot(session))
    restoreSnapshot(session, session.redoStack.pop()!)
    scheduleHistoryNotify(session)
    scheduleDeckBroadcast(session)
    return buildAllRenderSlides(session.opened, session.fitWidthPx)
  })

  ipcMain.handle('slides:is-dirty', (e) => {
    const session = sessions.get(e.sender.id)
    return !!session && sessionDirty(session)
  })
  return {
    release(id: number) {
      lastSlidePaste.delete(id)
    },
  }
}
