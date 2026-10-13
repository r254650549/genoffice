/** Electron host for the shared original Slides document runtime. */
import { BrowserWindow, webContents } from 'electron'
import type { WebContents } from 'electron'
import { join } from 'node:path'
import type { FontMetricsProvider } from '@genoffice/pptx-render'
import { createSystemFontMetrics, resetFontRegistry } from './fonts'
import { createSlidesDocumentRuntime, type Session } from './session-document'
export type { Session, OpLogEntry, HistorySnapshot } from './session-document'

export interface RuntimePaths {
  preloadPath: string
  rendererDevUrl?: string | undefined
  rendererFilePath?: string | undefined
  /** Shell router used to open exported PDFs in a new GenOffice tab. */
  openGeneratedPath?: (path: string) => boolean
}

export const runtime: RuntimePaths = {
  preloadPath: join(__dirname, '../preload/index.js'),
  rendererDevUrl: process.env.ELECTRON_RENDERER_URL,
  rendererFilePath: join(__dirname, '../renderer/index.html'),
}

export function configureSlidesRuntime(paths: RuntimePaths): void {
  runtime.preloadPath = paths.preloadPath
  runtime.rendererDevUrl = paths.rendererDevUrl
  runtime.rendererFilePath = paths.rendererFilePath
  runtime.openGeneratedPath = paths.openGeneratedPath
}

export const sessions = new Map<number, Session>()

/** webContents ids currently mapped to this session (aliased entries included). */
export function attachedIds(session: Session): number[] {
  const ids: number[] = []
  for (const [id, s] of sessions) if (s === session) ids.push(id)
  return ids
}

/**
 * View-only attachments (presenter audience windows). They receive broadcasts
 * but cannot save, so they must not count as "someone still holds this
 * document" in close guards.
 */
export const viewerWcIds = new Set<number>()

/** Attached windows that can actually edit/save the session. */
export function editorAttachedIds(session: Session): number[] {
  return attachedIds(session).filter((id) => !viewerWcIds.has(id))
}

// ── Window references (shell tab mode + active renderer tracking) ──────
export const windowRefs = {
  /** Parent window for dialogs in tab mode (the shell's single BrowserWindow) */
  shellWindow: null as BrowserWindow | null,
  /** Currently active slides renderer (window or tab view) — target of menu commands; the shell updates it on tab switch */
  activeWebContents: null as WebContents | null,
}

export function setSlidesShellWindow(win: BrowserWindow | null): void {
  windowRefs.shellWindow = win
}

/** Shell-registered hook (aggregate/tab mode only): cover the tab strip with a tab's
 *  view during a slideshow without going through HTML fullscreen. Standalone slides
 *  windows have no tab strip and leave this null. */
export const showChrome = {
  setBleed: null as ((wc: WebContents, on: boolean) => void) | null,
  /** the window hosting a tab's WebContentsView when BrowserWindow.fromWebContents
   *  cannot tell (a detached "Open in New Window" / torn-off editor) */
  hostWindow: null as ((wc: WebContents) => BrowserWindow | undefined) | null,
}

export function setSlidesShowBleed(cb: (wc: WebContents, on: boolean) => void): void {
  showChrome.setBleed = cb
}

export function setSlidesHostWindowHook(
  fn: ((wc: WebContents) => BrowserWindow | undefined) | null,
): void {
  showChrome.hostWindow = fn
}

/** the window a tab view lives in: its own BrowserWindow, the shell-registered
 *  detached window, else the shell window */
export function hostWindowFor(wc: WebContents): BrowserWindow | undefined {
  const own = BrowserWindow.fromWebContents(wc) ?? showChrome.hostWindow?.(wc)
  if (own && !own.isDestroyed()) return own
  return windowRefs.shellWindow && !windowRefs.shellWindow.isDestroyed()
    ? windowRefs.shellWindow
    : undefined
}

export function setActiveSlidesWebContents(wc: WebContents | null): void {
  windowRefs.activeWebContents = wc
}

export function dialogParent(): BrowserWindow | undefined {
  // Focused window first: a detached editor window must parent its own dialogs
  // (the shell's tab views live inside the shell window, so tab mode is unchanged)
  return BrowserWindow.getFocusedWindow() ?? windowRefs.shellWindow ?? undefined
}

// ── RenderSlide rebuild helpers ─────────────────────────────────────────

/** Precise system-font metrics (lazily built, shared process-wide; unmatched fonts fall back to heuristics per run). */
let fontMetrics: FontMetricsProvider | null = null
export function getFontMetrics(): FontMetricsProvider {
  if (!fontMetrics) fontMetrics = createSystemFontMetrics()
  return fontMetrics
}

/** Drop the cached metrics (and its font registry) after the user font store changes. */
export function resetFontMetrics(): void {
  resetFontRegistry()
  fontMetrics = null
}

const documentRuntime = createSlidesDocumentRuntime({
  getFontMetrics,
  historyChanged(session) {
    for (const id of attachedIds(session))
      webContents.fromId(id)?.send('slides:history-changed', {
        canUndo: session.undoStack.length > 0,
        canRedo: session.redoStack.length > 0,
      })
  },
  deckChanged(session) {
    // The shared native runtime coalesces this callback after handler processing.
    const ids = attachedIds(session)
    if (ids.length < 2) return
    const payload = {
      slides: buildAllRenderSlides(session.opened, session.fitWidthPx),
      size: { cx: session.opened.deck.size.cx, cy: session.opened.deck.size.cy },
    }
    for (const id of ids) webContents.fromId(id)?.send('slides:deck-changed', payload)
  },
})
export const {
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
} = documentRuntime
export { documentRuntime }
