/** Node host for the original Slides document model. No Electron or alternate OOXML writer. */
import { createHash, randomUUID } from 'node:crypto'
import { createBlankPptx, openPptx, savePptx, commitSaved } from '@genoffice/pptx-engine'
import type { ApplyTxnOp, ApplyTxnResult } from '../shared/ipc'
import {
  createSlidesDocumentRuntime,
  deckDefaultFont,
  type Session,
  type SlidesDocumentHost,
} from './session-document'
import { createSlidesDocumentOperations } from './document-operations'
import {
  registerSlidesDocumentHandlers,
  type SlidesDocumentEvent,
  type SlidesDocumentPorts,
} from './document-handlers'
import { readDeckModel } from './deck-reader'

export const presentationReadChannels = new Set([
  'slides:get-render-slides',
  'slides:get-layouts',
  'slides:get-slide-size',
  'slides:chart-color-schemes',
  'slides:get-chart-data',
  'slides:get-link',
  'slides:get-slide-links',
  'slides:get-run-links',
  'slides:get-header-footer',
  'slides:get-transition',
  'slides:get-animations',
  'slides:get-shape-keys',
  'slides:get-sections',
  'slides:get-notes',
  'slides:get-comments',
  'slides:is-dirty',
  'slides:clipboard-external',
  'slides:clipboard-probe',
  'slides:has-slide-clipboard',
  'slides:copy-elements',
  'slides:copy-elements-image',
  'slides:copy-slides',
  'slides:media-data',
])

type Mutation = { kind: 'apply'; transaction: ApplyTxnOp } | { kind: 'undo' | 'redo' }
type Result = { revision: number; result: ApplyTxnResult }
type Receipt = { fingerprint: string; result?: Result; error?: unknown }

/** Transport owners authenticate and serialize access to one document instance. */
export class PresentationRuntime {
  private readonly document
  private readonly operations
  private readonly session: Session
  private revision = 0
  private readonly receipts = new Map<string, Receipt>()
  private prepared: { token: string; revision: number; sha256: string } | null = null
  private serialized: { revision: number; bytes: Uint8Array; sha256: string } | null = null
  private handlers = new Map<string, (event: SlidesDocumentEvent, ...args: any[]) => any>()
  private senderId = 1
  private baseline: string

  private constructor(session: Session, host: SlidesDocumentHost) {
    this.session = session
    this.document = createSlidesDocumentRuntime(host)
    this.operations = createSlidesDocumentOperations(this.document)
    this.baseline = this.fingerprint()
  }
  private fingerprint() {
    return createHash('sha256')
      .update(
        JSON.stringify(
          {
            slides: this.session.opened.deck.slides,
            size: this.session.opened.deck.size,
            entries: [...this.session.opened.archive.entries],
          },
          (key, value) =>
            key === 'metaDirty' || /^dirty/.test(key)
              ? undefined
              : value instanceof Uint8Array
                ? createHash('sha256').update(value).digest('hex')
                : value,
        ),
      )
      .digest('hex')
  }
  private dirty() {
    return this.document.sessionDirty(this.session) || this.fingerprint() !== this.baseline
  }

  static async open(
    bytes: Uint8Array | null,
    host: SlidesDocumentHost,
    fitWidthPx = 1280,
    revision = 0,
    restoredDirty = false,
  ) {
    if (!Number.isFinite(fitWidthPx) || fitWidthPx < 1 || fitWidthPx > 8192)
      throw new Error('PPTX_VIEWPORT_INVALID')
    const opened = await openPptx(bytes ?? (await createBlankPptx()))
    const runtime = new PresentationRuntime(
      { path: '', opened, fitWidthPx, undoStack: [], redoStack: [] },
      host,
    )
    runtime.revision = revision
    runtime.session.metaDirty = restoredDirty
    return runtime
  }

  configureHandlers(
    ports: SlidesDocumentPorts,
    shared?: {
      handlers: Map<string, (event: SlidesDocumentEvent, ...args: any[]) => any>
      sessions: Map<number, Session>
      id: number
    },
  ) {
    if (shared) {
      this.handlers = shared.handlers
      this.senderId = shared.id
      shared.sessions.set(shared.id, this.session)
    } else
      registerSlidesDocumentHandlers(
        { handle: (channel, handler) => this.handlers.set(channel, handler) },
        new Map([[1, this.session]]),
        this.document,
        ports,
      )
  }
  readDeck() {
    return { ...readDeckModel(this.session), revision: this.revision, dirty: this.dirty() }
  }
  async invoke(channel: string, args: unknown[], expectedRevision: number) {
    if (expectedRevision !== this.revision) throw new Error('PPTX_DRAFT_CONFLICT')
    const handler = this.handlers.get(channel)
    if (!handler) throw new Error('PPTX_COMMAND_UNSUPPORTED')
    const before = presentationReadChannels.has(channel) ? undefined : this.fingerprint()
    let value: any
    try {
      value = await handler(
        { sender: { id: this.senderId, cut() {}, copy() {}, paste() {} } },
        ...args,
      )
    } catch (error) {
      if (before !== undefined && this.fingerprint() !== before) {
        this.revision++
        this.prepared = null
        this.serialized = null
        throw new Error('PPTX_OPERATION_OUTCOME_UNKNOWN', { cause: error })
      }
      throw error
    }
    if (
      !presentationReadChannels.has(channel) &&
      value !== null &&
      value !== false &&
      !(channel === 'slides:apply-txn' && !value?.applied) &&
      !channel.startsWith('slides:history-batch')
    ) {
      this.revision++
      this.prepared = null
      this.serialized = null
    }
    return { value: channel === 'slides:is-dirty' ? this.dirty() : value, state: this.read() }
  }

  read() {
    const { session } = this
    return {
      revision: this.revision,
      defaultFont: deckDefaultFont(session.opened),
      dirty: this.dirty(),
      canUndo: session.undoStack.length > 0,
      canRedo: session.redoStack.length > 0,
      size: { ...session.opened.deck.size },
      slides: this.document.buildAllRenderSlides(session.opened, session.fitWidthPx),
    }
  }

  /** Duplicate IDs return the original outcome even after later document edits. */
  mutate(requestId: string, expectedRevision: number, mutation: Mutation): Result {
    if (!requestId || requestId.length > 128) throw new Error('PPTX_REQUEST_ID_INVALID')
    const fingerprint = createHash('sha256')
      .update(JSON.stringify({ expectedRevision, mutation }))
      .digest('hex')
    const previous = this.receipts.get(requestId)
    if (previous) {
      if (previous.fingerprint !== fingerprint) throw new Error('PPTX_REQUEST_ID_REUSED')
      if (previous.error) throw previous.error
      return structuredClone(previous.result!)
    }
    // The enclosing Office service persists receipts. Never evict a local outcome
    // and silently re-execute a reused ID; require a fresh document attachment.
    if (this.receipts.size >= 512) throw new Error('PPTX_RECEIPT_CAPACITY')
    const receipt: Receipt = { fingerprint }
    this.receipts.set(requestId, receipt)
    try {
      if (expectedRevision !== this.revision) throw new Error('PPTX_DRAFT_CONFLICT')
      let result: ApplyTxnResult
      if (mutation.kind === 'apply') {
        result = this.operations.applySessionTxn(this.session, mutation.transaction) ?? {
          applied: false,
        }
      } else {
        const { session, document } = this
        document.settleStaleHistoryBatch(session)
        const from = mutation.kind === 'undo' ? session.undoStack : session.redoStack
        const to = mutation.kind === 'undo' ? session.redoStack : session.undoStack
        if (session.masterEdit || from.length === 0) result = { applied: false }
        else {
          to.push(document.takeSnapshot(session))
          document.restoreSnapshot(session, from.pop()!)
          document.scheduleHistoryNotify(session)
          document.scheduleDeckBroadcast(session)
          result = {
            applied: true,
            slides: document.buildAllRenderSlides(session.opened, session.fitWidthPx),
          }
        }
      }
      if (result.applied) {
        this.revision += 1
        this.prepared = null
        this.serialized = null
      }
      receipt.result = { revision: this.revision, result }
      return structuredClone(receipt.result)
    } catch (error) {
      // Retain exceptions too: an unknown post-processing outcome must not be
      // repeated by a reconnect using the same request ID.
      receipt.error = error
      throw error
    }
  }

  /** Serialize a private draft; only the authenticated owner publishes its bytes. */
  async prepareSave(expectedRevision: number) {
    if (expectedRevision !== this.revision) throw new Error('PPTX_DRAFT_CONFLICT')
    const bytes =
      this.serialized?.revision === expectedRevision
        ? this.serialized.bytes
        : await savePptx(this.session.opened)
    if (expectedRevision !== this.revision) throw new Error('PPTX_DRAFT_CONFLICT')
    const sha256 = createHash('sha256').update(bytes).digest('hex')
    this.serialized = { revision: expectedRevision, bytes, sha256 }
    const token = randomUUID()
    this.prepared = { token, revision: expectedRevision, sha256 }
    return { token, revision: expectedRevision, sha256, bytes: new Uint8Array(bytes) }
  }

  /** A failed cloud/workspace write never clears dirty state. */
  acceptSaved(token: string, expectedRevision: number, publishedSha256: string) {
    const prepared = this.prepared
    if (
      !prepared ||
      prepared.token !== token ||
      prepared.revision !== expectedRevision ||
      expectedRevision !== this.revision ||
      prepared.sha256 !== publishedSha256
    )
      throw new Error('PPTX_SAVE_CONFLICT')
    commitSaved(this.session.opened)
    this.session.metaDirty = false
    this.baseline = this.fingerprint()
    this.prepared = null
    return this.read()
  }
}
