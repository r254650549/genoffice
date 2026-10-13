/** Node host for the original Slides document model. No Electron or alternate OOXML writer. */
import { createHash, randomUUID } from 'node:crypto'
import { createBlankPptx, openPptx, savePptx, commitSaved } from '@genoffice/pptx-engine'
import type { ApplyTxnOp, ApplyTxnResult } from '../shared/ipc'
import {
  createSlidesDocumentRuntime,
  type Session,
  type SlidesDocumentHost,
} from './session-document'
import { createSlidesDocumentOperations } from './document-operations'

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

  private constructor(session: Session, host: SlidesDocumentHost) {
    this.session = session
    this.document = createSlidesDocumentRuntime(host)
    this.operations = createSlidesDocumentOperations(this.document)
  }

  static async open(bytes: Uint8Array | null, host: SlidesDocumentHost, fitWidthPx = 1280) {
    if (!Number.isFinite(fitWidthPx) || fitWidthPx < 1 || fitWidthPx > 8192)
      throw new Error('PPTX_VIEWPORT_INVALID')
    const opened = await openPptx(bytes ?? (await createBlankPptx()))
    return new PresentationRuntime(
      { path: '', opened, fitWidthPx, undoStack: [], redoStack: [] },
      host,
    )
  }

  read() {
    const { session } = this
    return {
      revision: this.revision,
      dirty: this.document.sessionDirty(session),
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
    const bytes = await savePptx(this.session.opened)
    if (expectedRevision !== this.revision) throw new Error('PPTX_DRAFT_CONFLICT')
    const sha256 = createHash('sha256').update(bytes).digest('hex')
    const token = randomUUID()
    this.prepared = { token, revision: expectedRevision, sha256 }
    return { token, revision: expectedRevision, sha256, bytes }
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
    this.prepared = null
    return this.read()
  }
}
