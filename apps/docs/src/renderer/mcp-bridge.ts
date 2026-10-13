import {getOfficeHost} from './office-host'
import type {McpCommandMessage} from '../shared/ipc'
import {buildDocBytes} from './file-actions'
import {runDocumentCommand, type DocumentCommandContext} from './mcp-document-tools'
export type McpBridgeDeps = DocumentCommandContext

/**
 * A fresh tab boots asynchronously (`newFile()` calls setContent, then setDoc),
 * so a command that lands before `doc` exists would be wiped by that blank
 * reset. Wait for the loaded document before announcing readiness.
 *
 * The wait never stops retrying, only slows down: a tab that gave up while its
 * renderer was still booting would look fine in the UI yet stay unaddressable
 * over MCP for the rest of its life, because readiness is announced exactly
 * once and never re-derived.
 */
const READY_POLL_MS = 50
const READY_SLOW_POLL_MS = 1_000
const READY_FAST_WINDOW_MS = 20_000

async function announceWhenLoaded(
  deps: McpBridgeDeps,
  signal: () => void,
  isCancelled: () => boolean,
): Promise<void> {
  const startedAt = Date.now()
  for (;;) {
    if (isCancelled()) return
    const ctx = deps.getCtx()
    if (ctx?.editor && ctx.doc) {
      signal()
      return
    }
    const slow = Date.now() - startedAt > READY_FAST_WINDOW_MS
    await new Promise((resolve) => setTimeout(resolve, slow ? READY_SLOW_POLL_MS : READY_POLL_MS))
  }
}

/** Subscribe the live editor to MCP commands. Returns the unsubscribe function. */
export function installMcpBridge(deps: McpBridgeDeps): () => void {
  const desktop = getOfficeHost()
  if (!desktop?.onMcpCommand || !desktop.reportMcpResult) return () => {}
  let cancelled = false
  let queue: Promise<void> = Promise.resolve()
  const unsubscribe = desktop.onMcpCommand((message: McpCommandMessage) => {
    if (!message || typeof message.requestId !== 'string') return
    queue = queue.then(async () => {
      try {
        desktop.beforeMcpCommand?.(message)
        const result = await runDocumentCommand(deps, message.command, message.payload)
        await desktop.reportMcpResult({ requestId: message.requestId, ok: true, result })
      } catch (error) {
        await desktop.reportMcpResult({
          requestId: message.requestId,
          ok: false,
          error: error instanceof Error ? error.message : String(error),
        })
      }
    })
  })
  // Let the shell know this tab can accept commands (device for targeted routing).
  void announceWhenLoaded(
    deps,
    () => {
      desktop.bindDocument?.(() => buildDocBytes(deps.getCtx()), (listener) => {
        const editor = deps.getCtx().editor!
        const changed = ({ transaction }: { transaction: { docChanged: boolean } }) => {
          if (transaction.docChanged && !deps.getCtx().saveInFlightRef.current) listener()
        }
        editor.on('transaction', changed)
        return () => { editor.off('transaction', changed) }
      })
      desktop.signalMcpReady?.()
    },
    () => cancelled,
  )
  return () => {
    cancelled = true
    unsubscribe()
  }
}
