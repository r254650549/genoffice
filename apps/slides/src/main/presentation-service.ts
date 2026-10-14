import { createServer, createConnection } from 'node:net'
import { spawn } from 'node:child_process'
import { readFile, writeFile, mkdir, unlink, stat } from 'node:fs/promises'
import { createHash, randomUUID } from 'node:crypto'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { z } from 'zod'
import { PresentationRuntime, presentationReadChannels } from './presentation-runtime'
import { createSlidesDocumentRuntime, type Session } from './session-document'
import { registerSlidesDocumentHandlers, type SlidesDocumentEvent } from './document-handlers'
import { createNodeDocumentPorts } from './node-document-ports'
import {
  createSystemFontMetrics,
  listPrivateFontFaces,
  getPrivateFontData,
  registerEmbeddedFonts,
  resetFontRegistry,
} from './fonts'
import { listEmbeddedFonts } from '@genoffice/pptx-engine'
import { OP_DOCS, opSignatureIndex, opGuide, opGuideCatalog } from '@genoffice/pptx-ops/op-docs'

const schema = z
  .object({
    requestId: z.uuid(),
    owner: z.string().regex(/^[a-f0-9-]{36}:[a-f0-9-]{36}$/),
    method: z.enum([
      'open',
      'rpc',
      'read',
      'describe',
      'guide',
      'apply',
      'prepareSave',
      'acceptSaved',
      'content',
      'close',
      'touch',
      'query',
      'fontFaces',
      'fontData',
    ]),
    payload: z.record(z.string(), z.unknown()).default({}),
    bytes: z
      .string()
      .max(90 * 1024 * 1024)
      .optional(),
  })
  .strict()
type Request = z.infer<typeof schema>
const limit = 96 * 1024 * 1024
const hash = (value: string | Uint8Array) => createHash('sha256').update(value).digest('hex')
type Document = {
  owner: string
  runtime: PresentationRuntime
  sourceVersion: string
  numericId: number
}
type Receipt = {
  digest: string
  owner: string
  pending?: Promise<unknown>
  value?: any
  error?: unknown
}
/** A single native actor in the user's existing environment. All imported document handlers are upstream. */
export class PresentationService {
  private readonly documents = new Map<string, Document>()
  private readonly sessions = new Map<number, Session>()
  private readonly receipts = new Map<string, Receipt>()
  private queue: Promise<unknown> = Promise.resolve()
  private nextId = 1
  private readonly nodePorts = createNodeDocumentPorts()
  private readonly handlers = new Map<string, (event: SlidesDocumentEvent, ...args: any[]) => any>()
  private readonly host = { getFontMetrics: () => createSystemFontMetrics() }
  constructor() {
    registerSlidesDocumentHandlers(
      { handle: (name, handler) => this.handlers.set(name, handler) },
      this.sessions,
      createSlidesDocumentRuntime(this.host),
      this.nodePorts.ports,
    )
  }
  get openSessionCount() {
    return this.documents.size
  }
  dispatch(raw: unknown): Promise<any> {
    const input = schema.parse(raw),
      digest = hash(JSON.stringify(input))
    const read =
      ['read', 'describe', 'guide', 'content', 'touch', 'query', 'fontFaces', 'fontData'].includes(
        input.method,
      ) ||
      (input.method === 'rpc' && presentationReadChannels.has(String(input.payload.channel)))
    if (read) {
      const result = this.queue.then(() => this.execute(input))
      this.queue = result.catch(() => {})
      return result
    }
    const known = this.receipts.get(input.requestId)
    if (known) {
      if (known.digest !== digest) throw new Error('PPTX_REQUEST_CONFLICT')
      if (known.pending) return known.pending
      if (known.error) throw known.error
      return Promise.resolve({ ...known.value, recovered: true })
    }
    const receipt: Receipt = { digest, owner: input.owner }
    this.receipts.set(input.requestId, receipt)
    const result = this.queue.then(() => this.execute(input))
    receipt.pending = result
    this.queue = result.catch(() => {})
    void result.then(
      (value) => {
        receipt.value = compact(value)
        receipt.pending = undefined
      },
      (error) => {
        receipt.error = error
        receipt.pending = undefined
      },
    )
    return result
  }
  private async execute(input: Request): Promise<any> {
    const { owner, method, payload } = input
    if (method === 'query') {
      const known = this.receipts.get(String(payload.requestId))
      if (!known || known.owner.split(':')[0] !== owner.split(':')[0]) return { state: 'unknown' }
      return known.pending
        ? { state: 'pending' }
        : known.error
          ? {
              state: 'failed',
              error: known.error instanceof Error ? known.error.message : 'PPTX_OPERATION_FAILED',
            }
          : { state: 'completed', value: known.value }
    }
    if (method === 'open') {
      const revision = z.number().int().min(0).parse(payload.revision),
        version = String(payload.sourceVersion ?? '')
      for (const [id, document] of this.documents)
        if (document.owner.split(':')[0] === owner.split(':')[0]) {
          if (document.sourceVersion === version && document.runtime.read().revision >= revision) {
            document.owner = owner
            return {
              sessionId: id,
              channels: [...this.handlers.keys()],
              ...document.runtime.read(),
            }
          }
          this.documents.delete(id)
          this.sessions.delete(document.numericId)
        }
      if (this.documents.size >= 16) throw new Error('PPTX_DOCUMENT_LIMIT')
      const bytes = input.bytes === undefined ? null : Buffer.from(input.bytes, 'base64')
      const runtime = await PresentationRuntime.open(
        bytes,
        this.host,
        Number(payload.fitWidthPx ?? 1280),
        revision,
        !!payload.restoredDirty,
      )
      const id = randomUUID(),
        numericId = this.nextId++
      runtime.configureHandlers(this.nodePorts.ports, {
        handlers: this.handlers,
        sessions: this.sessions,
        id: numericId,
      })
      if (registerEmbeddedFonts(listEmbeddedFonts(this.sessions.get(numericId)!.opened.archive)))
        resetFontRegistry()
      this.documents.set(id, { owner, runtime, sourceVersion: version, numericId })
      return { sessionId: id, channels: [...this.handlers.keys()], ...runtime.read() }
    }
    const id = String(payload.sessionId),
      document = this.documents.get(id)
    if (!document || document.owner !== owner) throw new Error('PPTX_SESSION_EXPIRED')
    const { runtime } = document
    if (method === 'touch') return null
    if (method === 'close') {
      this.documents.delete(id)
      this.sessions.delete(document.numericId)
      for (const [key, r] of this.receipts) if (r.owner === owner) this.receipts.delete(key)
      return null
    }
    if (method === 'read')
      return { ...runtime.readDeck(), sourceVersion: document.sourceVersion, state: runtime.read() }
    if (method === 'describe')
      return {
        commands: {
          read_deck: {},
          read_notes: {
            slideIndices: 'Optional array of zero-based slide indices; defaults to all slides',
          },
          apply_slide_ops: {
            ops: 'Native op objects. At most 50. Coordinates are document EMU.',
            isolation: 'atomic (default) or per_op',
            dryRun: 'optional boolean',
          },
          insert_image: {
            url: 'Immutable drive://file/version/sha256 reference or image data URL',
            slideIndex: 'Optional zero-based slide index; defaults to current slide',
          },
          load_guide: { group: 'text | element | insert | table | slide | deck' },
          save_deck: {},
          undo: {},
          redo: {},
        },
        signatures: opSignatureIndex(),
        operations: OP_DOCS,
        guides: opGuideCatalog(),
      }
    if (method === 'guide')
      return { group: payload.group, guide: opGuide(String(payload.group)) ?? null }
    if (method === 'fontFaces') return listPrivateFontFaces()
    if (method === 'fontData') {
      const bytes = getPrivateFontData(String(payload.id))
      return bytes ? { base64: Buffer.from(bytes).toString('base64') } : null
    }
    if (method === 'rpc') {
      const channel = z.string().startsWith('slides:').parse(payload.channel),
        args = z
          .array(z.unknown())
          .max(8)
          .parse(payload.args ?? [])
      this.nodePorts.prepare(payload.asset as Parameters<typeof this.nodePorts.prepare>[0])
      try {
        return await runtime.invoke(
          channel,
          args,
          z.number().int().min(0).parse(payload.expectedRevision),
        )
      } finally {
        this.nodePorts.prepare()
      }
    }
    if (method === 'apply') {
      const value = runtime.mutate(
        input.requestId,
        z.number().int().min(0).parse(payload.expectedRevision),
        {
          kind: 'apply',
          transaction: {
            ops: z.array(z.unknown()).max(50).parse(payload.ops),
            isolation: payload.isolation === 'per_op' ? 'per_op' : 'atomic',
            dryRun: payload.dryRun === true,
          },
        },
      )
      return { ...value, state: runtime.read() }
    }
    if (method === 'prepareSave' || method === 'content') {
      const saved = await runtime.prepareSave(
        method === 'content'
          ? runtime.read().revision
          : z.number().int().min(0).parse(payload.expectedRevision),
      )
      return { ...saved, bytes: Buffer.from(saved.bytes).toString('base64') }
    }
    if (method === 'acceptSaved') {
      const state = runtime.acceptSaved(
        String(payload.token),
        Number(payload.expectedRevision),
        String(payload.sha256),
      )
      document.sourceVersion = String(payload.sourceVersion ?? document.sourceVersion)
      return state
    }
    throw new Error('PPTX_COMMAND_UNSUPPORTED')
  }
}
// Receipts keep operation IDs and outcomes, not repeated deck/media trees. Querying
// a completed effect is followed by a fresh read; it never repeats the mutation.
function compact(value: any): any {
  if (value === null || typeof value !== 'object') return value
  if (Array.isArray(value)) return value.map(compact)
  return Object.fromEntries(
    Object.entries(value)
      .filter(([key]) => !['slides', 'nodes', 'slide', 'bytes', 'base64', 'state'].includes(key))
      .map(([key, item]) => [key, compact(item)]),
  )
}
async function main() {
  const [mode, root, id] = process.argv.slice(2)
  if (!root || !/^\/tmp\/kompagent-office-pptx-[a-f0-9]{24}$/.test(root))
    throw new Error('PPTX_INVALID_DIRECTORY')
  const socketPath = '\0' + root.slice(5)
  await mkdir(root, { recursive: true, mode: 0o700 })
  if (mode === 'serve') {
    const service = new PresentationService()
    let active = 0,
      lastRequest = Date.now()
    const server = createServer((socket) => {
      let line = ''
      socket.setTimeout(195000, () => socket.destroy())
      socket.on('error', () => {})
      socket.on('data', (chunk) => {
        line += chunk.toString()
        if (line.length > 40) {
          socket.destroy()
          return
        }
        if (!line.endsWith('\n')) return
        const requestId = line.trim()
        line = ''
        if (!z.uuid().safeParse(requestId).success) {
          socket.destroy()
          return
        }
        active++
        lastRequest = Date.now()
        void (async () => {
          const input = join(root, requestId + '.request'),
            output = join(root, requestId + '.result')
          let result
          try {
            if ((await stat(input)).size > limit) throw new Error('PPTX_REQUEST_TOO_LARGE')
            const request = schema.parse(JSON.parse(await readFile(input, 'utf8')))
            if (request.requestId !== requestId) throw new Error('PPTX_REQUEST_CONFLICT')
            result = { ok: true, value: await service.dispatch(request) }
          } catch (error) {
            const message = error instanceof Error ? error.message : 'PPTX_OPERATION_FAILED'
            result = {
              ok: false,
              error: /^PPTX_[A-Z_]+$/.test(message) ? message : 'PPTX_OPERATION_FAILED',
              detail: message.slice(0, 1500),
            }
          }
          await writeFile(output, JSON.stringify(result), { mode: 0o600 })
          await unlink(input).catch(() => {})
          socket.end('done\n')
        })()
          .catch(() => socket.destroy())
          .finally(() => {
            active--
            lastRequest = Date.now()
          })
      })
    })
    server.on('error', () => process.exit(1))
    server.listen(socketPath)
    const timer = setInterval(() => {
      if (!active && Date.now() - lastRequest > (service.openSessionCount ? 86400000 : 300000)) {
        clearInterval(timer)
        server.close()
        process.exit(0)
      }
    }, 10000)
    process.on('SIGTERM', () => process.exit(0))
    process.on('SIGINT', () => process.exit(0))
  } else if (mode === 'request') {
    if (!z.uuid().safeParse(id).success) throw new Error('PPTX_INVALID_REQUEST')
    const send = () =>
      new Promise<void>((resolve, reject) => {
        const socket = createConnection(socketPath)
        let connected = false
        socket.setTimeout(190000, () => socket.destroy(new Error('PPTX_REQUEST_UNKNOWN')))
        socket.on('connect', () => {
          connected = true
          socket.write(id + '\n')
        })
        socket.on('data', (chunk) => {
          if (chunk.toString().includes('done')) {
            socket.destroy()
            resolve()
          }
        })
        socket.on('error', (error) => reject(Object.assign(error, { connected })))
        socket.on('end', () => reject(new Error('PPTX_REQUEST_UNKNOWN')))
      })
    try {
      await send()
    } catch (error) {
      if ((error as { connected?: boolean }).connected) throw error
      const code = (error as NodeJS.ErrnoException).code
      if (code !== 'ENOENT' && code !== 'ECONNREFUSED') throw error
      const child = spawn(process.execPath, [fileURLToPath(import.meta.url), 'serve', root], {
        detached: true,
        stdio: 'ignore',
      })
      child.unref()
      let done = false
      for (let n = 0; n < 50; n++) {
        await new Promise((resolve) => setTimeout(resolve, 100))
        try {
          await send()
          done = true
          break
        } catch (next) {
          if ((next as { connected?: boolean }).connected) throw next
        }
      }
      if (!done) throw new Error('PPTX_RUNTIME_UNAVAILABLE')
    }
  } else throw new Error('PPTX_INVALID_MODE')
}
if (process.argv[2] === 'serve' || process.argv[2] === 'request')
  void main().catch(() => {
    process.exitCode = 1
  })
