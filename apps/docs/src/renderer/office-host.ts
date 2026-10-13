import type { DesktopApi, McpCommandMessage } from '../shared/ipc'
/** Explicit host injection shared by the complete native renderer and its Web entry. */
export interface OfficeHost extends DesktopApi {
  kind?: 'electron' | 'web'
  readOnly?: boolean
  documentStateChanged?(values: readonly unknown[], saving: boolean): void
  openAgent?(prompt?: string): void
  bindDocument?(snapshot: () => Promise<Uint8Array | null>, subscribe: (listener: () => void) => () => void): void
  beforeMcpCommand?(message: McpCommandMessage): void
}
let host: OfficeHost
export function configureOfficeHost(value: OfficeHost): void { host = value }
export function getOfficeHost(): OfficeHost { return host }
