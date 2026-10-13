import type { DesktopApi, McpCommandMessage, McpCommandResult } from '../shared/ipc'
import { configureOfficeHost, type OfficeHost } from './office-host'
import { OFFICE_WEB_CHANNEL as channel, type OfficeWebResponse } from './web-protocol'
const pending = new Map<string, {resolve(value: unknown): void; reject(error: Error): void; timer: ReturnType<typeof setTimeout>}>()
const listeners = new Map<string, Set<(value: unknown) => void>>()
let externalState: readonly unknown[] | undefined
let revision = 0
let snapshot: (() => Promise<Uint8Array | null>) | undefined
let stopObserving: (() => void) | undefined
let draftTimer: ReturnType<typeof setTimeout> | undefined
let publishing: Promise<void> = Promise.resolve()
let draftRetry: ReturnType<typeof setTimeout> | undefined
let retryDelay = 1000
let tornDown = false
const origin = window.location.origin
function notify(method: string, args: unknown[]): void {
  window.parent.postMessage({ channel, kind: 'notify', id: crypto.randomUUID(), method, args, revision }, origin)
}
function request(method: string, args: unknown[], capturedRevision = revision): Promise<unknown> {
  const id = crypto.randomUUID()
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {pending.delete(id);reject(new Error('宿主响应超时，请查询保存结果。'))}, 130000)
    pending.set(id, {resolve, reject, timer})
    window.parent.postMessage({channel, kind: 'request', id, method, args, revision: capturedRevision}, origin)
  })
}
function call<K extends keyof DesktopApi>(method: K, args: Parameters<DesktopApi[K]>): ReturnType<DesktopApi[K]> {
  return request(method, args) as ReturnType<DesktopApi[K]>
}
function subscribe<K extends keyof DesktopApi>(method: K, listener: unknown): ReturnType<DesktopApi[K]> {
  const callback = listener as (value: unknown) => void
  const set = listeners.get(method) ?? new Set<(value: unknown) => void>()
  set.add(callback);listeners.set(method, set)
  return (() => {set.delete(callback)}) as ReturnType<DesktopApi[K]>
}
window.addEventListener('message', event => {
  if (event.source !== window.parent || event.origin !== origin) return
  const message = event.data as OfficeWebResponse
  if (message?.channel !== channel) return
  if (message.kind === 'response' && typeof message.id === 'string') {
    const item = pending.get(message.id);if (!item) return
    clearTimeout(item.timer);pending.delete(message.id)
    if (message.error) item.reject(new Error(message.error));else item.resolve(message.value)
  } else if (message.kind === 'event' && typeof message.method === 'string') {
    if (message.method === 'onTeardown') {tornDown=true;if(draftRetry)clearTimeout(draftRetry);if(draftTimer)clearTimeout(draftTimer);stopObserving?.()}
    if (String(message.method) === 'flushDraft') {void flushDraft().then(() => notify('flushed', [message.id])).catch(error => notify('flushFailed', [message.id,String(error)]));return}
    for (const listener of listeners.get(message.method) ?? []) listener(message.value)
  }
})
function flushDraft(): Promise<void> {
  if (session.readOnly || tornDown) return Promise.resolve()
  if (draftTimer) clearTimeout(draftTimer)
  if (draftRetry) clearTimeout(draftRetry)
  publishing = publishing.catch(() => {}).then(async () => {
    // Serialize a stable revision. A change during OOXML generation must not
    // acknowledge a flush without persisting the newer document.
    for (;;) {
      const captured = revision, bytes = await snapshot?.()
      if (captured !== revision) continue
      if (bytes) await request('draft', [{revision: captured, bytes}])
      retryDelay = 1000
      return
    }
  }).catch(error => {
    const message = error instanceof Error ? error.message : String(error)
    if (!tornDown && /OFFICE_UNAVAILABLE|OFFICE_BROWSER_REQUIRED|Failed to fetch|fetch failed|network/i.test(message)) {
      draftRetry = setTimeout(() => {void flushDraft().catch(() => {})}, retryDelay)
      retryDelay = Math.min(retryDelay * 2, 10000)
    }
    throw error
  })
  return publishing
}
const session = await request('session', []) as {revision: number;readOnly:boolean}
revision = session.revision
const host: OfficeHost = {
  kind: 'web',
  readOnly: session.readOnly,
  documentStateChanged: (values, saving) => {const previous=externalState;externalState=values;if(!snapshot||saving||!previous||values.every((value,index)=>Object.is(value,previous[index])))return;revision++;if(draftTimer)clearTimeout(draftTimer);draftTimer=setTimeout(()=>{void flushDraft().catch(error=>notify('draftError',[String(error)]))},750)},
  getLanguage: (...args) => call('getLanguage', args),
  getSystemLocale: (...args) => call('getSystemLocale', args),
  onLanguageChanged: (listener) => subscribe('onLanguageChanged', listener),
  getTheme: (...args) => call('getTheme', args),
  onThemeChanged: (listener) => subscribe('onThemeChanged', listener),
  getDocumentTheme: (...args) => call('getDocumentTheme', args),
  onDocumentThemeChanged: (listener) => subscribe('onDocumentThemeChanged', listener),
  getAutoSaveDefault: (...args) => call('getAutoSaveDefault', args),
  onAutoSaveDefaultChanged: (listener) => subscribe('onAutoSaveDefaultChanged', listener),
  getAiPanelPrefs: (...args) => call('getAiPanelPrefs', args),
  setAiPanelPrefs: (...args) => call('setAiPanelPrefs', args),
  onAiPanelPrefsChanged: (listener) => subscribe('onAiPanelPrefsChanged', listener),
  onChromePressed: (listener) => subscribe('onChromePressed', listener),
  zoteroCommand: (...args) => call('zoteroCommand', args),
  onZoteroRequest: (listener) => subscribe('onZoteroRequest', listener),
  respondToZotero: (...args) => notify('respondToZotero', args),
  openDocx: (...args) => call('openDocx', args),
  openDocxPath: (...args) => call('openDocxPath', args),
  confirmDocumentReplace: (...args) => call('confirmDocumentReplace', args),
  openDocxDecrypt: (...args) => call('openDocxDecrypt', args),
  convertAltChunkHtml: (...args) => call('convertAltChunkHtml', args),
  setDocPassword: (...args) => call('setDocPassword', args),
  docPasswordIntentRevision: (...args) => call('docPasswordIntentRevision', args),
  discardDocPasswordIntents: (...args) => call('discardDocPasswordIntents', args),
  consumePendingOpenDocx: (...args) => call('consumePendingOpenDocx', args),
  consumeNewBlankDoc: (...args) => call('consumeNewBlankDoc', args),
  consumeAiDocContent: (...args) => call('consumeAiDocContent', args),
  consumeHeadlessExport: (...args) => call('consumeHeadlessExport', args),
  headlessExportDone: (...args) => notify('headlessExportDone', args),
  createDocument: (...args) => call('createDocument', args),
  onOpenDocx: (listener) => subscribe('onOpenDocx', listener),
  onRenamedDocx: (listener) => subscribe('onRenamedDocx', listener),
  saveDocx: async (...args) => {const captured=revision;await flushDraft();return request('saveDocx', args, captured) as ReturnType<DesktopApi['saveDocx']>},
  writeRecoveryCopy: (...args) => call('writeRecoveryCopy', args),
  onTeardown: (listener) => subscribe('onTeardown', listener),
  respellKick: (...args) => call('respellKick', args),
  spellDiag: (...args) => notify('spellDiag', args),
  armContextMenu: (...args) => notify('armContextMenu', args),
  claimContextMenu: (...args) => notify('claimContextMenu', args),
  onContextMenuRequest: (listener) => subscribe('onContextMenuRequest', listener),
  spellAddWord: (...args) => call('spellAddWord', args),
  spellIgnoreWord: (...args) => call('spellIgnoreWord', args),
  spellReplace: (...args) => call('spellReplace', args),
  spellLanguages: (...args) => call('spellLanguages', args),
  spellSetLanguages: (...args) => call('spellSetLanguages', args),
  saveDocxAs: async (...args) => {const captured=revision;await flushDraft();return request('saveDocxAs', args, captured) as ReturnType<DesktopApi['saveDocxAs']>},
  saveDocxNew: async (...args) => {const captured=revision;await flushDraft();return request('saveDocxNew', args, captured) as ReturnType<DesktopApi['saveDocxNew']>},
  saveDocxTo: async (...args) => {const captured=revision;await flushDraft();return request('saveDocxTo', args, captured) as ReturnType<DesktopApi['saveDocxTo']>},
  onMcpCommand: (listener) => subscribe('onMcpCommand', listener),
  getRecentFiles: (...args) => call('getRecentFiles', args),
  pickImage: (...args) => call('pickImage', args),
  fontMetrics: (...args) => call('fontMetrics', args),
  getAiSettings: (...args) => call('getAiSettings', args),
  setAiSettings: (...args) => call('setAiSettings', args),
  onAiSettingsChanged: (listener) => subscribe('onAiSettingsChanged', listener),
  openAiModelSettings: (...args) => call('openAiModelSettings', args),
  print: (...args) => call('print', args),
  exportPdf: (...args) => call('exportPdf', args),
  exportHtml: (...args) => call('exportHtml', args),
  printPdfBuffer: (...args) => call('printPdfBuffer', args),
  saveMergedPdf: (...args) => call('saveMergedPdf', args),
  pickExportImagesTarget: (...args) => call('pickExportImagesTarget', args),
  takeExportPdf: (...args) => call('takeExportPdf', args),
  writeExportImage: (...args) => call('writeExportImage', args),
  saveImageAs: (...args) => call('saveImageAs', args),
  onViewImage: (listener) => subscribe('onViewImage', listener),
  aiChat: (...args) => call('aiChat', args),
  aiStream: (...args) => call('aiStream', args),
  aiStreamCancel: (...args) => call('aiStreamCancel', args),
  aiGskStatus: (...args) => call('aiGskStatus', args),
  aiGskLogin: (...args) => call('aiGskLogin', args),
  webSearch: (...args) => call('webSearch', args),
  imageSearch: (...args) => call('imageSearch', args),
  analyzeMedia: (...args) => call('analyzeMedia', args),
  fetchImage: (...args) => call('fetchImage', args),
  aiGenerateImage: (...args) => call('aiGenerateImage', args),
  pickAttachments: (...args) => call('pickAttachments', args),
  addAttachmentPaths: (...args) => call('addAttachmentPaths', args),
  addPastedImage: (...args) => call('addPastedImage', args),
  copyImageToClipboard: (...args) => call('copyImageToClipboard', args),
  readAttachment: (...args) => call('readAttachment', args),
  readAttachmentImage: (...args) => call('readAttachmentImage', args),
  kbList: (...args) => call('kbList', args),
  kbSearch: (...args) => call('kbSearch', args),
  kbRead: (...args) => call('kbRead', args),
  getPathForFile: (file) => file.name,
  openNewTab: (...args) => call('openNewTab', args),
  openSourcePath: (...args) => call('openSourcePath', args),
  listDocsTabs: (...args) => call('listDocsTabs', args),
  focusDocsTab: (...args) => call('focusDocsTab', args),
  onAiStream: (listener) => subscribe('onAiStream', listener),
  onMenuCommand: (listener) => subscribe('onMenuCommand', listener),
  onCloseCheck: (listener) => subscribe('onCloseCheck', listener),
  reportCloseCheck: (...args) => notify('reportCloseCheck', args),
  onCloseSaveRequest: (listener) => subscribe('onCloseSaveRequest', listener),
  reportCloseSaveResult: (...args) => notify('reportCloseSaveResult', args),
  reportViewMenuState: (...args) => notify('reportViewMenuState', args),
  openAgent: (prompt) => notify('openAgent', [prompt]),
  bindDocument: (build, observe) => {
    snapshot = build;stopObserving?.()
    stopObserving = observe(() => {revision++;if (draftTimer) clearTimeout(draftTimer);draftTimer = setTimeout(() => {void flushDraft().catch(error => notify('draftError', [error.message]))}, 750)})
  },
  beforeMcpCommand: (message: McpCommandMessage) => {
    const expected = (message as McpCommandMessage & {expectedRevision?: number}).expectedRevision
    if (expected !== undefined && expected !== revision) throw new Error('OFFICE_REVISION_CONFLICT: read_document again before editing')
    document.body.inert = true
  },
  reportMcpResult: (result: McpCommandResult) => {
    const captured=revision
    void flushDraft().then(() => request('reportMcpResult', [{...result, revision:captured}])).catch(error => notify('draftError', [String(error)])).finally(() => {document.body.inert=false})
  },
  signalMcpReady: () => {notify('ready', [revision]);void flushDraft().catch(error => notify('draftError', [String(error)]))},
}
configureOfficeHost(host)
