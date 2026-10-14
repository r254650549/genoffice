import type { SlidesApi, OpenResult } from '../shared/ipc'
import type { RenderSlide } from '@genoffice/pptx-render'
import { configureOfficeHost, type OfficeHost, type PresentationControls } from './office-host'
import { slidesApiChannels } from './api-channels'

const channel = 'kompagent.genoffice.host/v1',
  origin = window.location.origin
const pending = new Map<
  string,
  { resolve(value: any): void; reject(error: Error): void; timer: ReturnType<typeof setTimeout> }
>()
const listeners = new Map<string, Set<(value: any) => void>>()
let controls: PresentationControls | undefined
let revision = 0,
  published = 0,
  tornDown = false,
  commandInFlight = false
let state:
  | { revision: number; dirty: boolean; slides: RenderSlide[]; canUndo: boolean; canRedo: boolean }
  | undefined
let timer: ReturnType<typeof setTimeout> | undefined
let nativeTail: Promise<unknown> = Promise.resolve(),
  publishing: Promise<void> = Promise.resolve()
let channels = new Set<string>()
let resolveOpen: (value: void) => void, rejectOpen: (error: unknown) => void
const opened = new Promise<void>((resolve, reject) => {
  resolveOpen = resolve
  rejectOpen = reject
})
void opened.catch(() => {})
const seen = new Set<string>()
const previewTickets = new Map<string, number>()
function notify(method: string, args: unknown[]) {
  window.parent.postMessage(
    { channel, kind: 'notify', id: crypto.randomUUID(), method, args, revision },
    origin,
  )
}
function request(method: string, args: unknown[], captured = revision): Promise<any> {
  const id = crypto.randomUUID()
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pending.delete(id)
      reject(new Error('文档操作结果尚未确认，请保留当前页面。'))
    }, 215000)
    pending.set(id, { resolve, reject, timer })
    window.parent.postMessage(
      { channel, kind: 'request', id, method, args, revision: captured },
      origin,
    )
  })
}
function emit(method: string, value: any) {
  for (const callback of listeners.get(method) ?? []) callback(value)
}
function update(next: typeof state) {
  if (!next || next.revision < revision) return
  state = next
  revision = next.revision
  emit('onHistoryChanged', { canUndo: next.canUndo, canRedo: next.canRedo })
}
function bytes(base64: string) {
  return Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))
}
function scheduleDraft() {
  if (timer) clearTimeout(timer)
  timer = setTimeout(() => {
    void flushDraft().catch((error) => notify('draftError', [String(error)]))
  }, 750)
}
function nativeCall(nativeChannel: string, args: unknown[], asset?: unknown): Promise<any> {
  const preview = !!(args[0] as { preview?: boolean } | undefined)?.preview,
    key = nativeChannel + ':' + String((args[0] as { sourceId?: string })?.sourceId ?? ''),
    ticket = preview ? (previewTickets.get(key) ?? 0) + 1 : 0
  if (preview) previewTickets.set(key, ticket)
  const result = nativeTail
    .catch(() => {})
    .then(async () => {
      await opened
      if (preview && previewTickets.get(key) !== ticket) return null
      const response = await request('presentationRpc', [
        { channel: nativeChannel, args, expectedRevision: revision, ...(asset ? { asset } : {}) },
      ])
      update(response.state)
      if (response.recovered && response.state) controls?.apply(response.state)
      if (revision !== published && !preview) scheduleDraft()
      return response.value
    })
  nativeTail = result
  return result
}
function nativeMethod(method: string, payload: Record<string, unknown> = {}): Promise<any> {
  const result = nativeTail
    .catch(() => {})
    .then(async () => {
      await opened
      const value = await request('presentationMethod', [
        method,
        { ...payload, expectedRevision: revision },
      ])
      if (value?.state) update(value.state)
      return value
    })
  nativeTail = result
  return result
}
function flushDraft(): Promise<void> {
  if (timer) clearTimeout(timer)
  publishing = publishing
    .catch(() => {})
    .then(async () => {
      for (let attempt = 0; attempt < 3; attempt++) {
        await nativeTail
        if (tornDown || session.readOnly || revision === published) return
        if (!state?.dirty) {
          const captured = revision
          await request('presentationClean', [{ revision: captured }])
          published = captured
        } else {
          const snapshot = await nativeMethod('prepareSave'),
            captured = snapshot.revision
          // Wait for a stable read-only snapshot; never publish over newer input.
          if (captured !== revision) continue
          await request('draft', [{ revision: captured, bytes: bytes(snapshot.bytes) }], captured)
          published = captured
        }
        if (revision === published) return
      }
      scheduleDraft()
      throw new Error('OFFICE_REVISION_CONFLICT')
    })
  return publishing
}
async function command(message: {
  requestId: string
  command: string
  expectedRevision?: number
  payload?: Record<string, unknown>
}) {
  if (seen.has(message.requestId) || !controls || !state) return
  seen.add(message.requestId)
  commandInFlight = true
  let result: any
  try {
    await controls.flush()
    await flushDraft()
    if (message.expectedRevision !== undefined && message.expectedRevision !== revision)
      throw new Error('OFFICE_REVISION_CONFLICT')
    document.body.inert = true
    const payload = message.payload ?? {}
    if (message.command === 'read_deck') result = await nativeMethod('read')
    else if (message.command === 'describe_tools') result = await nativeMethod('describe', payload)
    else if (message.command === 'load_guide') result = await nativeMethod('guide', payload)
    else if (message.command === 'apply_slide_ops') {
      if (session.readOnly && payload.dryRun !== true) throw new Error('OFFICE_READ_ONLY')
      result = await nativeMethod('apply', payload)
      if (result.result?.failures?.length && !result.result?.applied)
        throw new Error(result.result.failures.map((x: { error: string }) => x.error).join('; '))
      if (result.state) controls.apply(result.state)
    } else if (message.command === 'read_notes') {
      const indices = payload.slideIndices ?? state.slides.map((_, index) => index)
      if (
        !Array.isArray(indices) ||
        indices.some(
          (index) => !Number.isInteger(index) || index < 0 || index >= state!.slides.length,
        )
      )
        throw new Error('PPTX_SLIDE_INDEX_INVALID')
      const notes = []
      for (const slideIndex of indices)
        notes.push({ slideIndex, text: await nativeCall('slides:get-notes', [slideIndex]) })
      result = { notes }
    } else if (message.command === 'insert_image') {
      if (session.readOnly) throw new Error('OFFICE_READ_ONLY')
      const image = /^data:image\/(png|jpeg|gif);base64,([a-zA-Z0-9+/=]+)$/.exec(
        String(payload.url ?? ''),
      )
      if (!image) throw new Error('OFFICE_IMAGE_REFERENCE_REQUIRED')
      const inserted = await nativeCall(
        'slides:insert-image',
        [
          Number(payload.slideIndex ?? controls.selection().slideIndex),
          state.slides[0]?.widthPx ?? 1280,
        ],
        { name: 'image.' + (image[1] === 'jpeg' ? 'jpg' : image[1]), base64: image[2] },
      )
      if (state) controls.apply(state)
      result = { inserted: !!inserted, sourceId: inserted?.sourceId }
    } else if (message.command === 'undo' || message.command === 'redo') {
      if (session.readOnly) throw new Error('OFFICE_READ_ONLY')
      const changed = await nativeCall('slides:' + message.command, [])
      result = { applied: changed !== null && changed !== false }
      if (state) controls.apply(state)
    } else if (message.command === 'save_deck') {
      const ok = await controls.save()
      if (!ok) throw new Error('文件尚未保存。')
      result = { saved: true }
    } else throw new Error('OFFICE_COMMAND_UNSUPPORTED')
    await flushDraft()
    const { state: renderState, ...receipt } = result ?? {}
    await request('reportMcpResult', [
      {
        requestId: message.requestId,
        ok: true,
        result: { ...receipt, selection: controls.selection() },
        revision,
      },
    ])
  } catch (error) {
    if (String(error).includes('UNKNOWN')) {
      const fresh = await nativeMethod('read').catch(() => null)
      if (fresh?.state) controls.apply(fresh.state)
    }
    await request('reportMcpResult', [
      { requestId: message.requestId, ok: false, error: String(error), revision },
    ]).catch(() => {})
  } finally {
    commandInFlight = false
    document.body.inert = false
  }
}
window.addEventListener('message', (event) => {
  if (event.source !== window.parent || event.origin !== origin || event.data?.channel !== channel)
    return
  const m = event.data
  if (m.kind === 'response') {
    const item = pending.get(m.id)
    if (!item) return
    clearTimeout(item.timer)
    pending.delete(m.id)
    if (m.error) item.reject(new Error(m.error))
    else item.resolve(m.value)
    return
  }
  if (m.kind !== 'event') return
  if (m.method === 'pauseEditing') {
    document.body.inert = true
    return
  }
  if (m.method === 'resumeEditing') {
    document.body.inert = false
    return
  }
  if (m.method === 'onTeardown') {
    tornDown = true
    if (timer) clearTimeout(timer)
    return
  }
  if (m.method === 'onMcpCommand') {
    void command(m.value)
    return
  }
  if (m.method === 'flushDraft') {
    void (async () => {
      await controls?.flush()
      await flushDraft()
    })()
      .then(() => notify('flushed', [m.id]))
      .catch((error) => notify('flushFailed', [m.id, String(error)]))
    return
  }
  if (m.method === 'onCloseCheck') {
    void (async () => {
      await controls?.flush()
      await nativeTail
      return !!state?.dirty
    })()
      .then((dirty) => notify('reportCloseCheck', [{ dirty }]))
      .catch(() => notify('reportCloseCheck', [{ dirty: true }]))
    return
  }
  for (const callback of listeners.get(m.method) ?? []) callback(m.value)
})
const session = (await request('session', [])) as {
  revision: number
  readOnly: boolean
  name: string
  hasDraft: boolean
}
revision = published = session.revision
async function invoke(method: keyof SlidesApi, args: any[]): Promise<any> {
  if (method === 'consumePendingOpen') {
    const value = await request('openPresentation', [args[0]]).catch((error) => {
      rejectOpen(error)
      throw error
    })
    channels = new Set(value.channels)
    update(value)
    resolveOpen()
    if (revision !== published) scheduleDraft()
    notify('ready', [])
    return { ...value, path: value.path } satisfies OpenResult
  }
  if (method === 'save' || method === 'saveAs') {
    await flushDraft()
    const captured = revision,
      response = await request(
        method === 'save' ? 'savePresentation' : 'copyPresentation',
        [args[0]],
        captured,
      )
    if (response.ok && !response.copied) {
      published = response.state?.revision ?? captured
      update(response.state)
      if (revision !== published) scheduleDraft()
    }
    return response
  }
  if (method === 'privateFontFaces') return nativeMethod('fontFaces')
  if (method === 'privateFontData') {
    const value = await nativeMethod('fontData', { id: args[0] })
    return value ? bytes(value.base64).buffer : null
  }
  if (method === 'nativeClipboard') return false
  if (method === 'newBlank') {
    notify('newPresentation', [])
    return null
  }
  if (method === 'openPptx' || method === 'openPptxPath') return null
  if (method === 'pickPictureFile') return request('pickOfficeAsset', ['image/*'])
  if (method === 'setShowFullScreen') {
    if (args[0]) await document.documentElement.requestFullscreen()
    else if (document.fullscreenElement) await document.exitFullscreen()
    return
  }
  if (method === 'getAiSettings') return { provider: 'anthropic', providers: {} }
  if (method === 'fontCatalog' || method === 'fontMissing' || method === 'listStyleTemplates')
    return []
  if (method === 'cloudGenStatus') return { enabled: false }
  if (method === 'consumeHeadlessExport') return null
  if (
    method === 'setAutoSavePref' ||
    method === 'headlessExportDone' ||
    method === 'aiLogRunFailure'
  )
    return
  if (method === 'aiStream' || method === 'cloudGeneratePage' || method === 'localGeneratePage') {
    notify('openAgent', [typeof args[0] === 'string' ? args[0] : JSON.stringify(args[0])])
    return
  }
  if (method === 'aiGskStatus' || method === 'gskStatus') return { loggedIn: false }
  const nativeChannel = slidesApiChannels[method as keyof typeof slidesApiChannels]
  if (nativeChannel?.startsWith('slides:')) await opened
  if (nativeChannel && channels.has(nativeChannel)) {
    let asset
    if (
      ['insertImage', 'insertMedia', 'insertModel3d'].includes(method) ||
      (method === 'editBackground' && (args[0] as { kind?: string })?.kind === 'image') ||
      (method === 'editImageFill' && !(args[0] as { source?: unknown })?.source)
    ) {
      asset = await request('pickOfficeAsset', [
        method === 'insertImage' || method === 'editBackground' || method === 'editImageFill'
          ? 'image/*'
          : method === 'insertMedia'
            ? 'audio/*,video/*'
            : '.glb,.gltf',
      ])
      if (!asset) return null
    }
    return nativeCall(nativeChannel, args, asset)
  }
  return request(method, args)
}
function call<K extends keyof SlidesApi>(
  method: K,
  args: Parameters<SlidesApi[K]>,
): ReturnType<SlidesApi[K]> {
  if (String(method).startsWith('on')) {
    const callback = args[0] as (value: any) => void,
      set = listeners.get(String(method)) ?? new Set()
    set.add(callback)
    listeners.set(String(method), set)
    return (() => set.delete(callback)) as ReturnType<SlidesApi[K]>
  }
  if (method === 'reportCloseSaveResult') {
    notify('reportCloseSaveResult', args)
    return undefined as ReturnType<SlidesApi[K]>
  }
  return invoke(method, args) as ReturnType<SlidesApi[K]>
}
const host: OfficeHost = {
  kind: 'web',
  readOnly: session.readOnly,
  bindPresentation(value) {
    controls = value
  },
  openAgent(prompt) {
    notify('openAgent', [prompt])
  },
  draftDirty() {
    return !!state?.dirty
  },
  getLanguage: (...args) => call('getLanguage', args),
  onLanguageChanged: (...args) => call('onLanguageChanged', args),
  getTheme: (...args) => call('getTheme', args),
  onThemeChanged: (...args) => call('onThemeChanged', args),
  getAutoSaveDefault: (...args) => call('getAutoSaveDefault', args),
  onAutoSaveDefaultChanged: (...args) => call('onAutoSaveDefaultChanged', args),
  getAiPanelPrefs: (...args) => call('getAiPanelPrefs', args),
  setAiPanelPrefs: (...args) => call('setAiPanelPrefs', args),
  onAiPanelPrefsChanged: (...args) => call('onAiPanelPrefsChanged', args),
  onChromePressed: (...args) => call('onChromePressed', args),
  setShowFullScreen: (...args) => call('setShowFullScreen', args),
  privateFontFaces: (...args) => call('privateFontFaces', args),
  privateFontData: (...args) => call('privateFontData', args),
  fontCatalog: (...args) => call('fontCatalog', args),
  fontDownload: (...args) => call('fontDownload', args),
  fontInstallLocal: (...args) => call('fontInstallLocal', args),
  fontMissing: (...args) => call('fontMissing', args),
  onFontsChanged: (...args) => call('onFontsChanged', args),
  openPptx: (...args) => call('openPptx', args),
  openPptxPath: (...args) => call('openPptxPath', args),
  consumePendingOpen: (...args) => call('consumePendingOpen', args),
  consumeHeadlessExport: (...args) => call('consumeHeadlessExport', args),
  headlessExportDone: (...args) => call('headlessExportDone', args),
  newBlank: (...args) => call('newBlank', args),
  landGeneratedPages: (...args) => call('landGeneratedPages', args),
  cloudGenStatus: (...args) => call('cloudGenStatus', args),
  cloudPageCancel: (...args) => call('cloudPageCancel', args),
  cloudGeneratePage: (...args) => call('cloudGeneratePage', args),
  localGeneratePage: (...args) => call('localGeneratePage', args),
  editText: (...args) => call('editText', args),
  setElementFont: (...args) => call('setElementFont', args),
  setElementParagraphFormat: (...args) => call('setElementParagraphFormat', args),
  findReplace: (...args) => call('findReplace', args),
  setSlideLayout: (...args) => call('setSlideLayout', args),
  setSlideSize: (...args) => call('setSlideSize', args),
  getSlideSize: (...args) => call('getSlideSize', args),
  editTransform: (...args) => call('editTransform', args),
  cancelTransformPreview: (...args) => call('cancelTransformPreview', args),
  editTransformMulti: (...args) => call('editTransformMulti', args),
  editConnectorEndpoints: (...args) => call('editConnectorEndpoints', args),
  editPictureSrcRect: (...args) => call('editPictureSrcRect', args),
  editPictureOpacity: (...args) => call('editPictureOpacity', args),
  editImageFill: (...args) => call('editImageFill', args),
  changeShape: (...args) => call('changeShape', args),
  setShapeAdjust: (...args) => call('setShapeAdjust', args),
  setShapeGeometry: (...args) => call('setShapeGeometry', args),
  setTextAnchor: (...args) => call('setTextAnchor', args),
  setTextBodyProps: (...args) => call('setTextBodyProps', args),
  setEffects: (...args) => call('setEffects', args),
  clipboardExternal: (...args) => call('clipboardExternal', args),
  groupElements: (...args) => call('groupElements', args),
  ungroupElement: (...args) => call('ungroupElement', args),
  batchEditTransform: (...args) => call('batchEditTransform', args),
  getRenderSlides: (...args) => call('getRenderSlides', args),
  addElement: (...args) => call('addElement', args),
  deleteElement: (...args) => call('deleteElement', args),
  deleteElements: (...args) => call('deleteElements', args),
  addSlide: (...args) => call('addSlide', args),
  addBlankSlide: (...args) => call('addBlankSlide', args),
  addSlideWithLayout: (...args) => call('addSlideWithLayout', args),
  getLayouts: (...args) => call('getLayouts', args),
  masterEnter: (...args) => call('masterEnter', args),
  masterOpen: (...args) => call('masterOpen', args),
  masterClose: (...args) => call('masterClose', args),
  masterEditText: (...args) => call('masterEditText', args),
  masterEditTransform: (...args) => call('masterEditTransform', args),
  masterEditFill: (...args) => call('masterEditFill', args),
  masterEditStroke: (...args) => call('masterEditStroke', args),
  masterDeleteElement: (...args) => call('masterDeleteElement', args),
  editFill: (...args) => call('editFill', args),
  editStroke: (...args) => call('editStroke', args),
  flipElements: (...args) => call('flipElements', args),
  editBackground: (...args) => call('editBackground', args),
  pickPictureFile: (...args) => call('pickPictureFile', args),
  insertImage: (...args) => call('insertImage', args),
  copySlides: (...args) => call('copySlides', args),
  pasteSlide: (...args) => call('pasteSlide', args),
  repasteSlide: (...args) => call('repasteSlide', args),
  hasSlideClipboard: (...args) => call('hasSlideClipboard', args),
  clipboardProbe: (...args) => call('clipboardProbe', args),
  deleteSlide: (...args) => call('deleteSlide', args),
  deleteSlides: (...args) => call('deleteSlides', args),
  duplicateSlides: (...args) => call('duplicateSlides', args),
  reorderElement: (...args) => call('reorderElement', args),
  editTableCell: (...args) => call('editTableCell', args),
  tableStructure: (...args) => call('tableStructure', args),
  tableMerge: (...args) => call('tableMerge', args),
  setTableColWidth: (...args) => call('setTableColWidth', args),
  setTableRowHeight: (...args) => call('setTableRowHeight', args),
  setTableCellAnchor: (...args) => call('setTableCellAnchor', args),
  editTableStyle: (...args) => call('editTableStyle', args),
  editChart: (...args) => call('editChart', args),
  getChartColorSchemes: (...args) => call('getChartColorSchemes', args),
  getChartData: (...args) => call('getChartData', args),
  copyElements: (...args) => call('copyElements', args),
  copyElementsImage: (...args) => call('copyElementsImage', args),
  pasteElements: (...args) => call('pasteElements', args),
  duplicateElements: (...args) => call('duplicateElements', args),
  addTable: (...args) => call('addTable', args),
  addInk: (...args) => call('addInk', args),
  addChart: (...args) => call('addChart', args),
  addSmartArt: (...args) => call('addSmartArt', args),
  addImageBytes: (...args) => call('addImageBytes', args),
  replacePictureBytes: (...args) => call('replacePictureBytes', args),
  insertMedia: (...args) => call('insertMedia', args),
  addMediaBytes: (...args) => call('addMediaBytes', args),
  getMediaData: (...args) => call('getMediaData', args),
  insertModel3d: (...args) => call('insertModel3d', args),
  setLink: (...args) => call('setLink', args),
  getLink: (...args) => call('getLink', args),
  getSlideLinks: (...args) => call('getSlideLinks', args),
  getRunLinks: (...args) => call('getRunLinks', args),
  applyHeaderFooter: (...args) => call('applyHeaderFooter', args),
  getHeaderFooter: (...args) => call('getHeaderFooter', args),
  applyTheme: (...args) => call('applyTheme', args),
  setTransition: (...args) => call('setTransition', args),
  getTransition: (...args) => call('getTransition', args),
  setAdvanceTimes: (...args) => call('setAdvanceTimes', args),
  getAnimations: (...args) => call('getAnimations', args),
  getShapeKeys: (...args) => call('getShapeKeys', args),
  setAnimations: (...args) => call('setAnimations', args),
  setSlideHidden: (...args) => call('setSlideHidden', args),
  setSlidesHidden: (...args) => call('setSlidesHidden', args),
  getSections: (...args) => call('getSections', args),
  setSections: (...args) => call('setSections', args),
  addSection: (...args) => call('addSection', args),
  renameSection: (...args) => call('renameSection', args),
  removeSection: (...args) => call('removeSection', args),
  removeSectionSlides: (...args) => call('removeSectionSlides', args),
  moveSection: (...args) => call('moveSection', args),
  moveSlide: (...args) => call('moveSlide', args),
  moveSlides: (...args) => call('moveSlides', args),
  getNotes: (...args) => call('getNotes', args),
  setNotes: (...args) => call('setNotes', args),
  getComments: (...args) => call('getComments', args),
  addComment: (...args) => call('addComment', args),
  deleteComment: (...args) => call('deleteComment', args),
  nativeClipboard: (...args) => call('nativeClipboard', args),
  beginHistoryBatch: (...args) => call('beginHistoryBatch', args),
  endHistoryBatch: (...args) => call('endHistoryBatch', args),
  applyEditScript: (...args) => call('applyEditScript', args),
  applyTxn: (...args) => call('applyTxn', args),
  aiSnapshotRestore: (...args) => call('aiSnapshotRestore', args),
  undo: (...args) => call('undo', args),
  redo: (...args) => call('redo', args),
  pickExportDir: (...args) => call('pickExportDir', args),
  exportImages: (...args) => call('exportImages', args),
  pickExportPdfPath: (...args) => call('pickExportPdfPath', args),
  savePicture: (...args) => call('savePicture', args),
  exportPdf: (...args) => call('exportPdf', args),
  printSlides: (...args) => call('printSlides', args),
  save: (...args) => call('save', args),
  saveAs: (...args) => call('saveAs', args),
  onCloseSaveRequest: (...args) => call('onCloseSaveRequest', args),
  onHistoryChanged: (...args) => call('onHistoryChanged', args),
  onDeckChanged: (...args) => call('onDeckChanged', args),
  reportCloseSaveResult: (...args) => call('reportCloseSaveResult', args),
  setAutoSavePref: (...args) => call('setAutoSavePref', args),
  isDirty: (...args) => call('isDirty', args),
  getRecentFiles: (...args) => call('getRecentFiles', args),
  onMenuCommand: (...args) => call('onMenuCommand', args),
  onOpened: (...args) => call('onOpened', args),
  onRenamed: (...args) => call('onRenamed', args),
  getAiSettings: (...args) => call('getAiSettings', args),
  setAiSettings: (...args) => call('setAiSettings', args),
  onAiSettingsChanged: (...args) => call('onAiSettingsChanged', args),
  openAiModelSettings: (...args) => call('openAiModelSettings', args),
  aiStream: (...args) => call('aiStream', args),
  aiStreamCancel: (...args) => call('aiStreamCancel', args),
  aiGskStatus: (...args) => call('aiGskStatus', args),
  aiGskLogin: (...args) => call('aiGskLogin', args),
  aiLogRunFailure: (...args) => call('aiLogRunFailure', args),
  webSearch: (...args) => call('webSearch', args),
  imageSearch: (...args) => call('imageSearch', args),
  insertImageUrl: (...args) => call('insertImageUrl', args),
  replacePictureUrl: (...args) => call('replacePictureUrl', args),
  generateImage: (...args) => call('generateImage', args),
  analyzeMedia: (...args) => call('analyzeMedia', args),
  gskStatus: (...args) => call('gskStatus', args),
  onAiStream: (...args) => call('onAiStream', args),
  saveStyleSidecar: (...args) => call('saveStyleSidecar', args),
  saveStyleTemplate: (...args) => call('saveStyleTemplate', args),
  listStyleTemplates: (...args) => call('listStyleTemplates', args),
  loadStyleTemplate: (...args) => call('loadStyleTemplate', args),
  presenterStart: (...args) => call('presenterStart', args),
  presenterSync: (...args) => call('presenterSync', args),
  presenterInk: (...args) => call('presenterInk', args),
  presenterSwap: (...args) => call('presenterSwap', args),
  presenterEnd: (...args) => call('presenterEnd', args),
  audienceReady: (...args) => call('audienceReady', args),
  audienceNav: (...args) => call('audienceNav', args),
  onShowSync: (...args) => call('onShowSync', args),
  onShowInk: (...args) => call('onShowInk', args),
  onAudienceNav: (...args) => call('onAudienceNav', args),
}
configureOfficeHost(host)
window.addEventListener('beforeunload', (event) => {
  if (
    state?.dirty ||
    commandInFlight ||
    (document.activeElement instanceof HTMLElement && document.activeElement.isContentEditable)
  ) {
    event.preventDefault()
    event.returnValue = ''
  }
})
