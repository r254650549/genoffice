import type {DesktopApi,McpCommandMessage} from '../shared/desktop-api'
import {configureOfficeHost,type OfficeHost} from './office-host'
const channel='kompagent.genoffice.host/v1',origin=window.location.origin
const pending=new Map<string,{resolve(value:unknown):void;reject(error:Error):void;timer:ReturnType<typeof setTimeout>}>()
const listeners=new Map<string,Set<(value:unknown)=>void>>()
let revision=0,published=0,loading=true,tornDown=false,commands=false
let controls:Parameters<NonNullable<OfficeHost['bindWorkbook']>>[0]|undefined
let timer:ReturnType<typeof setTimeout>|undefined
let publishing:Promise<void>=Promise.resolve()
function notify(method:string,args:unknown[]){window.parent.postMessage({channel,kind:'notify',id:crypto.randomUUID(),method,args,revision},origin)}
function request(method:string,args:unknown[],captured=revision):Promise<unknown>{const id=crypto.randomUUID();return new Promise((resolve,reject)=>{const timer=setTimeout(()=>{pending.delete(id);reject(new Error('宿主响应超时，编辑内容已保留。'))},215000);pending.set(id,{resolve,reject,timer});window.parent.postMessage({channel,kind:'request',id,method,args,revision:captured},origin)})}
function call<K extends keyof DesktopApi>(method:K,args:Parameters<DesktopApi[K]>):ReturnType<DesktopApi[K]>{return request(method,args) as ReturnType<DesktopApi[K]>}
function subscribe<K extends keyof DesktopApi>(method:K,listener:unknown):ReturnType<DesktopApi[K]>{const callback=listener as (value:unknown)=>void,set=listeners.get(method)??new Set();set.add(callback);listeners.set(method,set);return (()=>{set.delete(callback)}) as ReturnType<DesktopApi[K]>}
function flush():Promise<void>{
 if(timer)clearTimeout(timer)
 publishing=publishing.catch(()=>{}).then(async()=>{if(tornDown||session.readOnly||revision===published)return;if(!controls)throw new Error('编辑器尚未就绪');if(!await controls.flush())throw new Error('草稿尚未保存，请保持文档打开。')})
 return publishing
}
window.addEventListener('message',event=>{
 if(event.source!==window.parent||event.origin!==origin||event.data?.channel!==channel)return
 const m=event.data
 if(m.kind==='response'){const item=pending.get(m.id);if(!item)return;clearTimeout(item.timer);pending.delete(m.id);if(m.error)item.reject(new Error(m.error));else item.resolve(m.value);return}
 if(m.kind!=='event')return
 if(m.method==='pauseEditing'){document.body.inert=true;return}
 if(m.method==='resumeEditing'){document.body.inert=false;return}
 if(m.method==='onTeardown'){tornDown=true;if(timer)clearTimeout(timer);return}
 if(m.method==='flushDraft'){void (async()=>{if(!controls||!await controls.commit())throw new Error('当前单元格尚未提交，请保持文档打开。');await flush()})().then(()=>notify('flushed',[m.id])).catch(error=>notify('flushFailed',[m.id,String(error)]));return}
 if(m.method==='onCloseCheck'){void (async()=>{if(!controls||!await controls.commit())return true;return controls.dirty()})().then(dirty=>notify('reportCloseCheck',[{dirty}])).catch(()=>notify('reportCloseCheck',[{dirty:true}]));return}
 if(m.method==='onCloseSaveRequest'){void controls?.save().then(ok=>notify('reportCloseSaveResult',[ok])).catch(()=>notify('reportCloseSaveResult',[false]));return}
 for(const listener of listeners.get(m.method)??[])listener(m.value)
})
const session=await request('session',[]) as {revision:number;readOnly:boolean;source:{kind:string}}
revision=session.revision;published=revision
const host:OfficeHost={
 kind:'web',readOnly:session.readOnly,
 bindWorkbook(value){controls=value},
 history(command){controls?.[command]()},
 workbookLoading(value){loading=value},
 documentChanged(){if(loading||tornDown||session.readOnly)return;revision++;if(timer)clearTimeout(timer);timer=setTimeout(()=>{void flush().catch(error=>notify('draftError',[String(error)]))},1500)},
 openAgent(prompt){notify('openAgent',[prompt])},
 async beforeMcpCommand(message:McpCommandMessage){if(!controls||!await controls.commit())throw new Error('当前单元格尚未提交');await flush();if(message.expectedRevision!==undefined&&message.expectedRevision!==revision)throw new Error('OFFICE_REVISION_CONFLICT');if(session.readOnly&&!['read_sheet','describe_tools'].includes(message.command))throw new Error('OFFICE_READ_ONLY');commands=true;document.body.inert=true},
 getPathForFile(){return ''},
 signalMcpReady(){loading=false;notify('ready',[])},
 notifyPendingEdits(){},
 reportCloseSaveResult(ok){notify('reportCloseSaveResult',[ok])},
 async reportMcpResult(result){try{
  try{await flush()}catch(error){
   // An applied edit whose draft publication failed is not a completed
   // operation. Still deliver a durable failure receipt instead of letting
   // the caller time out and repeat an edit with an unknown outcome.
   await request('reportMcpResult',[{...result,ok:false,error:`草稿未持久化，当前修改仍在编辑器中：${String(error)}`,revision}]);return
  }
  await request('reportMcpResult',[{...result,revision}])
 }finally{commands=false;document.body.inert=false}},
 selectWorkbook:()=>call('selectWorkbook',[]),
 saveWorkbookEdits:async value=>{const captured=revision,result=await call('saveWorkbookEdits',[value]);if(!result.canceled)published=captured;return result},
 writeWorkbookRecovery:async value=>{const captured=revision,result=await call('writeWorkbookRecovery',[value]);if(result.ok)published=captured;return result},
 // Explicit DesktopApi bindings are generated once in the fork source, never at build time.
   openSourcePath: (...args) => call('openSourcePath', args),
  getLanguage: (...args) => call('getLanguage', args),
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
  selectWorkbooksForMerge: (...args) => call('selectWorkbooksForMerge', args),
  openWorkbooksForMerge: (...args) => call('openWorkbooksForMerge', args),
  reopenWorkbook: (...args) => call('reopenWorkbook', args),
  onSidecarCrashed: (listener) => subscribe('onSidecarCrashed', listener),
  readWorkbookRange: (...args) => call('readWorkbookRange', args),
  readWorkbookFormulas: (...args) => call('readWorkbookFormulas', args),
  findWorkbookCells: (...args) => call('findWorkbookCells', args),
  readWorkbookRowOutline: (...args) => call('readWorkbookRowOutline', args),
  recalcWorkbook: (...args) => call('recalcWorkbook', args),
  readWorkbookMedia: (...args) => call('readWorkbookMedia', args),
  readPivotDefinition: (...args) => call('readPivotDefinition', args),
  readLocalImage: (...args) => call('readLocalImage', args),
  captureScreenSources: (...args) => call('captureScreenSources', args),
  captureScreenSource: (...args) => call('captureScreenSource', args),
  beginSaveEditsTransfer: (...args) => call('beginSaveEditsTransfer', args),
  sendSaveEditsChunk: (...args) => call('sendSaveEditsChunk', args),
  abortSaveEditsTransfer: (...args) => call('abortSaveEditsTransfer', args),
  autoRenameWorkbook: (...args) => call('autoRenameWorkbook', args),
  exportPdf: (...args) => call('exportPdf', args),
  printWorkbook: (...args) => call('printWorkbook', args),
  listPrinters: (...args) => call('listPrinters', args),
  exportCsv: (...args) => call('exportCsv', args),
  confirmCsvSave: (...args) => call('confirmCsvSave', args),
  createDocument: (...args) => call('createDocument', args),
  closeWorkbook: (...args) => call('closeWorkbook', args),
  openExternal: (...args) => call('openExternal', args),
  onMenuAction: (listener) => subscribe('onMenuAction', listener),
  onWorkbookRenamed: (listener) => subscribe('onWorkbookRenamed', listener),
  onCloseSaveRequest: (listener) => subscribe('onCloseSaveRequest', listener),
  onRecoveryPrompt: (listener) => subscribe('onRecoveryPrompt', listener),
  replyRecoveryPrompt: (...args) => call('replyRecoveryPrompt', args),
  consumeNewBlankWorkbook: (...args) => call('consumeNewBlankWorkbook', args),
  onMcpCommand: (listener) => subscribe('onMcpCommand', listener),
  hasQueuedWorkbook: (...args) => call('hasQueuedWorkbook', args),
  consumeHeadlessExport: (...args) => call('consumeHeadlessExport', args),
  headlessExportDone: (...args) => call('headlessExportDone', args),
  getUserDisplayName: (...args) => call('getUserDisplayName', args),
  getAiSettings: (...args) => call('getAiSettings', args),
  setAiSettings: (...args) => call('setAiSettings', args),
  onAiSettingsChanged: (listener) => subscribe('onAiSettingsChanged', listener),
  openAiModelSettings: (...args) => call('openAiModelSettings', args),
  aiChat: (...args) => call('aiChat', args),
  aiStream: (...args) => call('aiStream', args),
  aiStreamCancel: (...args) => call('aiStreamCancel', args),
  aiGskStatus: (...args) => call('aiGskStatus', args),
  aiGskLogin: (...args) => call('aiGskLogin', args),
  webSearch: (...args) => call('webSearch', args),
  imageSearch: (...args) => call('imageSearch', args),
  generateImage: (...args) => call('generateImage', args),
  fetchImage: (...args) => call('fetchImage', args),
  onAiStream: (listener) => subscribe('onAiStream', listener),
  pickAttachments: (...args) => call('pickAttachments', args),
  addAttachmentPaths: (...args) => call('addAttachmentPaths', args),
  addPastedImage: (...args) => call('addPastedImage', args),
  readAttachment: (...args) => call('readAttachment', args),
  readAttachmentImage: (...args) => call('readAttachmentImage', args),
}
configureOfficeHost(host)
window.addEventListener('beforeunload',event=>{if(controls?.dirty()||revision!==published||commands){event.preventDefault();event.returnValue=''}})
