import {createHash,randomUUID} from 'node:crypto'
import {mkdir,readFile,writeFile,unlink,realpath} from 'node:fs/promises'
import {join,resolve} from 'node:path'
import {csvToXlsxBufferForOpen} from '@genoffice/xlsx-gateway/gateway/csv-import'
import {readArchiveEntryText} from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
import {parsePivotDefinition} from '@genoffice/xlsx-gateway/gateway/xlsx-pivot'
import * as schemas from '../shared/desktop-api'
import {XlsxSidecarClient} from './xlsx-sidecar-client'
import {writeWorkbookTo} from './workbook-save'
import {SaveEditsTransferStore} from './save-edits-transfer'

type Session={snapshotPath:string;sheetNames:ReadonlyMap<string,string>;file:schemas.WorkbookFile;transfers:SaveEditsTransferStore}
/** Native workbook host without Electron. One instance belongs to one execution environment.
 * The outer host authorizes files; this service only opens supplied bytes and never caller paths.
 * Saving prepares a new immutable snapshot. Publishing it is the outer host's versioned commit.
 */
export class WorkbookRuntime {
 private readonly sessions=new Map<string,Session>()
 private readonly client:XlsxSidecarClient
 constructor(readonly directory:string,binaryPath:string){
  this.client=new XlsxSidecarClient(binaryPath)
  this.client.onProcessExit(()=>{for(const session of this.sessions.values()){session.transfers.dispose();void unlink(session.snapshotPath).catch(()=>{})}this.sessions.clear()})
 }
 private session(id:string){const value=this.sessions.get(id);if(!value)throw new Error('XLSX_SESSION_EXPIRED');return value}
 hasSession(id:string){return this.sessions.has(id)}
 async open(bytes?:Buffer):Promise<schemas.WorkbookFile>{
  if(this.sessions.size>=32)throw new Error('XLSX_SESSION_LIMIT')
  const content=bytes??(await csvToXlsxBufferForOpen('', 'Sheet1')).buffer
  if(content.length>64*1024*1024||content[0]!==80||content[1]!==75)throw new Error('XLSX_INVALID_CONTENT')
  await mkdir(this.directory,{recursive:true})
  if(await realpath(this.directory)!==resolve(this.directory))throw new Error('XLSX_INVALID_DIRECTORY')
  const snapshotPath=join(this.directory,randomUUID()+'.xlsx')
  await writeFile(snapshotPath,content,{flag:'wx'})
  try{
   const opened=await this.client.open(snapshotPath,'zh','yyyy/m/d') as Record<string,unknown>
   const file=schemas.workbookFileSchema.parse({...opened,path:snapshotPath,sha256:createHash('sha256').update(content).digest('hex'),fileBytes:content.length,readOnly:false,needsSaveAs:false})
   this.sessions.set(file.sessionId,{snapshotPath,sheetNames:new Map(file.sheets.map(sheet=>[sheet.id,sheet.name])),file,transfers:new SaveEditsTransferStore()})
   return file
  }catch(error){await unlink(snapshotPath).catch(()=>{});throw error}
 }
 async content(id:string){return readFile(this.session(id).snapshotPath)}
 async prepareSave(raw:unknown){
  let request=schemas.workbookSaveRequestSchema.parse(raw)
  const session=this.session(request.sessionId)
  if(request.editsTransferId){request={...request,edits:session.transfers.take(request.editsTransferId,request.sessionId)};delete request.editsTransferId}
  const target=join(this.directory,randomUUID()+'.xlsx')
  let intermediate:schemas.WorkbookFile|undefined
  try{
   const shifts=request.structuralOps.length>0
   const sheetShifts=request.sheetOps.some(op=>op.kind!=='set-sheet-tab-color')
   const heldTables=shifts?request.tableAdditions:[]
   const heldPivots=shifts||sheetShifts?request.pivotAdditions:[]
   const heldEdits=shifts?(request.tableEdits??[]).filter(edit=>edit.area!==undefined):[]
   const heldNames=shifts||sheetShifts?request.definedNamesState:null
   const staged=heldTables.length>0||heldPivots.length>0||heldEdits.length>0||heldNames!==null
   let mutation=await writeWorkbookTo(this.client,session,staged?{...request,
    tableAdditions:heldTables.length?[]:request.tableAdditions,
    pivotAdditions:heldPivots.length?[]:request.pivotAdditions,
    tableEdits:(request.tableEdits??[]).filter(edit=>!heldEdits.includes(edit)),
    definedNamesState:null,
   }:request,target)
   if(staged){
    // Use the same original two-pass writer as the desktop save path, but
    // keep its intermediate snapshot private. A failed second pass cannot
    // publish a partial cloud version or replace the browser's base file.
    intermediate=await this.open(await readFile(target))
    const names=new Map(session.sheetNames)
    for(const op of request.sheetOps){
     if(op.kind==='add-sheet'||op.kind==='duplicate-sheet')names.set(op.sheetId,op.name)
     else if(op.kind==='rename-sheet')names.set(op.sheetId,op.newName)
    }
    const mapped=(id:string)=>{const found=intermediate!.sheets.find(sheet=>sheet.name===names.get(id));if(!found)throw new Error('XLSX_SHEET_NOT_FOUND');return found.id}
    const cleared=['edits','bulkConstantFills','structuralOps','chartEdits','visualEdits','visualAdditions','sheetOps','sheetOrder','filterStates','hyperlinkEdits','cfStates','dvStates','pageSetupStates','noteStates','formulaValues','pivotCacheRefreshPaths','pivotRefreshUpdates','sheetProtections','sparklineAdditions','protectedRangeStates']
    const second=schemas.workbookSaveRequestSchema.parse({...request,...Object.fromEntries(cleared.map(key=>[key,[]])),
     sessionId:intermediate.sessionId,mode:'save-as',editsTransferId:undefined,themeState:null,workbookProtectionState:null,
     tableAdditions:heldTables.map(table=>({...table,sheetId:mapped(table.sheetId)})),
     pivotAdditions:heldPivots.map(pivot=>({...pivot,sheetId:mapped(pivot.sheetId),sourceSheetId:mapped(pivot.sourceSheetId)})),
     tableEdits:heldEdits.map(edit=>({...edit,sheetId:mapped(edit.sheetId)})),definedNamesState:heldNames})
    const final=await writeWorkbookTo(this.client,this.session(intermediate.sessionId),second,target)
    mutation={...final,touchedEntries:[...new Set([...mutation.touchedEntries,...final.touchedEntries])]}
   }
   const bytes=await readFile(target)
   const file=await this.open(bytes)
   return {canceled:false as const,file,touchedEntries:mutation.touchedEntries}
  }finally{if(intermediate)await this.close(intermediate.sessionId).catch(()=>{});await unlink(target).catch(()=>{})}
 }
 async call(method:string,raw:unknown):Promise<unknown>{
  // Every request is validated with the original schemas and bound to an owned session.
  const id=(raw as {sessionId?:unknown})?.sessionId
  if(typeof id!=='string')throw new Error('XLSX_SESSION_REQUIRED')
  const session=this.session(id)
  switch(method){
   case 'readWorkbookRange':return schemas.workbookRangeResultSchema.parse(await this.client.readRange(schemas.workbookRangeRequestSchema.parse(raw)))
   case 'readWorkbookFormulas':return schemas.workbookFormulaCellsResultSchema.parse(await this.client.readFormulaCells(schemas.workbookFormulaCellsRequestSchema.parse(raw)))
   case 'readWorkbookRowOutline':return schemas.workbookRowOutlineResultSchema.parse(await this.client.readRowOutline(schemas.workbookFormulaCellsRequestSchema.parse(raw)))
   case 'findWorkbookCells':return schemas.workbookFindCellsResultSchema.parse(await this.client.findCells(schemas.workbookFindCellsRequestSchema.parse(raw)))
   case 'readWorkbookMedia':return schemas.workbookMediaResultSchema.parse(await this.client.readMedia(schemas.workbookMediaRequestSchema.parse(raw)))
   case 'readPivotDefinition':{const r=schemas.workbookPivotRequestSchema.parse(raw);const [pivot,cache]=await Promise.all([readArchiveEntryText(this.client,session.snapshotPath,r.path),readArchiveEntryText(this.client,session.snapshotPath,r.cachePath)]);return schemas.workbookPivotDefinitionSchema.parse(parsePivotDefinition(pivot,cache))}
   case 'recalcWorkbook':{
    const r=schemas.workbookRecalcRequestSchema.parse(raw)
    const name=(sheetId:string)=>{const n=session.sheetNames.get(sheetId);if(!n)throw new Error('XLSX_SHEET_NOT_FOUND');return n}
    const result=await this.client.recalcCells({path:session.snapshotPath,edits:r.edits.map(({sheetId,...edit})=>({...edit,sheet:name(sheetId)})),reads:r.reads.map(({sheetId,...read})=>({...read,sheet:name(sheetId)}))}) as {cells:Array<{sheet:string;row:number;column:number;formatted:string;number?:number;isError?:boolean;isFormula:boolean}>}
    const ids=new Map([...session.sheetNames].map(([key,value])=>[value,key]))
    return schemas.workbookRecalcResultSchema.parse({cells:result.cells.map(({sheet,...cell})=>({...cell,sheetId:ids.get(sheet)}))})
   }
   case 'beginSaveEditsTransfer':session.transfers.begin(schemas.workbookSaveEditsBeginSchema.parse(raw));return null
   case 'sendSaveEditsChunk':{const r=schemas.workbookSaveEditsChunkSchema.parse(raw);session.transfers.addChunk({...r,edits:schemas.saveEditsChunkArraySchema.parse(JSON.parse(r.editsJson))});return null}
   case 'abortSaveEditsTransfer':{const r=schemas.workbookSaveEditsAbortSchema.parse(raw);session.transfers.discard(r.transferId,r.sessionId);return null}
   case 'saveWorkbookEdits':return this.prepareSave(raw)
   default:throw new Error('XLSX_METHOD_UNSUPPORTED')
  }
 }
 async close(id:string){const session=this.session(id);this.sessions.delete(id);session.transfers.dispose();try{await this.client.close(id)}finally{await unlink(session.snapshotPath).catch(()=>{})}}
 async dispose(){for(const id of [...this.sessions.keys()])await this.close(id).catch(()=>{});this.client.stop()}
}
