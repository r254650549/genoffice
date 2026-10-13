import assert from 'node:assert/strict'
import {mkdtemp,writeFile} from 'node:fs/promises'
import {join} from 'node:path'
import JSZip from 'jszip'
import {WorkbookRuntime} from '../src/main/workbook-runtime'
import {workbookSaveRequestSchema,type WorkbookRangeResult,type WorkbookRecalcResult} from '../src/shared/desktop-api'

const directory=await mkdtemp('/tmp/kompagent-xlsx-')
const runtime=new WorkbookRuntime(directory,process.env.XLSX_SIDECAR!)
const arrays=['edits','structuralOps','chartEdits','visualEdits','visualAdditions','tableAdditions','pivotAdditions','sheetOps','sheetOrder','filterStates','hyperlinkEdits','cfStates','dvStates','pageSetupStates','noteStates','formulaValues','pivotCacheRefreshPaths','pivotRefreshUpdates','sheetProtections']
const request=(sessionId:string,extra:Record<string,unknown>)=>workbookSaveRequestSchema.parse({...Object.fromEntries(arrays.map(name=>[name,[]])),sessionId,mode:'save',definedNamesState:null,...extra})
try{
 const first=await runtime.open(),sheetId=first.sheets[0]!.id
 const original=await runtime.content(first.sessionId)
 const blankSaved=await runtime.prepareSave(request(first.sessionId,{mode:'save-as'}))
 const blankReopened=await runtime.open(await runtime.content(blankSaved.file.sessionId))
 assert.equal(blankReopened.sheets.length,1)
 assert.equal(blankReopened.sheets[0]!.name,'Sheet1')
 await runtime.close(blankReopened.sessionId)
 await runtime.close(blankSaved.file.sessionId)
 const staged=await runtime.prepareSave(request(first.sessionId,{
  structuralOps:[{kind:'insert-rows',sheetId,index:9,count:1}],
  definedNamesState:{names:[{name:'验收金额',formula:'Sheet1!$B$2:$B$4'}],preserveNames:[]},
 }))
 const stagedArchive=await JSZip.loadAsync(await runtime.content(staged.file.sessionId))
 assert((await stagedArchive.file('xl/workbook.xml')!.async('string')).includes('验收金额'))
 assert.deepEqual(await runtime.content(first.sessionId),original)
 await runtime.close(staged.file.sessionId)
 const stagedTable=await runtime.prepareSave(request(first.sessionId,{
  sheetOps:[{kind:'add-sheet',sheetId:'added-table-sheet',name:'新表结构'}],sheetOrder:[sheetId,'added-table-sheet'],
  structuralOps:[{kind:'insert-rows',sheetId:'added-table-sheet',index:9,count:1}],
  edits:[{sheetId:'added-table-sheet',row:0,column:0,writeValue:true,value:'类别'},{sheetId:'added-table-sheet',row:1,column:0,writeValue:true,value:'甲'}],
  tableAdditions:[{sheetId:'added-table-sheet',name:'StagedTable',area:{startRow:0,endRow:1,startColumn:0,endColumn:0},columnNames:['类别'],style:'TableStyleMedium2',bandedRows:true}],
 }))
 const stagedTableZip=await JSZip.loadAsync(await runtime.content(stagedTable.file.sessionId))
 assert(Object.keys(stagedTableZip.files).some(path=>/^xl\/tables\/table\d+\.xml$/.test(path)))
 assert.deepEqual(await runtime.content(first.sessionId),original)
 await runtime.close(stagedTable.file.sessionId)
 const saved=await runtime.prepareSave(request(first.sessionId,{
  edits:[{sheetId,row:0,column:0,writeValue:true,value:10,style:{bold:true}},{sheetId,row:1,column:0,writeValue:true,value:20},{sheetId,row:2,column:0,writeValue:true,value:30,formula:'SUM(A1:A2)'},{sheetId:'new-sheet',row:0,column:0,writeValue:true,value:'第二张工作表'}],
  sheetOps:[{kind:'add-sheet',sheetId:'new-sheet',name:'验收明细'}],sheetOrder:[sheetId,'new-sheet'],
 }))
 const cells=await runtime.call('readWorkbookRange',{sessionId:saved.file.sessionId,sheetId:saved.file.sheets[0]!.id,range:{startRow:0,endRow:2,startColumn:0,endColumn:0}}) as WorkbookRangeResult
 assert.equal(saved.file.sheets.length,2)
 assert(cells.cells.some(cell=>cell.row===2&&cell.formula?.replace(/^=/,'')==='SUM(A1:A2)'),JSON.stringify(cells.cells))
 const firstCell=cells.cells.find(cell=>cell.row===0&&cell.column===0)
 assert.equal(saved.file.styles[firstCell!.styleIndex!]!.bold,true)
 const calc=await runtime.call('recalcWorkbook',{sessionId:saved.file.sessionId,edits:[],reads:[{sheetId:saved.file.sheets[0]!.id,range:{startRow:2,endRow:2,startColumn:0,endColumn:0}}]}) as WorkbookRecalcResult
 assert.equal(calc.cells[0]?.number,30)
 const savedBytes=await runtime.content(saved.file.sessionId)
 assert.deepEqual(await runtime.content(first.sessionId),original,'preparing save must retain old snapshot for failed publishes')
 const opened=await runtime.open(savedBytes)
 assert.equal(opened.sheets[1]?.name,'验收明细')
 const before=await JSZip.loadAsync(original),after=await JSZip.loadAsync(savedBytes),unchanged:string[]=[]
 for(const name of Object.keys(before.files)){if(before.files[name]?.dir||!after.file(name))continue;if((await before.file(name)!.async('nodebuffer')).equals(await after.file(name)!.async('nodebuffer')))unchanged.push(name)}
 assert(unchanged.includes('xl/theme/theme1.xml'))
 await assert.rejects(runtime.call('readWorkbookRange',{sessionId:first.sessionId,sheetId,range:{startRow:0,endRow:100000,startColumn:0,endColumn:100000}}))
 await assert.rejects(runtime.call('readWorkbookRange',{sessionId:crypto.randomUUID(),sheetId,range:{startRow:0,endRow:1,startColumn:0,endColumn:1}}),/XLSX_SESSION_EXPIRED/)
 await runtime.close(first.sessionId)
 await assert.rejects(runtime.content(first.sessionId),/XLSX_SESSION_EXPIRED/)
 await writeFile('/out/xlsx-runtime-sample.xlsx',savedBytes)
 await writeFile('/out/xlsx-runtime-result.json',JSON.stringify({nativeRust:true,formulaResult:calc.cells[0]?.number,sheets:opened.sheets.map(s=>s.name),unchangedParts:unchanged,checks:{blankFirstSave:true,immutableSave:true,rangeLimit:true,expiredSession:true,reopen:true}},null,2))
 console.log('Native XLSX open/read/formula/style/multi-sheet/save/reopen passed.')
}finally{await runtime.dispose()}
