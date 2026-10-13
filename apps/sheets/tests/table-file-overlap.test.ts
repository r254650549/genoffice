import {expect,it,vi} from 'vitest'
import {applyAiTableAdd} from '../src/renderer/workbook-ops'
import {createEditJournal,recordStructuralOp} from '../src/renderer/edit-journal'
import type {LazyWorkbookState,UniverRuntime} from '../src/renderer/univer-state'
function fixture(){
 const journal=createEditJournal(),worksheet={getRange:vi.fn()}
 const runtime={univerAPI:{getActiveWorkbook:()=>({getSheetBySheetId:()=>worksheet})}} as unknown as UniverRuntime
 const state={file:{sheets:[{id:'sheet-1',name:'Sheet1',tables:[{name:'Existing',range:{startRow:0,endRow:3,startColumn:0,endColumn:1},headerRowCount:1}]}]},editJournal:journal} as unknown as LazyWorkbookState
 return {runtime,state,worksheet,journal}
}
it('rejects overlapping a file-native table before normalizing headers or inserting a journal entry',()=>{
 const {runtime,state,worksheet,journal}=fixture()
 expect(()=>applyAiTableAdd(runtime,state,{op:'add_table',sheetId:'sheet-1',range:'A2:C4',name:'Overlap'})).toThrow('overlaps the existing file table')
 expect(worksheet.getRange).not.toHaveBeenCalled();expect(journal.tableAdds).toHaveLength(0)
})
it('checks the current table coordinates after a journaled row insertion',()=>{
 const {runtime,state,worksheet,journal}=fixture()
 recordStructuralOp(journal,'sheet-1',{kind:'insert-rows',index:0,count:10})
 expect(()=>applyAiTableAdd(runtime,state,{op:'add_table',sheetId:'sheet-1',range:'A11:C14',name:'Overlap'})).toThrow('overlaps the existing file table')
 expect(worksheet.getRange).not.toHaveBeenCalled()
})
