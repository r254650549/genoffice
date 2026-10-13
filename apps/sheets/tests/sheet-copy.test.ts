import {expect,it,vi} from 'vitest'
import {duplicateWorksheet} from '../src/renderer/sheet-copy'
it('rejects a refused native copy before a caller can rename its source',()=>{
 const source={getSheetId:()=> 'source',setName:vi.fn()},workbook={getSheets:()=>[source],duplicateSheet:()=>source}
 expect(()=>duplicateWorksheet(workbook,source).setName('copy')).toThrow('source was not renamed')
 expect(source.setName).not.toHaveBeenCalled()
})
it('identifies the inserted worksheet even when the native facade returns a stale source',()=>{
 const source={getSheetId:()=> 'source',setName:vi.fn()},copy={getSheetId:()=> 'copy',setName:vi.fn()};let sheets=[source]
 const workbook={getSheets:()=>sheets,duplicateSheet:()=>{sheets=[source,copy];return source}}
 duplicateWorksheet(workbook,source).setName('copy name')
 expect(copy.setName).toHaveBeenCalledWith('copy name');expect(source.setName).not.toHaveBeenCalled()
})
