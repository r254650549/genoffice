import {beforeEach,expect,it,vi} from 'vitest'
const load=vi.hoisted(()=>vi.fn())
const cellRead=vi.hoisted(()=>vi.fn())
vi.mock('../src/renderer/univer-sync',()=>({ensureLazyRangeLoaded:load,lazyCellReader:()=>cellRead,modelCellValue:(cell:{value:unknown;rawValue?:unknown})=>cell.rawValue??cell.value}))
vi.mock('../src/renderer/grid-growth',()=>({ensureAddressableGrid:vi.fn()}))
import {readLiveDraftCells,getActiveSheetInfo,type WorkbookReadContext} from '../src/renderer/ai/workbook-readers'
import {createEditJournal,recordSetRangeValues} from '../src/renderer/edit-journal'
function fixture(newSheet=false){
 const journal=createEditJournal(),worksheet={getSheetId:()=> 'sheet-1'},workbook={getActiveSheet:()=>worksheet,getSheetBySheetId:()=>worksheet}
 const ctx={univerRef:{current:{univerAPI:{getActiveWorkbook:()=>workbook}}},lazyWorkbookRef:{current:{file:{sheets:newSheet?[]:[{id:'sheet-1',rowCount:3,columnCount:1}]},editJournal:journal}},adapterRef:{current:{}}} as unknown as WorkbookReadContext
 return {ctx,journal}
}
beforeEach(()=>{load.mockReset().mockResolvedValue(true);cellRead.mockReset().mockReturnValue({value:77,rawValue:77})})
it('waits for the real formula engine value after hydration instead of returning a transient blank',async()=>{
 const {ctx}=fixture();let calculated=false
 cellRead.mockImplementation(()=>({value:calculated?84:null,rawValue:calculated?84:null,formula:'=A1*2'}))
 const read=readLiveDraftCells(ctx,['A1'],'sheet-1',()=>{})
 await Promise.resolve();calculated=true
 await expect(read).resolves.toMatchObject({A1:{value:84,rawValue:84,formula:'=A1*2'}})
})
it('keeps workbook discovery usable when a deleted row leaves a stale full-column selection',()=>{
 const ctx={univerRef:{current:{univerAPI:{getActiveWorkbook:()=>({getActiveRange:()=>{throw new Error('Range is out of bounds')},getActiveSheet:()=>({getSheetId:()=> 'sheet-1'})})}}},lazyWorkbookRef:{current:null},adapterRef:{current:{getSnapshot:()=>({revision:1,sheets:[{id:'sheet-1',name:'Sheet1',cells:{A1:{value:10}}}]})}}} as unknown as WorkbookReadContext
 expect(getActiveSheetInfo(ctx)).toMatchObject({sheetId:'sheet-1',sheets:[{id:'sheet-1',name:'Sheet1'}],selection:undefined})
})
it('reads an unsaved remote cell from the real edit journal without fetching nonexistent base bytes',async()=>{
 const {ctx,journal}=fixture();recordSetRangeValues(journal,'sheet-1',{1999:{1:{v:99}}})
 expect(await readLiveDraftCells(ctx,['B2000'],'sheet-1',()=>{})).toMatchObject({B2000:{value:99,rawValue:99}})
 expect(load).not.toHaveBeenCalled()
})
it('reads an unsaved new sheet before its first file publication',async()=>{
 const {ctx,journal}=fixture(true);recordSetRangeValues(journal,'sheet-1',{0:{0:{v:'draft'}}})
 expect(await readLiveDraftCells(ctx,['A1'],'sheet-1',()=>{})).toMatchObject({A1:{value:'draft'}})
 expect(load).not.toHaveBeenCalled()
})
it('hydrates existing file cells before returning their values',async()=>{
 const {ctx}=fixture()
 expect(await readLiveDraftCells(ctx,['A1'],'sheet-1',()=>{})).toMatchObject({A1:{value:77}})
 expect(load).toHaveBeenCalledTimes(1)
})
it('reports a genuine base-file load failure rather than inventing an empty value',async()=>{
 const {ctx}=fixture();load.mockResolvedValue(false)
 await expect(readLiveDraftCells(ctx,['A1'],'sheet-1',()=>{})).rejects.toThrow('could not be loaded')
})
