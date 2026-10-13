import {beforeEach,expect,it,vi} from 'vitest'
const load=vi.hoisted(()=>vi.fn())
vi.mock('../src/renderer/univer-sync',()=>({ensureLazyRangeLoaded:load,lazyCellReader:()=>()=>({value:77,rawValue:77})}))
vi.mock('../src/renderer/grid-growth',()=>({ensureAddressableGrid:vi.fn()}))
import {readLiveDraftCells,type WorkbookReadContext} from '../src/renderer/ai/workbook-readers'
import {createEditJournal,recordSetRangeValues} from '../src/renderer/edit-journal'
function fixture(newSheet=false){
 const journal=createEditJournal(),worksheet={getSheetId:()=> 'sheet-1'},workbook={getActiveSheet:()=>worksheet,getSheetBySheetId:()=>worksheet}
 const ctx={univerRef:{current:{univerAPI:{getActiveWorkbook:()=>workbook}}},lazyWorkbookRef:{current:{file:{sheets:newSheet?[]:[{id:'sheet-1',rowCount:3,columnCount:1}]},editJournal:journal}},adapterRef:{current:{}}} as unknown as WorkbookReadContext
 return {ctx,journal}
}
beforeEach(()=>load.mockReset().mockResolvedValue(true))
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
