import type {RangeBounds} from '@genoffice/xlsx-gateway/domain/cell-address'
import {areasOverlap} from '@genoffice/xlsx-gateway/domain/pivot-layout'
import type {LazyWorkbookState} from './univer-state'
import {fileRangeToScreenRanges} from './view-transform'
export function fileTableOverlapError(state:LazyWorkbookState,sheetId:string,bounds:RangeBounds):string|null {
 for(const table of state.file.sheets.find(sheet=>sheet.id===sheetId)?.tables??[]){
  const ranges=fileRangeToScreenRanges(state.editJournal.structuralOps.get(sheetId)??[],table.range)
  if(ranges.some(range=>areasOverlap(bounds,range)))return `Table range overlaps the existing file table "${table.name??''}"; choose a different range.`
 }
 return null
}
