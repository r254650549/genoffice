/** Univer can return the source when a guarded copy command does not create a
 * worksheet. Identify the actual newly inserted sheet before any rename. */
export function duplicateWorksheet<T extends {getSheetId():string}>(workbook:{getSheets():T[];duplicateSheet(sheet:T):T},source:T):T {
 const before=new Set(workbook.getSheets().map(sheet=>sheet.getSheetId()))
 workbook.duplicateSheet(source)
 const copy=workbook.getSheets().find(sheet=>!before.has(sheet.getSheetId()))
 if(!copy)throw new Error('The worksheet copy was refused; the source was not renamed.')
 return copy
}
