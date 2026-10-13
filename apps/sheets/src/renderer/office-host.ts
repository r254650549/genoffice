import type {DesktopApi,McpCommandMessage} from '../shared/desktop-api'
/** Explicit host injection; Electron and Web share the complete native renderer. */
export interface OfficeHost extends DesktopApi {
 kind?:'electron'|'web'
 readOnly?:boolean
 workbookLoading?(loading:boolean):void
 documentChanged?():void
 openAgent?(prompt?:string):void
 bindWorkbook?(controls:{commit():Promise<boolean>;flush():Promise<boolean>;dirty():boolean;save():Promise<boolean>;undo():void;redo():void}):void
 beforeMcpCommand?(message:McpCommandMessage):Promise<void>
 history?(command:'undo'|'redo'):void
}
let host:OfficeHost
export function configureOfficeHost(value:OfficeHost){host=value}
export function getOfficeHost():OfficeHost{return host}
