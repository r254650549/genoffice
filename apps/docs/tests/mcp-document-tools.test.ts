import {afterEach,expect,it} from 'vitest'
import {Editor} from '@tiptap/core'
import {editorExtensions} from '../src/renderer/editor/extensions'
import {configureOfficeHost,type OfficeHost} from '../src/renderer/office-host'
import {describeDocumentTools,runDocumentCommand,type McpAccess} from '../src/renderer/mcp-document-tools'
import {AGENT_TOOLS} from '../src/renderer/ai/tools'
import {DOCUMENT_COMMANDS} from '../src/shared/document-commands'
import type {FileActionContext} from '../src/renderer/file-actions'
import type {NoteInfo} from '@genoffice/docx-engine'

const editors:Editor[]=[]
afterEach(()=>{editors.splice(0).forEach(editor=>editor.destroy())})
function fixture(access:McpAccess={}){
 configureOfficeHost({kind:'web',readOnly:false} as OfficeHost)
 const editor=new Editor({element:document.createElement('div'),extensions:editorExtensions,content:{type:'doc',content:[{type:'docParagraph',attrs:{docxIndex:null},content:[{type:'text',text:'Alpha beta gamma'}]}]}})
 editors.push(editor)
 const ctx={editor,doc:{isBlank:true,parsed:{blocks:[]}}} as unknown as FileActionContext
 return {editor,run:(command:typeof DOCUMENT_COMMANDS[number],payload:unknown={})=>runDocumentCommand({getCtx:()=>ctx,getAccess:()=>access},command,payload)}
}
it('exposes native schemas and complete operation help, excluding second-agent/service tools',()=>{
 const catalog=describeDocumentTools()
 const excluded=['write_document','web_search','image_search','analyze_media','generate_image','create_document']
 expect(catalog.tools.map(tool=>tool.name).sort()).toEqual([...AGENT_TOOLS.filter(tool=>!excluded.includes(tool.name)).map(tool=>tool.name==='get_document_context'?'read_document':tool.name),'save_document','undo','redo','describe_tools'].sort())
 expect(catalog.operations?.some(op=>op.name==='setFont'&&op.signature.includes('target'))).toBe(true)
 expect(describeDocumentTools(['add_comment']).tools[0].inputSchema).toEqual(AGENT_TOOLS.find(tool=>tool.name==='add_comment')!.inputSchema)
})
it('freezes a selected range, rejects stale selections, and uses native undo/redo',async()=>{
 const f=fixture();f.editor.commands.setTextSelection({from:1,to:6})
 const read=await f.run('read_document') as {selectionToken:string}
 f.editor.commands.setTextSelection({from:7,to:11})
 await f.run('replace_selection',{html:'ONE',selectionToken:read.selectionToken})
 expect(f.editor.state.doc.textContent).toBe('ONE beta gamma')
 await expect(f.run('replace_selection',{html:'bad',selectionToken:read.selectionToken})).rejects.toThrow('OFFICE_SELECTION_STALE')
 await f.run('undo');expect(f.editor.state.doc.textContent).toBe('Alpha beta gamma')
 await f.run('redo');expect(f.editor.state.doc.textContent).toBe('ONE beta gamma')
})
it('validates a whole batch before dispatch and dryRun leaves content unchanged',async()=>{
 const f=fixture();await f.run('read_document');const before=f.editor.getJSON()
 await f.run('apply_ops',{dryRun:true,ops:[{op:'setFont',target:{blockIndexes:[0]},bold:true}]})
 expect(f.editor.getJSON()).toEqual(before)
 await expect(f.run('apply_ops',{ops:[{op:'setFont',target:{blockIndexes:[0]},bold:true},{op:'invalid'}]})).rejects.toThrow('Nothing was applied')
 expect(f.editor.getJSON()).toEqual(before)
})
it('passes native note and header stores through and reads their changes',async()=>{
 const notes:{footnote:NoteInfo[];endnote:NoteInfo[]}={footnote:[],endnote:[]};let header=''
 const f=fixture({hf:{read:()=>({header,footer:'',headerFirst:null,footerFirst:null,headerEven:null,footerEven:null,titlePg:false,evenOddHf:false,multiSection:false}),set:(_kind,_view,text)=>{header=text;return null}},notes:{list:kind=>notes[kind],add:(kind,text)=>{const id=String(notes[kind].length+1);notes[kind].push({id,text});return id},remove:()=>false,replace:(kind,id,next)=>{const i=notes[kind].findIndex(n=>n.id===id);if(i<0)return false;notes[kind][i]=next;return true}}})
 await f.run('read_document');await f.run('insert_footnote',{blockIndex:0,text:'Source A'})
 await f.run('edit_note',{kind:'footnote',id:'1',find:'A',replace:'B'})
 expect(JSON.stringify(await f.run('read_notes'))).toContain('Source B')
 await f.run('set_header_footer',{kind:'header',text:'Project header'})
 expect(JSON.stringify(await f.run('read_document'))).toContain('Project header')
})
it('allows reads but refuses edits for read-only documents and protected comments',async()=>{
 const f=fixture({locked:true});await expect(f.run('read_document')).resolves.toBeTruthy()
 await expect(f.run('insert_content',{html:'<p>bad</p>'})).rejects.toThrow('OFFICE_READ_ONLY')
 const c=fixture({commentsAllowed:false});await expect(c.run('add_comment',{blockIndex:0,comment:'bad'})).rejects.toThrow('comments are not allowed')
})
