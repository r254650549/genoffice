import {createServer,createConnection} from 'node:net'
import {spawn} from 'node:child_process'
import {readFile,writeFile,mkdir,unlink,stat} from 'node:fs/promises'
import {createHash} from 'node:crypto'
import {join,dirname} from 'node:path'
import {fileURLToPath} from 'node:url'
import {z} from 'zod'
import {WorkbookRuntime} from './workbook-runtime'

const requestSchema=z.object({requestId:z.uuid(),owner:z.string().regex(/^[a-f0-9-]{36}:[a-f0-9-]{36}$/),method:z.enum(['open','content','close','touch','readWorkbookRange','readWorkbookFormulas','readWorkbookRowOutline','findWorkbookCells','readWorkbookMedia','readPivotDefinition','recalcWorkbook','beginSaveEditsTransfer','sendSaveEditsChunk','abortSaveEditsTransfer','saveWorkbookEdits']),payload:z.record(z.string(),z.unknown()).default({}),bytes:z.string().max(90*1024*1024).optional()}).strict()
const limit=96*1024*1024
type Request=z.infer<typeof requestSchema>
export function workbookServiceIdleExpired(active:number,openSessions:number,lastRequest:number,now:number){
 // A frozen browser cannot send a keepalive. Retain its bounded native
 // snapshots while documents remain open; released services exit promptly.
 return active===0&&now-lastRequest>(openSessions===0?300000:86400000)
}
/** One service per principal environment. The outer authenticated host supplies the owner;
 * native session ids never grant access across Office documents or browser instances. */
export class WorkbookService {
 private readonly owners=new Map<string,string>()
 private readonly receipts=new Map<string,{digest:string;result:Promise<unknown>}>()
 private queue:Promise<unknown>=Promise.resolve()
 constructor(readonly runtime:WorkbookRuntime){}
 get openSessionCount(){for(const id of this.owners.keys())if(!this.runtime.hasSession(id))this.owners.delete(id);return this.owners.size}
 dispatch(raw:unknown):Promise<unknown>{
  const request=requestSchema.parse(raw),digest=createHash('sha256').update(JSON.stringify(request)).digest('hex')
  // Range/content/keepalive reads are repeatable and must not retain large
  // base64 results or consume the mutation-receipt budget while a tab is open.
  if(['touch','content','readWorkbookRange','readWorkbookFormulas','readWorkbookRowOutline','findWorkbookCells','readWorkbookMedia','readPivotDefinition','recalcWorkbook'].includes(request.method)){
   const result=this.queue.then(()=>this.execute(request));this.queue=result.catch(()=>{});return result
  }
  const known=this.receipts.get(request.requestId)
  if(known){if(known.digest!==digest)throw new Error('XLSX_REQUEST_CONFLICT');return known.result}
  // Do not evict unknown mutation receipts and silently execute them again.
  if(this.receipts.size>=4096)throw new Error('XLSX_RECEIPT_LIMIT')
  const result=this.queue.then(()=>this.execute(request))
  this.queue=result.catch(()=>{})
  this.receipts.set(request.requestId,{digest,result})
  return result
 }
 private async execute(request:Request):Promise<unknown>{
  const {owner,method,payload}=request
  if(method==='open'){
   const file=await this.runtime.open(request.bytes===undefined?undefined:Buffer.from(request.bytes,'base64'))
   this.owners.set(file.sessionId,owner);return file
  }
  const id=payload.sessionId
  if(typeof id!=='string'||this.owners.get(id)!==owner)throw new Error('XLSX_SESSION_EXPIRED')
  if(!this.runtime.hasSession(id)){this.owners.delete(id);throw new Error('XLSX_SESSION_EXPIRED')}
  if(method==='touch')return null
  if(method==='close'){await this.runtime.close(id);this.owners.delete(id);return null}
  if(method==='content')return {bytes:(await this.runtime.content(id)).toString('base64')}
  const result=await this.runtime.call(method,payload)
  if(method==='saveWorkbookEdits')this.owners.set((result as {file:{sessionId:string}}).file.sessionId,owner)
  return result
 }
 async dispose(){await this.queue;await this.runtime.dispose();this.owners.clear();this.receipts.clear()}
}

// The service uses a Unix socket inside the existing execution container. No host
// port, credentials, arbitrary filenames or shell fragments come from the browser.
async function main(){
 const [mode,root,id]=process.argv.slice(2)
 if(!root||!/^\/tmp\/kompagent-office-[a-f0-9]{24}$/.test(root))throw new Error('XLSX_INVALID_DIRECTORY')
 // Linux abstract sockets disappear on process exit and are atomically bound.
 // Concurrent starters cannot unlink another process's live socket.
 const socketPath='\0'+root.slice(5)
 await mkdir(root,{recursive:true,mode:0o700})
 if(mode==='serve'){
  const service=new WorkbookService(new WorkbookRuntime(join(root,'snapshots'),join(dirname(fileURLToPath(import.meta.url)),'xlsx-sidecar')))
  let lastRequest=Date.now(),active=0
  const server=createServer(socket=>{
   let line='';socket.setTimeout(195000,()=>socket.destroy())
   socket.on('data',chunk=>{line+=chunk.toString();if(line.length>40){socket.destroy();return}if(!line.endsWith('\n'))return
    const requestId=line.trim();line='';if(!z.uuid().safeParse(requestId).success){socket.destroy();return}
    active++;lastRequest=Date.now()
    void (async()=>{
     const input=join(root,requestId+'.request'),output=join(root,requestId+'.result')
     let result:unknown
     try{if((await stat(input)).size>limit)throw new Error('XLSX_REQUEST_TOO_LARGE');const request=requestSchema.parse(JSON.parse(await readFile(input,'utf8')));if(request.requestId!==requestId)throw new Error('XLSX_REQUEST_CONFLICT');result={ok:true,value:await service.dispatch(request)}}
     catch(error){const message=error instanceof Error?error.message:'';result={ok:false,error:/^XLSX_[A-Z_]+$/.test(message)?message:'XLSX_OPERATION_FAILED',...(!/^XLSX_[A-Z_]+$/.test(message)?{detail:message.slice(0,1500)}:{})}}
     await writeFile(output,JSON.stringify(result),{mode:0o600});await unlink(input).catch(()=>{});socket.end('done\n')
    })().catch(()=>socket.destroy()).finally(()=>{active--;lastRequest=Date.now()})
   })
   socket.on('error',()=>{})
  })
  // listen is the startup lock. Never unlink a socket owned by another live service.
  server.on('error',()=>{void service.dispose().finally(()=>process.exit(1))})
  server.listen(socketPath)
  let stopping=false
  const stop=async()=>{if(stopping)return;stopping=true;clearInterval(timer);server.close();await service.dispose();process.exit(0)}
  const timer=setInterval(()=>{if(workbookServiceIdleExpired(active,service.openSessionCount,lastRequest,Date.now()))void stop()},10000)
  process.on('SIGTERM',()=>void stop());process.on('SIGINT',()=>void stop())
 }else if(mode==='request'){
  if(!z.uuid().safeParse(id).success)throw new Error('XLSX_INVALID_REQUEST')
  const send=()=>new Promise<void>((resolve,reject)=>{const socket=createConnection(socketPath);let connected=false;socket.setTimeout(190000,()=>socket.destroy(new Error('XLSX_REQUEST_UNKNOWN')));socket.on('connect',()=>{connected=true;socket.write(id+'\n')});socket.on('data',chunk=>{if(chunk.toString().includes('done')){socket.destroy();resolve()}});socket.on('error',error=>reject(Object.assign(error,{connected})));socket.on('end',()=>reject(new Error('XLSX_REQUEST_UNKNOWN')))})
  try{await send()}catch(error){
   if((error as {connected?:boolean}).connected)throw error
   const code=(error as NodeJS.ErrnoException).code;if(code!=='ENOENT'&&code!=='ECONNREFUSED')throw error
   const child=spawn(process.execPath,[fileURLToPath(import.meta.url),'serve',root],{detached:true,stdio:'ignore'});child.unref()
   let done=false;for(let n=0;n<50;n++){await new Promise(resolve=>setTimeout(resolve,100));try{await send();done=true;break}catch(next){if((next as {connected?:boolean}).connected)throw next}}
   if(!done)throw new Error('XLSX_RUNTIME_UNAVAILABLE')
  }
 }else throw new Error('XLSX_INVALID_MODE')
}
if(process.argv[2]==='serve'||process.argv[2]==='request')void main().catch(()=>{process.exitCode=1})
