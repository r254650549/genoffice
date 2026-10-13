import {describe,expect,it,vi} from 'vitest'
import {WorkbookService,workbookServiceIdleExpired} from '../src/main/workbook-service'
import type {WorkbookRuntime} from '../src/main/workbook-runtime'

const owner=crypto.randomUUID()+':'+crypto.randomUUID(),sessionId=crypto.randomUUID()
function setup(){
 const runtime={hasSession:vi.fn().mockReturnValue(true),open:vi.fn().mockResolvedValue({sessionId}),content:vi.fn().mockResolvedValue(Buffer.from('content')),close:vi.fn().mockResolvedValue(null),call:vi.fn(),dispose:vi.fn()}
 const service=new WorkbookService(runtime as unknown as WorkbookRuntime)
 const request=(method:string,payload:Record<string,unknown>={sessionId})=>({requestId:crypto.randomUUID(),owner,method,payload})
 return {runtime,service,request}
}
describe('owned workbook service',()=>{
 it('does not expire an open document when a background browser pauses timers',()=>{
  expect(workbookServiceIdleExpired(0,1,0,360000)).toBe(false)
  expect(workbookServiceIdleExpired(0,0,0,360000)).toBe(true)
  expect(workbookServiceIdleExpired(1,0,0,360000)).toBe(false)
  expect(workbookServiceIdleExpired(0,1,0,86400001)).toBe(true)
 })
 it('keeps reads and keepalives outside the mutation receipt budget',async()=>{
  const {service,request,runtime}=setup()
  await service.dispatch(request('open',{}))
  for(let n=0;n<4100;n++)await service.dispatch(request(n%2?'touch':'content'))
  expect(runtime.content).toHaveBeenCalledTimes(2050)
  await service.dispatch(request('close'))
  await expect(service.dispatch(request('touch'))).rejects.toThrow('XLSX_SESSION_EXPIRED')
 })
 it('refuses another document/browser owner for keepalives and content',async()=>{
  const {service,request,runtime}=setup()
  await service.dispatch(request('open',{}))
  for(const method of ['touch','content'])await expect(service.dispatch({...request(method),owner:crypto.randomUUID()+':'+crypto.randomUUID()})).rejects.toThrow('XLSX_SESSION_EXPIRED')
  expect(runtime.content).not.toHaveBeenCalled()
 })
 it('replays a mutation once and rejects a reused id with different content',async()=>{
  const {service,request,runtime}=setup(),open=request('open',{})
  await service.dispatch(open)
  await service.dispatch(open)
  expect(runtime.open).toHaveBeenCalledTimes(1)
  expect(()=>service.dispatch({...open,bytes:'UEs='})).toThrow('XLSX_REQUEST_CONFLICT')
 })
})
