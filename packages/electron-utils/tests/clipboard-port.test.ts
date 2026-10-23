import {beforeEach,expect,it,vi} from 'vitest'
const mock=vi.hoisted(()=>({items:[] as any[],nativeImage:{createFromBuffer:(bytes:Buffer)=>({isEmpty:()=>bytes.length===0,toPNG:()=>bytes})}}))
vi.mock('electron',()=>({
  nativeImage:mock.nativeImage,
  ClipboardItem:class {types:string[];constructor(private values:Record<string,Blob|string>){this.types=Object.keys(values)}async getType(type:string){const value=this.values[type];return typeof value==='string'?new Blob([value],{type}):value}},
  clipboard:{read:async()=>mock.items,write:async(items:any[])=>{mock.items=items},readText:async()=>''},
}))
import {nativeClipboardPort} from '../src/clipboard-port'
beforeEach(()=>{mock.items=[]})
it('round trips custom ownership markers through the async ClipboardItem API and notices external replacement',async()=>{
  const clipboard=nativeClipboardPort();await clipboard.writeBuffer('io.genoffice.slides.elements',Buffer.from('owner-token'));
  expect((await clipboard.readBuffer('io.genoffice.slides.elements')).toString()).toBe('owner-token');
  mock.items=[];expect(await clipboard.readBuffer('io.genoffice.slides.elements')).toHaveLength(0);
})
it('writes the PNG and HTML together and reads them through asynchronous blobs',async()=>{
  const clipboard=nativeClipboardPort();await clipboard.write({image:{toPNG:()=>Buffer.from('png-bytes')},html:'<img data-owner="current">'});
  expect(mock.items).toHaveLength(1);expect(await clipboard.availableFormats()).toEqual(['image/png','text/html']);
  expect(await clipboard.readHTML()).toBe('<img data-owner="current">');expect((await clipboard.readImage()).toPNG().toString()).toBe('png-bytes');
})
