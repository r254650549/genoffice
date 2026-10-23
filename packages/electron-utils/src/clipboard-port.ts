import * as electron from 'electron'
/** Public clipboard adapter for the synchronous API and Electron's ClipboardItem API. */
export function nativeClipboardPort() {
  const native = electron.clipboard as any
  const modern = typeof native.readBuffer !== 'function'
  const raw = (format:string) => `electron application/osclipboard;format="${format}"`
  const read = async(type:string):Promise<Buffer> => {
    for(const item of await native.read()) if(item.types.includes(type)) return Buffer.from(await (await item.getType(type)).arrayBuffer())
    return Buffer.alloc(0)
  }
  return {
    readBuffer:async(format:string):Promise<Buffer> => modern ? read(raw(format)) : native.readBuffer(format),
    writeBuffer:async(format:string,value:Buffer) => {if(!modern){native.writeBuffer(format,value);return}const Item=(electron as any).ClipboardItem;await native.write([new Item({[raw(format)]:new Blob([new Uint8Array(value)])})])},
    readHTML:async():Promise<string> => modern ? (await read('text/html')).toString('utf8') : native.readHTML(),
    readText:async():Promise<string> => native.readText(),
    readImage:async() => modern ? electron.nativeImage.createFromBuffer(await read('image/png')) : native.readImage(),
    availableFormats:async():Promise<string[]> => modern ? (await native.read()).flatMap((item:any)=>item.types) : native.availableFormats(),
    write:async(value:{image:any;html:string}) => {if(!modern){native.write(value);return}const Item=(electron as any).ClipboardItem;await native.write([new Item({'image/png':new Blob([new Uint8Array(value.image.toPNG())],{type:'image/png'}),'text/html':value.html})])},
  }
}
