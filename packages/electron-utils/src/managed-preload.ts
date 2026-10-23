import {contextBridge,ipcRenderer} from 'electron';
// Only an embedding process that adds this preload receives these capabilities.
contextBridge.exposeInMainWorld('kompagentOffice',{
 event:(value:unknown)=>ipcRenderer.send('kompagent:office:event',value),
 request:(value:unknown)=>ipcRenderer.invoke('kompagent:office:request',value),
 reply:(value:unknown)=>ipcRenderer.send('kompagent:office:reply',value),
 onControl:(callback:(value:unknown)=>void)=>{const listener=(_event:unknown,value:unknown)=>callback(value);ipcRenderer.on('kompagent:office:control',listener);return()=>ipcRenderer.removeListener('kompagent:office:control',listener);},
});
