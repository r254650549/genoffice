import {DEFAULT_AI_PANEL_PREFS} from '../../ui/src/ai-panel-prefs';
/** Explicit embedding capability. Native editing/file APIs remain native; the host owns documents and Agent. */
export function managedOfficeHost<T extends object>(original:T):T & {hostManagedAgent?:boolean} {
 const bridge=(window as unknown as {kompagentOffice?:{event(value:unknown):void;request(value:unknown):Promise<any>;onControl(callback:(value:any)=>void):()=>void;reply(value:unknown):void}}).kompagentOffice;
 if(!bridge)return original;
 const native=original as any;
 let controls:any,snapshot:(()=>Promise<Uint8Array|null>)|undefined,dirty=false,revision=0,loading=true;
 let stopObserving:(()=>void)|undefined,externalState:readonly unknown[]|undefined;
 const changed=()=>{if(loading)return;dirty=true;revision++;bridge.event({type:'changed',revision});};
 const wrapped:any={...original,hostManagedAgent:true,
  getAutoSaveDefault:async()=>false,onAutoSaveDefaultChanged:()=>()=>{},
  getAiPanelPrefs:async()=>DEFAULT_AI_PANEL_PREFS,onAiPanelPrefsChanged:()=>()=>{},
  history:(command:'undo'|'redo')=>controls?.[command]?.(),
  openAgent:(prompt?:string)=>bridge.event({type:'agent',prompt}),
  bindDocument:(get:()=>Promise<Uint8Array|null>,subscribe:(callback:()=>void)=>()=>void)=>{snapshot=get;stopObserving?.();stopObserving=subscribe(changed);loading=false;},
  documentStateChanged:(values:readonly unknown[],saving:boolean)=>{const old=externalState;externalState=values;if(!saving&&old&&!values.every((v,i)=>Object.is(v,old[i])))changed();},
  documentChanged:changed,workbookLoading:(value:boolean)=>{loading=value;},
  bindWorkbook:(value:any)=>{controls=value;},
  bindPresentation:(value:any)=>{controls=value;loading=false;bridge.event({type:'ready'});},
  beforeMcpCommand:async(message:any)=>{if(controls?.commit&&!await controls.commit())throw Error('OFFICE_EDIT_NOT_COMMITTED');if(message.expectedRevision!==undefined&&message.expectedRevision!==revision)throw Error('OFFICE_REVISION_CONFLICT');document.body.inert=true;},
  reportMcpResult:async(result:any)=>{try{return await native.reportMcpResult({...result,revision});}finally{document.body.inert=false;}},
  signalMcpReady:()=>{loading=false;native.signalMcpReady?.();bridge.event({type:'ready'});},
 };
 const saveTarget=(asCopy=false)=>bridge.request({action:'saveTarget',asCopy});
 const completed=async(path:string,result:any)=>{if(result?.ok===false||result?.canceled===true||result===null)return result;await bridge.request({action:'saved',path,revision});dirty=false;return result;};
 if(native.saveDocxTo){
  const save=async(data:ArrayBuffer,asCopy=false)=>{const path=await saveTarget(asCopy);if(!path)return {ok:false};const result=await native.saveDocxTo(path,data,true);return completed(path,result);};
  wrapped.saveDocx=(_path:string,data:ArrayBuffer)=>save(data);
  wrapped.saveDocxNew=(_name:string,data:ArrayBuffer)=>save(data);
  wrapped.saveDocxAs=(_name:string,data:ArrayBuffer)=>save(data,true);
  wrapped.saveDocxTo=(_path:string,data:ArrayBuffer)=>save(data);
 }
 if(native.saveWorkbookEdits)wrapped.saveWorkbookEdits=async(request:any)=>{const path=await saveTarget(request.mode==='save-as'&&!request.targetPath);if(!path)return {canceled:true};return completed(path,await native.saveWorkbookEdits({...request,mode:'save-as',targetPath:path,overwrite:true}));};
 if(native.save&&native.saveAs){wrapped.save=()=>bridge.request({action:'savePresentation'});wrapped.saveAs=()=>bridge.request({action:'savePresentation',asCopy:true});}
 for(const name of ['openNewTab','openSourcePath','createDocument','openDocx','openDocxPath','selectWorkbook','openPptx'])if(native[name])wrapped[name]=async(...args:unknown[])=>loading?native[name](...args):bridge.request({action:'navigate',method:name,args});
 bridge.onControl(async request=>{try{let value:unknown;
  if(request.method==='state')value={dirty:controls?.dirty?.()??dirty,revision,ready:!!controls||!!snapshot};
  else if(request.method==='flush'){if(controls?.commit&&!await controls.commit())throw Error('OFFICE_EDIT_NOT_COMMITTED');if(controls?.flush&&(!controls.dirty||controls.dirty())&&await controls.flush()===false)throw Error('OFFICE_DRAFT_NOT_PERSISTED');value=snapshot?await snapshot():null;}
  else if(request.method==='save'){if(!controls?.save)throw Error('HOST_DOCUMENT_SAVE_REQUIRED');value=await controls.save();if(value)dirty=false;}
  else if(request.method==='saved'){dirty=false;value=true;}
  else if(request.method==='pause'){document.body.inert=true;value=revision;}
  else if(request.method==='resume'){document.body.inert=false;value=true;}
  else throw Error('UNKNOWN_OFFICE_CONTROL');
  bridge.reply({id:request.id,value});
 }catch(error){bridge.reply({id:request.id,error:error instanceof Error?error.message:String(error)});}});
 return wrapped;
}
