(function(){
  const ACTIVITIES='soaActivities', OUTBOX='soaSyncOutbox', AUDIT='soaAuditLog';
  const read=(key,fallback)=>{try{return JSON.parse(localStorage.getItem(key))??fallback}catch{return fallback}};
  const write=(key,value)=>localStorage.setItem(key,JSON.stringify(value));
  const uid=()=>crypto.randomUUID?.()||`${Date.now()}-${Math.random().toString(16).slice(2)}`;
  class LocalProvider{
    constructor(seed){if(!localStorage.getItem(ACTIVITIES))write(ACTIVITIES,seed)}
    async listActivities(){return read(ACTIVITIES,[])}
    async addActivity(input){const items=read(ACTIVITIES,[]);const record={...input,_recordId:uid(),_updatedAt:new Date().toISOString(),_syncState:'pending'};items.unshift(record);write(ACTIVITIES,items);const outbox=read(OUTBOX,[]);outbox.push({operation:'create',entity:'activity',record,queuedAt:new Date().toISOString()});write(OUTBOX,outbox);this.audit('activity.created',record._recordId,{qcIndex:record.id});return record}
    async pendingCount(){return read(OUTBOX,[]).length}
    async auditLog(){return read(AUDIT,[])}
    audit(action,recordId,details={}){const log=read(AUDIT,[]),session=read('soaDemoSession',{});log.unshift({id:uid(),action,recordId,details,at:new Date().toISOString(),actor:session.name||'مستخدم تجريبي'});write(AUDIT,log.slice(0,500))}
  }
  /* يفعّل بعد إنشاء قوائم SharePoint وتطبيق Entra. المصادقة: Authorization Code + PKCE. */
  class SharePointProvider{
    constructor(config){this.config=config}
    async listActivities(){throw new Error('SHAREPOINT_NOT_CONFIGURED')}
    async addActivity(){throw new Error('SHAREPOINT_NOT_CONFIGURED')}
    async pendingCount(){return 0}
    async auditLog(){return []}
  }
  window.SOAData={create:seed=>window.SOA_CONFIG?.dataProvider==='sharepoint'?new SharePointProvider(window.SOA_CONFIG):new LocalProvider(seed)};
})();
