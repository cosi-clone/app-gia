// In-memory adapter for domain/service tests, never loaded by production code.
export class MemoryDB {
  data=new Map(); queue=Promise.resolve();
  collection(name){return new Query(this,name);}
  async runTransaction(fn){const run=async()=>{const ops=[];const tx={get:async ref=>ref.get(),create:(ref,value)=>ops.push(['create',ref,value]),set:(ref,value)=>ops.push(['set',ref,value]),update:(ref,value)=>ops.push(['update',ref,value])};const result=await fn(tx);this.apply(ops);return result;};const task=this.queue.then(run,run);this.queue=task.catch(()=>{});return task;}
  apply(ops){const draft=new Map(this.data);for(const [type,ref,value]of ops){if(type==='create'&&draft.has(ref.path))throw Error('already exists');if(type==='update'&&!draft.has(ref.path))throw Error('not found');draft.set(ref.path,structuredClone(type==='update'?{...draft.get(ref.path),...value}:value));}this.data=draft;}
  batch(){const ops=[];return{create:(ref,value)=>ops.push(['create',ref,value]),commit:async()=>this.apply(ops)};}
}
class Ref {
  constructor(db,path){this.db=db;this.path=path;this.id=path.split('/').at(-1);}
  async get(){const value=this.db.data.get(this.path);return {exists:value!==undefined,id:this.id,ref:this,data:()=>structuredClone(value)};}
  async create(value){this.db.apply([['create',this,value]]);}
}
class Query {
  constructor(db,name,filters=[],order=null,after=null,take=Infinity){Object.assign(this,{db,name,filters,order,after,take});}
  doc(id){return new Ref(this.db,this.name+'/'+id);}
  where(key,op,value){if(op!=='==')throw Error('unsupported operator');return new Query(this.db,this.name,[...this.filters,[key,value]],this.order,this.after,this.take);}
  orderBy(key,direction){return new Query(this.db,this.name,this.filters,[key,direction],this.after,this.take);}
  startAfter(doc){return new Query(this.db,this.name,this.filters,this.order,doc.id,this.take);}
  limit(value){return new Query(this.db,this.name,this.filters,this.order,this.after,value);}
  async get(){let items=[...this.db.data.entries()].filter(([path])=>path.startsWith(this.name+'/')&&path.split('/').length===2).map(([path,value])=>({id:path.split('/')[1],data:()=>structuredClone(value)})).filter(doc=>this.filters.every(([key,value])=>doc.data()[key]===value));if(this.order){const[key,direction]=this.order;items.sort((a,b)=>((a.data()[key]-b.data()[key])||a.id.localeCompare(b.id))*(direction==='desc'?-1:1));}if(this.after){const index=items.findIndex(doc=>doc.id===this.after);items=items.slice(index+1);}return{docs:items.slice(0,this.take)};}
  count(){return{get:async()=>{const docs=await this.get();return{data:()=>({count:docs.docs.length})};}};}
}
