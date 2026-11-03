import { AppError, assert } from './domain.mjs';
import { backend, adminChecker, webConfig } from './firebase.mjs';
import { createService } from './service.mjs';

async function readBody(req) {
  assert((req.headers['content-type'] || '').includes('application/json'),415,'Yêu cầu phải dùng JSON.');
  if (req.body && typeof req.body === 'object') { assert(Buffer.byteLength(JSON.stringify(req.body)) <= 16384,413,'Yêu cầu quá lớn.'); return req.body; }
  if (typeof req.body === 'string') { assert(Buffer.byteLength(req.body) <= 16384,413,'Yêu cầu quá lớn.'); try { return JSON.parse(req.body); } catch { throw new AppError(400,'JSON không hợp lệ.'); } }
  const chunks=[];let length=0;
  for await (const chunk of req) { length+=chunk.length;assert(length<=16384,413,'Yêu cầu quá lớn.');chunks.push(chunk); }
  try { return JSON.parse(Buffer.concat(chunks).toString() || '{}'); } catch { throw new AppError(400,'JSON không hợp lệ.'); }
}
export function makeHandler({ getBackend = backend, getWebConfig = webConfig, isAdmin = adminChecker() } = {}) {
  return async function handler(req,res) {
    const send=(status,data)=>{res.statusCode=status;res.setHeader('Content-Type','application/json; charset=utf-8');res.setHeader('Cache-Control','no-store');res.setHeader('X-Content-Type-Options','nosniff');res.end(JSON.stringify(data));};
    try {
      const url=new URL(req.url,'http://localhost'),route=url.searchParams.get('route')||'';
      if(route==='config'){assert(req.method==='GET',405,'Method không được hỗ trợ.');return send(200,getWebConfig());}
      const allowed={me:'GET',access:'GET',redeem:'POST','admin/summary':'GET','admin/codes':'GET','admin/create-codes':'POST','admin/revoke-code':'POST','admin/members':'GET','admin/create-member':'POST','admin/block-member':'POST'};
      assert(allowed[route],404,'Không tìm thấy API.');assert(req.method===allowed[route],405,'Method không được hỗ trợ.');
      const header=req.headers.authorization||'';assert(header.startsWith('Bearer '),401,'Vui lòng đăng nhập.','LOGIN_REQUIRED');
      const dependencies=getBackend();let actor;
      try{actor=await dependencies.auth.verifyIdToken(header.slice(7),true);}catch{throw new AppError(401,'Phiên đăng nhập hết hạn hoặc tài khoản đã bị vô hiệu hóa. Đăng nhập lại.','LOGIN_REQUIRED');}
      const service=createService({...dependencies,isAdmin});
      let result;
      if(req.method==='POST'){
        const input=await readBody(req);
        assert(input && !Array.isArray(input) && typeof input==='object',400,'JSON phải là một object.');
        const methods={redeem:'redeem','admin/create-codes':'createCodes','admin/revoke-code':'revokeCode','admin/create-member':'createMember','admin/block-member':'setBlocked'};
        result=await service[methods[route]](actor,input);
      }else{
        const methods={me:'me',access:'requireActive','admin/summary':'summary','admin/codes':'listCodes','admin/members':'listMembers'};
        result=await service[methods[route]](actor,Object.fromEntries(url.searchParams));
      }
      send(200,result);
    }catch(error){
      if(error instanceof AppError)return send(error.status,{error:error.message,code:error.code});
      console.error('[api]',error.code||error.name,error.message);
      const message=error.code===9?'Firestore đang thiếu index. Hãy deploy firestore.indexes.json.':'Không kết nối được dịch vụ. Kiểm tra cấu hình Firebase hoặc thử lại.';
      send(503,{error:message,code:'SERVICE_UNAVAILABLE'});
    }
  };
}
export default makeHandler();
