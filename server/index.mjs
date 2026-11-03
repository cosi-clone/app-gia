import 'dotenv/config';
import { createServer } from 'node:http';
import { createReadStream } from 'node:fs';
import { stat } from 'node:fs/promises';
import { dirname, extname, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import handler from './handler.mjs';

export function createApp() {
  const root=resolve(dirname(fileURLToPath(import.meta.url)),'../public');
  const mime={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.json':'application/json','.svg':'image/svg+xml','.png':'image/png','.ico':'image/x-icon'};
  return createServer(async(req,res)=>{
    let path;
    try{path=decodeURIComponent(new URL(req.url,'http://localhost').pathname);}catch{res.statusCode=400;return res.end('Bad URL');}
    if(path==='/api/service')return handler(req,res);
    if(!['GET','HEAD'].includes(req.method)){res.statusCode=405;return res.end();}
    const file=resolve(root,'.'+(path==='/'?'/index.html':path));
    if(!file.startsWith(root+sep)){res.statusCode=403;return res.end();}
    try{const info=await stat(file);if(!info.isFile())throw Error();res.setHeader('Content-Type',mime[extname(file)]||'application/octet-stream');res.setHeader('X-Content-Type-Options','nosniff');res.setHeader('Cache-Control','no-store');if(req.method==='HEAD')return res.end();createReadStream(file).on('error',()=>res.destroy()).pipe(res);}catch{res.statusCode=404;res.end('Not found');}
  });
}
if(process.argv[1]&&resolve(process.argv[1])===fileURLToPath(import.meta.url)){
  const port=Number(process.env.PORT)||3000;createApp().listen(port,'127.0.0.1',()=>console.log(`Excel Reader: http://localhost:${port}\nCấu hình Firebase trong .env theo SETUP_FIREBASE.md`));
}
