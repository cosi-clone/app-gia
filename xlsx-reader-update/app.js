'use strict';
const state = { allChunks: [], fileName: '', isPdfProcessing: false };
function ensureJsZipLoaded() {
  if (!window.JSZip) throw new Error('Thiếu static/jszip.min.js. Hãy giải nén đầy đủ gói mã.');
  return Promise.resolve(window.JSZip);
}
let pdfPromise;
function ensurePdfLibrary() {
  if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  if (!pdfPromise) pdfPromise = new Promise((resolve,reject) => {
    const s = document.createElement('script'); s.src = 'static/pdf.min.js';
    s.onload = () => resolve(window.pdfjsLib);
    s.onerror = () => { pdfPromise = null; s.remove(); reject(Error('Không tải được thư viện PDF. Kiểm tra thư mục static.')); };
    document.head.append(s);
  });
  return pdfPromise;
}
function splitStory(text) {
  const chunks = [];
  for (const paragraph of text.replace(/\r\n?/g,'\n').split(/\n\s*\n|\n/)) {
    let rest = paragraph.replace(/[\t ]+/g,' ').trim();
    while (rest.length > 500) {
      const boundary = rest.lastIndexOf(' ',500), at = boundary > 200 ? boundary : 500;
      chunks.push(rest.slice(0,at)); rest = rest.slice(at).trim();
    }
    if (rest) chunks.push(rest);
  }
  return chunks;
}
async function processPdfUrl(value, progress = () => {}) {
  let url;
  try { url = new URL(value.trim()); } catch { throw Error('Link không hợp lệ. Dán URL đầy đủ bắt đầu bằng https:// hoặc http://.'); }
  if (!['http:','https:'].includes(url.protocol) || url.username || url.password) throw Error('Chỉ hỗ trợ link HTTP/HTTPS không chứa thông tin đăng nhập.');
  const limit = 100 * 1024 * 1024;
  async function download(target, milliseconds) {
    const controller = new AbortController(), timer = setTimeout(()=>controller.abort(), milliseconds);
    try {
      const response = await fetch(target, {signal:controller.signal,credentials:'omit',referrerPolicy:'no-referrer'});
      if(!response.ok) {
        if (target.startsWith('/api/')) {
          const error = await response.json().catch(()=>({}));
          throw Error(error.error || `API tải PDF trả HTTP ${response.status}.`);
        }
        throw Error(`Nguồn PDF trả HTTP ${response.status}.`);
      }
      if(Number(response.headers.get('content-length')) > limit) throw Error('PDF vượt giới hạn 100 MB.');
      const reader = response.body?.getReader();
      let bytes;
      if (reader) {
        const chunks = []; let size = 0;
        try {
          while(true) { const {done,value} = await reader.read(); if(done)break; size += value.byteLength; if(size>limit) {controller.abort();throw Error('PDF vượt giới hạn 100 MB.');} chunks.push(value);progress(`Đang tải PDF: ${(size/1048576).toFixed(1)} MB…`); }
        } finally { reader.releaseLock(); }
        bytes = new Uint8Array(size); let offset = 0;for(const chunk of chunks){bytes.set(chunk,offset);offset+=chunk.byteLength;}
      } else { bytes = new Uint8Array(await response.arrayBuffer()); if(bytes.byteLength>limit)throw Error('PDF vượt giới hạn 100 MB.'); }
      if(!Array.from(bytes.slice(0,1024), b=>String.fromCharCode(b)).join('').includes('%PDF-')) throw Error('Link không trả về file PDF; có thể là trang web, trang đăng nhập hoặc link hết hạn.');
      return bytes;
    } finally { clearTimeout(timer); }
  }
  progress('Đang tải PDF từ link…');
  let bytes;
  try { bytes = await download(url.href,15000); }
  catch (directError) {
    progress('Không tải trực tiếp được; đang thử qua API…');
    try { bytes = await download('/api/pdf?url='+encodeURIComponent(url.href),100000); }
    catch(error) {throw Error(`Không tải được PDF: ${error.message} Có thể tải PDF về máy rồi dùng chức năng upload PDF.`);}
  }
  let name;
  try {name=decodeURIComponent(url.pathname.split('/').pop() || 'online.pdf');}catch{name='online.pdf';}
  if(!name.toLowerCase().endsWith('.pdf')) name += '.pdf';
  const file = new File([bytes],name,{type:'application/pdf'});
  await processDocumentFile(file,progress);
}
async function processDocumentFile(file, progress = () => {}) {
  if (file.size > 100 * 1024 * 1024) throw Error('Tài liệu tối đa 100 MB.');
  const extension = file.name.split('.').pop().toLowerCase();
  if (!['pdf','txt','epub'].includes(extension)) throw Error('Chọn PDF, TXT hoặc EPUB.');
  const result = []; state.isPdfProcessing = true;
  try {
    if (extension === 'txt') {
      result.push(...splitStory(await file.text()).map(text => ({text,page:1})));
    } else if (extension === 'pdf') {
      const pdfjs = await ensurePdfLibrary();
      pdfjs.GlobalWorkerOptions.workerSrc = 'static/pdf.worker.min.js';
      const task = pdfjs.getDocument({ data:await file.arrayBuffer(), isEvalSupported:false });
      let pdf;
      try {
        pdf = await task.promise;
        for (let pageNumber=1; pageNumber<=pdf.numPages; pageNumber++) {
          progress(`Đang đọc PDF: trang ${pageNumber}/${pdf.numPages}…`);
          const page = await pdf.getPage(pageNumber), content = await page.getTextContent();
          let text = '', line = '', lastY = null;
          for (const item of content.items) {
            if (!('str' in item)) continue;
            const y = item.transform?.[5];
            if (lastY !== null && Math.abs(y-lastY)>3 && line) {text += line+'\n';line='';}
            line += item.str + ' '; lastY = y;
            if(item.hasEOL){text += line+'\n';line='';lastY=null;}
          }
          text += line;
          result.push(...splitStory(text).map(text => ({text,page:pageNumber})));
          page.cleanup();
          await new Promise(resolve => setTimeout(resolve,0));
        }
      } finally { await task.destroy(); }
    } else {
      const Zip = await ensureJsZipLoaded(), zip = await Zip.loadAsync(await file.arrayBuffer());
      const parse = text => new DOMParser().parseFromString(text,'application/xml');
      const container = zip.file('META-INF/container.xml');
      if(!container) throw Error('EPUB không có container.xml.');
      const root = parse(await container.async('string')).getElementsByTagName('rootfile')[0];
      const path = root?.getAttribute('full-path'); if(!path || !zip.file(path))throw Error('Không tìm thấy cấu trúc EPUB.');
      const opf = parse(await zip.file(path).async('string'));
      const items = new Map([...opf.getElementsByTagName('item')].map(i=>[i.getAttribute('id'),i.getAttribute('href')]));
      let page = 0;
      for (const item of opf.getElementsByTagName('itemref')) {
        const href = items.get(item.getAttribute('idref')); if(!href)continue;
        const target = decodeURIComponent(new URL(href,'https://local/'+path).pathname.slice(1));
        const entry = zip.file(target);if(!entry)continue;
        const doc = new DOMParser().parseFromString(await entry.async('string'),'text/html');
        doc.querySelectorAll('script,style,nav').forEach(e=>e.remove());
        const paragraphs = [...doc.querySelectorAll('p,h1,h2,h3,h4,li,pre')].filter(e=>!e.querySelector('p,li,pre')).map(e=>e.textContent).join('\n');
        result.push(...splitStory(paragraphs || doc.body.textContent).map(text=>({text,page:++page})));
      }
    }
    if (!result.length) throw Error('Không tìm thấy văn bản. PDF dạng ảnh/scan cần OCR trước.');
    state.allChunks = result; state.fileName = file.name;
  } finally { state.isPdfProcessing = false; }
}
