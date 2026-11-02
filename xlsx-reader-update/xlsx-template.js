/* Excel-only reader: XLSX template view and story insertion. */
(() => {
  'use strict';
  const M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const children = (e, tag) => e ? [...e.children].filter(x => x.localName === tag) : [];
  const first = (e, tag) => children(e, tag)[0];
  const txt = e => e ? [...e.getElementsByTagNameNS(M, 't')].map(t => t.textContent).join('') : '';
  const xml = s => { const d = new DOMParser().parseFromString(s, 'application/xml'); if (d.querySelector('parsererror')) throw Error('XML không hợp lệ trong XLSX.'); return d; };
  const colNum = s => [...s.toUpperCase()].reduce((n,c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const colName = n => { let s = ''; for (; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };
  const addr = s => { const m = /^([A-Z]+)(\d+)$/i.exec(s); if (!m) throw Error('Địa chỉ ô không hợp lệ.'); return {c:colNum(m[1]), r:+m[2]}; };
  const range = s => { const a = s.split(':'); return {a:addr(a[0]), b:addr(a[1] || a[0])}; };
  const node = (tag, text) => { const e = document.createElement(tag); if (text !== undefined) e.textContent = text; return e; };
  let book = null, sheet = null, page = 0, overlay = new Map(), inserted = false;
  let styles = [], theme = [];
  const pageSize = 100;
  const panel = node('section'); panel.id = 'xlsx-template-panel'; panel.hidden = false;
  panel.innerHTML = `<div class="xt-controls"><strong>Excel Reader</strong><label>1. PDF / TXT / EPUB <input id="xt-story" type="file" accept=".pdf,.txt,.epub"></label><label>Hoặc link PDF <input id="xt-url" type="url" placeholder="https://…/sach.pdf" aria-label="Link PDF online"></label><button id="xt-load-url">Tải PDF từ link</button><label>2. Template XLSX <input id="xt-file" type="file" accept=".xlsx"></label><label>Sheet <select id="xt-sheet"></select></label><label>Cột <select id="xt-column"></select></label><label>Dòng đầu <input id="xt-start" type="number" value="6" min="1" max="1048576"></label><button id="xt-insert">Chèn truyện</button><button id="xt-reset">Khôi phục template</button></div><p id="xt-status" role="status">Nạp PDF/TXT/EPUB và XLSX để bắt đầu. Chọn sheet, cột và dòng chèn; file gốc không đổi.</p><div class="xt-scroll" id="xt-scroll"></div><div class="xt-footer"><button id="xt-prev">← Trước</button><span id="xt-pages"></span><button id="xt-next">Sau →</button></div>`;
  document.body.append(panel);
  const $ = id => document.getElementById('xt-' + id);
  const status = s => { $('status').textContent = s; };
  const setBusy = busy => { for (const id of ['story','file','sheet','column','start','insert','reset','url','load-url']) $(id).disabled = busy; };
  $('scroll').append(node('div','Chọn file .xlsx của bạn ở thanh công cụ. Nội dung và định dạng của file sẽ hiện tại đây.'));
  function color(e) {
    if (!e) return '';
    if (e.hasAttribute('rgb')) return '#' + e.getAttribute('rgb').slice(-6);
    if (e.hasAttribute('theme')) return theme[+e.getAttribute('theme')] || '';
    const indexed = {0:'#000000',1:'#ffffff',2:'#ff0000',3:'#00ff00',4:'#0000ff',5:'#ffff00',6:'#ff00ff',7:'#00ffff',64:'#000000',65:'#ffffff'};
    return indexed[e.getAttribute('indexed')] || '';
  }
  function loadStyles(d) {
    const root = d.documentElement, fonts = children(first(root,'fonts'),'font'), fills = children(first(root,'fills'),'fill'), borders = children(first(root,'borders'),'border');
    return children(first(root,'cellXfs'),'xf').map(x => {
      const f = fonts[+x.getAttribute('fontId')], fill = fills[+x.getAttribute('fillId')], border = borders[+x.getAttribute('borderId')], a = first(x,'alignment');
      const css = {};
      if (f) { css.fontFamily = first(f,'name')?.getAttribute('val') || 'Calibri'; css.fontSize = (first(f,'sz')?.getAttribute('val') || 11) + 'pt'; css.fontWeight = first(f,'b') && first(f,'b').getAttribute('val') !== '0' ? 'bold' : 'normal'; css.fontStyle = first(f,'i') ? 'italic' : 'normal'; css.color = color(first(f,'color')); if(first(f,'u')) css.textDecoration = 'underline'; }
      const pattern = first(fill,'patternFill'); if (pattern?.getAttribute('patternType') === 'solid') css.backgroundColor = color(first(pattern,'fgColor'));
      if (a) { const h = a.getAttribute('horizontal'); if(['left','right','center','justify'].includes(h)) css.textAlign = h; const v = a.getAttribute('vertical'); css.verticalAlign = v === 'center' ? 'middle' : v || 'bottom'; if(a.getAttribute('wrapText') === '1') css.whiteSpace = 'pre-wrap'; }
      for(const side of ['left','right','top','bottom']) { const b = first(border,side), s = b?.getAttribute('style'); if (s) css['border' + side[0].toUpperCase() + side.slice(1)] = `${/medium|thick/.test(s) ? 2 : 1}px ${/dash/i.test(s) ? 'dashed' : s === 'dotted' ? 'dotted' : 'solid'} ${color(first(b,'color')) || '#222'}`; }
      return css;
    });
  }
  async function readBook(file) {
    if (file.size > 50 * 1024 * 1024) throw Error('Template tối đa 50 MB.');
    const Zip = await ensureJsZipLoaded(), z = await Zip.loadAsync(await file.arrayBuffer());
    const read = async p => { const f = z.file(p); if (!f) throw Error('Thiếu thành phần XLSX: ' + p); return xml(await f.async('string')); };
    const wb = await read('xl/workbook.xml'), rels = await read('xl/_rels/workbook.xml.rels');
    const links = new Map([...rels.documentElement.children].map(r => [r.getAttribute('Id'), r.getAttribute('Target')]));
    const strings = z.file('xl/sharedStrings.xml') ? children((await read('xl/sharedStrings.xml')).documentElement,'si').map(txt) : [];
    theme = [];
    if(z.file('xl/theme/theme1.xml')) { const td = await read('xl/theme/theme1.xml'), cs = [...td.getElementsByTagName('*')].find(e => e.localName === 'clrScheme'); const order = ['lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink']; theme = order.map(k => { const c = first(cs,k)?.firstElementChild; return c ? '#' + (c.getAttribute('lastClr') || c.getAttribute('val')) : ''; }); }
    styles = z.file('xl/styles.xml') ? loadStyles(await read('xl/styles.xml')) : [];
    const sheets = [];
    for (const s of children(first(wb.documentElement,'sheets'),'sheet')) {
      const target = links.get(s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'));
      if(!target) continue;
      const path = new URL(target, 'https://xlsx.local/xl/workbook.xml').pathname.slice(1);
      const d = (await read(path)).documentElement, cells = new Map(), rows = new Map(), cols = new Map(); let maxR = 1, maxC = 1;
      for(const row of children(first(d,'sheetData'),'row')) { rows.set(+row.getAttribute('r'), {height:+row.getAttribute('ht') || 15, hidden:row.getAttribute('hidden') === '1'}); for(const c of children(row,'c')) { const ref=c.getAttribute('r'), p=addr(ref), v=first(c,'v')?.textContent || '', t=c.getAttribute('t'); const value=t === 's' ? strings[+v] || '' : t === 'inlineStr' ? txt(first(c,'is')) : t === 'b' ? (v === '1' ? 'TRUE' : 'FALSE') : v; cells.set(ref,{value,style:+c.getAttribute('s') || 0}); maxR=Math.max(maxR,p.r); maxC=Math.max(maxC,p.c); } }
      for(const c of children(first(d,'cols'),'col')) for(let i=+c.getAttribute('min');i<=+c.getAttribute('max');i++) cols.set(i,{width:Math.max(20,(+c.getAttribute('width') || 9)*7+5),hidden:c.getAttribute('hidden')==='1',style:+c.getAttribute('style')||0});
      const merges=children(first(d,'mergeCells'),'mergeCell').map(m=>range(m.getAttribute('ref')));
      for(const m of merges) {maxR=Math.max(maxR,m.b.r);maxC=Math.max(maxC,m.b.c);}
      if(maxC>500 || maxR>100000) throw Error('Template vượt giới hạn hiển thị 500 cột / 100.000 dòng.');
      let headerRow=1,best=-1;for(let r=1;r<=Math.min(maxR,20);r++){let score=0;for(let c=1;c<=maxC;c++){const v=cells.get(colName(c)+r)?.value||'';if(v && !/^[-\d.,% ]+$/.test(v))score++;}if(score>best){best=score;headerRow=r;}}sheets.push({name:s.getAttribute('name'),cells,rows,cols,merges,maxR,maxC,baseMaxR:maxR,headerRow});
    }
    if(!sheets.length) throw Error('Không có sheet đọc được.');
    return {name:file.name,sheets};
  }
  function clearExtension(){if(!sheet)return;for(let r=sheet.baseMaxR+1;r<=sheet.maxR;r++){sheet.rows.delete(r);for(let c=1;c<=sheet.maxC;c++)sheet.cells.delete(colName(c)+r);}sheet.maxR=sheet.baseMaxR;}
  function selectSheet() {
    clearExtension();
    sheet=book.sheets[+$('sheet').value]; page=0; overlay=new Map();inserted=false;
    $('column').replaceChildren(); for(let c=1;c<=sheet.maxC;c++) { const name=colName(c), title=sheet.cells.get(name+sheet.headerRow)?.value || ''; const o=node('option',name+(title?' — '+title:''));o.value=c;$('column').append(o); }
    $('column').value=String(Math.min(4,sheet.maxC)); $('start').value=String(Math.min(sheet.maxR,sheet.headerRow+1)); render(); status(`${book.name} • ${sheet.name} • ${sheet.maxR} dòng, ${sheet.maxC} cột. Chọn cột và dòng bắt đầu; nhãn cột tham khảo từ dòng ${sheet.headerRow} (tự nhận diện).`);
  }
  function render() {
    if(!sheet)return;
    const start=page*pageSize+1, end=Math.min(sheet.maxR,start+pageSize-1), table=node('table'), cg=node('colgroup'), header=node('tr');cg.append(node('col'));header.append(node('th',''));
    for(let c=1;c<=sheet.maxC;c++){if(sheet.cols.get(c)?.hidden)continue; const col=node('col');col.style.width=(sheet.cols.get(c)?.width||80)+'px';cg.append(col); const h=node('th',colName(c));h.onclick=()=>{$('column').value=String(c);};header.append(h);}
    const head=node('thead');head.append(header);table.append(cg,head);const body=node('tbody');
    for(let r=start;r<=end;r++) {if(sheet.rows.get(r)?.hidden)continue;const tr=node('tr');tr.style.height=(sheet.rows.get(r)?.height||15)+'pt';tr.append(node('th',r));
      for(let c=1;c<=sheet.maxC;c++){if(sheet.cols.get(c)?.hidden)continue;const merge=sheet.merges.find(m=>r>=m.a.r&&r<=m.b.r&&c>=m.a.c&&c<=m.b.c);if(merge&&(r!==Math.max(start,merge.a.r)||c!==merge.a.c))continue;
        const ref=colName(c)+r, sourceRef=merge?colName(merge.a.c)+merge.a.r:ref, cell=sheet.cells.get(sourceRef), td=node('td');
        Object.assign(td.style,styles[cell?.style ?? sheet.cols.get(c)?.style ?? 0]||{});td.textContent=overlay.has(sourceRef)?overlay.get(sourceRef):cell?.value||'';
        if(merge){td.rowSpan=Math.min(end,merge.b.r)-Math.max(start,merge.a.r)+1;td.colSpan=Array.from({length:merge.b.c-merge.a.c+1},(_,i)=>merge.a.c+i).filter(i=>!sheet.cols.get(i)?.hidden).length;}
        if(overlay.has(sourceRef)){td.classList.add('xt-story-cell');td.style.whiteSpace='pre-wrap';td.style.verticalAlign='top';}
        td.title=sourceRef;tr.append(td);
      }body.append(tr);
    }table.append(body);$('scroll').replaceChildren(table);$('pages').textContent=`Dòng ${start}–${end} / ${sheet.maxR} • Trang ${page+1}/${Math.ceil(sheet.maxR/pageSize)}`;$('prev').disabled=page===0;$('next').disabled=end>=sheet.maxR;
  }
  $('file').onchange=async e=>{const f=e.target.files[0];if(!f)return;status('Đang đọc template…');setBusy(true);try{const incoming=await readBook(f);clearExtension();book=incoming;$('sheet').replaceChildren(...book.sheets.map((s,i)=>{const o=node('option',s.name);o.value=i;return o;}));$('sheet').value='0';selectSheet();}catch(e){status(e.message);}finally{e.target.value='';setBusy(false);}};
  $('sheet').onchange=selectSheet;
  function documentLoaded(){clearExtension();overlay=new Map();page=0;render();status(`Đã nạp ${state.fileName}: ${state.allChunks.length} đoạn. Chọn XLSX và vị trí rồi nhấn Chèn truyện.`);}
  function documentFailed(error){status(error.message+(state.allChunks.length?' Nội dung truyện đã nạp trước vẫn được giữ.':''));}
  $('story').onchange=async e=>{const f=e.target.files[0];if(!f)return;setBusy(true);status('Đang nạp truyện…');try{await processDocumentFile(f,status);documentLoaded();}catch(e){documentFailed(e);}finally{e.target.value='';setBusy(false);}};
  $('load-url').onclick=async()=>{if(!$('url').value.trim())return status('Dán link PDF trước.');setBusy(true);try{await processPdfUrl($('url').value,status);documentLoaded();}catch(e){documentFailed(e);}finally{setBusy(false);}};
  $('url').addEventListener('keydown',e=>{if(e.key==='Enter'&&!$('load-url').disabled){e.preventDefault();$('load-url').click();}});
  $('insert').onclick=()=>{
    if(!sheet)return status('Chọn file XLSX trước.');
    const chunks=state.allChunks.filter(c=>c.text&&!c.text.includes('[[STEALTH_IMG:')).map(c=>c.text);if(!chunks.length)return status('Nạp PDF/TXT/EPUB có nội dung văn bản trước.');
    const c=+$('column').value,start=+$('start').value;if(!Number.isInteger(start)||start<1||start>sheet.maxR)return status(`Dòng đầu phải nằm trong 1–${sheet.maxR}.`);
    clearExtension();const targets=[];let skipped=0;for(let r=start;r<=sheet.maxR;r++){if(sheet.rows.get(r)?.hidden)continue;const m=sheet.merges.find(m=>r>=m.a.r&&r<=m.b.r&&c>=m.a.c&&c<=m.b.c);if(m){skipped++;continue;}targets.push(colName(c)+r);}
    if(chunks.length>targets.length){const count=chunks.length-targets.length;if(sheet.maxR+count>100000)return status('Truyện quá dài: vượt giới hạn 100.000 dòng.');const oldEnd=sheet.maxR;for(let i=1;i<=count;i++){const r=oldEnd+i,ref=colName(c)+r;targets.push(ref);const previous=sheet.cells.get(colName(c)+oldEnd);if(previous)sheet.cells.set(ref,{value:'',style:previous.style});sheet.rows.set(r,{height:sheet.rows.get(oldEnd)?.height||15,hidden:false});}sheet.maxR+=count;}overlay=new Map(targets.slice(0,chunks.length).map((ref,i)=>[ref,chunks[i]]));inserted=true;page=Math.floor((start-1)/pageSize);render();status(`Đã chèn ${overlay.size}/${chunks.length} đoạn vào ${sheet.name}, cột ${colName(c)}. Bỏ qua ${skipped} ô nằm trong vùng gộp. Có thể khôi phục template bất cứ lúc nào. File gốc không đổi.`);
  };
  $('reset').onclick=()=>{clearExtension();page=0;overlay=new Map();inserted=false;render();status('Đã khôi phục nội dung template gốc.');};
  $('prev').onclick=()=>{if(page>0){page--;render();$('scroll').scrollTop=0;}};
  $('next').onclick=()=>{if(sheet&&(page+1)*pageSize<sheet.maxR){page++;render();$('scroll').scrollTop=0;}};
})();
