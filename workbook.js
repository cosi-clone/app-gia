(function(root){
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
let styles=[],theme=[];
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
    return {name:file.name,sheets,styles};
  }

root.ExcelWorkbook={readBook,colName,colNum};
})(typeof window==='undefined'?globalThis:window);
