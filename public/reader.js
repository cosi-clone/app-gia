(() => {
  'use strict';
  const $=id=>document.getElementById(id),grid=$('reader-grid'),scroll=$('grid-scroll-container');
  let session,sheet,plan,activeSheet,prefix,columns=[],heights=[],selected={r:1,c:1},scheduled=false,fontSize=11,bold=false,italic=false;
  $('excel-select-font-family').value='Calibri';$('excel-select-font-size').value='11';
  function notice(message){$('reader-notice').textContent=message;$('reader-notice').hidden=false;clearTimeout(notice.timer);notice.timer=setTimeout(()=>{$('reader-notice').hidden=true;},4500);}
  function lowerBound(y){let lo=0,hi=plan.totalRows;while(lo<hi){const mid=(lo+hi)>>1;if(prefix[mid]<=y)lo=mid+1;else hi=mid;}return Math.max(1,Math.min(plan.totalRows,lo));}
  function buildGeometry(){
    columns=[];for(let c=1;c<=sheet.maxC;c++)if(!sheet.cols.get(c)?.hidden)columns.push(c);
    prefix=new Float64Array(plan.totalRows+1);heights=new Float64Array(plan.totalRows+1);
    const width=sheet.cols.get(plan.column)?.width||80;
    for(let r=1;r<=plan.totalRows;r++){
      const source=ExcelLayout.sourceRow(sheet,plan,r),meta=sheet.rows.get(source);
      let height=meta?.hidden?0:Math.max(22,(meta?.height||15)*96/72);
      if(height&&plan.storyAt.has(r)){
        const text=session.chunks[plan.storyAt.get(r)],perLine=Math.max(4,Math.floor((width-12)/(fontSize*.72)));
        const lines=text.split('\n').reduce((n,line)=>n+Math.max(1,Math.ceil(line.length/perLine)),0);
        height=Math.max(height,Math.ceil(lines*fontSize*96/72*1.45+12));
      }
      heights[r]=height;prefix[r]=prefix[r-1]+height;
    }
  }
  function cellElement(row,column,from,to){
    const merge=ExcelLayout.mergeAt(sheet,plan,row,column);
    let anchorRow=row,anchorCol=column,spanRows=1,spanCols=1,height=heights[row];
    if(merge){
      const visibleColumns=columns.filter(c=>c>=merge.a.c&&c<=merge.b.c);
      let first=Math.max(from,merge.a.r);while(first<=merge.b.r&&!heights[first])first++;
      if(row!==first||column!==visibleColumns[0])return null;
      anchorRow=merge.a.r;anchorCol=merge.a.c;spanCols=visibleColumns.length;spanRows=0;height=0;
      for(let r=first;r<=Math.min(to,merge.b.r);r++)if(heights[r]){spanRows++;height+=heights[r];}
    }
    const cell=ExcelLayout.cellAt(sheet,plan,session.chunks,anchorRow,anchorCol),td=document.createElement('td'),inner=document.createElement('div');
    Object.assign(td.style,session.book.styles[cell.style]||{});
    if(merge){td.rowSpan=spanRows;td.colSpan=spanCols;}
    inner.className='reader-cell-inner';inner.style.height=Math.max(1,height-1)+'px';
    inner.style.whiteSpace=td.style.whiteSpace||'pre';inner.textContent=cell.value;
    if(cell.story){td.classList.add('reader-story');td.style.verticalAlign='top';inner.style.whiteSpace='pre-wrap';}
    td.dataset.row=row;td.dataset.column=column;td.title=ExcelWorkbook.colName(column)+row;
    if(selected.r===row&&selected.c===column)td.classList.add('active-cell');
    td.append(inner);return td;
  }
  function spacer(body,height){if(height<=0)return;const tr=document.createElement('tr'),td=document.createElement('td');tr.className='reader-spacer';td.colSpan=columns.length+1;td.style.height=height+'px';tr.append(td);body.append(tr);}
  function render(){
    if(!sheet)return;
    const from=Math.max(1,lowerBound(Math.max(0,scroll.scrollTop-25))-8);
    let to=lowerBound(scroll.scrollTop+(scroll.clientHeight||650)+600);to=Math.min(plan.totalRows,Math.max(to,from+35));
    const fragment=document.createDocumentFragment(),cg=document.createElement('colgroup'),numberCol=document.createElement('col');numberCol.style.width='48px';cg.append(numberCol);
    const thead=document.createElement('thead'),header=document.createElement('tr'),corner=document.createElement('th');corner.textContent='◢';header.append(corner);
    for(const c of columns){const col=document.createElement('col');col.style.width=(sheet.cols.get(c)?.width||80)+'px';cg.append(col);const th=document.createElement('th');th.textContent=ExcelWorkbook.colName(c);if(c===selected.c)th.className='active-column';header.append(th);}
    thead.append(header);fragment.append(cg,thead);const body=document.createElement('tbody');spacer(body,prefix[from-1]);
    for(let r=from;r<=to;r++){if(!heights[r])continue;const tr=document.createElement('tr');tr.style.height=heights[r]+'px';const number=document.createElement('th');number.textContent=r;if(selected.r===r)number.className='active-row';tr.append(number);for(const c of columns){const td=cellElement(r,c,from,to);if(td)tr.append(td);}body.append(tr);}
    spacer(body,prefix[plan.totalRows]-prefix[to]);fragment.append(body);grid.replaceChildren(fragment);
    const at=lowerBound(scroll.scrollTop),done=plan.storyRows.filter(r=>r<=at).length;
    $('status-reading-info').textContent=`Dòng: ${at} | ${plan.storyRows.length?Math.round(done/plan.storyRows.length*100):0}%`;
  }
  function selectCell(row,column,reveal=false){
    selected={r:Math.max(1,Math.min(plan.totalRows,row)),c:column};
    const cell=ExcelLayout.cellAt(sheet,plan,session.chunks,selected.r,selected.c);
    $('active-cell-address').textContent=ExcelWorkbook.colName(column)+selected.r;$('formula-input').value=cell.value;
    if(reveal){const y=prefix[selected.r-1];if(y<scroll.scrollTop||y+heights[selected.r]>scroll.scrollTop+scroll.clientHeight-30)scroll.scrollTop=y;}
    render();
  }
  function activate(index){
    activeSheet=index;sheet=session.book.sheets[index];
    plan=index===session.plan.sheetIndex?session.plan:{...session.plan,column:-1,startRow:1,blockLength:sheet.maxR,totalRows:sheet.maxR,storyAt:new Map(),storyRows:[]};
    buildGeometry();scroll.scrollTop=0;scroll.scrollLeft=0;
    document.querySelectorAll('.sheet-tab').forEach((button,i)=>button.classList.toggle('active',i===index));
    $('status-ready-label').textContent='Sẵn sàng';
    const count=document.querySelector('.count-item');if(count)count.textContent='Count: '+plan.totalRows.toLocaleString('vi-VN');
    const sum=document.querySelector('.sum-item');if(sum)sum.hidden=true;
    const statusBar=document.querySelector('.excel-status-bar');if(statusBar.lastElementChild&&!statusBar.lastElementChild.id&&!statusBar.lastElementChild.classList.contains('count-item'))statusBar.lastElementChild.hidden=true;
    const row=index===session.plan.sheetIndex?session.plan.storyRows[0]:1;
    selectCell(row||1,columns.includes(plan.column)?plan.column:columns[0]);scroll.scrollTop=prefix[(row||1)-1];render();
  }
  grid.onclick=e=>{const td=e.target.closest('td[data-row]');if(td)selectCell(+td.dataset.row,+td.dataset.column);};
  scroll.onscroll=()=>{if(!scheduled){scheduled=true;requestAnimationFrame(()=>{scheduled=false;render();});}};
  scroll.onkeydown=e=>{
    if(!['ArrowDown','ArrowUp','ArrowLeft','ArrowRight'].includes(e.key))return;e.preventDefault();
    let r=selected.r,c=selected.c;if(e.key==='ArrowDown'||e.key==='ArrowUp'){const step=e.key==='ArrowDown'?1:-1;r+=step;while(r>0&&r<=plan.totalRows&&!heights[r])r+=step;}else{const i=columns.indexOf(c)+(e.key==='ArrowRight'?1:-1);c=columns[Math.max(0,Math.min(columns.length-1,i))];}
    selectCell(r,c,true);
  };
  $('reader-fullscreen').onclick=async()=>{try{if(document.fullscreenElement)await document.exitFullscreen();else if(document.documentElement.requestFullscreen)await document.documentElement.requestFullscreen();else notice('Dùng chế độ toàn màn hình của trình duyệt (Mac: Control + Command + F).');}catch{notice('Trình duyệt chưa cho bật toàn màn hình. Dùng nút toàn màn hình của trình duyệt.');}};
  document.querySelector('.win-close').onclick=()=>window.close();
  function changeFont(){const family=$('excel-select-font-family').value;document.documentElement.style.setProperty('--reader-font-family',family);document.documentElement.style.setProperty('--reader-font-size',fontSize+'pt');document.documentElement.style.setProperty('--reader-font-weight',bold?'bold':'normal');document.documentElement.style.setProperty('--reader-font-style',italic?'italic':'normal');if(sheet){const row=lowerBound(scroll.scrollTop);buildGeometry();scroll.scrollTop=prefix[row-1];render();}}
  $('excel-select-font-family').onchange=changeFont;$('excel-select-font-size').onchange=()=>{fontSize=Math.max(8,Math.min(24,+$('excel-select-font-size').value||11));changeFont();};
  $('excel-btn-font-inc').onclick=()=>{fontSize=Math.min(24,fontSize+1);$('excel-select-font-size').value=fontSize;changeFont();};$('excel-btn-font-dec').onclick=()=>{fontSize=Math.max(8,fontSize-1);$('excel-select-font-size').value=fontSize;changeFont();};
  $('excel-btn-bold').classList.remove('active-state');$('excel-btn-bold').onclick=()=>{bold=!bold;$('excel-btn-bold').classList.toggle('active-state',bold);changeFont();};$('excel-btn-italic').onclick=()=>{italic=!italic;$('excel-btn-italic').classList.toggle('active-state',italic);changeFont();};
  window.addEventListener('resize',()=>{if(sheet)render();});
  (async()=>{try{const id=location.hash.slice(1);if(!id)throw Error('Chưa có phiên đọc. Quay lại trang chuẩn bị và bấm Mở tab Excel để đọc.');session=await ReaderSessions.get(id);if(!session)throw Error('Không tìm thấy phiên đọc trong trình duyệt này. Quay lại trang chuẩn bị và mở lại.');document.title=session.book.name+' - Excel';$('excel-doc-title').textContent=session.book.name;const tabs=document.querySelector('.sheet-tabs-list');tabs.replaceChildren(...session.book.sheets.map((s,i)=>{const b=document.createElement('button');b.className='sheet-tab';b.textContent=s.name;b.onclick=()=>activate(i);return b;}));activate(session.plan.sheetIndex);}catch(e){const message=document.createElement('p');message.className='reader-error';message.textContent=e.message;scroll.replaceChildren(message);}})();
})();
