(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const grid = $('reader-grid'), scroll = $('grid-scroll-container');
  const formula = $('formula-input');
  let session, sheet, plan, activeSheet, state;
  let columns = [], prefix, heights, scheduled = false;
  let selected = { r: 1, c: 1 }, editor = null, draft = null, dragging = null;
  let fontSize = 11, bold = false, italic = false;
  let showStory = true;
  let rangeEnd = null;
  let revision = 0, savedRevision = 0, saveTimer, saveQueue = Promise.resolve();
  const undoStack = [], redoStack = [];
  const key = (r, c) => ExcelWorkbook.colName(c) + r;
  const clamp = (value, min, max) => Math.max(min, Math.min(max, value));

  function notice(message) {
    $('reader-notice').textContent = message;
    $('reader-notice').hidden = false;
    clearTimeout(notice.timer);
    notice.timer = setTimeout(() => { $('reader-notice').hidden = true; }, 4000);
  }
  function dirty() {
    revision++;
    $('status-ready-label').textContent = 'Đang lưu…';
    clearTimeout(saveTimer);
    saveTimer = setTimeout(persist, 500);
  }
  function persist() {
    clearTimeout(saveTimer);
    if (!session || revision === savedRevision) return saveQueue;
    const version = revision, snapshot = structuredClone(session);
    saveQueue = saveQueue.catch(() => {}).then(async () => {
      if (session.ownerUid && session.ownerUid !== window.Access?.uid) {
        throw Error('Tài khoản đã thay đổi. Không lưu phiên của tài khoản khác.');
      }
      await ReaderSessions.save(snapshot);
      savedRevision = Math.max(savedRevision, version);
      if (revision === version) $('status-ready-label').textContent = 'Đã lưu';
    }).catch(error => {
      $('status-ready-label').textContent = 'Chưa lưu';
      notice(error.message + ' Bấm Ctrl/Cmd+S để thử lưu lại.');
    });
    return saveQueue;
  }
  function sheetState(index) {
    const value = session.editor.sheets[index] ||= { columns: {}, rows: {}, cells: {} };
    value.originalCells ||= {};
    value.borders ||= {};
    return value;
  }
  function canonical(r, c) {
    const merge = ExcelLayout.mergeAt(sheet, plan, r, c);
    return merge ? { r: merge.a.r, c: merge.a.c } : { r, c };
  }
  function cellAt(r, c) {
    const cell = ExcelLayout.cellAt(sheet, showStory ? plan : { ...plan, column: -1 }, session.chunks, r, c);
    const source=ExcelLayout.sourceRow(sheet,plan,r);
    if(!sheet.cells.has(key(source,c)))cell.style=sheet.rows.get(source)?.style??sheet.cols.get(c)?.style??0;
    const ref = key(r, c);
    const edits = showStory ? state.cells : state.originalCells;
    return Object.hasOwn(edits, ref)
      ? { ...cell, value: edits[ref], edited: true }
      : cell;
  }
  function colWidth(c) {
    return state.columns[c] ?? sheet.cols.get(c)?.width ?? 80;
  }
  function lowerBound(y) {
    let lo = 0, hi = plan.totalRows;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (prefix[mid] <= y) lo = mid + 1;
      else hi = mid;
    }
    return clamp(lo, 1, plan.totalRows);
  }
  function textHeight(value, width, size) {
    const perLine = Math.max(3, Math.floor((width - 12) / (size * .72)));
    const lines = String(value).split('\n').reduce((n, line) => {
      return n + Math.max(1, Math.ceil(line.length / perLine));
    }, 0);
    return Math.ceil(lines * size * 96 / 72 * 1.45 + 12);
  }
  function buildGeometry() {
    columns = [];
    for (let c = 1; c <= sheet.maxC; c++) {
      if (!sheet.cols.get(c)?.hidden) columns.push(c);
    }
    const editsByRow = new Map();
    for (const [ref, value] of Object.entries(showStory ? state.cells : state.originalCells)) {
      const match = /^([A-Z]+)(\d+)$/.exec(ref);
      if (!match) continue;
      const r = +match[2], c = ExcelWorkbook.colNum(match[1]);
      if (!editsByRow.has(r)) editsByRow.set(r, []);
      editsByRow.get(r).push({ c, value });
    }
    prefix = new Float64Array(plan.totalRows + 1);
    heights = new Float64Array(plan.totalRows + 1);
    for (let r = 1; r <= plan.totalRows; r++) {
      const source = ExcelLayout.sourceRow(sheet, plan, r);
      const meta = sheet.rows.get(source);
      let height = meta?.hidden ? 0 : Math.max(22, (meta?.height || 15) * 96 / 72);
      if (height && showStory && plan.storyAt.has(r)) {
        height = Math.max(height, textHeight(cellAt(r, plan.column).value, colWidth(plan.column), fontSize));
      }
      if (height) for (const edit of editsByRow.get(r) || []) {
        if (ExcelLayout.mergeAt(sheet, plan, r, edit.c)) continue;
        const cell = cellAt(r, edit.c), css = session.book.styles[cell.style] || {};
        if (cell.story || String(edit.value).includes('\n') || css.whiteSpace === 'pre-wrap') {
          height = Math.max(height, textHeight(edit.value, colWidth(edit.c), cell.story ? fontSize : parseFloat(css.fontSize) || 11));
        }
      }
      if (height && state.rows[r] !== undefined) height = state.rows[r];
      heights[r] = height;
      prefix[r] = prefix[r - 1] + height;
    }
  }
  function rebuild() {
    const row = lowerBound(scroll.scrollTop), offset = scroll.scrollTop - prefix[row - 1];
    buildGeometry();
    scroll.scrollTop = prefix[row - 1] + Math.max(0, offset);
    render();
  }
  function selectionBounds() {
    const end=rangeEnd||selected;
    const a=ExcelLayout.mergeAt(sheet,plan,selected.r,selected.c);
    const b=ExcelLayout.mergeAt(sheet,plan,end.r,end.c);
    let bounds={r1:Math.min(a?.a.r??selected.r,b?.a.r??end.r),r2:Math.max(a?.b.r??selected.r,b?.b.r??end.r),
      c1:Math.min(a?.a.c??selected.c,b?.a.c??end.c),c2:Math.max(a?.b.c??selected.c,b?.b.c??end.c)};
    // Include whole merged regions touched by the selection.
    let changed=true;
    while(changed){changed=false;for(let r=bounds.r1;r<=bounds.r2;r++)for(const c of columns){
      if(c<bounds.c1||c>bounds.c2)continue;
      const m=ExcelLayout.mergeAt(sheet,plan,r,c);if(!m)continue;
      const next={r1:Math.min(bounds.r1,m.a.r),r2:Math.max(bounds.r2,m.b.r),c1:Math.min(bounds.c1,m.a.c),c2:Math.max(bounds.c2,m.b.c)};
      if(JSON.stringify(next)!==JSON.stringify(bounds)){bounds=next;changed=true;}
    }}
    return bounds;
  }
  function syncSelection() {
    grid.querySelectorAll('.active-cell,.active-row,.active-column').forEach(e => {
      e.classList.remove('active-cell', 'active-row', 'active-column');
    });
    grid.querySelector(`[data-ref="${key(selected.r, selected.c)}"]`)?.classList.add('active-cell');
    grid.querySelector(`tbody th[data-row="${selected.r}"]`)?.classList.add('active-row');
    grid.querySelector(`thead th[data-column="${selected.c}"]`)?.classList.add('active-column');
    const bounds=selectionBounds();
    grid.querySelectorAll('td[data-ref]').forEach(td=>{
      const r=+td.dataset.row,c=+td.dataset.column;
      td.classList.toggle('reader-range-selected',!!rangeEnd&&r>=bounds.r1&&r<=bounds.r2&&c>=bounds.c1&&c<=bounds.c2);
    });
    $('active-cell-address').textContent = rangeEnd ? key(bounds.r1,bounds.c1)+':'+key(bounds.r2,bounds.c2) : key(selected.r,selected.c);
    if (!draft) formula.value = cellAt(selected.r, selected.c).value;
  }
  function selectCell(r, c, reveal = false) {
    rangeEnd = null;
    selected = canonical(clamp(r, 1, plan.totalRows), c);
    if (reveal) {
      const y = prefix[selected.r - 1];
      if (y < scroll.scrollTop || y + heights[selected.r] > scroll.scrollTop + scroll.clientHeight - 30) {
        scroll.scrollTop = y;
      }
      const index = columns.indexOf(selected.c);
      const left = columns.slice(0, index).reduce((sum, col) => sum + colWidth(col), 0);
      if (left < scroll.scrollLeft) scroll.scrollLeft = left;
      else if (left + colWidth(selected.c) > scroll.scrollLeft + scroll.clientWidth - 48) {
        scroll.scrollLeft = left + colWidth(selected.c) - scroll.clientWidth + 48;
      }
      render();
    }
    syncSelection();
  }
  function move(dr, dc) {
    let r = selected.r, c = selected.c;
    const merged = ExcelLayout.mergeAt(sheet, plan, r, c);
    if (dr) {
      r = dr > 0 ? (merged?.b.r ?? r) + 1 : r - 1;
      while (r >= 1 && r <= plan.totalRows && !heights[r]) r += dr;
    }
    if (dc) {
      const start = dc > 0 ? merged?.b.c ?? c : c;
      const index = columns.findIndex(col => col >= start);
      c = columns[clamp(index + dc, 0, columns.length - 1)];
    }
    selectCell(r, c, true);
  }
  function resizeHandle(type, index) {
    const handle = document.createElement('span');
    handle.className = 'reader-resize ' + type;
    handle.dataset.resize = type;
    handle.dataset.index = index;
    handle.title = type === 'column' ? 'Kéo để đổi độ rộng; double-click để tự vừa' : 'Kéo để đổi chiều cao; double-click để tự vừa';
    return handle;
  }
  function cellElement(r, c, from, to) {
    const merge = ExcelLayout.mergeAt(sheet, plan, r, c);
    let anchor = { r, c }, spanRows = 1, spanCols = 1, height = heights[r];
    if (merge) {
      const visible = columns.filter(col => col >= merge.a.c && col <= merge.b.c);
      let first = Math.max(from, merge.a.r);
      while (first <= merge.b.r && !heights[first]) first++;
      if (r !== first || c !== visible[0]) return null;
      anchor = merge.a; spanCols = visible.length; spanRows = 0; height = 0;
      for (let row = first; row <= Math.min(to, merge.b.r); row++) {
        if (heights[row]) { spanRows++; height += heights[row]; }
      }
    }
    const cell = cellAt(anchor.r, anchor.c);
    const td = document.createElement('td'), inner = document.createElement('div');
    Object.assign(td.style, session.book.styles[cell.style] || {});
    // Merged cells can store the final right/bottom borders on edge cells.
    if(merge){
      for(const [side,rr,cc] of [['Right',merge.a.r,merge.b.c],['Bottom',merge.b.r,merge.a.c]]){
        const edge=cellAt(rr,cc),css=session.book.styles[edge.style]||{};
        if(css['border'+side])td.style['border'+side]=css['border'+side];
      }
    }
    Object.assign(td.style,state.borders[key(anchor.r,anchor.c)]||{});
    if (merge) { td.rowSpan = spanRows; td.colSpan = spanCols; }
    inner.className = 'reader-cell-inner';
    inner.style.height = Math.max(1, height - 1) + 'px';
    inner.style.whiteSpace = td.style.whiteSpace || 'pre';
    inner.textContent = cell.value;
    if (cell.story) { td.classList.add('reader-story'); td.style.verticalAlign = 'top'; inner.style.whiteSpace = 'pre-wrap'; }
    if (String(cell.value).includes('\n')) inner.style.whiteSpace = 'pre-wrap';
    td.dataset.row = anchor.r; td.dataset.column = anchor.c;
    td.dataset.ref = key(anchor.r, anchor.c); td.title = td.dataset.ref;
    td.append(inner); return td;
  }
  function spacer(body, height) {
    if (height <= 0) return;
    const tr = document.createElement('tr'), td = document.createElement('td');
    tr.className = 'reader-spacer'; td.colSpan = columns.length + 1; td.style.height = height + 'px';
    tr.append(td); body.append(tr);
  }
  function render() {
    if (!sheet || editor) return;
    const from = Math.max(1, lowerBound(Math.max(0, scroll.scrollTop - 25)) - 8);
    const to = Math.min(plan.totalRows, Math.max(lowerBound(scroll.scrollTop + (scroll.clientHeight || 650) + 600), from + 35));
    const fragment = document.createDocumentFragment(), cg = document.createElement('colgroup');
    const numberCol = document.createElement('col'); numberCol.style.width = '48px'; cg.append(numberCol);
    const thead = document.createElement('thead'), header = document.createElement('tr');
    const corner = document.createElement('th'); corner.textContent = '◢'; header.append(corner);
    for (const c of columns) {
      const col = document.createElement('col'); col.dataset.column = c; col.style.width = colWidth(c) + 'px'; cg.append(col);
      const th = document.createElement('th'); th.dataset.column = c;
      th.append(document.createTextNode(ExcelWorkbook.colName(c)), resizeHandle('column', c)); header.append(th);
    }
    thead.append(header); fragment.append(cg, thead);
    const body = document.createElement('tbody'); spacer(body, prefix[from - 1]);
    for (let r = from; r <= to; r++) {
      if (!heights[r]) continue;
      const tr = document.createElement('tr'); tr.style.height = heights[r] + 'px';
      const number = document.createElement('th'); number.dataset.row = r;
      number.append(document.createTextNode(r), resizeHandle('row', r)); tr.append(number);
      for (const c of columns) { const td = cellElement(r, c, from, to); if (td) tr.append(td); }
      body.append(tr);
    }
    spacer(body, prefix[plan.totalRows] - prefix[to]); fragment.append(body);
    grid.style.width = (48 + columns.reduce((sum, c) => sum + colWidth(c), 0)) + 'px';
    grid.replaceChildren(fragment); syncSelection();
    const at = lowerBound(scroll.scrollTop), done = plan.storyRows.filter(r => r <= at).length;
    $('status-reading-info').textContent = `Dòng: ${at} | ${plan.storyRows.length ? Math.round(done / plan.storyRows.length * 100) : 0}%`;
  }
  function applyChange(change, forward) {
    for(const item of change.items||[change]){
      const target=sheetState(change.sheet)[item.group];
      const value=forward?item.after:item.before;
      if(value===undefined)delete target[item.key];else target[item.key]=value;
    }
    if (change.sheet === activeSheet) rebuild();
    dirty();
  }
  function record(group, ref, value) {
    const before = state[group][ref];
    if (before === value) return;
    const change = { sheet: activeSheet, group, key: ref, before, after: value };
    undoStack.push(change); if (undoStack.length > 100) undoStack.shift(); redoStack.length = 0;
    applyChange(change, true);
  }
  function setCell(r, c, value) {
    const anchor = canonical(r, c), text = String(value);
    if (text.length > 32767) { notice('Mỗi ô tối đa 32.767 ký tự.'); return; }
    if (text === String(cellAt(anchor.r, anchor.c).value)) return;
    record(showStory ? 'cells' : 'originalCells', key(anchor.r, anchor.c), text);
  }
  function finishEditor(commit = true) {
    if (!editor) return;
    const current = editor; editor = null; current.input.remove();
    if (commit) setCell(current.r, current.c, current.input.value);
    render();
  }
  function beginEdit(seed) {
    finishFormula(true);
    if (editor || !sheet) return;
    selectCell(selected.r, selected.c, true);
    const td = grid.querySelector(`[data-ref="${key(selected.r, selected.c)}"]`);
    if (!td) return;
    const rect = td.getBoundingClientRect(), input = document.createElement('textarea');
    input.className = 'reader-cell-editor'; input.setAttribute('aria-label', 'Sửa ô ' + td.dataset.ref);
    input.spellcheck = false; input.value = seed ?? String(cellAt(selected.r, selected.c).value);
    input.style.left = clamp(rect.left, 2, innerWidth - 100) + 'px';
    input.style.top = clamp(rect.top, 2, innerHeight - 40) + 'px';
    input.style.width = clamp(rect.width, 100, innerWidth - Math.max(2, rect.left) - 4) + 'px';
    input.style.height = clamp(rect.height, 32, Math.max(40, Math.min(320, innerHeight - rect.top - 5))) + 'px';
    editor = { input, r: selected.r, c: selected.c };
    input.onblur = () => finishEditor(true);
    input.onkeydown = event => {
      if (event.isComposing || event.keyCode === 229) return;
      if (event.key === 'Escape') { event.preventDefault(); finishEditor(false); scroll.focus(); }
      else if (event.key === 'Enter' && event.altKey) {
        event.preventDefault(); input.setRangeText('\n', input.selectionStart, input.selectionEnd, 'end');
      } else if (event.key === 'Enter' || event.key === 'Tab') {
        event.preventDefault(); const direction = event.shiftKey ? -1 : 1;
        finishEditor(true); move(event.key === 'Enter' ? direction : 0, event.key === 'Tab' ? direction : 0); scroll.focus();
      }
    };
    document.body.append(input); input.focus({ preventScroll: true });
    input.selectionStart = input.selectionEnd = input.value.length;
  }
  function finishFormula(commit) {
    if (!draft) return;
    const current = draft, value = formula.value; draft = null;
    if (commit) setCell(current.r, current.c, value);
    syncSelection();
  }
  formula.readOnly = false;
  formula.onfocus = () => { if (sheet) draft = { ...selected }; };
  formula.onblur = () => finishFormula(true);
  formula.onkeydown = event => {
    if (event.isComposing || event.keyCode === 229) return;
    if (event.key === 'Enter' || event.key === 'Escape') {
      event.preventDefault(); finishFormula(event.key === 'Enter'); scroll.focus();
      if (event.key === 'Enter') move(1, 0);
    }
  };
  for (const [selector, commit] of [['.formula-btn.enter', true], ['.formula-btn.cancel', false]]) {
    const button = document.querySelector(selector);
    if (button) { button.onpointerdown = event => event.preventDefault(); button.onclick = () => { finishFormula(commit); scroll.focus(); }; }
  }
  function undo(redo = false) {
    finishEditor(true); finishFormula(true);
    const from = redo ? redoStack : undoStack, to = redo ? undoStack : redoStack;
    const change = from.pop(); if (!change) return;
    to.push(change); applyChange(change, redo);
  }
  function autoFit(type, index) {
    finishEditor(true); finishFormula(true);
    if (type === 'row') { record('rows', index, undefined); return; }
    const canvas = document.createElement('canvas'), context = canvas.getContext('2d');
    let width = 36;
    for (let r = 1; r <= Math.min(plan.totalRows, 1500); r++) {
      const cell = cellAt(r, index), css = session.book.styles[cell.style] || {};
      context.font = cell.story ? `${fontSize}pt ${$('excel-select-font-family').value}` : `${css.fontWeight || 'normal'} ${css.fontSize || '11pt'} ${css.fontFamily || 'Calibri'}`;
      for (const line of String(cell.value).split('\n')) width = Math.max(width, context.measureText(line).width + 16);
    }
    record('columns', index, clamp(Math.ceil(width), 36, 1600));
  }
  function startResize(event, handle) {
    if (event.button !== 0) return;
    event.preventDefault(); finishEditor(true); finishFormula(true);
    const type = handle.dataset.resize, index = +handle.dataset.index;
    const group = type === 'column' ? 'columns' : 'rows';
    const start = type === 'column' ? event.clientX : event.clientY;
    const base = type === 'column' ? colWidth(index) : heights[index];
    dragging = { type, index, group, start, base, before: state[group][index], value: base, moved: false };
    document.body.classList.add('reader-resizing-' + type);
    function onMove(e) {
      const delta = (type === 'column' ? e.clientX : e.clientY) - start;
      if (!dragging || Math.abs(delta) < 2 && !dragging.moved) return;
      dragging.moved = true; dragging.value = Math.round(clamp(base + delta, type === 'column' ? 36 : 18, 1600));
      state[group][index] = dragging.value;
      if (type === 'column') {
        const col = grid.querySelector(`col[data-column="${index}"]`); if (col) col.style.width = dragging.value + 'px';
        grid.style.width = (48 + columns.reduce((sum, c) => sum + colWidth(c), 0)) + 'px';
      } else { rebuild(); }
    }
    function stop(e) {
      document.removeEventListener('pointermove', onMove);
      document.removeEventListener('pointerup', stop);
      document.removeEventListener('pointercancel', stop);
      document.body.classList.remove('reader-resizing-' + type);
      const current = dragging; dragging = null;
      if (!current) return;
      if (current.before === undefined) delete state[group][index]; else state[group][index] = current.before;
      if (current.moved && e.type !== 'pointercancel') record(group, index, current.value);
      else rebuild();
    }
    document.addEventListener('pointermove', onMove);
    document.addEventListener('pointerup', stop);
    document.addEventListener('pointercancel', stop);
  }
  grid.onpointerdown = event => {
    const handle = event.target.closest('[data-resize]');
    if (handle) startResize(event, handle);
  };
  grid.onclick = event => {
    if (event.target.closest('[data-resize]')) return;
    const td = event.target.closest('td[data-ref]');
    if (!td) return;
    finishEditor(true); finishFormula(true);
    if(event.shiftKey){rangeEnd=canonical(+td.dataset.row,+td.dataset.column);syncSelection();}
    else selectCell(+td.dataset.row,+td.dataset.column);
    scroll.focus({ preventScroll:true });
  };
  grid.ondblclick = event => {
    const handle = event.target.closest('[data-resize]');
    if (handle) { event.preventDefault(); autoFit(handle.dataset.resize, +handle.dataset.index); return; }
    const td = event.target.closest('td[data-ref]');
    if (td) { selectCell(+td.dataset.row, +td.dataset.column); beginEdit(); }
  };
  scroll.onscroll = () => {
    finishEditor(true);
    if (!scheduled) { scheduled = true; requestAnimationFrame(() => { scheduled = false; render(); }); }
  };
  scroll.onkeydown = event => {
    if (!sheet || event.isComposing) return;
    if (toggleShortcut(event)) return;
    const modifier = event.ctrlKey || event.metaKey;
    if (modifier && event.key.toLowerCase() === 'z') { event.preventDefault(); undo(event.shiftKey); return; }
    if (modifier && event.key.toLowerCase() === 'y') { event.preventDefault(); undo(true); return; }
    if (['ArrowDown','ArrowUp','ArrowLeft','ArrowRight','Enter','Tab'].includes(event.key)) {
      event.preventDefault();
      const dr = event.key === 'ArrowDown' || event.key === 'Enter' ? (event.shiftKey ? -1 : 1) : event.key === 'ArrowUp' ? -1 : 0;
      const dc = event.key === 'ArrowRight' || event.key === 'Tab' ? (event.shiftKey ? -1 : 1) : event.key === 'ArrowLeft' ? -1 : 0;
      move(dr, dc);
    } else if (event.key === 'F2') { event.preventDefault(); beginEdit(); }
    else if (event.key === 'Delete' || event.key === 'Backspace') { event.preventDefault(); setCell(selected.r, selected.c, ''); }
    else if (event.key.length === 1 && !modifier && !event.altKey) { event.preventDefault(); beginEdit(event.key); }
  };
  scroll.addEventListener('copy', event => {
    if (!sheet || editor) return;
    event.clipboardData.setData('text/plain', String(cellAt(selected.r, selected.c).value)); event.preventDefault();
  });
  scroll.addEventListener('paste', event => {
    if (!sheet || editor) return;
    event.preventDefault(); setCell(selected.r, selected.c, event.clipboardData.getData('text/plain'));
  });
  document.addEventListener('keydown', event => {
    if(event.key==='Escape')closeBorders();
    if (toggleShortcut(event)) return;
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's' && session) {
      event.preventDefault(); finishEditor(true); finishFormula(true); persist();
    }
  });
  function toggleShortcut(event) {
    if (!session || event.defaultPrevented || event.isComposing ||
        event.ctrlKey || event.metaKey || event.altKey || event.key.toLowerCase() !== 'b' ||
        editor || draft || event.target.closest('input,textarea,select,[contenteditable="true"]')) return false;
    event.preventDefault();
    if (event.repeat) return true;
    showStory = !showStory;
    rebuild();
    return true;
  }
  function activate(index) {
    finishEditor(true); finishFormula(true);
    rangeEnd=null;
    closeBorders();
    activeSheet = index; sheet = session.book.sheets[index]; state = sheetState(index);
    plan = index === session.plan.sheetIndex ? session.plan : { ...session.plan, column: -1, startRow: 1, endRow: sheet.maxR, blockLength: sheet.maxR, totalRows: sheet.maxR, storyAt: new Map(), storyRows: [] };
    buildGeometry(); scroll.scrollTop = 0; scroll.scrollLeft = 0;
    document.querySelectorAll('.sheet-tab').forEach((button, i) => button.classList.toggle('active', i === index));
    const count = document.querySelector('.count-item'); if (count) count.textContent = 'Count: ' + plan.totalRows.toLocaleString('vi-VN');
    const sum = document.querySelector('.sum-item'); if (sum) sum.hidden = true;
    const bar = document.querySelector('.excel-status-bar');
    if (bar.lastElementChild && !bar.lastElementChild.id && !bar.lastElementChild.classList.contains('count-item')) bar.lastElementChild.hidden = true;
    const row = index === session.plan.sheetIndex ? session.plan.storyRows[0] || 1 : 1;
    selected = canonical(row, columns.includes(plan.column) ? plan.column : columns[0]);
    scroll.scrollTop = prefix[row - 1]; render();
  }
  const bordersButton=$('reader-borders-button'), bordersMenu=$('reader-borders-menu');
  function closeBorders(){if(bordersMenu)bordersMenu.hidden=true;bordersButton?.setAttribute('aria-expanded','false');}
  if(bordersButton&&bordersMenu){
    bordersButton.onpointerdown=e=>e.preventDefault();
    bordersButton.onclick=()=>{
      finishEditor(true);finishFormula(true);
      if(!sheet)return;
      if(!bordersMenu.hidden){closeBorders();return;}
      bordersMenu.hidden=false;
      const box=bordersButton.getBoundingClientRect();
      bordersMenu.style.left=Math.max(4,Math.min(box.left,innerWidth-bordersMenu.offsetWidth-8))+'px';
      bordersMenu.style.top=Math.max(4,Math.min(box.bottom+5,innerHeight-bordersMenu.offsetHeight-8))+'px';
      bordersButton.setAttribute('aria-expanded','true');
    };
    bordersMenu.addEventListener('click',e=>{
      const button=e.target.closest('[data-border]');if(!button)return;
      applyBorders(button.dataset.border);closeBorders();scroll.focus({preventScroll:true});
    });
    document.addEventListener('pointerdown',e=>{if(!bordersMenu.contains(e.target)&&!bordersButton.contains(e.target))closeBorders();});
  }
  function applyBorders(mode){
    const bounds=selectionBounds();
    if((bounds.r2-bounds.r1+1)*(bounds.c2-bounds.c1+1)>20000){notice('Chọn tối đa 20.000 ô mỗi lần kẻ viền.');return;}
    const width=+$('reader-border-width').value;
    const type=$('reader-border-style').value;
    const line=`${type==='double'?Math.max(3,width):width}px ${type} ${$('reader-border-color').value}`;
    const values=new Map();
    const touched=new Set();
    function update(r,c,side,value){
      if(r<1||r>plan.totalRows||c<1||c>sheet.maxC)return;
      const cell=canonical(r,c),ref=key(cell.r,cell.c);
      const css=values.get(ref)||{...(state.borders[ref]||{})};
      if(value===undefined)delete css['border'+side];else css['border'+side]=value;
      values.set(ref,css);
    }
    for(let r=bounds.r1;r<=bounds.r2;r++){
      if(!heights[r])continue;
      for(const c of columns){
        if(c<bounds.c1||c>bounds.c2)continue;
        const anchor=canonical(r,c),ref=key(anchor.r,anchor.c);if(touched.has(ref))continue;touched.add(ref);
        const merge=ExcelLayout.mergeAt(sheet,plan,r,c);
        const top=merge?.a.r??r,bottom=merge?.b.r??r,left=merge?.a.c??c,right=merge?.b.c??c;
        for(const [side,outer,nr,nc,opposite] of [
          ['Top',top===bounds.r1,top-1,left,'Bottom'],['Bottom',bottom===bounds.r2,bottom+1,left,'Top'],
          ['Left',left===bounds.c1,top,left-1,'Right'],['Right',right===bounds.c2,top,right+1,'Left']]){
          const horizontal=side==='Top'||side==='Bottom';
          const apply=mode==='all'||mode==='none'||mode==='reset'||mode==='outside'&&outer||mode==='inside'&&!outer||
            mode==='horizontal'&&!outer&&horizontal||mode==='vertical'&&!outer&&!horizontal||mode===side.toLowerCase()&&outer;
          if(!apply)continue;
          update(top,left,side,mode==='reset'?undefined:mode==='none'?'none':line);
          // Remove the opposite edge to avoid a doubled or stale shared border.
          update(nr,nc,opposite,mode==='reset'?undefined:'none');
        }
      }
    }
    const items=[];
    for(const [ref,value] of values){const after=Object.keys(value).length?value:undefined,before=state.borders[ref];
      if(JSON.stringify(before)!==JSON.stringify(after))items.push({group:'borders',key:ref,before,after});}
    if(!items.length)return;
    const change={sheet:activeSheet,items};undoStack.push(change);if(undoStack.length>100)undoStack.shift();redoStack.length=0;applyChange(change,true);
  }
  $('reader-fullscreen').onclick = async () => {
    try {
      if (document.fullscreenElement) await document.exitFullscreen();
      else if (document.documentElement.requestFullscreen) await document.documentElement.requestFullscreen();
      else notice('Mac: Control + Command + F để bật toàn màn hình.');
    } catch { notice('Hãy dùng nút toàn màn hình của trình duyệt.'); }
  };
  document.querySelector('.win-close').onclick = async () => { finishEditor(true); finishFormula(true); await persist(); window.close(); };
  function changeFont(save = true) {
    finishEditor(true); finishFormula(true);
    document.documentElement.style.setProperty('--reader-font-family', $('excel-select-font-family').value);
    document.documentElement.style.setProperty('--reader-font-size', fontSize + 'pt');
    document.documentElement.style.setProperty('--reader-font-weight', bold ? 'bold' : 'normal');
    document.documentElement.style.setProperty('--reader-font-style', italic ? 'italic' : 'normal');
    $('excel-btn-bold').classList.toggle('active-state', bold);
    $('excel-btn-italic').classList.toggle('active-state', italic);
    if (sheet) rebuild();
    if (save && session) { session.editor.font = { family: $('excel-select-font-family').value, size: fontSize, bold, italic }; dirty(); }
  }
  $('excel-select-font-family').onchange = () => changeFont();
  $('excel-select-font-size').onchange = () => { fontSize = clamp(+$('excel-select-font-size').value || 11, 8, 24); changeFont(); };
  $('excel-btn-font-inc').onclick = () => { fontSize = Math.min(24, fontSize + 1); $('excel-select-font-size').value = fontSize; changeFont(); };
  $('excel-btn-font-dec').onclick = () => { fontSize = Math.max(8, fontSize - 1); $('excel-select-font-size').value = fontSize; changeFont(); };
  $('excel-btn-bold').onclick = () => { bold = !bold; changeFont(); };
  $('excel-btn-italic').onclick = () => { italic = !italic; changeFont(); };
  if ($('excel-btn-undo')) $('excel-btn-undo').onclick = () => undo();
  window.addEventListener('resize', () => { finishEditor(true); render(); });
  (async () => {
    try {
      const id = location.hash.slice(1);
      if (!id) throw Error('Quay lại trang chuẩn bị và mở tab Excel để đọc.');
      session = await ReaderSessions.get(id);
      if (!session) throw Error('Không tìm thấy phiên đọc. Hãy mở lại từ trang chuẩn bị.');
      session.editor ||= { sheets: {} }; session.editor.sheets ||= {};
      const font = session.editor.font || { family: 'Calibri', size: 11, bold: false, italic: false };
      fontSize = font.size; bold = font.bold; italic = font.italic;
      $('excel-select-font-family').value = font.family; $('excel-select-font-size').value = fontSize; changeFont(false);
      document.title = session.book.name + ' - Excel'; $('excel-doc-title').textContent = session.book.name;
      document.querySelector('.sheet-tabs-list').replaceChildren(...session.book.sheets.map((item, index) => {
        const button = document.createElement('button'); button.className = 'sheet-tab'; button.textContent = item.name; button.onclick = () => activate(index); return button;
      }));
      activate(session.plan.sheetIndex); $('status-ready-label').textContent = 'Sẵn sàng';
    } catch (error) {
      const message = document.createElement('p'); message.className = 'reader-error'; message.textContent = error.message; scroll.replaceChildren(message);
    }
  })();
})();
