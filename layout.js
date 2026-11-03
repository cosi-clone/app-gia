(function(root){
  'use strict';
  function makePlan(book, sheetIndex, column, startRow, chunks) {
    const sheet=book.sheets[sheetIndex];
    if(!sheet)throw Error('Chọn sheet hợp lệ.');
    if(!Number.isInteger(column)||column<1||column>sheet.maxC||sheet.cols.get(column)?.hidden)throw Error('Chọn một cột đang hiển thị.');
    if(!Number.isInteger(startRow)||startRow<1||startRow>sheet.maxR)throw Error(`Dòng bắt đầu phải nằm trong 1–${sheet.maxR}.`);
    if(!chunks.length)throw Error('Tải nội dung truyện từ link PDF trước.');
    if(chunks.length>100000)throw Error('Truyện vượt giới hạn 100.000 đoạn.');
    if(sheet.merges.some(m=>m.a.r<startRow&&m.b.r>=startRow))throw Error('Dòng bắt đầu cắt ngang một vùng ô gộp. Chọn dòng đầu vùng gộp hoặc dòng ngay sau vùng đó.');
    const eligible=[];
    for(let r=startRow;r<=sheet.maxR;r++){
      if(sheet.rows.get(r)?.hidden)continue;
      if(sheet.merges.some(m=>r>=m.a.r&&r<=m.b.r&&column>=m.a.c&&column<=m.b.c))continue;
      eligible.push(r);
    }
    if(!eligible.length)throw Error('Vùng chọn không có ô phù hợp: các dòng đang ẩn hoặc cột nằm trong ô gộp.');
    const blockLength=sheet.maxR-startRow+1;
    const copies=Math.max(1,Math.ceil(chunks.length/eligible.length));
    const totalRows=sheet.maxR+(copies-1)*blockLength;
    if(totalRows>100000)throw Error('Số hàng sau khi lặp vượt 100.000. Chọn vùng mẫu ngắn hơn hoặc chia truyện.');
    const storyRows=[],storyAt=new Map();
    for(let copy=0;copy<copies;copy++)for(const row of eligible){
      if(storyRows.length>=chunks.length)break;
      const target=row+copy*blockLength;
      storyAt.set(target,storyRows.length);storyRows.push(target);
    }
    return {sheetIndex,column,startRow,blockLength,copies,totalRows,addedRows:totalRows-sheet.maxR,storyRows,storyAt};
  }
  function sourceRow(sheet,plan,row){return row<=sheet.maxR?row:plan.startRow+(row-plan.startRow)%plan.blockLength;}
  function mergeAt(sheet,plan,row,column){
    const source=sourceRow(sheet,plan,row),offset=row-source;
    const m=sheet.merges.find(m=>source>=m.a.r&&source<=m.b.r&&column>=m.a.c&&column<=m.b.c);
    return m?{a:{r:m.a.r+offset,c:m.a.c},b:{r:m.b.r+offset,c:m.b.c}}:null;
  }
  function cellAt(sheet,plan,chunks,row,column){
    const source=sourceRow(sheet,plan,row),ref=root.ExcelWorkbook.colName(column)+source;
    const cell=sheet.cells.get(ref)||{value:'',style:sheet.cols.get(column)?.style||0};
    if(column===plan.column&&plan.storyAt.has(row))return {...cell,value:chunks[plan.storyAt.get(row)],story:true};
    return {...cell,story:false};
  }
  const api={makePlan,sourceRow,mergeAt,cellAt};
  root.ExcelLayout=api;
  if(typeof module!=='undefined')module.exports=api;
})(typeof window==='undefined'?globalThis:window);
