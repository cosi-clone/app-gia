(function(root) {
  'use strict';

  function repeatEnd(sheet, startRow) {
    const filledRows = new Set();

    for (const [ref, cell] of sheet.cells) {
      const value = String(cell.value ?? '').trim();

      // Ô trống hoặc chỉ có giá trị 0 không xác định
      // đây là một hàng dữ liệu thực.
      if (
        !value ||
        /^[+-]?0+(?:[.,]0+)?$/.test(value)
      ) {
        continue;
      }

      const row = +(ref.match(/\d+$/)?.[0] || 0);

      if (
        row >= startRow &&
        !sheet.rows.get(row)?.hidden
      ) {
        filledRows.add(row);
      }
    }

    let lastDataRow = 0;

    for (let row = startRow; row <= sheet.maxR; row++) {
      if (sheet.rows.get(row)?.hidden) continue;

      if (filledRows.has(row)) {
        lastDataRow = row;
      } else if (lastDataRow) {
        // Hết vùng dữ liệu đang hiển thị.
        // File mẫu: kết thúc ở 145, không kéo tới 940.
        return lastDataRow;
      }
    }

    return lastDataRow || sheet.maxR;
  }

  function makePlan(book, sheetIndex, column, startRow, chunks) {
    const sheet = book.sheets[sheetIndex];

    if (!sheet) {
      throw Error('Chọn sheet hợp lệ.');
    }

    if (
      !Number.isInteger(column) ||
      column < 1 ||
      column > sheet.maxC ||
      sheet.cols.get(column)?.hidden
    ) {
      throw Error('Chọn một cột đang hiển thị.');
    }

    if (
      !Number.isInteger(startRow) ||
      startRow < 1 ||
      startRow > sheet.maxR
    ) {
      throw Error(
        `Dòng bắt đầu phải nằm trong 1–${sheet.maxR}.`
      );
    }

    if (!chunks.length) {
      throw Error('Tải file PDF trước.');
    }

    if (chunks.length > 100000) {
      throw Error('Truyện vượt giới hạn 100.000 đoạn.');
    }

    if (
      sheet.merges.some(
        merge =>
          merge.a.r < startRow &&
          merge.b.r >= startRow
      )
    ) {
      throw Error(
        'Dòng bắt đầu cắt ngang ô gộp. ' +
        'Chọn dòng đầu ô gộp hoặc dòng ngay sau ô gộp.'
      );
    }

    const endRow = repeatEnd(sheet, startRow);

    if (
      sheet.merges.some(
        merge =>
          merge.a.r <= endRow &&
          merge.b.r > endRow
      )
    ) {
      throw Error(
        'Cuối khối lặp cắt ngang ô gộp. ' +
        'Hãy kiểm tra vùng mẫu.'
      );
    }

    const blockLength = endRow - startRow + 1;
    const eligible = [];

    for (let row = startRow; row <= endRow; row++) {
      if (sheet.rows.get(row)?.hidden) continue;

      const merged = sheet.merges.some(
        merge =>
          row >= merge.a.r &&
          row <= merge.b.r &&
          column >= merge.a.c &&
          column <= merge.b.c
      );

      if (!merged) {
        eligible.push(row);
      }
    }

    if (!eligible.length) {
      throw Error(
        'Vùng chọn không có ô phù hợp để chèn truyện.'
      );
    }

    const copies = Math.max(
      1,
      Math.ceil(chunks.length / eligible.length)
    );

    const totalRows =
      endRow + (copies - 1) * blockLength;

    if (totalRows > 100000) {
      throw Error(
        'Số hàng sau khi lặp vượt 100.000. ' +
        'Hãy chọn khối mẫu ngắn hơn hoặc chia truyện.'
      );
    }

    const storyRows = [];
    const storyAt = new Map();

    for (let copy = 0; copy < copies; copy++) {
      for (const source of eligible) {
        if (storyRows.length >= chunks.length) break;

        const target =
          source + copy * blockLength;

        storyAt.set(target, storyRows.length);
        storyRows.push(target);
      }
    }

    return {
      layoutVersion: 3,
      sheetIndex,
      column,
      startRow,
      endRow,
      blockLength,
      copies,
      totalRows,
      addedRows: totalRows - endRow,
      storyRows,
      storyAt
    };
  }

  function sourceRow(sheet, plan, row) {
    const endRow = plan.endRow ?? sheet.maxR;

    if (row <= endRow) {
      return row;
    }

    return (
      plan.startRow +
      ((row - endRow - 1) % plan.blockLength)
    );
  }

  function mergeAt(sheet, plan, row, column) {
    const source = sourceRow(sheet, plan, row);
    const offset = row - source;

    const merge = sheet.merges.find(
      item =>
        source >= item.a.r &&
        source <= item.b.r &&
        column >= item.a.c &&
        column <= item.b.c
    );

    if (!merge) return null;

    return {
      a: {
        r: merge.a.r + offset,
        c: merge.a.c
      },
      b: {
        r: merge.b.r + offset,
        c: merge.b.c
      }
    };
  }

  function cellAt(sheet, plan, chunks, row, column) {
    const source = sourceRow(sheet, plan, row);

    const ref =
      root.ExcelWorkbook.colName(column) + source;

    const original = sheet.cells.get(ref) || {
      value: '',
      style: sheet.cols.get(column)?.style || 0
    };

    if (
      column === plan.column &&
      plan.storyAt.has(row)
    ) {
      return {
        ...original,
        value: chunks[plan.storyAt.get(row)],
        story: true
      };
    }

    // Các cột khác luôn lặp nội dung Excel gốc.
    // Nhấn B: reader đặt column = -1 nên cột truyện
    // cũng đi qua đây và hiện nội dung Excel gốc.
    return {
      ...original,
      story: false
    };
  }

  const api = {
    makePlan,
    sourceRow,
    mergeAt,
    cellAt
  };

  root.ExcelLayout = api;

  if (typeof module !== 'undefined') {
    module.exports = api;
  }
})(typeof window === 'undefined' ? globalThis : window);