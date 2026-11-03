(function(root) {
  'use strict';

  function templateEnd(sheet, startRow) {
    let end = startRow;

    for (const [ref, cell] of sheet.cells) {
      if (String(cell.value ?? '').trim() === '') continue;

      const row = +(ref.match(/\d+$/)?.[0] || 0);
      end = Math.max(end, row);
    }

    // Không cắt ngang ô gộp ở cuối khối mẫu.
    let extended = true;

    while (extended) {
      extended = false;

      for (const merge of sheet.merges) {
        if (merge.a.r <= end && merge.b.r > end) {
          end = merge.b.r;
          extended = true;
        }
      }
    }

    return Math.min(sheet.maxR, end);
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
        'Dòng bắt đầu cắt ngang một vùng ô gộp. ' +
        'Chọn dòng đầu vùng gộp hoặc dòng ngay sau vùng đó.'
      );
    }

    const endRow = templateEnd(sheet, startRow);
    const eligible = [];

    for (let r = startRow; r <= endRow; r++) {
      if (sheet.rows.get(r)?.hidden) continue;

      const merged = sheet.merges.some(
        merge =>
          r >= merge.a.r &&
          r <= merge.b.r &&
          column >= merge.a.c &&
          column <= merge.b.c
      );

      if (merged) continue;

      eligible.push(r);
    }

    if (!eligible.length) {
      throw Error(
        'Vùng chọn không có ô phù hợp: ' +
        'các dòng đang ẩn hoặc cột nằm trong ô gộp.'
      );
    }

    const blockLength = endRow - startRow + 1;

    const copies = Math.max(
      1,
      Math.ceil(chunks.length / eligible.length)
    );

    const totalRows =
      endRow + (copies - 1) * blockLength;

    if (totalRows > 100000) {
      throw Error(
        'Số hàng sau khi lặp vượt 100.000. ' +
        'Chọn vùng mẫu ngắn hơn hoặc chia truyện.'
      );
    }

    const storyRows = [];
    const storyAt = new Map();

    for (let copy = 0; copy < copies; copy++) {
      for (const row of eligible) {
        if (storyRows.length >= chunks.length) break;

        const target = row + copy * blockLength;

        storyAt.set(target, storyRows.length);
        storyRows.push(target);
      }
    }

    return {
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
    const end = plan.endRow ?? sheet.maxR;

    if (row <= end) return row;

    return (
      plan.startRow +
      (row - plan.startRow) % plan.blockLength
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

    const cell = sheet.cells.get(ref) || {
      value: '',
      style: sheet.cols.get(column)?.style || 0
    };

    if (
      column === plan.column &&
      plan.storyAt.has(row)
    ) {
      return {
        ...cell,
        value: chunks[plan.storyAt.get(row)],
        story: true
      };
    }

    return {
      ...cell,
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