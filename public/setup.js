(() => {
  'use strict';

  const $ = id => document.getElementById(id);

  let book = null;
  let busy = false;
  let plan = null;

  const status = message => {
    $('status').textContent = message;
  };

  function controls(value) {
    busy = value;

    for (const id of ['story-file', 'template-file']) {
      $(id).disabled = value;
    }

    for (const id of ['sheet', 'column', 'start-row']) {
      $(id).disabled = value || !book;
    }

    $('open-reader').disabled = value || !plan;
  }

  function updatePlan() {
    plan = null;

    if (book && state.allChunks.length) {
      try {
        plan = ExcelLayout.makePlan(
          book,
          +$('sheet').value,
          +$('column').value,
          +$('start-row').value,
          state.allChunks.map(chunk => chunk.text)
        );

        const endRow =
          plan.endRow ??
          book.sheets[plan.sheetIndex].maxR;

        const repeatInfo = plan.addedRows
          ? (
              `Lặp khối hàng ${plan.startRow}–${endRow} ` +
              `thêm ${plan.copies - 1} lần ` +
              `(${plan.addedRows.toLocaleString('vi-VN')} hàng mới), ` +
              'kèm dữ liệu và định dạng các cột khác.'
            )
          : 'Template đã đủ hàng, không cần lặp thêm.';

        $('plan-info').textContent =
          `${state.allChunks.length.toLocaleString('vi-VN')} đoạn • ` +
          `${plan.totalRows.toLocaleString('vi-VN')} hàng Excel. ` +
          repeatInfo;
      } catch (error) {
        $('plan-info').textContent = error.message;
      }
    } else {
      $('plan-info').textContent =
        'Tải truyện và chọn XLSX để xem số hàng cần dùng.';
    }

    $('open-reader').disabled = busy || !plan;
  }

  function preview() {
    const sheet = book.sheets[+$('sheet').value];
    const table = document.createElement('table');
    const head = document.createElement('tr');

    for (let c = 1; c <= Math.min(sheet.maxC, 12); c++) {
      const th = document.createElement('th');

      th.textContent = ExcelWorkbook.colName(c);
      head.append(th);
    }

    table.append(head);

    const from = +$('start-row').value || 1;

    for (
      let r = from;
      r <= Math.min(sheet.maxR, from + 6);
      r++
    ) {
      const tr = document.createElement('tr');

      for (
        let c = 1;
        c <= Math.min(sheet.maxC, 12);
        c++
      ) {
        const td = document.createElement('td');

        td.textContent =
          sheet.cells.get(
            ExcelWorkbook.colName(c) + r
          )?.value || '';

        tr.append(td);
      }

      table.append(tr);
    }

    $('preview').replaceChildren(table);
  }

  function selectSheet() {
    const sheet = book.sheets[+$('sheet').value];

    $('column').replaceChildren();

    for (let c = 1; c <= sheet.maxC; c++) {
      if (sheet.cols.get(c)?.hidden) continue;

      const option = document.createElement('option');
      option.value = c;

      const header =
        sheet.cells.get(
          ExcelWorkbook.colName(c) + sheet.headerRow
        )?.value || '';

      option.textContent =
        ExcelWorkbook.colName(c) + ' — ' + header;

      $('column').append(option);
    }

    const defaultColumn = Math.min(4, sheet.maxC);

    $('column').value = String(
      sheet.cols.get(defaultColumn)?.hidden
        ? +$('column').options[0]?.value
        : defaultColumn
    );

    $('start-row').value = Math.min(
      sheet.maxR,
      sheet.headerRow + 1
    );

    $('template-meta').textContent =
      `${book.name} • ${book.sheets.length} sheet • ` +
      `${sheet.maxR} hàng, ${sheet.maxC} cột. ` +
      'Khối từ dòng chọn đến hàng cuối có dữ liệu ' +
      'sẽ được lặp nếu thiếu hàng.';

    preview();
    updatePlan();
  }

  $('story-file').onchange = async event => {
    const file = event.target.files[0];

    if (!file) return;

    if (!file.name.toLowerCase().endsWith('.pdf')) {
      event.target.value = '';
      return status('Chọn file PDF từ máy của bạn.');
    }

    controls(true);
    status('Đang đọc PDF…');

    try {
      await Access.requireActive();
      await processDocumentFile(file, status);

      $('story-meta').textContent =
        `${state.fileName} • ` +
        `${state.allChunks.length.toLocaleString('vi-VN')} ` +
        'đoạn đã sẵn sàng.';

      status(
        'Đã đọc PDF. Chọn vị trí trong template ' +
        'rồi mở tab Excel.'
      );
    } catch (error) {
      status(error.message);
    } finally {
      event.target.value = '';
      updatePlan();
      controls(false);
    }
  };

  $('template-file').onchange = async event => {
    const file = event.target.files[0];

    if (!file) return;

    controls(true);
    status('Đang đọc template…');

    try {
      await Access.requireActive();

      book = await ExcelWorkbook.readBook(file);

      $('sheet').replaceChildren(
        ...book.sheets.map((sheet, index) => {
          const option = document.createElement('option');

          option.value = index;
          option.textContent = sheet.name;

          return option;
        })
      );

      $('sheet').value = '0';

      selectSheet();

      status(
        'Đã đọc template. Kiểm tra cột và dòng bắt đầu ' +
        'trước khi mở tab.'
      );
    } catch (error) {
      status(error.message);
    } finally {
      event.target.value = '';
      controls(false);
    }
  };

  $('sheet').onchange = selectSheet;
  $('column').onchange = updatePlan;

  $('start-row').oninput = () => {
    preview();
    updatePlan();
  };

  $('open-reader').onclick = async () => {
    if (!plan || busy) return;

    const target = window.open('about:blank', '_blank');

    if (!target) {
      return status(
        'Trình duyệt đã chặn tab mới. ' +
        'Cho phép mở cửa sổ bật lên cho trang rồi bấm lại.'
      );
    }

    controls(true);

    try {
      await Access.requireActive();

      const id = crypto.randomUUID
        ? crypto.randomUUID()
        : Date.now().toString(36) +
          Math.random().toString(36).slice(2);

      await ReaderSessions.save({
        id,
        book,
        chunks: state.allChunks.map(chunk => chunk.text),
        storyName: state.fileName,
        plan,
        createdAt: Date.now()
      });

      target.location.replace('reader.html#' + id);

      status(
        'Đã mở tab Excel riêng. ' +
        'Bấm nút □ trong tab đó để bật toàn màn hình.'
      );
    } catch (error) {
      target.close();
      status(error.message);
    } finally {
      controls(false);
    }
  };
})();