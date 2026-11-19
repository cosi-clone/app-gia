(function(root) {
  'use strict';

  const M =
    'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

  const children = (e, tag) =>
    e ? [...e.children].filter(x => x.localName === tag) : [];

  const first = (e, tag) => children(e, tag)[0];

  const txt = e =>
    e
      ? [...e.getElementsByTagNameNS(M, 't')]
          .map(t => t.textContent)
          .join('')
      : '';

  const xml = s => {
    const d = new DOMParser().parseFromString(
      s,
      'application/xml'
    );

    if (d.querySelector('parsererror')) {
      throw Error('XML không hợp lệ trong XLSX.');
    }

    return d;
  };

  const colNum = s =>
    [...s.toUpperCase()].reduce(
      (n, c) => n * 26 + c.charCodeAt(0) - 64,
      0
    );

  const colName = n => {
    let s = '';

    for (; n; n = Math.floor((n - 1) / 26)) {
      s = String.fromCharCode(65 + (n - 1) % 26) + s;
    }

    return s;
  };

  const addr = s => {
    const m = /^([A-Z]+)(\d+)$/i.exec(s);

    if (!m) {
      throw Error('Địa chỉ ô không hợp lệ.');
    }

    return {
      c: colNum(m[1]),
      r: +m[2]
    };
  };

  const range = s => {
    const a = s.split(':');

    return {
      a: addr(a[0]),
      b: addr(a[1] || a[0])
    };
  };

  let styles = [];
  let theme = [];
  let palette = [];

  const color = (element, fallback = '') =>
    root.XlsxColors.resolve(
      element,
      theme,
      palette,
      fallback
    );

  function loadStyles(d) {
    const element = d.documentElement;

    const fonts = children(
      first(element, 'fonts'),
      'font'
    );

    const fills = children(
      first(element, 'fills'),
      'fill'
    );

    const borders = children(
      first(element, 'borders'),
      'border'
    );

    const bases = children(
      first(element, 'cellStyleXfs'),
      'xf'
    );

    palette = children(
      first(first(element, 'colors'), 'indexedColors'),
      'rgbColor'
    ).map(c => c.getAttribute('rgb'));

    return children(first(element, 'cellXfs'), 'xf').map(x => {
      const base = bases[+x.getAttribute('xfId') || 0];

      const id = name => +(
        x.getAttribute(
          'apply' + name[0].toUpperCase() + name.slice(1)
        ) === '0'
          ? base?.getAttribute(name + 'Id') || 0
          : x.getAttribute(name + 'Id') ??
            base?.getAttribute(name + 'Id') ??
            0
      );

      const font = fonts[id('font')];
      const fill = fills[id('fill')];
      const border = borders[id('border')];

      const a =
        first(x, 'alignment') ||
        first(base, 'alignment');

      const css = {
        color: '#000000'
      };

      if (font) {
        css.fontFamily =
          first(font, 'name')?.getAttribute('val') ||
          'Calibri';

        css.fontSize =
          (first(font, 'sz')?.getAttribute('val') || 11) +
          'pt';

        css.fontWeight =
          first(font, 'b') &&
          !['0', 'false'].includes(
            first(font, 'b').getAttribute('val')
          )
            ? 'bold'
            : 'normal';

        css.fontStyle =
          first(font, 'i') &&
          !['0', 'false'].includes(
            first(font, 'i').getAttribute('val')
          )
            ? 'italic'
            : 'normal';

        css.color = color(
          first(font, 'color'),
          '#000000'
        );

        if (
          first(font, 'u') &&
          first(font, 'u').getAttribute('val') !== 'none'
        ) {
          css.textDecoration = 'underline';
        }
      }

      const pattern = first(fill, 'patternFill');

      if (pattern?.getAttribute('patternType') === 'solid') {
        css.backgroundColor = color(
          first(pattern, 'fgColor'),
          '#FFFFFF'
        );
      }

      if (a) {
        const h = a.getAttribute('horizontal');

        if (['left', 'right', 'center', 'justify'].includes(h)) {
          css.textAlign = h;
        }

        const v = a.getAttribute('vertical');

        css.verticalAlign =
          v === 'center' ? 'middle' : v || 'bottom';

        if (a.getAttribute('wrapText') === '1') {
          css.whiteSpace = 'pre-wrap';
        }
      }

      for (const side of ['left', 'right', 'top', 'bottom']) {
        const edge = first(border, side);
        const style = edge?.getAttribute('style');

        if (!style || style === 'none') continue;

        const originalWidth =
          style === 'double' || style === 'thick'
            ? 3
            : /medium/i.test(style)
              ? 2
              : 1;

        const width = style === 'double'
          ? 3
          : originalWidth * 0.75;

        const type = style === 'double'
          ? 'double'
          : /dash/i.test(style)
            ? 'dashed'
            : style === 'dotted' || style === 'hair'
              ? 'dotted'
              : 'solid';

        const borderColor = color(
          first(edge, 'color'),
          '#000000'
        );

        const property =
          'border' +
          side[0].toUpperCase() +
          side.slice(1);

        css[property] =
          `${width}px ${type} ${borderColor}`;
      }

      return css;
    });
  }

  async function readBook(file) {
    if (file.size > 50 * 1024 * 1024) {
      throw Error('Template tối đa 50 MB.');
    }

    const Zip = await ensureJsZipLoaded();

    const z = await Zip.loadAsync(
      await file.arrayBuffer()
    );

    const read = async p => {
      const f = z.file(p);

      if (!f) {
        throw Error('Thiếu thành phần XLSX: ' + p);
      }

      return xml(await f.async('string'));
    };

    const wb = await read('xl/workbook.xml');

    const rels = await read(
      'xl/_rels/workbook.xml.rels'
    );

    const links = new Map(
      [...rels.documentElement.children].map(r => [
        r.getAttribute('Id'),
        r.getAttribute('Target')
      ])
    );

    const strings = z.file('xl/sharedStrings.xml')
      ? children(
          (await read('xl/sharedStrings.xml')).documentElement,
          'si'
        ).map(txt)
      : [];

    theme = [];
    palette = [];

    if (z.file('xl/theme/theme1.xml')) {
      const td = await read('xl/theme/theme1.xml');

      const cs = [...td.getElementsByTagName('*')]
        .find(e => e.localName === 'clrScheme');

      const order = [
        'lt1',
        'dk1',
        'lt2',
        'dk2',
        'accent1',
        'accent2',
        'accent3',
        'accent4',
        'accent5',
        'accent6',
        'hlink',
        'folHlink'
      ];

      theme = order.map(k => {
        const c = first(cs, k)?.firstElementChild;

        return c
          ? '#' + (
              c.getAttribute('lastClr') ||
              c.getAttribute('val')
            )
          : '';
      });
    }

    styles = z.file('xl/styles.xml')
      ? loadStyles(await read('xl/styles.xml'))
      : [];

    const sheets = [];

    for (
      const s of children(
        first(wb.documentElement, 'sheets'),
        'sheet'
      )
    ) {
      const target = links.get(
        s.getAttributeNS(
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          'id'
        )
      );

      if (!target) continue;

      const path = new URL(
        target,
        'https://xlsx.local/xl/workbook.xml'
      ).pathname.slice(1);

      const d = (await read(path)).documentElement;

      const cells = new Map();
      const rows = new Map();
      const cols = new Map();

      let maxR = 1;
      let maxC = 1;

      for (
        const row of children(
          first(d, 'sheetData'),
          'row'
        )
      ) {
        rows.set(+row.getAttribute('r'), {
          height: +row.getAttribute('ht') || 15,
          hidden: row.getAttribute('hidden') === '1',
          style: row.hasAttribute('s')
            ? +row.getAttribute('s')
            : undefined
        });

        for (const c of children(row, 'c')) {
          const ref = c.getAttribute('r');
          const p = addr(ref);
          const v = first(c, 'v')?.textContent || '';
          const t = c.getAttribute('t');

          const value = t === 's'
            ? strings[+v] || ''
            : t === 'inlineStr'
              ? txt(first(c, 'is'))
              : t === 'b'
                ? v === '1' ? 'TRUE' : 'FALSE'
                : v;

          cells.set(ref, {
            value,
            style: c.hasAttribute('s')
              ? +c.getAttribute('s')
              : row.hasAttribute('s')
                ? +row.getAttribute('s')
                : null
          });

          maxR = Math.max(maxR, p.r);
          maxC = Math.max(maxC, p.c);
        }
      }

      for (
        const c of children(
          first(d, 'cols'),
          'col'
        )
      ) {
        for (
          let i = +c.getAttribute('min');
          i <= +c.getAttribute('max');
          i++
        ) {
          cols.set(i, {
            width: Math.max(
              20,
              (+c.getAttribute('width') || 9) * 7 + 5
            ),
            hidden: c.getAttribute('hidden') === '1',
            style: +c.getAttribute('style') || 0
          });
        }
      }

      for (const [ref, cell] of cells) {
        if (cell.style === null) {
          cell.style = cols.get(addr(ref).c)?.style || 0;
        }
      }

      const merges = children(
        first(d, 'mergeCells'),
        'mergeCell'
      ).map(m => range(m.getAttribute('ref')));

      for (const m of merges) {
        maxR = Math.max(maxR, m.b.r);
        maxC = Math.max(maxC, m.b.c);
      }

      if (maxC > 500 || maxR > 100000) {
        throw Error(
          'Template vượt giới hạn hiển thị 500 cột / 100.000 dòng.'
        );
      }

      let headerRow = 1;
      let best = -1;

      for (let r = 1; r <= Math.min(maxR, 20); r++) {
        let score = 0;

        for (let c = 1; c <= maxC; c++) {
          const v =
            cells.get(colName(c) + r)?.value || '';

          if (v && !/^[-\d.,% ]+$/.test(v)) {
            score++;
          }
        }

        if (score > best) {
          best = score;
          headerRow = r;
        }
      }

      sheets.push({
        name: s.getAttribute('name'),
        cells,
        rows,
        cols,
        merges,
        maxR,
        maxC,
        baseMaxR: maxR,
        headerRow
      });
    }

    if (!sheets.length) {
      throw Error('Không có sheet đọc được.');
    }

    return {
      name: file.name,
      sheets,
      styles
    };
  }

  root.ExcelWorkbook = {
    readBook,
    colName,
    colNum
  };
})(typeof window === 'undefined' ? globalThis : window);