/**
 * generateTableA.js
 * 依案件資料生成符合「提報修繕明細」格式的 Excel（表A）。
 * 格式依據範本實測：微軟正黑體、外框 medium、內框 thin。
 */

async function generateTableA(cases, rocYear, month) {
  const wb = XLSX.utils.book_new();
  const ws = {};
  const merges = [];

  // ── 顏色 ──────────────────────────────────────────────────────────────────
  const BLACK = '000000';
  const WHITE = 'FFFFFF';
  const TEAL  = '46BDC6';
  const GRAY  = '808080';

  // ── 框線 ──────────────────────────────────────────────────────────────────
  const MED  = { style: 'medium', color: { rgb: BLACK } };
  const THIN = { style: 'thin',   color: { rgb: BLACK } };

  const allMed = { top: MED, bottom: MED, left: MED, right: MED };

  // 資料列：A欄左緣 / N欄右緣用 medium，其餘 thin
  function dataBorder(c) {
    return {
      top:    THIN,
      bottom: THIN,
      left:   c === 0  ? MED : THIN,
      right:  c === 13 ? MED : THIN,
    };
  }

  // ── 樣式 ──────────────────────────────────────────────────────────────────

  // 標題列（第1~2列）：24pt 青色底白字
  const titleStyle = {
    fill:      { patternType: 'solid', fgColor: { rgb: TEAL } },
    font:      { name: '微軟正黑體', sz: 24, bold: true, color: { rgb: WHITE } },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: false },
    border:    allMed,
  };

  // 欄位標題（第3列）：16pt 青色底白字
  const headerStyle = {
    fill:      { patternType: 'solid', fgColor: { rgb: TEAL } },
    font:      { name: '微軟正黑體', sz: 16, bold: true, color: { rgb: WHITE } },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border:    allMed,
  };

  // 欄位標題 J（品項）、M（單價）：灰底白字 12pt
  const grayHeaderStyle = {
    fill:      { patternType: 'solid', fgColor: { rgb: GRAY } },
    font:      { name: 'Arial', sz: 12, bold: true, color: { rgb: WHITE } },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: false },
    border:    allMed,
  };

  // 案號欄（A欄）：青色底白字置中
  const caseNoStyle = {
    fill:      { patternType: 'solid', fgColor: { rgb: TEAL } },
    font:      { name: '微軟正黑體', sz: 11, bold: true, color: { rgb: WHITE } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border:    { top: THIN, bottom: THIN, left: MED, right: THIN },
  };

  // B 欄（編號）：青色底白字
  const bColStyle = {
    fill:      { patternType: 'solid', fgColor: { rgb: TEAL } },
    font:      { name: '微軟正黑體', sz: 11, color: { rgb: WHITE } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border:    dataBorder(1),
  };

  // 一般文字資料欄（依欄位決定邊框）
  function dataStyle(c) {
    return {
      fill:      { patternType: 'none' },
      font:      { name: '微軟正黑體', sz: 11 },
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
      border:    dataBorder(c),
    };
  }

  // 數字欄（靠右、千分位）
  function numStyle(c) {
    return {
      fill:      { patternType: 'none' },
      font:      { name: '微軟正黑體', sz: 11 },
      alignment: { horizontal: 'right', vertical: 'center' },
      numFmt:    '#,##0',
      border:    dataBorder(c),
    };
  }

  // ── 輔助函式 ──────────────────────────────────────────────────────────────
  const setCell = (r, c, value, style, type) => {
    const addr = XLSX.utils.encode_cell({ r, c });
    const t = type !== undefined ? type : (typeof value === 'number' ? 'n' : 's');
    ws[addr] = { v: (value !== undefined && value !== null) ? value : '', t, s: style };
  };

  // ── Row 0：標題 "提報修繕明細"，合併 A1:N1 ───────────────────────────────
  setCell(0, 0, '提報修繕明細', titleStyle);
  for (let c = 1; c <= 13; c++) setCell(0, c, '', titleStyle);
  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 13 } });

  // ── Row 1：公司名稱（A2:F2）＋月份（G2:N2）────────────────────────────────
  setCell(1, 0, '晟晁物業管理顧問股份有限公司', titleStyle);
  for (let c = 1; c <= 5; c++) setCell(1, c, '', titleStyle);
  merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: 5 } });

  setCell(1, 6, `${rocYear}年${month}月份月結`, titleStyle);
  for (let c = 7; c <= 13; c++) setCell(1, c, '', titleStyle);
  merges.push({ s: { r: 1, c: 6 }, e: { r: 1, c: 13 } });

  // ── Row 2：欄位標題 ───────────────────────────────────────────────────────
  const colHeaders = [
    '案號', '編號', '報修日期', '報修單號', '站別', '地址',
    '故障說明', '完工日期', '合約項次', '品項', '單位', '數量', '單價', '未稅金額小計(1)',
  ];
  colHeaders.forEach((h, c) => {
    const style = (c === 9 || c === 12) ? grayHeaderStyle : headerStyle;
    setCell(2, c, h, style);
  });

  // ── 資料列（Row 3 起）────────────────────────────────────────────────────
  let currentRow = 3;

  cases.forEach((cas, caseIdx) => {
    const caseNo   = caseIdx + 1;
    const items    = cas.items;
    const startRow = currentRow;

    items.forEach((item, itemIdx) => {
      const r = currentRow;

      if (itemIdx === 0) {
        // 第一個品項列：填入 A–H 案件欄位
        setCell(r, 0, caseNo,       caseNoStyle,  'n');
        setCell(r, 1, 1,            bColStyle,     'n');
        setCell(r, 2, cas.報修日期,  dataStyle(2));
        setCell(r, 3, cas.JDM案號,   dataStyle(3));
        setCell(r, 4, cas.站別,      dataStyle(4));
        setCell(r, 5, cas.地址,      dataStyle(5));
        setCell(r, 6, cas.故障說明,  dataStyle(6));
        setCell(r, 7, cas.完工日期,  dataStyle(7));
      } else {
        // 合併範圍後續列：填空白儲存格保持框線正確
        for (let c = 0; c <= 7; c++) {
          setCell(r, c, '', {
            fill:   { patternType: 'none' },
            font:   { name: '微軟正黑體', sz: 11 },
            border: dataBorder(c),
          });
        }
        setCell(r, 1, itemIdx + 1, bColStyle, 'n');
      }

      // 每列都填：I–N 品項欄位
      setCell(r, 8,  item.合約項次, dataStyle(8));
      setCell(r, 9,  item.品項,     dataStyle(9));
      setCell(r, 10, item.單位,     dataStyle(10));
      setCell(r, 11, item.數量,     numStyle(11),  'n');
      setCell(r, 12, item.單價,     numStyle(12),  'n');
      setCell(r, 13, item.未稅小計, numStyle(13),  'n');

      currentRow++;
    });

    // 多品項案件：合併 A–H（欄 0–7）
    if (items.length > 1) {
      merges.push({
        s: { r: startRow, c: 0 },
        e: { r: startRow + items.length - 1, c: 7 },
      });
    }
  });

  // ── 工作表屬性 ────────────────────────────────────────────────────────────
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: currentRow - 1, c: 13 } });
  ws['!merges'] = merges;

  // 欄寬（依範本實測值）
  ws['!cols'] = [
    { wch: 8.332  },                      // A 案號
    { wch: 8.887,  hidden: true },        // B 編號（隱藏）
    { wch: 30.0   },                      // C 報修日期
    { wch: 27.109 },                      // D 報修單號
    { wch: 25.664 },                      // E 站別
    { wch: 30.332 },                      // F 地址
    { wch: 80.664 },                      // G 故障說明
    { wch: 25.664 },                      // H 完工日期
    { wch: 24.555 },                      // I 合約項次
    { wch: 24.555, hidden: true },        // J 品項（隱藏）
    { wch: 10.664 },                      // K 單位
    { wch: 14.109 },                      // L 數量
    { wch: 17.664, hidden: true },        // M 單價（隱藏）
    { wch: 30.664 },                      // N 未稅金額小計
  ];

  // 列高（依範本實測值）
  ws['!rows'] = [];
  ws['!rows'][0] = { hpt: 50.1  };   // 標題列
  ws['!rows'][1] = { hpt: 49.8  };   // 公司/月份列
  ws['!rows'][2] = { hpt: 49.95 };   // 欄位標題列
  for (let r = 3; r < currentRow; r++) {
    ws['!rows'][r] = { hpt: 49.95 };
  }

  const sheetName = `${rocYear}.${String(month).padStart(2, '0')}`;
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  return XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
}
