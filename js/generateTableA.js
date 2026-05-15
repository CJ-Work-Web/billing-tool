/**
 * generateTableA.js
 * 依案件資料生成符合「提報修繕明細」格式的 Excel（表A）。
 * 使用手動設定樣式，不依賴範本載入，確保在本機與 GitHub Pages 皆可運作。
 */

async function generateTableA(cases, rocYear, month) {
  const wb = XLSX.utils.book_new();
  const ws = {};
  const merges = [];

  // ── 樣式常數 ──────────────────────────────────────────────────────────────
  const tealFill = { patternType: 'solid', fgColor: { rgb: '46BDC6' } };
  const noFill   = { patternType: 'none' };

  const baseFont   = { name: '新細明體', sz: 10 };
  const boldWhite  = { name: '新細明體', sz: 10, bold: true, color: { rgb: 'FFFFFF' } };
  const thinBorder = {
    top:    { style: 'thin', color: { rgb: 'AAAAAA' } },
    bottom: { style: 'thin', color: { rgb: 'AAAAAA' } },
    left:   { style: 'thin', color: { rgb: 'AAAAAA' } },
    right:  { style: 'thin', color: { rgb: 'AAAAAA' } },
  };

  // 表頭樣式（青色底、白字、置中）
  const headerStyle = {
    fill: tealFill,
    font: boldWhite,
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: thinBorder,
  };
  // 資料儲存格（一般）
  const dataStyle = {
    fill: noFill,
    font: baseFont,
    alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
    border: thinBorder,
  };
  // 資料儲存格（數字，靠右）
  const numStyle = {
    fill: noFill,
    font: baseFont,
    alignment: { horizontal: 'right', vertical: 'center' },
    numFmt: '#,##0',
    border: thinBorder,
  };
  // 案號欄（青色底、白字、置中）
  const caseNoStyle = {
    fill: tealFill,
    font: { name: '新細明體', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: thinBorder,
  };

  // ── 輔助：設定儲存格 ──────────────────────────────────────────────────────
  const setCell = (r, c, value, style, type) => {
    const addr = XLSX.utils.encode_cell({ r, c });
    const t = type !== undefined ? type : (typeof value === 'number' ? 'n' : 's');
    ws[addr] = { v: value !== undefined && value !== null ? value : '', t, s: style || dataStyle };
  };

  // ── Row 0（第1列）：標題 "提報修繕明細"，合併 A1:N1 ─────────────────────
  const titleStyle = {
    fill: tealFill,
    font: { name: '新細明體', sz: 14, bold: true, color: { rgb: 'FFFFFF' } },
    alignment: { horizontal: 'center', vertical: 'center' },
  };
  setCell(0, 0, '提報修繕明細', titleStyle);
  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 13 } });

  // ── Row 1（第2列）：公司名稱 + 月份，合併 A2:F2 與 G2:N2 ───────────────
  setCell(1, 0, '晟晁物業管理顧問股份有限公司', headerStyle);
  merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: 5 } });
  setCell(1, 6, `${rocYear}年${month}月份月結`, headerStyle);
  merges.push({ s: { r: 1, c: 6 }, e: { r: 1, c: 13 } });

  // ── Row 2（第3列）：欄位標題 ──────────────────────────────────────────────
  const colHeaders = [
    '案號', '編號', '報修日期', '報修單號', '站別', '地址',
    '故障說明', '完工日期', '合約項次', '品項', '單位', '數量', '單價', '未稅金額小計(1)',
  ];
  colHeaders.forEach((h, c) => setCell(2, c, h, headerStyle));

  // ── 資料列（Row 3 起，0-indexed）────────────────────────────────────────
  let currentRow = 3;

  cases.forEach((cas, caseIdx) => {
    const caseNo  = caseIdx + 1;
    const items   = cas.items;
    const startRow = currentRow;

    items.forEach((item, itemIdx) => {
      const r = currentRow;

      if (itemIdx === 0) {
        // 第一個項次列：填入案件欄位 A–H
        setCell(r, 0, caseNo, caseNoStyle, 'n');       // A: 案號
        setCell(r, 1, 1, caseNoStyle, 'n');             // B: 編號（隱藏）
        setCell(r, 2, cas.報修日期, dataStyle);          // C: 報修日期
        setCell(r, 3, cas.JDM案號, dataStyle);           // D: 報修單號
        setCell(r, 4, cas.站別, dataStyle);              // E: 站別
        setCell(r, 5, cas.地址, dataStyle);              // F: 地址
        setCell(r, 6, cas.故障說明, dataStyle);           // G: 故障說明
        setCell(r, 7, cas.完工日期, dataStyle);           // H: 完工日期
      } else {
        // 非第一列：A–H 屬合併範圍，只填 B（編號）
        setCell(r, 1, itemIdx + 1, caseNoStyle, 'n');  // B: 編號（隱藏）
      }

      // 項次欄位（每列都填）I–N
      setCell(r, 8,  item.合約項次, dataStyle);          // I: 合約項次
      setCell(r, 9,  item.品項,     dataStyle);          // J: 品項（隱藏）
      setCell(r, 10, item.單位,     dataStyle);          // K: 單位
      setCell(r, 11, item.數量,     numStyle, 'n');      // L: 數量
      setCell(r, 12, item.單價,     numStyle, 'n');      // M: 單價（隱藏）
      setCell(r, 13, item.未稅小計, numStyle, 'n');      // N: 未稅金額小計

      currentRow++;
    });

    // 多項次案件：合併 A–H 欄（col 0–7）
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

  ws['!cols'] = [
    { wch: 7.5 },                      // A 案號
    { wch: 8,  hidden: true },         // B 編號（隱藏）
    { wch: 29.17 },                    // C 報修日期
    { wch: 26.33 },                    // D 報修單號
    { wch: 24.83 },                    // E 站別
    { wch: 29.5  },                    // F 地址
    { wch: 79.83 },                    // G 故障說明
    { wch: 24.83 },                    // H 完工日期
    { wch: 23.67 },                    // I 合約項次
    { wch: 23.67, hidden: true },      // J 品項（隱藏）
    { wch: 9.83  },                    // K 單位
    { wch: 13.33 },                    // L 數量
    { wch: 16.83, hidden: true },      // M 單價（隱藏）
    { wch: 29.83 },                    // N 未稅金額小計
  ];

  // 列高
  ws['!rows'] = [];
  ws['!rows'][0] = { hpt: 40 };  // 標題列
  ws['!rows'][1] = { hpt: 35 };  // 公司/月份列
  ws['!rows'][2] = { hpt: 40 };  // 欄位標題列
  for (let r = 3; r < currentRow; r++) {
    ws['!rows'][r] = { hpt: 50 };
  }

  const sheetName = `${rocYear}.${String(month).padStart(2, '0')}`;
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  return XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
}
