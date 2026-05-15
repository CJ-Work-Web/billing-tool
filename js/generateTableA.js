/**
 * generateTableA.js
 * 依案件資料生成符合「提報修繕明細」格式的 Excel（表A）。
 *
 * 完全不依賴範本載入，使用 ExcelJS 直接建立並套用硬編碼樣式。
 */

async function generateTableA(cases, rocYear, month) {

  const ExcelJS   = window.ExcelJS;
  const workbook  = new ExcelJS.Workbook();
  const sheetName = `${rocYear}.${String(month).padStart(2, '0')}`;
  const ws        = workbook.addWorksheet(sheetName);

  // ── 欄寬設定 ──────────────────────────────────────────────────────────────
  ws.columns = [
    { width: 7.5  },  // A 案號
    { width: 8    },  // B 序號
    { width: 15   },  // C 報修日期
    { width: 22   },  // D 報修單號
    { width: 20   },  // E 站別
    { width: 28   },  // F 地址
    { width: 60   },  // G 故障說明
    { width: 15   },  // H 完工日期
    { width: 18   },  // I 合約項次
    { width: 18   },  // J 品項
    { width: 8    },  // K 單位
    { width: 10   },  // L 數量
    { width: 12   },  // M 單價
    { width: 22   },  // N 未稅小計
  ];

  // ── 樣式定義 ──────────────────────────────────────────────────────────────
  const TEAL_FILL = {
    type: 'pattern', pattern: 'solid',
    fgColor: { argb: 'FF46BDC6' },
  };
  const WHITE_FONT = (bold = false) => ({
    name: '微軟正黑體', size: 14, bold, color: { argb: 'FFFFFFFF' },
  });
  const DATA_FONT = () => ({
    name: '微軟正黑體', size: 14,
  });
  const THIN_BORDER = {
    top:    { style: 'thin' },
    left:   { style: 'thin' },
    bottom: { style: 'thin' },
    right:  { style: 'thin' },
  };
  const CENTER = { horizontal: 'center', vertical: 'middle', wrapText: true };
  const LEFT   = { horizontal: 'left',   vertical: 'middle', wrapText: true };

  const applyHeader = (cell, val, align = CENTER) => {
    cell.value     = val;
    cell.font      = WHITE_FONT(true);
    cell.fill      = TEAL_FILL;
    cell.border    = THIN_BORDER;
    cell.alignment = align;
  };
  const applyData = (cell, val, align = CENTER) => {
    cell.value     = val != null ? val : null;
    cell.font      = DATA_FONT();
    cell.border    = THIN_BORDER;
    cell.alignment = align;
  };
  const applyStats = (cell, val, align = CENTER) => {
    cell.value     = val != null ? val : null;
    cell.font      = WHITE_FONT(true);
    cell.fill      = TEAL_FILL;
    cell.border    = THIN_BORDER;
    cell.alignment = align;
  };

  // ── Row 1：大標題 ──────────────────────────────────────────────────────────
  const r1 = ws.getRow(1);
  r1.height = 30;
  applyHeader(r1.getCell(1), 'B項修繕');
  for (let c = 2; c <= 14; c++) applyHeader(r1.getCell(c), null);
  ws.mergeCells(1, 1, 1, 14);
  r1.commit();

  // ── Row 2：月份標題 ──────────────────────────────────────────────────────
  const r2 = ws.getRow(2);
  r2.height = 30;
  for (let c = 1; c <= 6; c++) applyHeader(r2.getCell(c), null);
  applyHeader(r2.getCell(7), `${rocYear}年${month}月份月結`);
  for (let c = 8; c <= 14; c++) applyHeader(r2.getCell(c), null);
  ws.mergeCells(2, 1, 2, 6);
  ws.mergeCells(2, 7, 2, 14);
  r2.commit();

  // ── Row 3：欄標題 ──────────────────────────────────────────────────────────
  const r3 = ws.getRow(3);
  r3.height = 30;
  const headers = ['案號','序號','報修日期','報修單號','站別','地址','故障說明','完工日期','合約項次','品項','單位','數量','單價','未稅小計'];
  headers.forEach((h, i) => applyHeader(r3.getCell(i + 1), h));
  r3.commit();

  // ── 資料列 ────────────────────────────────────────────────────────────────
  let excelRow = 4;

  cases.forEach((cas, caseIdx) => {
    const caseNo   = caseIdx + 1;
    const items    = cas.items;
    const startRow = excelRow;

    items.forEach((item, itemIdx) => {
      const row = ws.getRow(excelRow);
      row.height = 50;

      if (itemIdx === 0) {
        applyData(row.getCell(1), caseNo);
        applyData(row.getCell(2), 1);
        applyData(row.getCell(3), cas.報修日期);
        applyData(row.getCell(4), cas.JDM案號);
        applyData(row.getCell(5), cas.站別);
        applyData(row.getCell(6), cas.地址,     LEFT);
        applyData(row.getCell(7), cas.故障說明, LEFT);
        applyData(row.getCell(8), cas.完工日期);
      } else {
        applyData(row.getCell(1), null);
        applyData(row.getCell(2), itemIdx + 1);
        for (let c = 3; c <= 8; c++) applyData(row.getCell(c), null);
      }

      applyData(row.getCell(9),  item.合約項次);
      applyData(row.getCell(10), item.品項,              LEFT);
      applyData(row.getCell(11), item.單位);
      applyData(row.getCell(12), Number(item.數量)     || 0);
      applyData(row.getCell(13), Number(item.單價)     || 0);
      applyData(row.getCell(14), Number(item.未稅小計) || 0);

      row.commit();
      excelRow++;
    });

    // A 獨立合併，C–H 各自合併，B 不合併
    if (items.length > 1) {
      const endRow = startRow + items.length - 1;
      ws.mergeCells(startRow, 1, endRow, 1);
      for (let c = 3; c <= 8; c++) {
        ws.mergeCells(startRow, c, endRow, c);
      }
    }
  });

  // ── 統計行 ────────────────────────────────────────────────────────────────
  const statsRow = ws.getRow(excelRow);
  statsRow.height = 50;

  for (let col = 1; col <= 8; col++)  applyStats(statsRow.getCell(col), null);
  for (let col = 9; col <= 12; col++) applyStats(statsRow.getCell(col), col === 9 ? '總計' : null);
  ws.mergeCells(excelRow, 9, excelRow, 12);

  applyStats(statsRow.getCell(13), null);

  const totalSum = Math.round(
    cases.reduce((s, cas) => s + cas.items.reduce((ss, item) => ss + (Number(item.未稅小計) || 0), 0), 0)
  );
  applyStats(statsRow.getCell(14), null);
  statsRow.getCell(14).value = {
    formula: `SUM(N4:N${excelRow - 1})`,
    result:  totalSum,
  };

  statsRow.commit();

  // ── 輸出 ──────────────────────────────────────────────────────────────────
  const outBuf = await workbook.xlsx.writeBuffer();
  return new Uint8Array(outBuf);
}
