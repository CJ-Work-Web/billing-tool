/**
 * generateTableA.js
 * 依案件資料生成符合「提報修繕明細」格式的 Excel（表A）。
 *
 * 完全不依賴範本載入，使用 ExcelJS 直接建立並套用硬編碼樣式。
 */

async function generateTableA(cases, rocYear, month) {

  const ExcelJS  = window.ExcelJS;
  const workbook = new ExcelJS.Workbook();
  const sheetName = `${rocYear}.${String(month).padStart(2, '0')}`;
  const ws = workbook.addWorksheet(sheetName);

  // ── 欄寬 & 隱藏 ──────────────────────────────────────────────────────────────
  ws.columns = [
    { width:  8.33 },                  // A  案號
    { width:  8.89, hidden: true  },   // B  序號（隱藏）
    { width: 30.0  },                  // C  報修日期
    { width: 27.11 },                  // D  報修單號
    { width: 25.66 },                  // E  站別
    { width: 30.33 },                  // F  地址
    { width: 80.66 },                  // G  故障說明
    { width: 25.66 },                  // H  完工日期
    { width: 24.55 },                  // I  合約項次
    { width: 24.55, hidden: true  },   // J  品項（隱藏）
    { width: 10.66 },                  // K  單位
    { width: 14.11 },                  // L  數量
    { width: 17.66, hidden: true  },   // M  單價（隱藏）
    { width: 30.66 },                  // N  未稅金額小計
  ];

  // ── 樣式定義 ──────────────────────────────────────────────────────────────────
  const TEAL_FILL = {
    type: 'pattern', pattern: 'solid',
    fgColor: { argb: 'FF46BDC6' },
  };
  const LB_FILL = {                    // 橙色（統計行）
    type: 'pattern', pattern: 'solid',
    fgColor: { argb: 'FFFFE1CC' },
  };

  const TITLE_FONT  = () => ({ name: '微軟正黑體', size: 24, bold: true });
  const HDR_FONT    = () => ({ name: '微軟正黑體', size: 16, bold: true, color: { argb: 'FFFFFFFF' } });
  const DATA_FONT   = () => ({ name: '微軟正黑體', size: 16 });
  const STATS_FONT  = () => ({ name: '微軟正黑體', size: 16, bold: true });

  const THIN_BORDER = {
    top:    { style: 'thin' },
    left:   { style: 'thin' },
    bottom: { style: 'thin' },
    right:  { style: 'thin' },
  };
  const CENTER = { horizontal: 'center', vertical: 'middle', wrapText: true };
  const LEFT   = { horizontal: 'left',   vertical: 'middle', wrapText: true };
  const RIGHT  = { horizontal: 'right',  vertical: 'middle', wrapText: true };

  // N 欄：會計格式（$符號 + 千位分號 + 右對齊）
  const N_FMT = '_-"$"* #,##0_-;\\-"$"* #,##0_-;_-"$"* "-"??_-;_-@_-';

  // ── Row 1：大標題「提報修繕明細」────────────────────────────────────────────
  const r1 = ws.getRow(1);
  r1.height = 75;    // 目標 50pt × 1.5 = 75（ExcelJS px→pt 換算）
  for (let c = 1; c <= 14; c++) {
    const cell = r1.getCell(c);
    cell.font      = TITLE_FONT();
    cell.border    = THIN_BORDER;
    cell.alignment = CENTER;
  }
  r1.getCell(1).value = '提報修繕明細';
  ws.mergeCells(1, 1, 1, 14);
  r1.commit();

  // ── Row 2：公司名稱 | 月份標題 ──────────────────────────────────────────────
  const r2 = ws.getRow(2);
  r2.height = 75;    // 目標 50pt × 1.5 = 75
  for (let c = 1; c <= 14; c++) {
    const cell = r2.getCell(c);
    cell.font      = TITLE_FONT();
    cell.border    = THIN_BORDER;
    cell.alignment = CENTER;
  }
  r2.getCell(1).value = '晟晁物業管理顧問股份有限公司';
  r2.getCell(7).value = `${rocYear}年${month}月份月結`;
  ws.mergeCells(2, 1, 2, 6);
  ws.mergeCells(2, 7, 2, 14);
  r2.commit();

  // ── Row 3：欄標題 ────────────────────────────────────────────────────────────
  const r3 = ws.getRow(3);
  r3.height = 75;    // 目標 50pt × 1.5 = 75
  const headers = [
    '案號','序號','報修日期','報修單號','站別','地址',
    '故障說明','完工日期','合約項次','品項','單位','數量','單價','未稅金額小計(1)',
  ];
  headers.forEach((h, i) => {
    const cell = r3.getCell(i + 1);
    cell.value     = h;
    cell.font      = HDR_FONT();
    cell.fill      = TEAL_FILL;
    cell.border    = THIN_BORDER;
    cell.alignment = CENTER;
  });
  r3.commit();

  // ── 資料列 ────────────────────────────────────────────────────────────────────
  let excelRow = 4;

  /**
   * 計算故障說明文字需要的額外列高（pt display）。
   * G 欄寬 80.66，中文字 2 單位寬 → 每行 ~40 字，行高 ~20pt，上下內距 10pt。
   * @param {string} text        故障說明文字
   * @param {number} availPt     合併儲存格可用總高度（display pt）
   * @returns {number} 需要額外增加的 pt（display），可能為 0
   */
  function extraHeightPt(text, availPt) {
    // G 欄寬 80.66，微軟正黑體 16pt 中文字 ≈ 21px，
    // Calibri 11pt digit ≈ 7px → 欄寬 ≈ 564px → 每行 ≈ 26 中文字
    const CHARS_PER_LINE = 24;   // 保守值（26×0.9），兼顧混排與寬字符
    const LINE_HEIGHT_PT = 24;   // 16pt 字體 × 1.5 行距
    const PADDING_PT     = 8;    // 儲存格上下內距
    const lines = text.split('\n').reduce(
      (s, l) => s + Math.max(1, Math.ceil(l.length / CHARS_PER_LINE)), 0
    );
    const requiredPt = lines * LINE_HEIGHT_PT + PADDING_PT;
    return Math.max(0, requiredPt - availPt);
  }

  cases.forEach((cas, caseIdx) => {
    const caseNo      = caseIdx + 1;
    const items       = cas.items;
    const startRow    = excelRow;
    const isSingle    = items.length === 1;
    const k           = items.length;

    // 計算每列需要的高度（pt display）
    const basePt      = isSingle ? 100 : 50;   // 預設每列顯示高度
    const availPt     = k * basePt;             // 合併儲存格可用總高度
    const extraTotal  = extraHeightPt(cas.故障說明, availPt);
    const extraPerRow = extraTotal / k;          // 平均分配給每列
    const heightPt    = basePt + extraPerRow;   // 每列最終顯示高度（pt）

    items.forEach((item, itemIdx) => {
      const row = ws.getRow(excelRow);
      row.height = heightPt * 1.5;             // ExcelJS px→pt × 1.5

      const dc = (col, val, align = CENTER, numFmt = null) => {
        const cell = row.getCell(col);
        cell.value     = val != null ? val : null;
        cell.font      = DATA_FONT();
        cell.border    = THIN_BORDER;
        cell.alignment = align;
        if (numFmt) cell.numFmt = numFmt;
      };

      if (itemIdx === 0) {
        dc(1, caseNo);
        dc(2, 1);
        dc(3, cas.報修日期);
        dc(4, cas.JDM案號);
        dc(5, cas.站別);
        dc(6, cas.地址,     LEFT);
        dc(7, cas.故障說明, LEFT);
        dc(8, cas.完工日期);
      } else {
        dc(1, null);
        dc(2, itemIdx + 1);
        for (let c = 3; c <= 8; c++) dc(c, null);
      }

      dc(9,  item.合約項次);
      dc(10, item.品項,              LEFT);
      dc(11, item.單位);
      dc(12, Number(item.數量)     || 0);
      dc(13, Number(item.單價)     || 0);
      dc(14, Number(item.未稅小計) || 0, RIGHT, N_FMT);

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

  // ── 統計行 ────────────────────────────────────────────────────────────────────
  const sr = ws.getRow(excelRow);
  sr.height = 97.5;  // 65pt × 1.5 = 97.5

  // A–H：只保留上格線（作為資料區邊界），無填色
  const TOP_BORDER = { top: { style: 'thin' } };
  for (let col = 1; col <= 8; col++) {
    const cell = sr.getCell(col);
    cell.font   = STATS_FONT();
    cell.border = TOP_BORDER;
  }

  // I–L：淺藍，合併，顯示「總計」
  for (let col = 9; col <= 12; col++) {
    const cell = sr.getCell(col);
    cell.font      = STATS_FONT();
    cell.fill      = LB_FILL;
    cell.border    = THIN_BORDER;
    cell.alignment = CENTER;
  }
  sr.getCell(9).value = '總計';
  ws.mergeCells(excelRow, 9, excelRow, 12);

  // M：淺藍（隱藏欄）
  const mCell = sr.getCell(13);
  mCell.font   = STATS_FONT();
  mCell.fill   = LB_FILL;
  mCell.border = THIN_BORDER;

  // N：淺藍，$ 格式，SUM 公式
  const totalSum = Math.round(
    cases.reduce((s, cas) =>
      s + cas.items.reduce((ss, it) => ss + (Number(it.未稅小計) || 0), 0), 0)
  );
  const nCell = sr.getCell(14);
  nCell.font      = STATS_FONT();
  nCell.fill      = LB_FILL;
  nCell.border    = THIN_BORDER;
  nCell.alignment = RIGHT;
  nCell.numFmt    = N_FMT;
  nCell.value     = { formula: `SUM(N4:N${excelRow - 1})`, result: totalSum };

  sr.commit();

  // ── 輸出 ──────────────────────────────────────────────────────────────────────
  const outBuf = await workbook.xlsx.writeBuffer();
  return new Uint8Array(outBuf);
}
