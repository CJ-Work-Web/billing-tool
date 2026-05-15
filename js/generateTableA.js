/**
 * generateTableA.js
 * 依案件資料生成符合「提報修繕明細」格式的 Excel（表A）。
 *
 * 策略：
 *   1. fetch 範本 → ExcelJS 解析，僅用來「讀取」樣式與標題列內容
 *   2. 建立全新 workbook + worksheet（避免 spliceRows 在有合併儲存格時損壞檔案）
 *   3. 將範本的標題列（rows 1–3）原樣複製到新 worksheet
 *   4. 從範本 row 4 讀取資料列樣式，row 32 讀取統計列樣式
 *   5. 在新 worksheet 填入資料並套用樣式
 */

async function generateTableA(cases, rocYear, month) {

  // ── 1. 載入範本（唯讀，不做任何修改）────────────────────────────────────
  const resp = await fetch('./template/tableA_template.xlsx');
  if (!resp.ok) throw new Error('無法載入表A範本（template/tableA_template.xlsx）');
  const buf = await resp.arrayBuffer();

  const ExcelJS     = window.ExcelJS;
  const templateWb  = new ExcelJS.Workbook();
  await templateWb.xlsx.load(buf);
  const templateWs  = templateWb.worksheets[0];

  // ── 2. 從範本讀取樣式（不修改範本）──────────────────────────────────────
  const saveRowStyles = (rowNum) => {
    const row    = templateWs.getRow(rowNum);
    const styles = {};
    for (let col = 1; col <= 14; col++) {
      const c = row.getCell(col);
      styles[col] = {
        font:      c.font      ? JSON.parse(JSON.stringify(c.font))      : null,
        fill:      c.fill      ? JSON.parse(JSON.stringify(c.fill))      : null,
        border:    c.border    ? JSON.parse(JSON.stringify(c.border))    : null,
        alignment: c.alignment ? JSON.parse(JSON.stringify(c.alignment)) : null,
        numFmt:    c.numFmt    || null,
      };
    }
    return styles;
  };

  const dataStyles  = saveRowStyles(4);   // Excel row 4 = 第一筆資料列
  const statsStyles = saveRowStyles(32);  // Excel row 32 = 統計列

  // ── 3. 建立全新 workbook（完全避開 spliceRows 的合併衝突問題）────────────
  const workbook   = new ExcelJS.Workbook();
  const sheetName  = `${rocYear}.${String(month).padStart(2, '0')}`;
  const ws         = workbook.addWorksheet(sheetName);

  // ── 4. 複製欄寬 ───────────────────────────────────────────────────────────
  templateWs.columns.forEach((col, idx) => {
    const dstCol = ws.getColumn(idx + 1);
    if (col.width  != null) dstCol.width  = col.width;
    if (col.hidden)         dstCol.hidden = col.hidden;
  });

  // ── 5. 複製標題列（rows 1–3）的值與樣式 ──────────────────────────────────
  const copyCell = (src, dst) => {
    dst.value = src.value;
    if (src.font)      dst.font      = JSON.parse(JSON.stringify(src.font));
    if (src.fill)      dst.fill      = JSON.parse(JSON.stringify(src.fill));
    if (src.border)    dst.border    = JSON.parse(JSON.stringify(src.border));
    if (src.alignment) dst.alignment = JSON.parse(JSON.stringify(src.alignment));
    if (src.numFmt)    dst.numFmt    = src.numFmt;
  };

  for (let rowNum = 1; rowNum <= 3; rowNum++) {
    const srcRow = templateWs.getRow(rowNum);
    const dstRow = ws.getRow(rowNum);
    if (srcRow.height) dstRow.height = srcRow.height;
    for (let col = 1; col <= 14; col++) {
      copyCell(srcRow.getCell(col), dstRow.getCell(col));
    }
    dstRow.commit();
  }

  // ── 6. 複製標題列的合併儲存格（僅 rows 1–3）────────────────────────────
  try {
    Object.values(templateWs._merges || {}).forEach(merge => {
      const m = merge && merge.model;
      if (m && m.bottom <= 3) {
        ws.mergeCells(m.top, m.left, m.bottom, m.right);
      }
    });
  } catch(e) { console.warn('複製標題合併時發生例外:', e); }

  // ── 7. 更新月份文字（G2）──────────────────────────────────────────────────
  ws.getCell('G2').value = `${rocYear}年${month}月份月結`;

  // ── 8. 輔助：套用樣式 ────────────────────────────────────────────────────
  const applyStyle = (cell, col, styleMap) => {
    const s = styleMap[col];
    if (!s) return;
    if (s.font)      cell.font      = JSON.parse(JSON.stringify(s.font));
    if (s.fill)      cell.fill      = JSON.parse(JSON.stringify(s.fill));
    if (s.border)    cell.border    = JSON.parse(JSON.stringify(s.border));
    if (s.alignment) cell.alignment = JSON.parse(JSON.stringify(s.alignment));
    if (s.numFmt)    cell.numFmt    = s.numFmt;
  };

  // ── 9. 插入資料列 ─────────────────────────────────────────────────────────
  let excelRow = 4;

  cases.forEach((cas, caseIdx) => {
    const caseNo   = caseIdx + 1;
    const items    = cas.items;
    const startRow = excelRow;

    items.forEach((item, itemIdx) => {
      const row = ws.getRow(excelRow);
      row.height = 49.95;

      for (let col = 1; col <= 14; col++) {
        applyStyle(row.getCell(col), col, dataStyles);
      }

      if (itemIdx === 0) {
        row.getCell(1).value  = caseNo;       // A 案號
        row.getCell(2).value  = 1;             // B 序號
        row.getCell(3).value  = cas.報修日期; // C
        row.getCell(4).value  = cas.JDM案號;  // D
        row.getCell(5).value  = cas.站別;     // E
        row.getCell(6).value  = cas.地址;     // F
        row.getCell(7).value  = cas.故障說明; // G
        row.getCell(8).value  = cas.完工日期; // H
      } else {
        row.getCell(1).value  = null;
        row.getCell(2).value  = itemIdx + 1;  // B 序號（不合併）
        for (let col = 3; col <= 8; col++) row.getCell(col).value = null;
      }

      // I–N 品項欄位
      row.getCell(9).value  = item.合約項次;
      row.getCell(10).value = item.品項;
      row.getCell(11).value = item.單位;
      row.getCell(12).value = Number(item.數量)     || 0;
      row.getCell(13).value = Number(item.單價)     || 0;
      row.getCell(14).value = Number(item.未稅小計) || 0;

      row.commit();
      excelRow++;
    });

    // 多品項：A 欄獨立合併，C–H 各自獨立合併，B 不合併
    if (items.length > 1) {
      const endRow = startRow + items.length - 1;
      ws.mergeCells(startRow, 1, endRow, 1);           // A
      for (let c = 3; c <= 8; c++) {
        ws.mergeCells(startRow, c, endRow, c);          // C D E F G H
      }
    }
  });

  // ── 10. 統計行 ────────────────────────────────────────────────────────────
  const statsRow = ws.getRow(excelRow);
  statsRow.height = 49.95;

  for (let col = 1; col <= 8; col++) {
    applyStyle(statsRow.getCell(col), col, statsStyles);
    statsRow.getCell(col).value = null;
  }
  for (let col = 9; col <= 12; col++) {
    applyStyle(statsRow.getCell(col), col, statsStyles);
    statsRow.getCell(col).value = col === 9 ? '總計' : null;
  }
  ws.mergeCells(excelRow, 9, excelRow, 12);

  applyStyle(statsRow.getCell(13), 13, statsStyles);
  statsRow.getCell(13).value = null;

  applyStyle(statsRow.getCell(14), 14, statsStyles);
  const totalSum = Math.round(
    cases.reduce((s, cas) => s + cas.items.reduce((ss, item) => ss + (Number(item.未稅小計) || 0), 0), 0)
  );
  statsRow.getCell(14).value = {
    formula: `SUM(N4:N${excelRow - 1})`,
    result:  totalSum,
  };
  statsRow.commit();

  // ── 11. 輸出 ──────────────────────────────────────────────────────────────
  const outBuf = await workbook.xlsx.writeBuffer();
  return new Uint8Array(outBuf);
}
