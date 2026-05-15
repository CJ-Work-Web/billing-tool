/**
 * generateTableA.js
 * 依案件資料生成符合「提報修繕明細」格式的 Excel（表A）。
 *
 * 使用 ExcelJS 載入範本並套用正確樣式：
 *   1. fetch 範本 → ExcelJS 解析（完整保留字型/填色/框線/對齊）
 *   2. 從範本第一筆資料列讀取各欄樣式
 *   3. 清除資料列，填入新資料並逐欄套用樣式
 *   4. 新增統計行（套用範本統計列樣式）
 */

async function generateTableA(cases, rocYear, month) {

  // ── 1. 載入範本 ────────────────────────────────────────────────────────────
  const resp = await fetch('./template/tableA_template.xlsx');
  if (!resp.ok) throw new Error('無法載入表A範本（template/tableA_template.xlsx）');
  const buf = await resp.arrayBuffer();

  const ExcelJS  = window.ExcelJS;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buf);

  const ws = workbook.worksheets[0];

  // ── 2. 在刪除資料前，讀取並儲存各列的樣式 ─────────────────────────────────
  // ExcelJS 的 cell.font / .fill / .border / .alignment 是完整的樣式物件，
  // 深拷貝後可直接指派給新儲存格，確保樣式完全正確。
  const saveRowStyles = (excelRowNum) => {
    const row    = ws.getRow(excelRowNum);
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

  const dataStyles  = saveRowStyles(4);  // Excel row 4 = 第一筆資料列
  const statsStyles = saveRowStyles(32); // Excel row 32 = 統計列

  // ── 3. 清除所有資料列（Excel row 4 以後）──────────────────────────────────
  const lastRow   = ws.rowCount;
  const deleteCount = Math.max(lastRow - 3 + 10, 10); // 確保清除乾淨
  ws.spliceRows(4, deleteCount);

  // ── 4. 更新月份文字（G2）──────────────────────────────────────────────────
  ws.getCell('G2').value = `${rocYear}年${month}月份月結`;

  // ── 5. 輔助：為儲存格套用樣式 ─────────────────────────────────────────────
  const applyStyle = (cell, col, styleMap) => {
    const s = styleMap[col];
    if (!s) return;
    if (s.font)      cell.font      = JSON.parse(JSON.stringify(s.font));
    if (s.fill)      cell.fill      = JSON.parse(JSON.stringify(s.fill));
    if (s.border)    cell.border    = JSON.parse(JSON.stringify(s.border));
    if (s.alignment) cell.alignment = JSON.parse(JSON.stringify(s.alignment));
    if (s.numFmt)    cell.numFmt    = s.numFmt;
  };

  // ── 6. 插入資料列 ─────────────────────────────────────────────────────────
  let excelRow = 4;

  cases.forEach((cas, caseIdx) => {
    const caseNo   = caseIdx + 1;
    const items    = cas.items;
    const startRow = excelRow;

    items.forEach((item, itemIdx) => {
      const row = ws.getRow(excelRow);
      row.height = 49.95;

      // 套用各欄樣式（從範本資料列複製）
      for (let col = 1; col <= 14; col++) {
        applyStyle(row.getCell(col), col, dataStyles);
      }

      if (itemIdx === 0) {
        // 第一個品項：填入 A–H 案件欄位
        row.getCell(1).value  = caseNo;        // A 案號
        row.getCell(2).value  = 1;              // B 序號
        row.getCell(3).value  = cas.報修日期;  // C
        row.getCell(4).value  = cas.JDM案號;   // D
        row.getCell(5).value  = cas.站別;      // E
        row.getCell(6).value  = cas.地址;      // F
        row.getCell(7).value  = cas.故障說明;  // G
        row.getCell(8).value  = cas.完工日期;  // H
      } else {
        // 後續品項：A 和 C–H 留空（被合併覆蓋）
        row.getCell(1).value  = null;
        row.getCell(2).value  = itemIdx + 1;   // B 序號（不合併）
        for (let col = 3; col <= 8; col++) row.getCell(col).value = null;
      }

      // I–N 品項欄位（每列都填）
      row.getCell(9).value  = item.合約項次;
      row.getCell(10).value = item.品項;
      row.getCell(11).value = item.單位;
      row.getCell(12).value = Number(item.數量)    || 0;
      row.getCell(13).value = Number(item.單價)    || 0;
      row.getCell(14).value = Number(item.未稅小計) || 0;

      row.commit();
      excelRow++;
    });

    // 多品項案件：A 欄獨立合併，C–H 各自獨立合併，B 不合併
    if (items.length > 1) {
      const endRow = startRow + items.length - 1;
      ws.mergeCells(startRow, 1, endRow, 1);          // A
      for (let c = 3; c <= 8; c++) {
        ws.mergeCells(startRow, c, endRow, c);         // C, D, E, F, G, H
      }
    }
  });

  // ── 7. 統計行 ─────────────────────────────────────────────────────────────
  const statsRow = ws.getRow(excelRow);
  statsRow.height = 49.95;

  // A–H：空白，套用範本統計列樣式
  for (let col = 1; col <= 8; col++) {
    applyStyle(statsRow.getCell(col), col, statsStyles);
    statsRow.getCell(col).value = null;
  }

  // I–L：合併「總計」，套用範本統計列樣式
  for (let col = 9; col <= 12; col++) {
    applyStyle(statsRow.getCell(col), col, statsStyles);
    statsRow.getCell(col).value = col === 9 ? '總計' : null;
  }
  ws.mergeCells(excelRow, 9, excelRow, 12);

  // M：空白，套用範本統計列樣式
  applyStyle(statsRow.getCell(13), 13, statsStyles);
  statsRow.getCell(13).value = null;

  // N：SUM 公式
  applyStyle(statsRow.getCell(14), 14, statsStyles);
  const totalSum = Math.round(
    cases.reduce((s, cas) => s + cas.items.reduce((ss, item) => ss + (Number(item.未稅小計) || 0), 0), 0)
  );
  statsRow.getCell(14).value = {
    formula: `SUM(N4:N${excelRow - 1})`,
    result:  totalSum,
  };

  statsRow.commit();

  // ── 8. 重命名工作表 ───────────────────────────────────────────────────────
  ws.name = `${rocYear}.${String(month).padStart(2, '0')}`;

  // ── 9. 輸出 ───────────────────────────────────────────────────────────────
  const outBuf = await workbook.xlsx.writeBuffer();
  return new Uint8Array(outBuf);
}
