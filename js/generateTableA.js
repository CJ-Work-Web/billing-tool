/**
 * generateTableA.js
 * 依案件資料生成符合「提報修繕明細」格式的 Excel（表A）。
 *
 * 採用「載入範本→清除資料→填入新資料」方式，
 * 確保字型、填色、框線、對齊等樣式完全繼承自範本，
 * 不依賴 SheetJS 從零建立樣式（社群版從零建立樣式不穩定）。
 */

async function generateTableA(cases, rocYear, month) {

  // ── 1. 載入範本 ────────────────────────────────────────────────────────────
  const resp = await fetch('./template/tableA_template.xlsx');
  if (!resp.ok) throw new Error('無法載入表A範本（template/tableA_template.xlsx）');
  const buf = await resp.arrayBuffer();
  const wb  = XLSX.read(new Uint8Array(buf), { type: 'array', cellStyles: true });

  const wsName = wb.SheetNames[0];
  const ws     = wb.Sheets[wsName];

  // ── 2. 在刪除資料前，儲存各需要位置的樣式索引 ─────────────────────────────
  //   SheetJS 讀取時，cell.s 是 styles.xml 的 xf 索引（整數），
  //   只要把這個整數指派給新儲存格，寫出時就會套用同一個樣式定義。
  const readS = (r, c) => {
    const cell = ws[XLSX.utils.encode_cell({ r, c })];
    return cell ? cell.s : undefined;
  };

  // 第一筆資料列（row index 3 = Excel row 4）：各欄樣式
  const S1 = {};
  for (let c = 0; c <= 13; c++) S1[c] = readS(3, c);

  // 多品項第二列（row index 8 = Excel row 9，範本 case 5 的第2個品項）：各欄樣式
  const S2 = {};
  for (let c = 0; c <= 13; c++) S2[c] = readS(8, c);

  // 統計列（row index 31 = Excel row 32）：各欄樣式
  const SS = {};
  for (let c = 0; c <= 13; c++) SS[c] = readS(31, c);

  // ── 3. 清除所有資料列（row 3 以後） ───────────────────────────────────────
  const origRange = XLSX.utils.decode_range(ws['!ref'] || 'A1:N50');
  for (let r = 3; r <= Math.max(origRange.e.r, 45); r++) {
    for (let c = 0; c <= 13; c++) {
      delete ws[XLSX.utils.encode_cell({ r, c })];
    }
  }

  // ── 4. 清除資料合併（保留表頭 row 0–2 的合併） ────────────────────────────
  ws['!merges'] = (ws['!merges'] || []).filter(m => m.e.r < 3);

  // ── 5. 更新月份文字（G2 = row 1, col 6） ─────────────────────────────────
  const g2 = ws[XLSX.utils.encode_cell({ r: 1, c: 6 })];
  if (g2) g2.v = `${rocYear}年${month}月份月結`;

  // ── 6. 插入資料列 ─────────────────────────────────────────────────────────
  // styleIdx 傳 undefined 則取 S1[c] 預設值
  const setCell = (r, c, value, styleIdx, type) => {
    const addr = XLSX.utils.encode_cell({ r, c });
    const s = styleIdx !== undefined ? styleIdx : S1[c];
    const t = type !== undefined ? type : (typeof value === 'number' ? 'n' : 's');
    ws[addr] = { v: (value !== undefined && value !== null) ? value : '', t, s };
  };

  let currentRow = 3;

  cases.forEach((cas, caseIdx) => {
    const caseNo   = caseIdx + 1;
    const items    = cas.items;
    const startRow = currentRow;

    items.forEach((item, itemIdx) => {
      const r = currentRow;

      if (itemIdx === 0) {
        // 第一個品項列：填入 A–H 案件資訊
        setCell(r, 0, caseNo,       S1[0],  'n');
        setCell(r, 1, 1,            S1[1],  'n');
        setCell(r, 2, cas.報修日期,  S1[2]);
        setCell(r, 3, cas.JDM案號,   S1[3]);
        setCell(r, 4, cas.站別,      S1[4]);
        setCell(r, 5, cas.地址,      S1[5]);
        setCell(r, 6, cas.故障說明,  S1[6]);
        setCell(r, 7, cas.完工日期,  S1[7]);
      } else {
        // 後續品項列：A 和 C–H 填空白（合併後不顯示）
        setCell(r, 0, '', S2[0] ?? S1[0]);
        setCell(r, 1, itemIdx + 1, S2[1] ?? S1[1], 'n');
        for (let c = 2; c <= 7; c++) {
          setCell(r, c, '', S2[c] ?? S1[c]);
        }
      }

      // I–N 每列都填品項資料
      setCell(r, 8,  item.合約項次, S1[8]);
      setCell(r, 9,  item.品項,     S1[9]);
      setCell(r, 10, item.單位,     S1[10]);
      setCell(r, 11, item.數量,     S1[11], 'n');
      setCell(r, 12, item.單價,     S1[12], 'n');
      setCell(r, 13, item.未稅小計, S1[13], 'n');

      currentRow++;
    });

    // 多品項案件：A 欄獨立合併，C–H 各自獨立合併，B 不合併
    if (items.length > 1) {
      const endRow = startRow + items.length - 1;
      ws['!merges'].push({ s: { r: startRow, c: 0 }, e: { r: endRow, c: 0 } });
      for (let c = 2; c <= 7; c++) {
        ws['!merges'].push({ s: { r: startRow, c }, e: { r: endRow, c } });
      }
    }
  });

  // ── 7. 統計行 ─────────────────────────────────────────────────────────────
  const statsRow      = currentRow;
  const lastDataExcel = currentRow; // 0-indexed statsRow = 1-indexed 最後資料列號

  const totalSum = Math.round(
    cases.reduce((s, cas) => s + cas.items.reduce((ss, item) => ss + (Number(item.未稅小計) || 0), 0), 0)
  );

  // A–H：空白，用範本統計列樣式
  for (let c = 0; c <= 7; c++) {
    setCell(statsRow, c, '', SS[c] ?? S1[c]);
  }

  // I–L：合併「總計」，用範本 I32 樣式
  const sTotalText = SS[8] ?? S1[0];
  for (let c = 8; c <= 11; c++) {
    setCell(statsRow, c, c === 8 ? '總計' : '', sTotalText);
  }
  ws['!merges'].push({ s: { r: statsRow, c: 8 }, e: { r: statsRow, c: 11 } });

  // M：空白，用範本 M32 樣式
  setCell(statsRow, 12, '', SS[12] ?? S1[0]);

  // N：SUM 公式（含預算值），用範本 N32 樣式
  const nAddr = XLSX.utils.encode_cell({ r: statsRow, c: 13 });
  ws[nAddr] = {
    f: `SUM(N4:N${lastDataExcel})`,
    v: totalSum,
    t: 'n',
    s: SS[13] ?? S1[0],
  };

  currentRow++; // 統計行佔一列

  // ── 8. 更新工作表屬性 ─────────────────────────────────────────────────────
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: currentRow - 1, c: 13 } });

  // 列高：保留表頭 3 列原始高度，資料列統一設 49.95pt
  const origRows = ws['!rows'] || [];
  const newRows  = [];
  for (let r = 0; r < 3; r++) newRows[r] = origRows[r] || { hpt: 49.95 };
  for (let r = 3; r < currentRow; r++) newRows[r] = { hpt: 49.95 };
  ws['!rows'] = newRows;

  // ── 9. 重命名工作表 ───────────────────────────────────────────────────────
  const newName = `${rocYear}.${String(month).padStart(2, '0')}`;
  if (wsName !== newName) {
    wb.SheetNames[0] = newName;
    wb.Sheets[newName] = wb.Sheets[wsName];
    delete wb.Sheets[wsName];
  }

  return XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
}
