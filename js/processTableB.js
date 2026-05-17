/**
 * processTableB.js
 * 1. 以表A的合約項次加總未稅小計，比對表B的複價
 * 2. 金額不符的行標記紅色底色
 * 3. 將表A案號填入表B的備註欄
 *
 * 改用 ExcelJS 讀寫，確保原始格式完整保留。
 */

const CONTRACT_CODE_RE = /^\d{2}-\d{2}-\d{2}$/;

async function processTableB(tableBArrayBuffer, cases) {
  const ExcelJS  = window.ExcelJS;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(tableBArrayBuffer);

  // ── 建立表A的合約項次彙總 map ──────────────────────────────────────────
  // key: 合約項次代碼  value: { sum: number, caseNos: Set<number>, units: Set<string> }
  const tableAMap = new Map();
  cases.forEach((cas, idx) => {
    const caseNo = idx + 1;
    cas.items.forEach(item => {
      const code = item.合約項次;
      if (!code) return;
      if (!tableAMap.has(code)) tableAMap.set(code, { sum: 0, caseNos: new Set(), units: new Set() });
      const entry = tableAMap.get(code);
      entry.sum += item.未稅小計;
      entry.caseNos.add(caseNo);
      if (item.單位) entry.units.add(item.單位);
    });
  });

  // ── 取第一個工作表（當月資料）─────────────────────────────────────────
  const ws = workbook.worksheets[0];
  const results  = [];
  const warnings = [];

  ws.eachRow((row, rowNumber) => {
    // B 欄（index 2）
    const bCell = row.getCell(2);
    const rawB  = bCell.value;
    if (rawB === null || rawB === undefined) return;
    const code = String(rawB).trim();
    if (!CONTRACT_CODE_RE.test(code)) return;

    // 取複價（H 欄 = 8）— 可能是公式 {formula, result}
    const hCell = row.getCell(8);
    let tableBAmount = 0;
    const hv = hCell.value;
    if (hv !== null && hv !== undefined) {
      tableBAmount = (typeof hv === 'object' && hv?.result !== undefined)
        ? Number(hv.result) || 0
        : Number(hv) || 0;
    }

    // 取項目名稱（C 欄 = 3）
    const cCell    = row.getCell(3);
    const itemName = cCell.value ? String(cCell.value) : '';

    // 取單位（D 欄 = 4）
    const dCell    = row.getCell(4);
    const tableBUnit = dCell.value ? String(dCell.value).trim() : '';

    const entry       = tableAMap.get(code);
    const tableASum   = entry ? Math.round(entry.sum)  : null;
    const tableBRound = Math.round(tableBAmount);
    const isDiscrepant = !!entry && tableASum !== tableBRound;
    const caseNos      = entry
      ? Array.from(entry.caseNos).sort((a, b) => a - b)
      : [];
    const tableAUnits  = entry ? Array.from(entry.units) : [];

    results.push({ r: rowNumber, code, itemName, tableASum, tableBAmount: tableBRound, tableBUnit, tableAUnits, isDiscrepant, hasMatch: !!entry, caseNos });

    // ── 標記紅色底色（A–I 欄，欄 1–9）────────────────────────────────────
    if (isDiscrepant) {
      const redFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCCCC' } };
      for (let c = 1; c <= 9; c++) {
        row.getCell(c).fill = redFill;
      }
    }

    // ── 填入備註（I 欄 = 9）──────────────────────────────────────────────
    if (caseNos.length > 0) {
      row.getCell(9).value = caseNos.map(n => `案號${n}`).join('、');
    }
  });

  // ── 檢查表A有但表B找不到的項次 ─────────────────────────────────────────
  tableAMap.forEach((entry, code) => {
    const found = results.some(res => res.code === code);
    if (!found) {
      warnings.push(
        `合約項次 ${code} 在表A有資料（案號 ${Array.from(entry.caseNos).map(n => `${n}`).join('、')}），但表B中找不到對應項次`
      );
    }
  });

  const outBuf = await workbook.xlsx.writeBuffer();
  return { results, warnings, outputBytes: new Uint8Array(outBuf) };
}
