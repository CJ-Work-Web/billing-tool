/**
 * processTableB.js
 * 1. 以表A的合約項次加總未稅小計，比對表B的複價
 * 2. 金額不符的行標記紅色底色
 * 3. 將表A案號填入表B的備註欄
 */

const CONTRACT_CODE_RE = /^\d{2}-\d{2}-\d{2}$/;

function processTableB(tableBArrayBuffer, cases) {
  const wb = XLSX.read(new Uint8Array(tableBArrayBuffer), { type: 'array', cellStyles: true });
  const wsName = wb.SheetNames[0];
  const ws = wb.Sheets[wsName];

  // ── 建立表A的合約項次彙總 map ──────────────────────────────────────────
  // key: 合約項次代碼  value: { sum: number, caseNos: Set<number> }
  const tableAMap = new Map();
  cases.forEach((cas, idx) => {
    const caseNo = idx + 1;
    cas.items.forEach(item => {
      const code = item.合約項次;
      if (!code) return;
      if (!tableAMap.has(code)) tableAMap.set(code, { sum: 0, caseNos: new Set() });
      const entry = tableAMap.get(code);
      entry.sum += item.未稅小計;
      entry.caseNos.add(caseNo);
    });
  });

  // ── 掃描表B的 B 欄，找出合約項次列 ─────────────────────────────────────
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1:Z200');
  const results  = [];
  const warnings = [];

  for (let r = range.s.r; r <= range.e.r; r++) {
    const bAddr = XLSX.utils.encode_cell({ r, c: 1 }); // B 欄
    const bCell = ws[bAddr];
    if (!bCell) continue;

    const code = String(bCell.v ?? '').trim();
    if (!CONTRACT_CODE_RE.test(code)) continue;

    // 取得複價（H 欄 = index 7）
    const hAddr = XLSX.utils.encode_cell({ r, c: 7 });
    const hCell = ws[hAddr];
    const tableBAmount = hCell ? (Number(hCell.v) || 0) : 0;

    // 取得項目名稱（C 欄 = index 2）
    const cAddr = XLSX.utils.encode_cell({ r, c: 2 });
    const itemName = ws[cAddr] ? String(ws[cAddr].v || '') : '';

    const entry = tableAMap.get(code);
    const tableASum   = entry ? Math.round(entry.sum)  : null;
    const tableBRound = Math.round(tableBAmount);
    const isDiscrepant = entry !== null && entry !== undefined && tableASum !== tableBRound;
    const caseNos = entry ? Array.from(entry.caseNos).sort((a, b) => a - b) : [];

    results.push({ r, code, itemName, tableASum, tableBAmount: tableBRound, isDiscrepant, hasMatch: !!entry, caseNos });

    // ── 標記紅色底色（整行 A–I，欄 0–8）────────────────────────────────
    if (isDiscrepant) {
      const redFill = { patternType: 'solid', fgColor: { rgb: 'FFCCCC' } };
      for (let c = 0; c <= 8; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        if (!ws[addr]) ws[addr] = { v: '', t: 's', s: {} };
        if (!ws[addr].s) ws[addr].s = {};
        ws[addr].s = { ...ws[addr].s, fill: redFill };
      }
    }

    // ── 填入備註（I 欄 = index 8）───────────────────────────────────────
    if (caseNos.length > 0) {
      const iAddr = XLSX.utils.encode_cell({ r, c: 8 });
      if (!ws[iAddr]) ws[iAddr] = { v: '', t: 's', s: {} };
      ws[iAddr].v = caseNos.map(n => `案號${n}`).join('、');
      ws[iAddr].t = 's';
    }
  }

  // ── 檢查表A有但表B找不到的項次 ─────────────────────────────────────────
  tableAMap.forEach((entry, code) => {
    const found = results.some(res => res.code === code);
    if (!found) {
      warnings.push(`合約項次 ${code} 在表A有資料（案號 ${Array.from(entry.caseNos).map(n=>`${n}`).join('、')}），但表B中找不到對應項次`);
    }
  });

  const outputBytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });

  return { results, warnings, outputBytes };
}
