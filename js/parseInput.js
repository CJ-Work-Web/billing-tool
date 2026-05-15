/**
 * parseInput.js
 * 解析「計價明細原始資料」Excel，依結報日期（民國年月）篩選，回傳案件陣列。
 */

function parseInputFile(arrayBuffer, rocYear, month) {
  const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  if (rows.length < 2) throw new Error('Excel 檔案內容為空或格式不正確');

  // 第一列為欄位標題
  const headers = rows[0].map(h => String(h).trim());
  const requiredCols = ['結報日期', '報修日期', 'JDM案號', '站別', '地址', '故障說明', '完工日期', '合約項次', '品項', '單位', '數量', '單價', '未稅小計'];
  const missing = requiredCols.filter(c => !headers.includes(c));
  if (missing.length > 0) {
    throw new Error(`Excel 格式不符，缺少欄位：${missing.join('、')}\n請確認已使用「計價明細」模式匯出`);
  }

  const idx = {};
  headers.forEach((h, i) => { idx[h] = i; });

  // 篩選條件：結報日期年月符合
  const ceYear = rocYear + 1911;
  const monthPadded = String(month).padStart(2, '0');
  const prefix = `${ceYear}/${monthPadded}`;
  const prefixAlt = `${ceYear}/${month}/`; // 非補零格式備用

  const allRows = rows.slice(1).filter(r => {
    const d = String(r[idx['結報日期']] || '').trim();
    return d.startsWith(prefix) || d.startsWith(prefixAlt);
  });

  if (allRows.length === 0) {
    throw new Error(`找不到 ${rocYear}年${month}月 的結報資料\n請確認結報日期欄位有值，且年月符合選擇`);
  }

  // 依 JDM案號 分組（相同案號合併為一個案件）
  // 若 JDM案號 為空，以「報修日期_地址」作為分組 key
  const caseMap = new Map();

  allRows.forEach(r => {
    const jdmNo = String(r[idx['JDM案號']] || '').trim();
    const reportDate = String(r[idx['報修日期']] || '').trim();
    const address = String(r[idx['地址']] || '').trim();
    const key = jdmNo || `${reportDate}_${address}`;

    if (!caseMap.has(key)) {
      caseMap.set(key, {
        結報日期: String(r[idx['結報日期']] || '').trim(),
        報修日期: reportDate,
        JDM案號: jdmNo,
        站別: String(r[idx['站別']] || '').trim(),
        地址: address,
        故障說明: String(r[idx['故障說明']] || '').trim(),
        完工日期: String(r[idx['完工日期']] || '').trim(),
        items: [],
      });
    }

    const qty = Number(r[idx['數量']]) || 0;
    const price = Number(r[idx['單價']]) || 0;
    caseMap.get(key).items.push({
      合約項次: String(r[idx['合約項次']] || '').trim(),
      品項: String(r[idx['品項']] || '').trim(),
      單位: String(r[idx['單位']] || '').trim(),
      數量: qty,
      單價: price,
      未稅小計: Number(r[idx['未稅小計']]) || (qty * price),
    });
  });

  // 依報修日期升冪排序
  const cases = Array.from(caseMap.values()).sort((a, b) =>
    String(a.報修日期).localeCompare(String(b.報修日期))
  );

  const itemCount = cases.reduce((s, c) => s + c.items.length, 0);
  return { cases, itemCount };
}
