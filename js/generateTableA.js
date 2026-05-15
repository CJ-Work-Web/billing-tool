/**
 * generateTableA.js
 * 依案件資料生成符合「提報修繕明細」格式的 Excel（表A）。
 * 格式依據範本實測：微軟正黑體 16pt 資料，外框 medium，內框 thin。
 * 合併方式：A 欄獨立合併、C–H 欄各自獨立合併、B 欄不合併（顯示品項序號）。
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

  // ── 欄位對齊（依範本實測） ────────────────────────────────────────────────
  // F(5)=地址、G(6)=故障說明、J(9)=品項 → 靠左 + 自動換行
  // E(4)=站別、K(10)=單位、L(11)=數量 → 置中 + 自動換行
  // 其餘 → 置中
  function colAlignment(c) {
    if (c === 5 || c === 6 || c === 9) {
      return { horizontal: 'left',   vertical: 'center', wrapText: true };
    }
    if (c === 4 || c === 10 || c === 11) {
      return { horizontal: 'center', vertical: 'center', wrapText: true };
    }
    return { horizontal: 'center', vertical: 'center', wrapText: false };
  }

  // ── 樣式 ──────────────────────────────────────────────────────────────────

  // 標題列（第1–2列）：24pt 青色底白字
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

  // A 欄（案號）：青色底白字，左緣 medium
  const caseNoStyle = {
    fill:      { patternType: 'solid', fgColor: { rgb: TEAL } },
    font:      { name: '微軟正黑體', sz: 16, bold: true, color: { rgb: WHITE } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border:    { top: THIN, bottom: THIN, left: MED, right: THIN },
  };

  // B 欄（品項序號）：青色底白字粗體（欄位隱藏）
  const bColStyle = {
    fill:      { patternType: 'solid', fgColor: { rgb: TEAL } },
    font:      { name: '微軟正黑體', sz: 16, bold: true, color: { rgb: WHITE } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border:    dataBorder(1),
  };

  // 一般資料欄（依欄位決定對齊與邊框）
  function dataStyle(c) {
    return {
      font:      { name: '微軟正黑體', sz: 16 },
      alignment: colAlignment(c),
      border:    dataBorder(c),
    };
  }

  // 數字欄（置中、千分位）
  function numStyle(c) {
    return {
      font:      { name: '微軟正黑體', sz: 16 },
      alignment: { horizontal: 'center', vertical: 'center' },
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
    setCell(2, c, h, (c === 9 || c === 12) ? grayHeaderStyle : headerStyle);
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
        // 非首列：A 和 C–H 填空白（保持框線，實際由合併決定顯示）
        setCell(r, 0, '', { font: { name: '微軟正黑體', sz: 16 }, border: dataBorder(0) });
        for (let c = 2; c <= 7; c++) {
          setCell(r, c, '', { font: { name: '微軟正黑體', sz: 16 }, border: dataBorder(c) });
        }
        // B 欄：顯示品項序號（不合併）
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

    // ── 多品項案件合併邏輯 ─────────────────────────────────────────────────
    // 範本做法：A 欄獨立合併、C–H 欄各自獨立合併、B 欄不合併
    if (items.length > 1) {
      const endRow = startRow + items.length - 1;
      // A 欄（案號）
      merges.push({ s: { r: startRow, c: 0 }, e: { r: endRow, c: 0 } });
      // C–H 欄（報修日期、報修單號、站別、地址、故障說明、完工日期）
      for (let c = 2; c <= 7; c++) {
        merges.push({ s: { r: startRow, c }, e: { r: endRow, c } });
      }
    }
  });

  // ── 統計行 ────────────────────────────────────────────────────────────────
  // 結構：A–H 空白（無填色）、I–L 合併顯示「總計」（青色）、M 空白（青色）、N SUM 公式（青色）
  const statsRow = currentRow;
  const lastDataRowExcel = currentRow; // 0-indexed currentRow = 1-indexed 最後資料列

  // 計算總金額（作為公式的預算值）
  const totalSum = Math.round(
    cases.reduce((s, cas) => s + cas.items.reduce((ss, item) => ss + (Number(item.未稅小計) || 0), 0), 0)
  );

  // A–H：空白，medium 下緣
  for (let c = 0; c <= 7; c++) {
    setCell(statsRow, c, '', {
      font:   { name: '微軟正黑體', sz: 16 },
      border: { top: THIN, bottom: MED, left: c === 0 ? MED : THIN, right: THIN },
    });
  }

  // I–L：合併，「總計」，青色底白字
  const statsTealStyle = {
    fill:      { patternType: 'solid', fgColor: { rgb: TEAL } },
    font:      { name: '微軟正黑體', sz: 16, bold: true, color: { rgb: WHITE } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border:    { top: THIN, bottom: MED, left: THIN, right: THIN },
  };
  setCell(statsRow, 8, '總計', statsTealStyle);
  for (let c = 9; c <= 11; c++) setCell(statsRow, c, '', statsTealStyle);
  merges.push({ s: { r: statsRow, c: 8 }, e: { r: statsRow, c: 11 } });

  // M：空白，青色底
  setCell(statsRow, 12, '', statsTealStyle);

  // N：SUM 公式，青色底，右緣 medium
  const nAddr = XLSX.utils.encode_cell({ r: statsRow, c: 13 });
  ws[nAddr] = {
    f: `SUM(N4:N${lastDataRowExcel})`,
    v: totalSum,
    t: 'n',
    s: {
      fill:      { patternType: 'solid', fgColor: { rgb: TEAL } },
      font:      { name: '微軟正黑體', sz: 16, bold: true, color: { rgb: WHITE } },
      alignment: { horizontal: 'center', vertical: 'center' },
      numFmt:    '#,##0',
      border:    { top: THIN, bottom: MED, left: THIN, right: MED },
    },
  };

  currentRow++; // 統計行佔一列

  // ── 工作表屬性 ────────────────────────────────────────────────────────────
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: currentRow - 1, c: 13 } });
  ws['!merges'] = merges;

  // 欄寬（依範本實測值）
  ws['!cols'] = [
    { wch: 8.332  },                      // A 案號
    { wch: 8.887,  hidden: true },        // B 品項序號（隱藏）
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
