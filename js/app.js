/**
 * app.js
 * 頁面狀態管理與事件處理。
 */

// ── 全域狀態 ────────────────────────────────────────────────────────────────
let appCases      = [];   // 篩選後的案件陣列
let processedBBytes = null; // 處理後的表B二進位資料

// ── 初始化：填入年份 / 月份選單 ─────────────────────────────────────────────
(function initSelects() {
  const yearSel  = document.getElementById('yearSelect');
  const monthSel = document.getElementById('monthSelect');
  const curRoc   = new Date().getFullYear() - 1911;
  const curMon   = new Date().getMonth() + 1;

  for (let y = 110; y <= curRoc + 1; y++) {
    const opt = document.createElement('option');
    opt.value = y;
    opt.textContent = `${y} 年`;
    if (y === curRoc) opt.selected = true;
    yearSel.appendChild(opt);
  }

  for (let m = 1; m <= 12; m++) {
    const opt = document.createElement('option');
    opt.value = m;
    opt.textContent = `${m} 月`;
    if (m === curMon) opt.selected = true;
    monthSel.appendChild(opt);
  }
})();

// ── 工具函式 ────────────────────────────────────────────────────────────────
function showStep1Error(msg)  {
  const el = document.getElementById('step1Error');
  el.textContent = msg;
  el.classList.remove('hidden');
  document.getElementById('step1Result').classList.add('hidden');
}
function showStep1Ok(msg) {
  const el = document.getElementById('step1Result');
  el.textContent = msg;
  el.classList.remove('hidden');
  document.getElementById('step1Error').classList.add('hidden');
}
function unlockCard(id, badgeId) {
  const card = document.getElementById(id);
  card.classList.remove('opacity-40', 'pointer-events-none');
  if (badgeId) {
    const badge = card.querySelector('.rounded-full');
    if (badge) {
      badge.classList.remove('bg-slate-300');
      badge.classList.add('bg-emerald-600');
    }
  }
}
function fmtNum(n) {
  if (n === null || n === undefined) return '—';
  return Number(n).toLocaleString('zh-TW');
}

// ── 步驟一：載入資料 ─────────────────────────────────────────────────────────
document.getElementById('loadBtn').addEventListener('click', async () => {
  const fileInput = document.getElementById('inputFile');
  if (!fileInput.files.length) {
    showStep1Error('請先選擇計價明細 Excel 檔案');
    return;
  }

  const rocYear = parseInt(document.getElementById('yearSelect').value, 10);
  const month   = parseInt(document.getElementById('monthSelect').value, 10);

  const btn = document.getElementById('loadBtn');
  btn.textContent = '載入中…';
  btn.disabled = true;

  try {
    const buf = await fileInput.files[0].arrayBuffer();
    const { cases, itemCount } = parseInputFile(buf, rocYear, month);

    appCases = cases;
    showStep1Ok(`✓ 已載入 ${rocYear}年${month}月 ｜ 案件 ${cases.length} 件 ／ 修繕項目 ${itemCount} 筆`);
    renderPreview(cases);
    unlockCard('step2Card');
  } catch (e) {
    showStep1Error(e.message || String(e));
  } finally {
    btn.textContent = '載入資料';
    btn.disabled = false;
  }
});

// ── 預覽表格 ──────────────────────────────────────────────────────────────────
function renderPreview(cases) {
  const tbody = document.getElementById('previewBody');
  const rows  = [];

  cases.forEach((cas, idx) => {
    const caseNo = idx + 1;
    cas.items.forEach((item, itemIdx) => {
      const isFirst = itemIdx === 0;
      const rowspan = cas.items.length;
      let html = '<tr class="border-b border-slate-100 hover:bg-slate-50">';

      if (isFirst) {
        const rs = rowspan > 1 ? ` rowspan="${rowspan}"` : '';
        html += `<td${rs} class="px-2 py-1.5 text-center font-bold text-[#46BDC6] bg-[#edf9fa] align-middle">${caseNo}</td>`;
        html += `<td${rs} class="px-2 py-1.5 align-middle whitespace-nowrap">${cas.報修日期}</td>`;
        html += `<td${rs} class="px-2 py-1.5 align-middle whitespace-nowrap text-slate-500">${cas.JDM案號 || '—'}</td>`;
        html += `<td${rs} class="px-2 py-1.5 align-middle whitespace-nowrap">${cas.站別}</td>`;
        html += `<td${rs} class="px-2 py-1.5 align-middle max-w-[140px] truncate" title="${cas.地址}">${cas.地址}</td>`;
        html += `<td${rs} class="px-2 py-1.5 align-middle max-w-[200px]" style="word-break:break-all">${cas.故障說明}</td>`;
        html += `<td${rs} class="px-2 py-1.5 align-middle whitespace-nowrap">${cas.完工日期}</td>`;
      }

      html += `<td class="px-2 py-1.5 whitespace-nowrap font-mono text-slate-700">${item.合約項次}</td>`;
      html += `<td class="px-2 py-1.5 whitespace-nowrap">${item.單位}</td>`;
      html += `<td class="px-2 py-1.5 text-right">${item.數量}</td>`;
      html += `<td class="px-2 py-1.5 text-right font-mono">${fmtNum(item.未稅小計)}</td>`;
      html += '</tr>';
      rows.push(html);
    });
  });

  tbody.innerHTML = rows.join('');
}

// ── 步驟二：下載表A ───────────────────────────────────────────────────────────
document.getElementById('downloadABtn').addEventListener('click', async () => {
  if (!appCases.length) return;

  const rocYear = parseInt(document.getElementById('yearSelect').value, 10);
  const month   = parseInt(document.getElementById('monthSelect').value, 10);
  const status  = document.getElementById('downloadAStatus');

  const btn = document.getElementById('downloadABtn');
  btn.textContent = '生成中…';
  btn.disabled = true;
  status.textContent = '';

  try {
    const bytes = await generateTableA(appCases, rocYear, month);
    const blob  = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url   = URL.createObjectURL(blob);
    const a     = document.createElement('a');
    a.href      = url;
    a.download  = `${rocYear}.${String(month).padStart(2,'0')}月份請款明細單_B項_提報修繕明細.xlsx`;
    a.click();
    URL.revokeObjectURL(url);

    status.textContent = '✓ 已下載';
    // 解鎖步驟三
    unlockCard('step3Card');
  } catch (e) {
    status.textContent = `錯誤：${e.message || e}`;
  } finally {
    btn.textContent = '下載表A Excel';
    btn.disabled = false;
  }
});

// ── 步驟三：比對表B ───────────────────────────────────────────────────────────
document.getElementById('compareBtn').addEventListener('click', async () => {
  const fileInput = document.getElementById('tableBFile');
  if (!fileInput.files.length) {
    alert('請先選擇表B 總表 Excel 檔案');
    return;
  }
  if (!appCases.length) {
    alert('請先完成步驟一載入資料');
    return;
  }

  const btn = document.getElementById('compareBtn');
  btn.textContent = '比對中…';
  btn.disabled = true;

  try {
    const buf = await fileInput.files[0].arrayBuffer();
    const { results, warnings, outputBytes } = processTableB(buf, appCases);

    processedBBytes = outputBytes;
    renderDiffTable(results, warnings);
    document.getElementById('compareResult').classList.remove('hidden');
  } catch (e) {
    alert(`比對失敗：${e.message || e}`);
  } finally {
    btn.textContent = '執行比對';
    btn.disabled = false;
  }
});

function renderDiffTable(results, warnings) {
  // 警告訊息
  const warnBox = document.getElementById('warningBox');
  if (warnings.length > 0) {
    warnBox.innerHTML = '<p class="font-bold mb-1">⚠️ 表A有項次在表B中找不到：</p>' +
      warnings.map(w => `<p class="pl-2 text-xs">${w}</p>`).join('');
    warnBox.classList.remove('hidden');
  } else {
    warnBox.classList.add('hidden');
  }

  // 比對結果列表
  const tbody = document.getElementById('diffBody');
  const rows  = results.map(res => {
    const bg = res.isDiscrepant ? 'bg-rose-50' : '';
    const statusBadge = res.isDiscrepant
      ? '<span class="px-2 py-0.5 rounded-full text-xs font-bold bg-rose-100 text-rose-700">金額不符 ⚠</span>'
      : res.hasMatch
        ? '<span class="px-2 py-0.5 rounded-full text-xs font-bold bg-emerald-100 text-emerald-700">符合 ✓</span>'
        : '<span class="px-2 py-0.5 rounded-full text-xs font-bold bg-slate-100 text-slate-500">無對應</span>';

    const tableADisplay = res.hasMatch ? fmtNum(res.tableASum) : '—';
    const notesDisplay  = res.caseNos.length > 0 ? res.caseNos.map(n=>`案號${n}`).join('、') : '—';

    return `<tr class="border-b border-slate-100 ${bg}">
      <td class="px-3 py-2 font-mono text-slate-700 whitespace-nowrap">${res.code}</td>
      <td class="px-3 py-2 text-slate-600 max-w-[200px] truncate" title="${res.itemName}">${res.itemName}</td>
      <td class="px-3 py-2 text-right font-mono">${tableADisplay}</td>
      <td class="px-3 py-2 text-right font-mono">${fmtNum(res.tableBAmount)}</td>
      <td class="px-3 py-2 text-center">${statusBadge}</td>
      <td class="px-3 py-2 text-slate-600 text-xs">${notesDisplay}</td>
    </tr>`;
  });
  tbody.innerHTML = rows.join('') || '<tr><td colspan="6" class="text-center py-6 text-slate-400">未找到 B 項修繕項次</td></tr>';
}

// ── 下載處理後的表B ─────────────────────────────────────────────────────────
document.getElementById('downloadBBtn').addEventListener('click', () => {
  if (!processedBBytes) return;

  const rocYear = parseInt(document.getElementById('yearSelect').value, 10);
  const month   = parseInt(document.getElementById('monthSelect').value, 10);
  const status  = document.getElementById('downloadBStatus');

  const blob = new Blob([processedBBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `${rocYear}.${String(month).padStart(2,'0')}月份請款明細單_B項_總表_已處理.xlsx`;
  a.click();
  URL.revokeObjectURL(url);
  status.textContent = '✓ 已下載';
});
