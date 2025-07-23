const form = document.getElementById('kpiForm');
const yearInput = document.getElementById('yearInput');
const recordTable = document.getElementById('recordTable');
const queryTable = document.getElementById('queryTable');
const data = [];
const exportBtn = document.getElementById('exportBtn');

// 年份加減
// function changeYear(delta) {
//   yearInput.value = parseInt(yearInput.value) + delta;
// }

// 表單送出
form.addEventListener('submit', (e) => {
  e.preventDefault();
  const formData = new FormData(form);
  const entry = {};
  for (let [key, value] of formData.entries()) {
    entry[key] = value;
  }
  data.push(entry);
  renderTables();
  form.reset();
  yearInput.value = "2025";
});

// 渲染只讀表格
function renderRecordTable() {
  recordTable.innerHTML = "";
  data.forEach((item, index) => {
    const row = document.createElement('tr');
    row.innerHTML = `
      <td>${index + 1}</td>
      <td>${item.vertical}</td>
      <td>${item.horizontal}</td>
      <td>${item.kpi || ''}</td>
      <td>${item.formula || ''}</td>
      <td>${item.year}</td>
      <td>${item.month}</td>
      <td>${item.quarter}</td>
      <td>${item.reviewDept}</td>
      <td>${item.value || ''}</td>
      <td>${item.unit}</td>
    `;

    recordTable.appendChild(row);
  });
}

// 渲染查詢（可編輯）表格
function renderQueryTable() {
  queryTable.innerHTML = "";
  data.forEach((item, index) => {
    const row = document.createElement('tr');
    row.setAttribute("data-row-index", index);
    row.innerHTML = `
      <td>${index + 1}</td>
      ${renderEditableCell('vertical', item.vertical, index)}
      ${renderEditableCell('horizontal', item.horizontal, index)}
      ${renderEditableCell('kpi', item.kpi, index)}
      ${renderEditableCell('formula', item.formula, index)}
      ${renderEditableCell('year', item.year, index)}
      ${renderEditableCell('month', item.month, index)}
      ${renderEditableCell('quarter', item.quarter, index)}
      ${renderEditableCell('reviewDept', item.reviewDept, index)}
      ${renderEditableCell('value', item.value, index)}
      ${renderEditableCell('unit', item.unit, index)}
      <td>
        <button class="edit-btn" onclick="toggleEdit(${index}, this)">修改</button>
        <button class="delete-btn" onclick="deleteRow(${index})">刪除</button>
      </td>
    `;
    queryTable.appendChild(row);
  });
}


function renderEditableCell(field, value, index) {
  return `<td contenteditable="false" data-index="${index}" data-field="${field}">${value || ''}</td>`;
}

// 編輯／儲存切換
function toggleEdit(index, btn) {
  const row = document.querySelector(`tr[data-row-index="${index}"]`);
  const isEditing = btn.textContent === "儲存";
  const cells = row.querySelectorAll(`td[contenteditable]`);

  if (!isEditing) {
    cells.forEach(cell => {
      cell.contentEditable = true;
      cell.style.backgroundColor = "#fff7cc";
    });
    btn.textContent = "儲存";
  } else {
    cells.forEach(cell => {
      cell.contentEditable = false;
      cell.style.backgroundColor = "";
      const field = cell.dataset.field;
      const newValue = cell.innerText.trim();
      data[index][field] = newValue;
    });
    btn.textContent = "修改";
    renderRecordTable();
  }
}

// 刪除列
function deleteRow(index) {
  if (confirm("確定要刪除這筆資料嗎？")) {
    data.splice(index, 1);
    renderTables();
  }
}

// 統一刷新表格
function renderTables() {
  renderRecordTable();
  renderQueryTable();
}

// 添加點擊事件處理程序
exportBtn.addEventListener('click', exportToExcel);

// Excel 匯出函數
function exportToExcel() {
  // 檢查是否有數據可導出
  if (data.length === 0) {
    alert('沒有數據可導出！');
    return;
  }

  // 準備工作表數據
  const wsData = [
    // 表頭
    ['序號', '縱向', '橫向', 'KPI', '統計公式', '年份', '月份', '季', '檢討部門', '數值', '單位'],
    // 數據行
    ...data.map((item, index) => [
      index + 1,
      item.vertical,
      item.horizontal,
      item.kpi || '',
      item.formula || '',
      item.year,
      item.month,
      item.quarter,
      item.reviewDept,
      item.value || '',
      item.unit
    ])
  ];

  // 創建工作表
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // 設置列寬（可選）
  ws['!cols'] = [
    { wch: 6 },  // 序號
    { wch: 8 },  // 縱向
    { wch: 8 },  // 橫向
    { wch: 20 }, // KPI
    { wch: 15 }, // 統計公式
    { wch: 6 },  // 年份
    { wch: 6 },  // 月份
    { wch: 6 },  // 季
    { wch: 15 }, // 檢討部門
    { wch: 10 }, // 數值
    { wch: 6 }   // 單位
  ];

  // 創建工作簿
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "部門績效數據");

  // 生成 Excel 文件並下載
  const fileName = `部門績效_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(wb, fileName);
}
