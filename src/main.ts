import './style.css';
import {
  loadFromLocalStorage, 
  saveToLocalStorage, 
  exportToJSONFile, 
  importFromJSONFile, 
  sanitizeDB,
  DEFAULT_HOLIDAYS_2026,
} from './dbSync';
import type {
  TimeStampDB,
  Holiday,
  RuleSettings,
  LeaveRecord
} from './dbSync';
import { 
  parseExcelFile, 
} from './excelParser';
import type {
  EmployeeData 
} from './excelParser';
import {
  loadGistConfig,
  saveGistConfig,
  clearGistConfig,
  fetchFromGist,
  updateGist,
} from './freeCloudSync';
import type {
  GistSyncConfig
} from './freeCloudSync';

// State Management
let dbState: TimeStampDB = loadFromLocalStorage();
let parsedEmployees: EmployeeData[] = [];
let gistConfig: GistSyncConfig | null = loadGistConfig();

// DOM Elements
const syncDot = document.querySelector('.status-dot') as HTMLElement;
const syncText = document.querySelector('.status-text') as HTMLElement;
const btnImportDB = document.getElementById('btn-import-db') as HTMLButtonElement;
const btnExportDB = document.getElementById('btn-export-db') as HTMLButtonElement;
const dbFileInput = document.getElementById('db-file-input') as HTMLInputElement;

const btnToggleGist = document.getElementById('btn-toggle-gist') as HTMLButtonElement;
const gistConfigPanel = document.getElementById('gist-config-panel') as HTMLElement;
const gistTokenInput = document.getElementById('gist-github-token') as HTMLInputElement;
const gistIdInput = document.getElementById('gist-id') as HTMLInputElement;
const btnSaveGistConfig = document.getElementById('btn-save-gist-config') as HTMLButtonElement;
const btnPullGist = document.getElementById('btn-pull-gist') as HTMLButtonElement;
const btnPushGist = document.getElementById('btn-push-gist') as HTMLButtonElement;
const btnClearGistConfig = document.getElementById('btn-clear-gist-config') as HTMLButtonElement;

const dropzone = document.getElementById('dropzone') as HTMLElement;
const excelFileInput = document.getElementById('excel-file-input') as HTMLInputElement;
const fileInfo = document.getElementById('file-info') as HTMLElement;
const fileInfoName = fileInfo.querySelector('.file-name') as HTMLElement;
const fileListContainer = document.getElementById('file-list-container') as HTMLElement;
const btnClearFile = document.getElementById('btn-clear-file') as HTMLButtonElement;
const btnAddMoreFiles = document.getElementById('btn-add-more-files') as HTMLButtonElement;

const statStaffCount = document.getElementById('stat-staff-count') as HTMLElement;
const statDaysCount = document.getElementById('stat-days-count') as HTMLElement;
const statsOverview = document.getElementById('stats-overview') as HTMLElement;

const rulesForm = document.getElementById('rules-form') as HTMLFormElement;
const morningWorkStartInput = document.getElementById('morningWorkStart') as HTMLInputElement;
const afternoonWorkEndInput = document.getElementById('afternoonWorkEnd') as HTMLInputElement;
const lateAllowanceInput = document.getElementById('lateAllowanceMinutes') as HTMLInputElement;
const earlyAllowanceInput = document.getElementById('earlyCheckoutAllowanceMinutes') as HTMLInputElement;
const halfDayLateInput = document.getElementById('halfDayLateMinutes') as HTMLInputElement;
const morningLeaveStartInput = document.getElementById('morningLeaveStart') as HTMLInputElement;
const morningLeaveEndInput = document.getElementById('morningLeaveEnd') as HTMLInputElement;
const afternoonLeaveStartInput = document.getElementById('afternoonLeaveStart') as HTMLInputElement;
const afternoonLeaveEndInput = document.getElementById('afternoonLeaveEnd') as HTMLInputElement;
const otWeekdayStartInput = document.getElementById('otWeekdayStart') as HTMLInputElement;
const otHolidayHoursInput = document.getElementById('otHolidayHours') as HTMLInputElement;

const holidayDateInput = document.getElementById('holiday-date') as HTMLInputElement;
const holidayNameInput = document.getElementById('holiday-name') as HTMLInputElement;
const btnAddHoliday = document.getElementById('btn-add-holiday') as HTMLButtonElement;
const btnLoadDefaultHolidays = document.getElementById('btn-load-default-holidays') as HTMLButtonElement;
const holidayListBody = document.getElementById('holiday-list-body') as HTMLElement;

const filterStartDate = document.getElementById('filter-start-date') as HTMLInputElement;
const filterEndDate = document.getElementById('filter-end-date') as HTMLInputElement;
const searchStaff = document.getElementById('search-staff') as HTMLInputElement;
const btnExportExcel = document.getElementById('btn-export-excel') as HTMLButtonElement;
const btnPrintReports = document.getElementById('btn-print-reports') as HTMLButtonElement;
const filterModeSelect = document.getElementById('filter-mode') as HTMLSelectElement;
const filterMonthSelect = document.getElementById('filter-month') as HTMLSelectElement;
const filterYearSelect = document.getElementById('filter-year') as HTMLSelectElement;
const filterDepartmentSelect = document.getElementById('filter-department') as HTMLSelectElement;
const printContainer = document.getElementById('print-container') as HTMLElement;
const printShowStatusCheckbox = document.getElementById('print-show-status') as HTMLInputElement;
const printShowVerificationCheckbox = document.getElementById('print-show-verification') as HTMLInputElement;
const selectAllPrint = document.getElementById('select-all-print') as HTMLInputElement;
const leaveSheetUrlInput = document.getElementById('leave-sheet-url') as HTMLInputElement;
const btnSyncLeave = document.getElementById('btn-sync-leave') as HTMLButtonElement;
const btnClearLeave = document.getElementById('btn-clear-leave') as HTMLButtonElement;

const summaryTableBody = document.getElementById('summary-table-body') as HTMLElement;

// Modal Elements
const detailModal = document.getElementById('detail-modal') as HTMLElement;
const modalTitle = document.getElementById('modal-title') as HTMLElement;
const detailTableBody = document.getElementById('detail-table-body') as HTMLElement;
const btnCloseModal = document.getElementById('btn-close-modal') as HTMLButtonElement;

interface SortState {
  column: string | null;
  direction: 'asc' | 'desc';
}

let currentSort: SortState = {
  column: null,
  direction: 'asc'
};


function mergeEmployeeRecords(existing: EmployeeData[], newEmployees: EmployeeData[]): EmployeeData[] {
  const mergedMap = new Map<string, EmployeeData>();
  
  existing.forEach(emp => {
    mergedMap.set(emp.id, {
      ...emp,
      records: { ...emp.records }
    });
  });
  
  newEmployees.forEach(newEmp => {
    if (mergedMap.has(newEmp.id)) {
      const existingEmp = mergedMap.get(newEmp.id)!;
      for (const [date, times] of Object.entries(newEmp.records)) {
        existingEmp.records[date] = Array.from(new Set([...(existingEmp.records[date] || []), ...times])).sort();
      }
      if (newEmp.name) existingEmp.name = newEmp.name;
      if (newEmp.position) existingEmp.position = newEmp.position;
      if (newEmp.group) existingEmp.group = newEmp.group;
      if (newEmp.department) existingEmp.department = newEmp.department;
    } else {
      mergedMap.set(newEmp.id, {
        ...newEmp,
        records: { ...newEmp.records }
      });
    }
  });
  
  return Array.from(mergedMap.values());
}

function applyEmployeesFromState() {
  if (dbState.employees && dbState.employees.length > 0) {
    parsedEmployees = dbState.employees;
    
    // Display file stats
    fileInfoName.textContent = `ข้อมูลพนักงานสะสม (${parsedEmployees.length} คน)`;
    dropzone.style.display = 'none';
    fileInfo.style.display = 'flex';

    // Render imported file list
    if (fileListContainer) {
      fileListContainer.innerHTML = '';
      if (dbState.importedFiles && dbState.importedFiles.length > 0) {
        dbState.importedFiles.forEach((filename) => {
          const item = document.createElement('div');
          item.style.display = 'flex';
          item.style.justifyContent = 'space-between';
          item.style.alignItems = 'center';
          item.style.background = 'rgba(255, 255, 255, 0.05)';
          item.style.padding = '4px 8px';
          item.style.borderRadius = '4px';
          item.innerHTML = `
            <span>📄 ${filename}</span>
            <button class="btn-delete-file" data-filename="${filename}" style="background: none; border: none; color: #f87171; cursor: pointer; font-size: 0.85rem; font-weight: bold; padding: 2px 6px;">✕</button>
          `;
          fileListContainer.appendChild(item);
        });

        // Attach delete handlers for individual files
        document.querySelectorAll('.btn-delete-file').forEach((btn) => {
          btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const filename = (e.currentTarget as HTMLElement).getAttribute('data-filename');
            if (filename) {
              deleteImportedFile(filename);
            }
          });
        });
      } else {
        fileListContainer.innerHTML = '<div style="font-style: italic; opacity: 0.6;">(ไม่มีรายชื่อไฟล์นำเข้า)</div>';
      }
    }

    // Extract Date limits
    let allDates: string[] = [];
    parsedEmployees.forEach(e => {
      allDates.push(...Object.keys(e.records));
    });
    
    if (allDates.length > 0) {
      allDates.sort();
      const minDate = allDates[0];
      const maxDate = allDates[allDates.length - 1];
      if (!filterStartDate.value) filterStartDate.value = minDate;
      if (!filterEndDate.value) filterEndDate.value = maxDate;

      // Extract Year and Month from minDate to preselect dropdowns
      const parts = minDate.split('-');
      if (parts.length === 3) {
        const year = parts[0];
        const month = parts[1];

        // Ensure year exists in year dropdown
        let hasYear = false;
        for (let i = 0; i < filterYearSelect.options.length; i++) {
          if (filterYearSelect.options[i].value === year) {
            hasYear = true;
            break;
          }
        }
        if (!hasYear) {
          const opt = document.createElement('option');
          opt.value = year;
          opt.textContent = year;
          filterYearSelect.appendChild(opt);
        }
        filterYearSelect.value = year;
        filterMonthSelect.value = month;
      }

      statDaysCount.textContent = String(new Set(allDates).size);
    }
    
    statStaffCount.textContent = String(parsedEmployees.length);
    statsOverview.style.display = 'grid';
    recalculateAndRender();
  } else {
    parsedEmployees = [];
    excelFileInput.value = '';
    dropzone.style.display = 'block';
    fileInfo.style.display = 'none';
    statsOverview.style.display = 'none';
    if (fileListContainer) fileListContainer.innerHTML = '';
    
    summaryTableBody.innerHTML = `
      <tr>
        <td colspan="8" class="text-center text-muted">กรุณานำเข้าไฟล์สแกนบัตร (Excel) เพื่อคำนวณและแสดงผลตารางสรุป</td>
      </tr>
    `;
  }
}

function deleteImportedFile(filename: string) {
  if (!confirm(`คุณต้องการลบไฟล์ "${filename}" และข้อมูลสแกนบัตรทั้งหมดในไฟล์นี้ใช่หรือไม่?`)) {
    return;
  }
  
  if (!dbState.importedFiles) return;
  
  const remainingFiles = dbState.importedFiles.filter(f => f !== filename);
  
  const remainingDates = new Set<string>();
  remainingFiles.forEach(f => {
    const dates = dbState.fileDates?.[f] || [];
    dates.forEach(d => remainingDates.add(d));
  });
  
  dbState.employees.forEach(emp => {
    const newRecords: { [date: string]: string[] } = {};
    for (const [date, times] of Object.entries(emp.records)) {
      if (remainingDates.has(date)) {
        newRecords[date] = times;
      }
    }
    emp.records = newRecords;
  });
  
  dbState.employees = dbState.employees.filter(emp => Object.keys(emp.records).length > 0);
  
  dbState.importedFiles = remainingFiles;
  if (dbState.fileDates) {
    delete dbState.fileDates[filename];
  }
  
  saveAndSync();
  applyEmployeesFromState();
}

function saveAndSync() {
  saveToLocalStorage(dbState);
  if (gistConfig) {
    pushToCloud();
  }
}

// App Setup
function init() {
  updateSyncStatus(false, 'ใช้งานแบบ Local (บันทึกลงเบราว์เซอร์ชั่วคราว)');
  renderRules();
  renderHolidays();
  
  // File Upload Handlers
  dropzone.addEventListener('click', () => excelFileInput.click());
  excelFileInput.addEventListener('change', handleExcelUpload);
  
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });
  
  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
  });
  
  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer?.files.length) {
      excelFileInput.files = e.dataTransfer.files;
      handleExcelUpload();
    }
  });

  btnClearFile.addEventListener('click', clearLoadedData);
  if (btnAddMoreFiles) {
    btnAddMoreFiles.addEventListener('click', () => {
      excelFileInput.value = '';
      excelFileInput.click();
    });
  }

  // DB Sync Handlers
  btnImportDB.addEventListener('click', () => dbFileInput.click());
  dbFileInput.addEventListener('change', handleDBImport);
  btnExportDB.addEventListener('click', handleDBExport);

  // GitHub Gist Handlers
  btnToggleGist.addEventListener('click', () => {
    const isHidden = gistConfigPanel.style.display === 'none';
    gistConfigPanel.style.display = isHidden ? 'block' : 'none';
  });

  if (gistConfig) {
    gistTokenInput.value = gistConfig.githubToken;
    gistIdInput.value = gistConfig.gistId;
    updateSyncStatus(true, 'เชื่อมต่อ GitHub Gist แล้ว');
    // Auto-pull on startup
    pullFromCloud();
  }

  btnSaveGistConfig.addEventListener('click', () => {
    const token = gistTokenInput.value.trim();
    const id = gistIdInput.value.trim();
    if (!token || !id) {
      alert('กรุณากรอกทั้ง GitHub Token และ Gist ID');
      return;
    }
    gistConfig = { githubToken: token, gistId: id };
    saveGistConfig(gistConfig);
    updateSyncStatus(true, 'บันทึกสิทธิ์ Gist แล้ว');
    alert('บันทึกสิทธิ์ Gist สำเร็จ! กำลังดึงข้อมูลจาก Cloud...');
    pullFromCloud();
  });

  btnPullGist.addEventListener('click', pullFromCloud);
  btnPushGist.addEventListener('click', pushToCloud);

  btnClearGistConfig.addEventListener('click', () => {
    clearGistConfig();
    gistConfig = null;
    gistTokenInput.value = '';
    gistIdInput.value = '';
    updateSyncStatus(false, 'Local Storage เท่านั้น');
    alert('ล้างสิทธิ์เชื่อมต่อ Gist เรียบร้อยแล้ว');
  });

  // Rules Handlers
  rulesForm.addEventListener('submit', (e) => {
    e.preventDefault();
    dbState.rules = {
      morningWorkStart: morningWorkStartInput.value,
      afternoonWorkEnd: afternoonWorkEndInput.value,
      lateAllowanceMinutes: parseInt(lateAllowanceInput.value, 10) || 0,
      earlyCheckoutAllowanceMinutes: parseInt(earlyAllowanceInput.value, 10) || 0,
      halfDayLateMinutes: parseInt(halfDayLateInput.value, 10) || 240,
      morningLeaveStart: morningLeaveStartInput.value,
      morningLeaveEnd: morningLeaveEndInput.value,
      afternoonLeaveStart: afternoonLeaveStartInput.value,
      afternoonLeaveEnd: afternoonLeaveEndInput.value,
      otWeekdayStart: otWeekdayStartInput.value,
      otHolidayHours: parseInt(otHolidayHoursInput.value, 10) || 8,
    };
    saveAndSync();
    alert('บันทึกเงื่อนไขและเกณฑ์เวลาแล้ว!');
    recalculateAndRender();
  });

  // Holiday Handlers
  btnAddHoliday.addEventListener('click', handleAddHoliday);
  btnLoadDefaultHolidays.addEventListener('click', handleLoadDefaultHolidays);
  
  // Filters Event Listeners
  const dateGroups = document.querySelectorAll('.filter-date-group') as NodeListOf<HTMLElement>;
  const monthlyGroups = document.querySelectorAll('.filter-monthly-group') as NodeListOf<HTMLElement>;

  const updateFilterVisibility = () => {
    const isMonthly = filterModeSelect.value === 'monthly';
    dateGroups.forEach(g => g.style.display = isMonthly ? 'none' : 'block');
    monthlyGroups.forEach(g => g.style.display = isMonthly ? 'flex' : 'none');
  };

  updateFilterVisibility();

  filterModeSelect.addEventListener('change', () => {
    updateFilterVisibility();
    recalculateAndRender();
  });
  filterMonthSelect.addEventListener('change', recalculateAndRender);
  filterYearSelect.addEventListener('change', recalculateAndRender);
  filterDepartmentSelect.addEventListener('change', recalculateAndRender);
  filterStartDate.addEventListener('change', recalculateAndRender);
  filterEndDate.addEventListener('change', recalculateAndRender);
  searchStaff.addEventListener('input', recalculateAndRender);

  selectAllPrint.addEventListener('change', () => {
    const isChecked = selectAllPrint.checked;
    document.querySelectorAll('.select-emp-print').forEach((cb) => {
      (cb as HTMLInputElement).checked = isChecked;
    });
  });

  // Attach Sort Handlers to Headers
  document.querySelectorAll('.sortable').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.getAttribute('data-sort');
      if (currentSort.column === col) {
        currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
      } else {
        currentSort.column = col;
        currentSort.direction = 'asc';
      }
      recalculateAndRender();
    });
  });

  // Detail Modal Close
  btnCloseModal.addEventListener('click', () => {
    detailModal.style.display = 'none';
  });
  window.addEventListener('click', (e) => {
    if (e.target === detailModal) {
      detailModal.style.display = 'none';
    }
  });

  // Export CSV Summary
  btnExportExcel.addEventListener('click', handleCSVExport);

  // Print Reports
  if (btnPrintReports) {
    btnPrintReports.addEventListener('click', handlePrintReports);
  }

  // Restore employee data from local storage if exists
  applyEmployeesFromState();

  if (dbState.leaveSheetUrl) {
    leaveSheetUrlInput.value = dbState.leaveSheetUrl;
  }
  renderLeaveSummaryTable();
  renderLeaveAlerts();

  btnSyncLeave.addEventListener('click', () => {
    const url = leaveSheetUrlInput.value.trim();
    if (!url) {
      alert('กรุณากรอก Google Sheets CSV URL');
      return;
    }
    btnSyncLeave.disabled = true;
    btnSyncLeave.textContent = '⏳ กำลังซิงค์...';

    fetch(url)
      .then(res => {
        if (!res.ok) throw new Error('Network response was not ok');
        return res.text();
      })
      .then(text => {
        const rows = parseCSV(text);
        const records = mapCSVToLeaveRecords(rows);
        dbState.leaveSheetUrl = url;
        dbState.leaveRecords = records;
        saveAndSync();
        renderLeaveSummaryTable();
        renderLeaveAlerts();
        recalculateAndRender();
        alert(`ซิงค์ข้อมูลวันลาสำเร็จ! โหลดได้ทั้งหมด ${records.length} รายการ`);
      })
      .catch(err => {
        console.error(err);
        alert('เกิดข้อผิดพลาดในการดึงข้อมูลจาก URL กรุณาตรวจสอบว่าคุณได้ "เผยแพร่ไปยังเว็บ" เป็นไฟล์ CSV เรียบร้อยแล้ว');
      })
      .finally(() => {
        btnSyncLeave.disabled = false;
        btnSyncLeave.textContent = '🔄 ซิงค์ข้อมูลวันลา';
      });
  });

  btnClearLeave.addEventListener('click', () => {
    if (confirm('คุณต้องการล้างข้อมูลวันลาทั้งหมดหรือไม่?')) {
      dbState.leaveSheetUrl = '';
      dbState.leaveRecords = [];
      leaveSheetUrlInput.value = '';
      saveAndSync();
      renderLeaveSummaryTable();
      renderLeaveAlerts();
      recalculateAndRender();
      alert('ล้างข้อมูลวันลาเรียบร้อยแล้ว');
    }
  });
}

// Update Database Sync Status Label
function updateSyncStatus(isOneDrive: boolean, text: string) {
  if (isOneDrive) {
    syncDot.className = 'status-dot online';
    syncText.textContent = text;
  } else {
    syncDot.className = 'status-dot offline';
    syncText.textContent = text;
  }
}

// Rules UI update
function renderRules() {
  const r = dbState.rules;
  morningWorkStartInput.value = r.morningWorkStart;
  afternoonWorkEndInput.value = r.afternoonWorkEnd;
  lateAllowanceInput.value = String(r.lateAllowanceMinutes);
  earlyAllowanceInput.value = String(r.earlyCheckoutAllowanceMinutes);
  halfDayLateInput.value = String(r.halfDayLateMinutes);
  morningLeaveStartInput.value = r.morningLeaveStart || "10:30";
  morningLeaveEndInput.value = r.morningLeaveEnd || "13:00";
  afternoonLeaveStartInput.value = r.afternoonLeaveStart || "12:00";
  afternoonLeaveEndInput.value = r.afternoonLeaveEnd || "13:00";
  otWeekdayStartInput.value = r.otWeekdayStart || "19:30";
  otHolidayHoursInput.value = String(r.otHolidayHours || 8);
}

// Holidays UI update
function renderHolidays() {
  // Sort holidays by date
  dbState.holidays.sort((a, b) => a.date.localeCompare(b.date));
  
  holidayListBody.innerHTML = '';
  if (dbState.holidays.length === 0) {
    holidayListBody.innerHTML = `<tr><td colspan="3" class="text-center text-muted">ไม่มีวันหยุดพิเศษ</td></tr>`;
    return;
  }

  dbState.holidays.forEach((h, index) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${h.date}</td>
      <td>${h.name}</td>
      <td><button class="btn btn-danger btn-sm btn-delete-holiday" data-index="${index}">✕</button></td>
    `;
    holidayListBody.appendChild(tr);
  });

  // Attach delete buttons
  document.querySelectorAll('.btn-delete-holiday').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      const idx = parseInt((e.currentTarget as HTMLElement).getAttribute('data-index') || '0', 10);
      dbState.holidays.splice(idx, 1);
      saveAndSync();
      renderHolidays();
      recalculateAndRender();
    });
  });
}

// Add Holiday
function handleAddHoliday() {
  const dateVal = holidayDateInput.value;
  const nameVal = holidayNameInput.value.trim();

  if (!dateVal || !nameVal) {
    alert('กรุณาระบุวันที่และชื่อวันหยุด');
    return;
  }

  // Check duplicate
  if (dbState.holidays.some(h => h.date === dateVal)) {
    alert('มีวันหยุดในวันที่เลือกแล้ว');
    return;
  }

  dbState.holidays.push({ date: dateVal, name: nameVal });
  saveAndSync();
  renderHolidays();
  
  // Clear input fields
  holidayDateInput.value = '';
  holidayNameInput.value = '';
  
  recalculateAndRender();
}

// Load default 2026 holidays
function handleLoadDefaultHolidays() {
  let count = 0;
  DEFAULT_HOLIDAYS_2026.forEach((dh) => {
    if (!dbState.holidays.some((h) => h.date === dh.date)) {
      dbState.holidays.push({ ...dh });
      count++;
    }
  });

  if (count === 0) {
    alert('มีวันหยุดราชการปี 2026 ทั้งหมดในระบบแล้ว');
    return;
  }

  saveAndSync();
  renderHolidays();
  recalculateAndRender();
  alert(`โหลดวันหยุดราชการประจำปี 2026 สำเร็จ (เพิ่มใหม่ ${count} วัน)`);
}

// JSON Database Export/Import (OneDrive integration support)
function handleDBExport() {
  exportToJSONFile(dbState);
  updateSyncStatus(true, 'ส่งออกไฟล์สำเร็จ! กรุณานำไปใส่ใน OneDrive โฟลเดอร์');
}

async function pullFromCloud() {
  if (!gistConfig) return;
  try {
    const data = await fetchFromGist(gistConfig);
    // Sanitize in case it's an empty {}
    const sanitized = sanitizeDB(data);
    if (sanitized) {
      dbState = sanitized;
      saveToLocalStorage(dbState);
      renderRules();
      renderHolidays();
      applyEmployeesFromState();
      leaveSheetUrlInput.value = dbState.leaveSheetUrl || '';
      renderLeaveSummaryTable();
      renderLeaveAlerts();
      updateSyncStatus(true, 'ซิงค์ข้อมูลจาก Gist Cloud สำเร็จ!');
    }
  } catch (err: any) {
    console.error(err);
    alert('ไม่สามารถดึงข้อมูลจาก Cloud Gist ได้: ' + err.message);
  }
}

async function pushToCloud() {
  if (!gistConfig) {
    alert('กรุณาตั้งค่าเชื่อมต่อ Gist ก่อน');
    return;
  }
  try {
    updateSyncStatus(true, 'กำลังส่งข้อมูลขึ้น Cloud...');
    await updateGist(gistConfig, dbState);
    updateSyncStatus(true, 'ส่งข้อมูลขึ้น Gist Cloud สำเร็จ!');
    alert('อัปเดตข้อมูลขึ้น Cloud Gist สำเร็จแล้ว!');
  } catch (err: any) {
    console.error(err);
    alert('ไม่สามารถส่งข้อมูลขึ้น Cloud Gist ได้: ' + err.message);
  }
}

async function handleDBImport(e: Event) {
  const file = (e.target as HTMLInputElement).files?.[0];
  if (!file) return;

  try {
    const imported = await importFromJSONFile(file);
    dbState = imported;
    saveAndSync();
    renderRules();
    renderHolidays();
    updateSyncStatus(true, `เชื่อมต่อสำเร็จ: ${file.name}`);
    alert('นำเข้าฐานข้อมูลจาก OneDrive สำเร็จแล้ว!');
    applyEmployeesFromState();
  } catch (err: any) {
    alert(err.message || 'การนำเข้าล้มเหลว');
  }
}

// Excel Upload Handler
import { debugExcelStructure } from './excelDebug';

async function handleExcelUpload() {
  const file = excelFileInput.files?.[0];
  if (!file) return;

  try {
    const debugRows = await debugExcelStructure(file);
    console.log("Raw Excel Rows (first 50):", debugRows);
    
    const freshlyParsed = await parseExcelFile(file);
    dbState.employees = mergeEmployeeRecords(dbState.employees || [], freshlyParsed);
    
    if (!dbState.importedFiles) {
      dbState.importedFiles = [];
    }
    if (!dbState.importedFiles.includes(file.name)) {
      dbState.importedFiles.push(file.name);
    }

    // Save date sources
    const fileDatesSet = new Set<string>();
    freshlyParsed.forEach(emp => {
      Object.keys(emp.records).forEach(date => {
        fileDatesSet.add(date);
      });
    });
    if (!dbState.fileDates) {
      dbState.fileDates = {};
    }
    dbState.fileDates[file.name] = Array.from(fileDatesSet);
    
    saveAndSync();
    applyEmployeesFromState();
  } catch (err) {
    console.error(err);
    alert('เกิดข้อผิดพลาดในการเปิดและสแกนไฟล์ Excel กรุณาตรวจสอบรูปแบบไฟล์ต้นฉบับ');
  }
}

function clearLoadedData() {
  parsedEmployees = [];
  dbState.employees = [];
  dbState.importedFiles = [];
  dbState.fileDates = {};
  saveAndSync();
  
  excelFileInput.value = '';
  dropzone.style.display = 'block';
  fileInfo.style.display = 'none';
  statsOverview.style.display = 'none';
  if (fileListContainer) fileListContainer.innerHTML = '';
  
  summaryTableBody.innerHTML = `
    <tr>
      <td colspan="10" class="text-center text-muted">กรุณานำเข้าไฟล์สแกนบัตร (Excel) เพื่อคำนวณและแสดงผลตารางสรุป</td>
    </tr>
  `;
}

// Business Logic: late, leave, early status calculation
interface ProcessedStaffSummary {
  id: string;
  name: string;
  department: string;
  position: string;
  workedDays: number;
  lateCount: number;
  holidayLateCount: number;
  leaveCount: number;
  earlyOutCount: number;
  ot3Count: number;
  ot8Count: number;
  records: ProcessedDayRecord[];
}

interface ProcessedDayRecord {
  date: string;
  checkIn: string;
  checkOut: string;
  status: string;
  lateMinutes: number;
  earlyMinutes: number;
}

// Convert time to minutes (e.g. "08:30" -> 510)
function timeToMinutes(timeStr: string): number {
  if (!timeStr) return 0;
  const [h, m] = timeStr.split(':').map(Number);
  return h * 60 + m;
}

function parseDateToISO(dateStr: string): string {
  if (!dateStr) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
    return dateStr;
  }
  const parts = dateStr.split('/');
  if (parts.length === 3) {
    const day = parts[0].padStart(2, '0');
    const month = parts[1].padStart(2, '0');
    let year = parseInt(parts[2], 10);
    if (year > 2500) {
      year -= 543;
    }
    return `${year}-${month}-${day}`;
  }
  return dateStr;
}

function expandLeaveRecordDays(r: LeaveRecord): { date: string, type: string, format: string, hourMinutes: number, missedScanType: string }[] {
  const days: { date: string, type: string, format: string, hourMinutes: number, missedScanType: string }[] = [];
  if (!r.startDate) return [];
  const start = new Date(r.startDate);
  const end = r.endDate ? new Date(r.endDate) : start;

  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    days.push({
      date: `${year}-${month}-${day}`,
      type: r.leaveType,
      format: r.leaveFormat,
      hourMinutes: r.hourMinutes || 0,
      missedScanType: r.missedScanType || ''
    });
  }
  return days;
}

function calculateStaffRecords(staff: EmployeeData, rules: RuleSettings, holidays: Holiday[], start: string, end: string): ProcessedStaffSummary {
  const workedDates = Object.keys(staff.records).filter(d => (!start || d >= start) && (!end || d <= end)).sort();
  
  // Build a map of dates we expect the user to work if they have data
  // In the real system, it would be every day within range. Let's list all days between start and end
  const daysInRange: string[] = [];
  if (start && end) {
    const curr = new Date(start);
    const stop = new Date(end);
    while (curr <= stop) {
      const yyyy = curr.getFullYear();
      const mm = String(curr.getMonth() + 1).padStart(2, '0');
      const dd = String(curr.getDate()).padStart(2, '0');
      daysInRange.push(`${yyyy}-${mm}-${dd}`);
      curr.setDate(curr.getDate() + 1);
    }
  } else {
    daysInRange.push(...workedDates);
  }

  // Load leave records for this employee
  const staffLeaves = (dbState.leaveRecords || [])
    .filter(r => r.staffId === staff.id)
    .flatMap(expandLeaveRecordDays);

  const processedRecords: ProcessedDayRecord[] = [];
  let workedDays = 0;
  let lateCount = 0;
  let holidayLateCount = 0;
  let leaveCount = 0;
  let earlyOutCount = 0;
  let ot3Count = 0;
  let ot8Count = 0;

  const normalInMinutes = timeToMinutes(rules.morningWorkStart);
  const normalOutMinutes = timeToMinutes(rules.afternoonWorkEnd);

  daysInRange.forEach(dateStr => {
    const jsDate = new Date(dateStr);
    const dayOfWeek = jsDate.getDay(); // 0 = Sunday, 6 = Saturday
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    const isHoliday = holidays.some(h => h.date === dateStr) || isWeekend;

    const scans = staff.records[dateStr] || [];
    const leaveDay = staffLeaves.find(l => l.date === dateStr);
    
    // Sort scan times
    scans.sort();

    if (scans.length === 0) {
      if (isHoliday) {
        processedRecords.push({
          date: dateStr,
          checkIn: '-',
          checkOut: '-',
          status: 'วันหยุด',
          lateMinutes: 0,
          earlyMinutes: 0
        });
      } else if (leaveDay) {
        if (leaveDay.type === 'แจ้งลืมประทับตราบัตร' && leaveDay.missedScanType === 'ลืมทั้งวัน') {
          workedDays++;
          processedRecords.push({
            date: dateStr,
            checkIn: '08:30 (แจ้ง)',
            checkOut: '16:30 (แจ้ง)',
            status: 'ลืมรูดบัตร',
            lateMinutes: 0,
            earlyMinutes: 0
          });
        } else {
          let statusText = leaveDay.type;
          if (leaveDay.format === 'ครึ่งวันเช้า') {
            statusText = `${leaveDay.type}(เช้า)`;
            leaveCount += 0.5;
          } else if (leaveDay.format === 'ครึ่งวันบ่าย') {
            statusText = `${leaveDay.type}(บ่าย)`;
            leaveCount += 0.5;
          } else {
            leaveCount += 1;
          }
          processedRecords.push({
            date: dateStr,
            checkIn: '-',
            checkOut: '-',
            status: statusText,
            lateMinutes: 0,
            earlyMinutes: 0
          });
        }
      } else {
        // Only classify as absent if we have imported data for this date in the system
        const anyEmployeeHasData = parsedEmployees.some(emp => emp.records[dateStr] && emp.records[dateStr].length > 0);
        if (anyEmployeeHasData) {
          leaveCount++;
          processedRecords.push({
            date: dateStr,
            checkIn: '-',
            checkOut: '-',
            status: 'ลา/ขาดงาน',
            lateMinutes: 0,
            earlyMinutes: 0
          });
        } else {
          processedRecords.push({
            date: dateStr,
            checkIn: '-',
            checkOut: '-',
            status: 'ไม่มีข้อมูล',
            lateMinutes: 0,
            earlyMinutes: 0
          });
        }
      }
      return;
    }

    // Has scans, so worked!
    workedDays++;
    
    // We assume the first scan is check-in, and the last scan (if any other exists) is check-out.
    // If only 1 scan exists:
    // If it's in the morning, it's check-in (missing check-out).
    // If it's in the afternoon, it's check-out (missing check-in).
    let checkIn = '';
    let checkOut = '';

    if (scans.length >= 2) {
      checkIn = scans[0];
      checkOut = scans[scans.length - 1];
    } else if (scans.length === 1) {
      const scanMin = timeToMinutes(scans[0]);
      if (scanMin < normalInMinutes + 120) {
        checkIn = scans[0];
      } else {
        checkOut = scans[0];
      }
    }

    // Apply missed scan stamping adjustments
    if (leaveDay && leaveDay.type === 'แจ้งลืมประทับตราบัตร') {
      if (leaveDay.missedScanType === 'รอบเข้างาน' && !checkIn) {
        checkIn = '08:30 (แจ้ง)';
      } else if (leaveDay.missedScanType === 'รอบออกงาน' && !checkOut) {
        checkOut = '16:30 (แจ้ง)';
      }
    }

    let lateMinutes = 0;
    let earlyMinutes = 0;
    let isLate = false;
    let isHalfDayLate = false;
    let isEarlyOut = false;

    // Calculate Late
    if (checkIn) {
      if (checkIn.includes('(แจ้ง)')) {
        lateMinutes = 0;
        isLate = false;
      } else {
        const inMin = timeToMinutes(checkIn);
        const diff = inMin - normalInMinutes;
        
        let adjustedDiff = diff;
        if (leaveDay && leaveDay.type === 'ลาชั่วโมง') {
          adjustedDiff = Math.max(0, diff - leaveDay.hourMinutes);
        }

        if (adjustedDiff > rules.lateAllowanceMinutes) {
          lateMinutes = adjustedDiff;
          isLate = true;
          if (adjustedDiff >= rules.halfDayLateMinutes) {
            isHalfDayLate = true;
          }
        }
      }
    } else {
      isLate = false;
    }

    if (leaveDay && leaveDay.type === 'ปฏิบัติงานนอกสถานที่') {
      isLate = false;
      isHalfDayLate = false;
      lateMinutes = 0;
    }

    // Calculate Early Checkout
    let isHalfDayAfternoonLeave = false;
    let isHalfDayMorningLeave = false;

    if (leaveDay) {
      if (leaveDay.format === 'ครึ่งวันเช้า') {
        isHalfDayMorningLeave = true;
      } else if (leaveDay.format === 'ครึ่งวันบ่าย') {
        isHalfDayAfternoonLeave = true;
      }
    }

    // Detect "ลาครึ่งวันเช้า" (Morning Leave) automatically from scans if not set
    if (!isHalfDayMorningLeave && scans.length >= 1) {
      const firstScanMin = timeToMinutes(scans[0]);
      if (firstScanMin >= 630 && firstScanMin <= 780) {
        checkIn = scans[0];
        checkOut = scans[scans.length - 1] !== checkIn ? scans[scans.length - 1] : '';
        isHalfDayMorningLeave = true;
      }
    }

    if (checkOut && !isHalfDayMorningLeave) {
      if (checkOut.includes('(แจ้ง)')) {
        earlyMinutes = 0;
        isEarlyOut = false;
      } else {
        const outMin = timeToMinutes(checkOut);
        
        // If check-out is around midday break (e.g. 12:00 - 13:00) and check-in was normal, it counts as "Half-Day Leave (afternoon)"
        if (outMin >= 720 && outMin <= 780) {
          isHalfDayAfternoonLeave = true;
          earlyMinutes = 0;
        } else {
          const diff = normalOutMinutes - outMin;
          
          let adjustedDiff = diff;
          if (leaveDay && leaveDay.type === 'ลาชั่วโมง') {
            adjustedDiff = Math.max(0, diff - leaveDay.hourMinutes);
          }

          if (adjustedDiff > rules.earlyCheckoutAllowanceMinutes) {
            earlyMinutes = adjustedDiff;
            isEarlyOut = true;
          }
        }
      }
    } else if (!checkOut && !isHalfDayMorningLeave) {
      isEarlyOut = true;
    }

    // If it's half day morning leave, checkIn is around noon. Don't mark it as late
    if (isHalfDayMorningLeave && checkIn) {
      isLate = false;
      lateMinutes = 0;
      // Calculate early checkout if checkOut exists
      if (checkOut) {
        if (checkOut.includes('(แจ้ง)')) {
          earlyMinutes = 0;
          isEarlyOut = false;
        } else {
          const outMin = timeToMinutes(checkOut);
          const diff = normalOutMinutes - outMin;
          
          let adjustedDiff = diff;
          if (leaveDay && leaveDay.type === 'ลาชั่วโมง') {
            adjustedDiff = Math.max(0, diff - leaveDay.hourMinutes);
          }

          if (adjustedDiff > rules.earlyCheckoutAllowanceMinutes) {
            earlyMinutes = adjustedDiff;
            isEarlyOut = true;
          }
        }
      } else {
        isEarlyOut = true; // Missing checkout
      }
    }

    if (leaveDay && leaveDay.type === 'ปฏิบัติงานนอกสถานที่') {
      isEarlyOut = false;
      earlyMinutes = 0;
    }

    // Determine final daily status
    let status: ProcessedDayRecord['status'] = 'ปกติ';
    
    // Check if scans are missing entirely to set specific warning status
    let isMissingCheckOut = !checkOut && scans.length === 1;
    let isMissingCheckIn = !checkIn && scans.length === 1;

    // Detect OT3 (Weekday Overtime >= 3 hours, after 16:30, meaning checkOut >= 19:30 on working days)
    // 19:30 in minutes is 19 * 60 + 30 = 1170. Normal workdays (not weekend/holiday)
    let isOT3 = false;
    if (checkOut && !isHoliday) {
      const outMin = timeToMinutes(checkOut);
      if (outMin >= 1170) {
        isOT3 = true;
      }
    }

    // Detect OT8 (Holiday Overtime >= 8 hours, 8:30 to 16:30 on weekends/holidays)
    let isOT8 = false;
    if (checkIn && checkOut && isHoliday) {
      const inMin = timeToMinutes(checkIn);
      const outMin = timeToMinutes(checkOut);
      
      // Calculate work duration excluding 1 hour break (normally 8:30 to 16:30 is 8 hours total duration)
      // Check if check-in is around work start (e.g. <= 08:30 + allowance) and check-out is around work end (e.g. >= 16:30 - allowance)
      const workDurationMinutes = outMin - inMin;
      // 8 hours is 480 minutes (excluding 60 min break means 8:30 to 16:30 is 8 hours total span)
      if (workDurationMinutes >= 480 - rules.earlyCheckoutAllowanceMinutes) {
        isOT8 = true;
      }
    }

    if (isHalfDayAfternoonLeave) {
      status = leaveDay ? `${leaveDay.type}(บ่าย)` : 'ลาครึ่งวันบ่าย';
    } else if (isHalfDayMorningLeave) {
      status = leaveDay ? `${leaveDay.type}(เช้า)` : 'ลาครึ่งวันเช้า';
    } else if (isMissingCheckOut) {
      status = 'ไม่สแกนออก';
      earlyMinutes = 0; // Clear early minutes since it was a missing scan, not early checkout
    } else if (isMissingCheckIn) {
      status = 'ไม่สแกนเข้า';
      lateMinutes = 0; // Clear late minutes
    } else if (isHalfDayLate) {
      status = 'สายครึ่งวัน';
      if (isHoliday) {
        holidayLateCount++;
      } else {
        lateCount++;
      }
    } else if (isLate) {
      status = 'สาย';
      if (isHoliday) {
        holidayLateCount++;
      } else {
        lateCount++;
      }
    } else if (isEarlyOut) {
      status = 'ออกก่อนเวลา';
      earlyOutCount++;
    } else if (isOT3) {
      status = 'OT3';
      ot3Count++;
    } else if (isOT8) {
      status = 'OT8';
      ot8Count++;
    } else if (checkIn.includes('(แจ้ง)') || checkOut.includes('(แจ้ง)')) {
      status = 'ลืมรูดบัตร';
    } else if (leaveDay && leaveDay.type === 'ปฏิบัติงานนอกสถานที่') {
      status = 'นอกสถานที่';
    } else if (leaveDay && leaveDay.type === 'ลาชั่วโมง') {
      status = 'ลาชั่วโมง';
    }

    processedRecords.push({
      date: dateStr,
      checkIn: checkIn || 'ไม่มี',
      checkOut: checkOut || 'ไม่มี',
      status,
      lateMinutes,
      earlyMinutes
    });

    // Accumulate leave count: if half-day leave, add 0.5 to leaveCount
    if (isHalfDayAfternoonLeave || isHalfDayMorningLeave) {
      leaveCount += 0.5;
    }
  });

  return {
    id: staff.id,
    name: staff.name,
    department: staff.department,
    position: staff.position,
    workedDays,
    lateCount,
    holidayLateCount,
    leaveCount,
    earlyOutCount,
    ot3Count,
    ot8Count,
    records: processedRecords
  };
}

// Main table renderer
let currentProcessedSummaries: ProcessedStaffSummary[] = [];

function populateDepartmentFilter() {
  const departments = new Set<string>();
  parsedEmployees.forEach(e => {
    if (e.department) {
      departments.add(e.department.trim());
    }
  });

  const sortedDepts = Array.from(departments).sort();
  const currentOptions = Array.from(filterDepartmentSelect.options)
    .map(opt => opt.value)
    .filter(val => val !== 'all');
    
  const isSame = currentOptions.length === sortedDepts.length && 
    currentOptions.every((val, index) => val === sortedDepts[index]);
    
  if (isSame) return;

  const selectedValue = filterDepartmentSelect.value || 'all';
  filterDepartmentSelect.innerHTML = '<option value="all">ทั้งหมด</option>';
  sortedDepts.forEach(dept => {
    const opt = document.createElement('option');
    opt.value = dept;
    opt.textContent = dept;
    filterDepartmentSelect.appendChild(opt);
  });

  if (departments.has(selectedValue)) {
    filterDepartmentSelect.value = selectedValue;
  } else {
    filterDepartmentSelect.value = 'all';
  }
}

function recalculateAndRender() {
  if (parsedEmployees.length === 0) return;

  populateDepartmentFilter();

  let startVal = filterStartDate.value;
  let endVal = filterEndDate.value;

  if (filterModeSelect.value === 'monthly') {
    const year = filterYearSelect.value;
    const month = filterMonthSelect.value;
    startVal = `${year}-${month}-01`;
    const lastDay = new Date(Number(year), Number(month), 0).getDate();
    endVal = `${year}-${month}-${String(lastDay).padStart(2, '0')}`;
  }

  const searchVal = searchStaff.value.trim().toLowerCase();
  const deptVal = filterDepartmentSelect.value;

  // Run business calculations for all employees
  currentProcessedSummaries = parsedEmployees.map(emp => 
    calculateStaffRecords(emp, dbState.rules, dbState.holidays, startVal, endVal)
  );

  // Filter based on search input and department select
  let filtered = currentProcessedSummaries.filter(summary => 
    summary.name.toLowerCase().includes(searchVal) || 
    summary.id.toLowerCase().includes(searchVal)
  );

  if (deptVal && deptVal !== 'all') {
    filtered = filtered.filter(summary => summary.department === deptVal);
  }

  // Sort based on currentSort state
  if (currentSort.column) {
    const col = currentSort.column;
    const dir = currentSort.direction === 'asc' ? 1 : -1;
    filtered.sort((a, b) => {
      let valA: string | number = '';
      let valB: string | number = '';

      if (col === 'id') {
        valA = a.id;
        valB = b.id;
      } else if (col === 'name') {
        valA = a.name;
        valB = b.name;
      } else if (col === 'workedDays') {
        valA = a.workedDays;
        valB = b.workedDays;
      } else if (col === 'lateCount') {
        valA = a.lateCount + a.holidayLateCount;
        valB = b.lateCount + b.holidayLateCount;
      } else if (col === 'leaveCount') {
        valA = a.leaveCount;
        valB = b.leaveCount;
      } else if (col === 'earlyOutCount') {
        valA = a.earlyOutCount;
        valB = b.earlyOutCount;
      } else if (col === 'ot3Count') {
        valA = a.ot3Count;
        valB = b.ot3Count;
      } else if (col === 'ot8Count') {
        valA = a.ot8Count;
        valB = b.ot8Count;
      }

      if (typeof valA === 'string' && typeof valB === 'string') {
        return valA.localeCompare(valB, 'th') * dir;
      }
      return ((valA as number) - (valB as number)) * dir;
    });
  }

  // Update header icons
  document.querySelectorAll('.sortable').forEach(th => {
    const col = th.getAttribute('data-sort');
    const iconSpan = th.querySelector('.sort-icon');
    if (iconSpan) {
      if (currentSort.column === col) {
        iconSpan.textContent = currentSort.direction === 'asc' ? '▲' : '▼';
        (th as HTMLElement).style.color = '#60a5fa'; // Light blue sort highlight
      } else {
        iconSpan.textContent = '↕';
        (th as HTMLElement).style.color = '';
      }
    }
  });

  summaryTableBody.innerHTML = '';
  
  if (filtered.length === 0) {
    summaryTableBody.innerHTML = `
      <tr>
        <td colspan="10" class="text-center text-muted">ไม่พบข้อมูลตามคำค้นหาที่ระบุ</td>
      </tr>
    `;
    return;
  }

  const getShortDate = (dateStr: string) => {
    const p = dateStr.split('-');
    return p.length === 3 ? `${p[2]}/${p[1]}` : dateStr;
  };

  filtered.forEach(s => {
    // Extract Late dates (separated into normal workday and holiday/weekend)
    const normalLateDates: string[] = [];
    const holidayLateDates: string[] = [];
    s.records.forEach(r => {
      if (r.status === 'สาย' || r.status === 'สายครึ่งวัน') {
        const isWeekend = new Date(r.date).getDay() === 0 || new Date(r.date).getDay() === 6;
        const isHoliday = dbState.holidays.some(h => h.date === r.date) || isWeekend;
        const shortD = getShortDate(r.date);
        if (isHoliday) {
          holidayLateDates.push(shortD);
        } else {
          normalLateDates.push(shortD);
        }
      }
    });

    let lateDatesHTML = '';
    if (normalLateDates.length > 0 && holidayLateDates.length > 0) {
      lateDatesHTML = `${normalLateDates.join(', ')} (${holidayLateDates.join(', ')})`;
    } else if (normalLateDates.length > 0) {
      lateDatesHTML = normalLateDates.join(', ');
    } else if (holidayLateDates.length > 0) {
      lateDatesHTML = `(${holidayLateDates.join(', ')})`;
    }
      
    // Extract Leave dates
    const leaveDates = s.records
      .filter(r => r.status === 'ลา/ขาดงาน' || r.status === 'ลาครึ่งวันเช้า' || r.status === 'ลาครึ่งวันบ่าย')
      .map(r => {
        const shortD = getShortDate(r.date);
        if (r.status === 'ลาครึ่งวันเช้า') return `${shortD}(เช้า)`;
        if (r.status === 'ลาครึ่งวันบ่าย') return `${shortD}(บ่าย)`;
        return shortD;
      });

    // Extract Early dates
    const earlyDates = s.records
      .filter(r => r.status === 'ออกก่อนเวลา')
      .map(r => getShortDate(r.date));

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align: center;"><input type="checkbox" class="select-emp-print" data-id="${s.id}" checked /></td>
      <td><strong>${s.id}</strong></td>
      <td>${s.name}</td>
      <td><span class="badge badge-info">${s.workedDays} วัน</span></td>
      
      <td>
        <span class="badge ${s.lateCount > 0 || s.holidayLateCount > 0 ? 'badge-danger' : 'badge-success'}">
          ${s.holidayLateCount > 0 ? `${s.lateCount}(+${s.holidayLateCount}) ครั้ง` : `${s.lateCount} ครั้ง`}
        </span>
        ${lateDatesHTML ? `<div style="font-size: 0.75rem; opacity: 0.8; margin-top: 4px; max-width: 140px; word-wrap: break-word; color: #f87171;">${lateDatesHTML}</div>` : ''}
      </td>
      
      <td>
        <span class="badge ${s.leaveCount > 0 ? 'badge-warning' : 'badge-success'}">${s.leaveCount} วัน</span>
        ${leaveDates.length > 0 ? `<div style="font-size: 0.75rem; opacity: 0.8; margin-top: 4px; max-width: 140px; word-wrap: break-word; color: #fbbf24;">${leaveDates.join(', ')}</div>` : ''}
      </td>
      
      <td>
        <span class="badge ${s.earlyOutCount > 0 ? 'badge-warning' : 'badge-success'}">${s.earlyOutCount} ครั้ง</span>
        ${earlyDates.length > 0 ? `<div style="font-size: 0.75rem; opacity: 0.8; margin-top: 4px; max-width: 140px; word-wrap: break-word; color: #fbbf24;">${earlyDates.join(', ')}</div>` : ''}
      </td>
      
      <td>
        <span class="badge ${s.ot3Count > 0 ? 'badge-info' : 'badge-success'}">${s.ot3Count} วัน</span>
      </td>

      <td>
        <span class="badge ${s.ot8Count > 0 ? 'badge-info' : 'badge-success'}">${s.ot8Count} วัน</span>
      </td>
      
      <td><button class="btn btn-secondary btn-sm btn-view-detail" data-id="${s.id}">🔍 รายละเอียด</button></td>
    `;
    summaryTableBody.appendChild(tr);
  });

  // Sync select all status checkbox to newly rendered rows
  const isAllChecked = selectAllPrint.checked;
  document.querySelectorAll('.select-emp-print').forEach((cb) => {
    (cb as HTMLInputElement).checked = isAllChecked;
  });

  // Attach Detail handlers
  document.querySelectorAll('.btn-view-detail').forEach((btn) => {
    btn.addEventListener('click', (e) => {
      const id = (e.currentTarget as HTMLElement).getAttribute('data-id');
      const staffSummary = currentProcessedSummaries.find(s => s.id === id);
      if (staffSummary) {
        showStaffDetail(staffSummary);
      }
    });
  });
}

// Show Detail Modal
function showStaffDetail(staff: ProcessedStaffSummary) {
  modalTitle.textContent = `รายละเอียดเวลาปฏิบัติงาน: ${staff.name} (${staff.id}) - ${staff.position}`;
  detailTableBody.innerHTML = '';

    // Helper to format date with Thai day prefix: e.g. "จ. 03/10/2022"
    const formatDateWithThaiDay = (dateStr: string): string => {
      const days = ['อา.', 'จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.'];
      const dateObj = new Date(dateStr);
      const dayName = days[dateObj.getDay()];
      
      const parts = dateStr.split('-');
      if (parts.length === 3) {
        const yyyy = parts[0];
        const mm = parts[1];
        const dd = parts[2];
        return `${dayName} ${dd}/${mm}/${yyyy}`;
      }
      return dateStr;
    };

    staff.records.forEach(r => {
      let badgeClass = 'badge-success';
      if (r.status === 'สาย') badgeClass = 'badge-danger';
      else if (r.status === 'สายครึ่งวัน') badgeClass = 'badge-danger';
      else if (r.status === 'ไม่สแกนออก' || r.status === 'ไม่สแกนเข้า') badgeClass = 'badge-danger';
      else if (r.status === 'ออกก่อนเวลา' || r.status === 'ลา/ขาดงาน' || r.status === 'ลาครึ่งวันเช้า' || r.status === 'ลาครึ่งวันบ่าย') badgeClass = 'badge-warning';
      else if (r.status === 'วันหยุด' || r.status.startsWith('OT')) badgeClass = 'badge-info';
      else if (r.status === 'ไม่มีข้อมูล') badgeClass = 'badge-secondary';

      // Check if weekend or holiday
      const dateObj = new Date(r.date);
      const dayOfWeek = dateObj.getDay();
      const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
      const isHoliday = dbState.holidays.some(h => h.date === r.date);

      const tr = document.createElement('tr');
      if (isWeekend || isHoliday) {
        tr.classList.add('highlight-weekend');
      }
      tr.innerHTML = `
        <td>${formatDateWithThaiDay(r.date)}</td>
        <td>${r.checkIn}</td>
        <td>${r.checkOut}</td>
        <td><span class="badge ${badgeClass}">${r.status}</span></td>
        <td>${r.lateMinutes > 0 ? r.lateMinutes + ' นาที' : '-'}</td>
        <td>${r.earlyMinutes > 0 ? r.earlyMinutes + ' นาที' : '-'}</td>
      `;
      detailTableBody.appendChild(tr);
    });

  detailModal.style.display = 'flex';
}

// CSV Export logic
function handleCSVExport() {
  if (currentProcessedSummaries.length === 0) {
    alert('ไม่มีข้อมูลสำหรับส่งออกรายงาน');
    return;
  }

  // Column Headers
  const csvHeaders = ['รหัสพนักงาน', 'ชื่อพนักงาน', 'ตำแหน่งงาน', 'ฝ่าย/หน่วยงาน', 'มาทำงาน (วัน)', 'สาย (ครั้ง)', 'ลา/ขาดงาน (วัน)', 'ออกก่อนเวลา (ครั้ง)'];
  const csvRows = [csvHeaders.join(',')];

  currentProcessedSummaries.forEach(s => {
    const row = [
      `"${s.id}"`,
      `"${s.name}"`,
      `"${s.position}"`,
      `"${s.department}"`,
      s.workedDays,
      s.lateCount,
      s.leaveCount,
      s.earlyOutCount
    ];
    csvRows.push(row.join(','));
  });

  // Convert to Blob with UTF-8 BOM for Thai encoding in Excel
  const csvContent = '\uFEFF' + csvRows.join('\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  
  const link = document.createElement("a");
  link.href = url;
  link.download = `รายงานสรุปเวลาทำงาน_${new Date().toISOString().split('T')[0]}.csv`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

// Print Report Generation
function handlePrintReports() {
  if (currentProcessedSummaries.length === 0) {
    alert('ไม่มีข้อมูลสำหรับพิมพ์รายงาน');
    return;
  }

  const checkedBoxes = document.querySelectorAll('.select-emp-print:checked');
  const checkedIds = Array.from(checkedBoxes).map(cb => (cb as HTMLInputElement).getAttribute('data-id'));
  
  if (checkedIds.length === 0) {
    alert('กรุณาเลือกเจ้าหน้าที่อย่างน้อย 1 คนเพื่อพิมพ์รายงาน');
    return;
  }

  const printableSummaries = currentProcessedSummaries.filter(s => checkedIds.includes(s.id));

  // Get active month and year text for header
  let periodText = '';
  if (filterModeSelect.value === 'monthly') {
    const monthsThai = [
      'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
      'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
    ];
    const monthIndex = parseInt(filterMonthSelect.value, 10) - 1;
    const yearThai = parseInt(filterYearSelect.value, 10) + 543; // Convert to Buddhist Era
    periodText = `${monthsThai[monthIndex]} พ.ศ. ${yearThai}`;
  } else {
    const startParts = filterStartDate.value.split('-');
    const endParts = filterEndDate.value.split('-');
    const formatPart = (p: string[]) => p.length === 3 ? `${p[2]}/${p[1]}/${parseInt(p[0], 10) + 543}` : '';
    periodText = `ระหว่างวันที่ ${formatPart(startParts)} ถึง ${formatPart(endParts)}`;
  }

  printContainer.innerHTML = '';

  const getShortDate = (dateStr: string) => {
    const p = dateStr.split('-');
    return p.length === 3 ? `${p[2]}/${p[1]}` : dateStr;
  };

  const formatDateWithThaiDayPrint = (dateStr: string): string => {
    const days = ['อา.', 'จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.'];
    const dateObj = new Date(dateStr);
    const dayName = days[dateObj.getDay()];
    const parts = dateStr.split('-');
    if (parts.length === 3) {
      const thaiYear = String(parseInt(parts[0], 10) + 543).slice(-2);
      return `${dayName} ${parts[2]}/${parts[1]}/${thaiYear}`;
    }
    return dateStr;
  };

  printableSummaries.forEach(s => {
    const page = document.createElement('div');
    page.className = 'print-page';

    // Split records into 2 halves for side-by-side view
    const half = Math.ceil(s.records.length / 2);
    const leftRecords = s.records.slice(0, half);
    const rightRecords = s.records.slice(half);

    const showStatusCol = printShowStatusCheckbox.checked;
    const showVerification = printShowVerificationCheckbox.checked;

    const renderTableColumn = (records: typeof s.records) => {
      let rowsHTML = '';
      records.forEach(r => {
        let statusStyle = '';
        if (r.status === 'สาย' || r.status === 'สายครึ่งวัน' || r.status === 'ไม่สแกนออก' || r.status === 'ไม่สแกนเข้า' || r.status === 'ลา/ขาดงาน') {
          statusStyle = 'font-weight: bold; color: black;';
        }
        rowsHTML += `
          <tr>
            <td style="padding: 3px 4px !important; font-size: 13px;">${formatDateWithThaiDayPrint(r.date)}</td>
            <td style="padding: 3px 4px !important; font-size: 13px;">${r.checkIn || '-'}</td>
            <td style="padding: 3px 4px !important; font-size: 13px;">${r.checkOut || '-'}</td>
            ${showStatusCol ? `<td style="${statusStyle} padding: 3px 4px !important; font-size: 13px;">${r.status}</td>` : ''}
            <td style="padding: 3px 4px !important; font-size: 13px;"></td>
          </tr>
        `;
      });

      return `
        <table class="print-table" style="width: 100%; margin-bottom: 0; font-size: 13px; border-collapse: collapse;">
          <thead>
            <tr>
              <th style="padding: 3px 4px !important; font-size: 13px; width: ${showStatusCol ? '28%' : '38%'};">วันที่</th>
              <th style="padding: 3px 4px !important; font-size: 13px; width: ${showStatusCol ? '17%' : '22%'};">สแกนเข้า</th>
              <th style="padding: 3px 4px !important; font-size: 13px; width: ${showStatusCol ? '17%' : '22%'};">สแกนออก</th>
              ${showStatusCol ? `<th style="padding: 3px 4px !important; font-size: 13px; width: 18%;">สถานะ</th>` : ''}
              <th style="padding: 3px 4px !important; font-size: 13px; width: 20%;">หมายเหตุ</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHTML}
          </tbody>
        </table>
      `;
    };

    const sideBySideTablesHTML = `
      <div style="display: flex; gap: 12px; margin-bottom: 10px; align-items: start;">
        <div style="flex: 1;">
          ${renderTableColumn(leftRecords)}
        </div>
        <div style="flex: 1;">
          ${renderTableColumn(rightRecords)}
        </div>
      </div>
    `;

    // Differentiate print late dates
    const normalLateDatesPrint: string[] = [];
    const holidayLateDatesPrint: string[] = [];
    s.records.forEach(r => {
      if (r.status === 'สาย' || r.status === 'สายครึ่งวัน') {
        const isWeekend = new Date(r.date).getDay() === 0 || new Date(r.date).getDay() === 6;
        const isHoliday = dbState.holidays.some(h => h.date === r.date) || isWeekend;
        const shortD = getShortDate(r.date);
        if (isHoliday) {
          holidayLateDatesPrint.push(shortD);
        } else {
          normalLateDatesPrint.push(shortD);
        }
      }
    });



    const verificationTableHTML = showVerification ? `
      <div class="print-verification-title">2. ตารางสรุปเพื่อลงข้อมูลยืนยันจากเจ้าหน้าที่</div>
      <table class="print-verification-table">
        <thead>
          <tr>
            <th style="width: 25%;">ประเภทวันลา/สาย</th>
            <th style="width: 55%;">วันที่</th>
            <th style="width: 20%;">รวม (วัน/ครั้ง)</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="font-weight: bold;">พักร้อน</td>
            <td></td>
            <td></td>
          </tr>
          <tr>
            <td style="font-weight: bold;">ลากิจ</td>
            <td></td>
            <td></td>
          </tr>
          <tr>
            <td style="font-weight: bold;">ลาป่วย</td>
            <td></td>
            <td></td>
          </tr>
          <tr>
            <td style="font-weight: bold;">สาย (วันธรรมดา)</td>
            <td style="font-size: 11px;">${normalLateDatesPrint.join(', ')}</td>
            <td style="text-align: center; font-weight: bold;">${s.lateCount > 0 ? s.lateCount + ' ครั้ง' : '-'}</td>
          </tr>
          <tr>
            <td style="font-weight: bold;">สาย (วันหยุด)</td>
            <td style="font-size: 11px;">${holidayLateDatesPrint.join(', ')}</td>
            <td style="text-align: center; font-weight: bold;">${s.holidayLateCount > 0 ? s.holidayLateCount + ' ครั้ง' : '-'}</td>
          </tr>
          <tr>
            <td style="font-weight: bold;">อื่นๆ</td>
            <td></td>
            <td></td>
          </tr>
        </tbody>
      </table>
    ` : '';

    page.innerHTML = `
      <div class="print-header" style="margin-bottom: 10px; padding-bottom: 5px;">
        <h1 style="font-size: 16px; margin: 0 0 3px 0;">ใบตรวจสอบเวลาปฏิบัติงานและวันลาเจ้าหน้าที่</h1>
        <div style="font-size: 12px; font-weight: bold;">ประจำเดือน ${periodText}</div>
      </div>
      
      <div class="print-info-grid" style="margin-bottom: 10px; font-size: 12px; grid-template-columns: repeat(4, 1fr);">
        <div><strong>ชื่อ-นามสกุล:</strong> ${s.name}</div>
        <div><strong>รหัสประจำตัว:</strong> ${s.id}</div>
        <div><strong>ตำแหน่ง:</strong> ${s.position}</div>
        <div><strong>ฝ่าย/หน่วยงาน:</strong> ${s.department}</div>
      </div>

      <h3 style="font-size: 12px; font-weight: bold; margin-top: 0; margin-bottom: 5px; text-align: left;">1. ประวัติเวลาปฏิบัติงานจริงประจำเดือน</h3>
      ${sideBySideTablesHTML}

      ${verificationTableHTML}
    `;

    printContainer.appendChild(page);
  });

  window.print();
}

function parseCSV(text: string): string[][] {
  const lines: string[][] = [];
  let row: string[] = [];
  let inQuotes = false;
  let currentVal = '';

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];

    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        currentVal += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      row.push(currentVal.trim());
      currentVal = '';
    } else if ((char === '\r' || char === '\n') && !inQuotes) {
      row.push(currentVal.trim());
      if (row.length > 0 && row.some(cell => cell !== '')) {
        lines.push(row);
      }
      row = [];
      currentVal = '';
      if (char === '\r' && nextChar === '\n') {
        i++;
      }
    } else {
      currentVal += char;
    }
  }
  if (currentVal || row.length > 0) {
    row.push(currentVal.trim());
    if (row.some(cell => cell !== '')) {
      lines.push(row);
    }
  }
  return lines;
}

function mapCSVToLeaveRecords(csvRows: string[][]): LeaveRecord[] {
  if (csvRows.length < 2) return [];
  const headers = csvRows[0].map(h => h.toLowerCase());

  const idxId = headers.findIndex(h => h.includes('รหัส') || h.includes('id') || h.includes('staff'));
  const idxName = headers.findIndex(h => h.includes('ชื่อ') || h.includes('name'));
  const idxType = headers.findIndex(h => h.includes('ประเภท') || h.includes('leave type') || h.includes('การลา'));
  const idxFormat = headers.findIndex(h => h.includes('รูปแบบ') || h.includes('ลักษณะ') || h.includes('ครึ่งวัน') || h.includes('format'));
  const idxMin = headers.findIndex(h => h.includes('นาที') || h.includes('ชั่วโมง') || h.includes('minute') || h.includes('hour'));
  const idxMissed = headers.findIndex(h => h.includes('ลืม') || h.includes('ประทับตรา') || h.includes('รูด') || h.includes('missed'));
  const idxStart = headers.findIndex(h => h.includes('เริ่ม') || h.includes('วันที่ลา') || h.includes('วันที่') || h.includes('start'));
  const idxEnd = headers.findIndex(h => h.includes('สิ้นสุด') || h.includes('ถึงวันที่') || h.includes('end'));

  const records: LeaveRecord[] = [];

  for (let i = 1; i < csvRows.length; i++) {
    const row = csvRows[i];
    const staffId = idxId !== -1 ? row[idxId] : '';
    const staffName = idxName !== -1 ? row[idxName] : '';
    if (!staffId) continue;

    const leaveType = idxType !== -1 ? row[idxType] : '';
    const leaveFormat = idxFormat !== -1 ? row[idxFormat] : 'เต็มวัน';
    
    let hourMinutes = 0;
    if (idxMin !== -1 && row[idxMin]) {
      hourMinutes = parseInt(row[idxMin], 10) || 0;
    }

    const missedScanType = idxMissed !== -1 ? row[idxMissed] : '';
    const rawStart = idxStart !== -1 ? row[idxStart] : '';
    const rawEnd = idxEnd !== -1 ? row[idxEnd] : '';

    const startDate = parseDateToISO(rawStart);
    let endDate = parseDateToISO(rawEnd);
    if (!endDate) {
      endDate = startDate;
    }

    records.push({
      timestamp: row[0] || '',
      staffId,
      staffName,
      leaveType,
      leaveFormat,
      hourMinutes,
      missedScanType,
      startDate,
      endDate
    });
  }

  return records;
}

function renderLeaveSummaryTable() {
  const tbody = document.getElementById('leave-summary-table-body');
  if (!tbody) return;

  const records = dbState.leaveRecords || [];
  if (records.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="8" class="text-center text-muted" style="padding: 12px; border: 1px solid var(--border-color); background-color: #0f172a;">ยังไม่มีการเชื่อมต่อหรือข้อมูลดิบของการขอลา</td>
      </tr>
    `;
    const countBadge = document.getElementById('leave-sync-count');
    if (countBadge) countBadge.textContent = '0 รายการ';
    return;
  }

  const statsMap = new Map<string, {
    name: string,
    vacation: number,
    personal: number,
    sick: number,
    hourMinutes: number,
    fieldwork: number,
    forgot: number
  }>();

  records.forEach(r => {
    if (!statsMap.has(r.staffId)) {
      statsMap.set(r.staffId, {
        name: r.staffName,
        vacation: 0,
        personal: 0,
        sick: 0,
        hourMinutes: 0,
        fieldwork: 0,
        forgot: 0
      });
    }

    const s = statsMap.get(r.staffId)!;
    let weight = 1;
    if (r.leaveFormat === 'ครึ่งวันเช้า' || r.leaveFormat === 'ครึ่งวันบ่าย') {
      weight = 0.5;
    }

    const type = r.leaveType;
    if (type.includes('พักร้อน')) {
      s.vacation += weight;
    } else if (type.includes('กิจ')) {
      s.personal += weight;
    } else if (type.includes('ป่วย')) {
      s.sick += weight;
    } else if (type.includes('ชั่วโมง')) {
      s.hourMinutes += r.hourMinutes;
    } else if (type.includes('นอกสถานที่')) {
      s.fieldwork += 1;
    } else if (type.includes('ลืม') || type.includes('ประทับตรา')) {
      s.forgot += 1;
    }
  });

  tbody.innerHTML = '';
  statsMap.forEach((s, id) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="padding: 8px; text-align: center; border: 1px solid var(--border-color);">${id}</td>
      <td style="padding: 8px; text-align: left; border: 1px solid var(--border-color);">${s.name}</td>
      <td style="padding: 8px; text-align: center; border: 1px solid var(--border-color);">${s.vacation} วัน</td>
      <td style="padding: 8px; text-align: center; border: 1px solid var(--border-color);">${s.personal} วัน</td>
      <td style="padding: 8px; text-align: center; border: 1px solid var(--border-color);">${s.sick} วัน</td>
      <td style="padding: 8px; text-align: center; border: 1px solid var(--border-color);">${s.hourMinutes} นาที</td>
      <td style="padding: 8px; text-align: center; border: 1px solid var(--border-color);">${s.fieldwork} ครั้ง</td>
      <td style="padding: 8px; text-align: center; border: 1px solid var(--border-color);">${s.forgot} ครั้ง</td>
    `;
    tbody.appendChild(tr);
  });

  const countBadge = document.getElementById('leave-sync-count');
  if (countBadge) {
    countBadge.textContent = `${records.length} รายการ`;
  }
}

function renderLeaveAlerts() {
  const alertPanel = document.getElementById('leave-alert-panel');
  const alertList = document.getElementById('leave-alert-list');
  if (!alertPanel || !alertList) return;

  const records = dbState.leaveRecords || [];
  if (records.length === 0) {
    alertPanel.style.display = 'none';
    return;
  }

  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = tomorrow.toISOString().split('T')[0];
  
  const tomorrowLeaves: { name: string, type: string, format: string }[] = [];

  records.forEach(r => {
    const days = expandLeaveRecordDays(r);
    const hasTomorrow = days.some(d => d.date === tomorrowStr);
    if (hasTomorrow) {
      tomorrowLeaves.push({
        name: r.staffName,
        type: r.leaveType,
        format: r.leaveFormat
      });
    }
  });

  if (tomorrowLeaves.length > 0) {
    alertPanel.style.display = 'block';
    alertList.innerHTML = tomorrowLeaves
      .map(l => {
        const formatText = l.format && l.format !== 'เต็มวัน' ? ` (${l.format})` : '';
        return `• <strong>${l.name}</strong> ขอ <strong>${l.type}${formatText}</strong> ในวันพรุ่งนี้ (${formatDateThai(tomorrowStr)})`;
      })
      .join('<br/>');
      
    // Trigger push notification once
    triggerBrowserNotification(tomorrowLeaves.length);
  } else {
    alertPanel.style.display = 'none';
  }
}

function formatDateThai(dateStr: string): string {
  const p = dateStr.split('-');
  if (p.length !== 3) return dateStr;
  const monthsThai = [
    'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
    'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'
  ];
  const day = parseInt(p[2], 10);
  const month = monthsThai[parseInt(p[1], 10) - 1];
  const year = parseInt(p[0], 10) + 543;
  return `${day} ${month} ${year}`;
}

function triggerBrowserNotification(tomorrowLeavesCount: number) {
  if (!("Notification" in window)) return;
  if (Notification.permission === "granted") {
    new Notification("🔔 มีเจ้าหน้าที่ลาปฏิบัติงานวันพรุ่งนี้", {
      body: `วันพรุ่งนี้มีผู้แจ้งขอลาจำนวน ${tomorrowLeavesCount} คน โปรดวางแผนกำลังพลล่วงหน้า`,
    });
  } else if (Notification.permission !== "denied") {
    Notification.requestPermission().then(permission => {
      if (permission === "granted") {
        new Notification("🔔 มีเจ้าหน้าที่ลาปฏิบัติงานวันพรุ่งนี้", {
          body: `วันพรุ่งนี้มีผู้แจ้งขอลาจำนวน ${tomorrowLeavesCount} คน โปรดวางแผนกำลังพลล่วงหน้า`
        });
      }
    });
  }
}

// Bootstrap
document.addEventListener('DOMContentLoaded', init);
