import './style.css';
import {
  loadFromLocalStorage, 
  saveToLocalStorage, 
  exportToJSONFile, 
  importFromJSONFile, 
  sanitizeDB,
} from './dbSync';
import type {
  TimeStampDB,
  Holiday,
  RuleSettings
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
const btnClearFile = document.getElementById('btn-clear-file') as HTMLButtonElement;

const statStaffCount = document.getElementById('stat-staff-count') as HTMLElement;
const statDaysCount = document.getElementById('stat-days-count') as HTMLElement;
const statsOverview = document.getElementById('stats-overview') as HTMLElement;

const rulesForm = document.getElementById('rules-form') as HTMLFormElement;
const morningWorkStartInput = document.getElementById('morningWorkStart') as HTMLInputElement;
const afternoonWorkEndInput = document.getElementById('afternoonWorkEnd') as HTMLInputElement;
const lateAllowanceInput = document.getElementById('lateAllowanceMinutes') as HTMLInputElement;
const earlyAllowanceInput = document.getElementById('earlyCheckoutAllowanceMinutes') as HTMLInputElement;
const halfDayLateInput = document.getElementById('halfDayLateMinutes') as HTMLInputElement;

const holidayDateInput = document.getElementById('holiday-date') as HTMLInputElement;
const holidayNameInput = document.getElementById('holiday-name') as HTMLInputElement;
const btnAddHoliday = document.getElementById('btn-add-holiday') as HTMLButtonElement;
const holidayListBody = document.getElementById('holiday-list-body') as HTMLElement;

const filterStartDate = document.getElementById('filter-start-date') as HTMLInputElement;
const filterEndDate = document.getElementById('filter-end-date') as HTMLInputElement;
const searchStaff = document.getElementById('search-staff') as HTMLInputElement;
const btnExportExcel = document.getElementById('btn-export-excel') as HTMLButtonElement;

const summaryTableBody = document.getElementById('summary-table-body') as HTMLElement;

// Modal Elements
const detailModal = document.getElementById('detail-modal') as HTMLElement;
const modalTitle = document.getElementById('modal-title') as HTMLElement;
const detailTableBody = document.getElementById('detail-table-body') as HTMLElement;
const btnCloseModal = document.getElementById('btn-close-modal') as HTMLButtonElement;

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
    };
    saveToLocalStorage(dbState);
    alert('บันทึกเงื่อนไขและเกณฑ์เวลาแล้ว!');
    recalculateAndRender();
  });

  // Holiday Handlers
  btnAddHoliday.addEventListener('click', handleAddHoliday);
  
  // Filters Event Listeners
  filterStartDate.addEventListener('change', recalculateAndRender);
  filterEndDate.addEventListener('change', recalculateAndRender);
  searchStaff.addEventListener('input', recalculateAndRender);

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
      saveToLocalStorage(dbState);
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
  saveToLocalStorage(dbState);
  renderHolidays();
  
  // Clear input fields
  holidayDateInput.value = '';
  holidayNameInput.value = '';
  
  recalculateAndRender();
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
      recalculateAndRender();
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
    saveToLocalStorage(dbState);
    renderRules();
    renderHolidays();
    updateSyncStatus(true, `เชื่อมต่อสำเร็จ: ${file.name}`);
    alert('นำเข้าฐานข้อมูลจาก OneDrive สำเร็จแล้ว!');
    recalculateAndRender();
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
    
    parsedEmployees = await parseExcelFile(file);
    
    // Display file stats
    fileInfoName.textContent = `${file.name} (${parsedEmployees.length} พนักงาน)`;
    dropzone.style.display = 'none';
    fileInfo.style.display = 'flex';

    // Extract Date limits
    let allDates: string[] = [];
    parsedEmployees.forEach(e => {
      allDates.push(...Object.keys(e.records));
    });
    
    if (allDates.length > 0) {
      allDates.sort();
      const minDate = allDates[0];
      const maxDate = allDates[allDates.length - 1];
      filterStartDate.value = minDate;
      filterEndDate.value = maxDate;
      statDaysCount.textContent = String(new Set(allDates).size);
    }
    
    statStaffCount.textContent = String(parsedEmployees.length);
    statsOverview.style.display = 'grid';

    recalculateAndRender();
  } catch (err) {
    console.error(err);
    alert('เกิดข้อผิดพลาดในการเปิดและสแกนไฟล์ Excel กรุณาตรวจสอบรูปแบบไฟล์ต้นฉบับ');
  }
}

function clearLoadedData() {
  parsedEmployees = [];
  excelFileInput.value = '';
  dropzone.style.display = 'block';
  fileInfo.style.display = 'none';
  statsOverview.style.display = 'none';
  
  summaryTableBody.innerHTML = `
    <tr>
      <td colspan="8" class="text-center text-muted">กรุณานำเข้าไฟล์สแกนบัตร (Excel) เพื่อคำนวณและแสดงผลตารางสรุป</td>
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
  leaveCount: number;
  earlyOutCount: number;
  records: ProcessedDayRecord[];
}

interface ProcessedDayRecord {
  date: string;
  checkIn: string;
  checkOut: string;
  status: 'ปกติ' | 'สาย' | 'สายครึ่งวัน' | 'ลาครึ่งวัน' | 'ออกก่อนเวลา' | 'วันหยุด' | 'ลา/ขาดงาน';
  lateMinutes: number;
  earlyMinutes: number;
}

// Convert time to minutes (e.g. "08:30" -> 510)
function timeToMinutes(timeStr: string): number {
  if (!timeStr) return 0;
  const [h, m] = timeStr.split(':').map(Number);
  return h * 60 + m;
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

  const processedRecords: ProcessedDayRecord[] = [];
  let workedDays = 0;
  let lateCount = 0;
  let leaveCount = 0;
  let earlyOutCount = 0;

  const normalInMinutes = timeToMinutes(rules.morningWorkStart);
  const normalOutMinutes = timeToMinutes(rules.afternoonWorkEnd);

  daysInRange.forEach(dateStr => {
    const jsDate = new Date(dateStr);
    const dayOfWeek = jsDate.getDay(); // 0 = Sunday, 6 = Saturday
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    const isHoliday = holidays.some(h => h.date === dateStr) || isWeekend;

    const scans = staff.records[dateStr] || [];
    
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
      } else {
        leaveCount++;
        processedRecords.push({
          date: dateStr,
          checkIn: '-',
          checkOut: '-',
          status: 'ลา/ขาดงาน',
          lateMinutes: 0,
          earlyMinutes: 0
        });
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

    let lateMinutes = 0;
    let earlyMinutes = 0;
    let isLate = false;
    let isHalfDayLate = false;
    let isEarlyOut = false;

    // Calculate Late
    if (checkIn) {
      const inMin = timeToMinutes(checkIn);
      const diff = inMin - normalInMinutes;
      if (diff > rules.lateAllowanceMinutes) {
        lateMinutes = diff;
        isLate = true;
        if (diff >= rules.halfDayLateMinutes) {
          isHalfDayLate = true;
        }
      }
    } else {
      // Missing Check-in
      isLate = true;
    }

    // Calculate Early Checkout
    let isHalfDayLeave = false;
    if (checkOut) {
      const outMin = timeToMinutes(checkOut);
      
      // If check-out is around midday break (e.g. 12:00 - 13:00) and check-in was normal, it counts as "Half-Day Leave (afternoon)"
      // Noon break is 12:00 to 13:00 (720 to 780 minutes)
      if (outMin >= 720 && outMin <= 780) {
        isHalfDayLeave = true;
        earlyMinutes = 0; // It's scheduled half-day leave, not accidental early check-out
      } else {
        const diff = normalOutMinutes - outMin;
        if (diff > rules.earlyCheckoutAllowanceMinutes) {
          earlyMinutes = diff;
          isEarlyOut = true;
        }
      }
    } else {
      // Missing Check-out
      isEarlyOut = true;
    }

    // Determine final daily status
    let status: ProcessedDayRecord['status'] = 'ปกติ';
    if (isHalfDayLeave) {
      status = 'ลาครึ่งวัน';
    } else if (isHalfDayLate) {
      status = 'สายครึ่งวัน';
      lateCount++;
    } else if (isLate) {
      status = 'สาย';
      lateCount++;
    } else if (isEarlyOut) {
      status = 'ออกก่อนเวลา';
      earlyOutCount++;
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
    if (isHalfDayLeave) {
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
    leaveCount,
    earlyOutCount,
    records: processedRecords
  };
}

// Main table renderer
let currentProcessedSummaries: ProcessedStaffSummary[] = [];

function recalculateAndRender() {
  if (parsedEmployees.length === 0) return;

  const startVal = filterStartDate.value;
  const endVal = filterEndDate.value;
  const searchVal = searchStaff.value.trim().toLowerCase();

  // Run business calculations for all employees
  currentProcessedSummaries = parsedEmployees.map(emp => 
    calculateStaffRecords(emp, dbState.rules, dbState.holidays, startVal, endVal)
  );

  // Filter based on search input
  const filtered = currentProcessedSummaries.filter(summary => 
    summary.name.toLowerCase().includes(searchVal) || 
    summary.id.toLowerCase().includes(searchVal)
  );

  summaryTableBody.innerHTML = '';
  
  if (filtered.length === 0) {
    summaryTableBody.innerHTML = `
      <tr>
        <td colspan="8" class="text-center text-muted">ไม่พบข้อมูลตามคำค้นหาที่ระบุ</td>
      </tr>
    `;
    return;
  }

  filtered.forEach(s => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${s.id}</strong></td>
      <td>${s.name}</td>
      <td><span class="text-muted">${s.department}</span></td>
      <td><span class="badge badge-info">${s.workedDays} วัน</span></td>
      <td><span class="badge ${s.lateCount > 0 ? 'badge-danger' : 'badge-success'}">${s.lateCount} ครั้ง</span></td>
      <td><span class="badge ${s.leaveCount > 0 ? 'badge-warning' : 'badge-success'}">${s.leaveCount} วัน</span></td>
      <td><span class="badge ${s.earlyOutCount > 0 ? 'badge-warning' : 'badge-success'}">${s.earlyOutCount} ครั้ง</span></td>
      <td><button class="btn btn-secondary btn-sm btn-view-detail" data-id="${s.id}">🔍 รายละเอียด</button></td>
    `;
    summaryTableBody.appendChild(tr);
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

  staff.records.forEach(r => {
    let badgeClass = 'badge-success';
    if (r.status === 'สาย') badgeClass = 'badge-danger';
    else if (r.status === 'สายครึ่งวัน') badgeClass = 'badge-danger';
    else if (r.status === 'ออกก่อนเวลา' || r.status === 'ลา/ขาดงาน' || r.status === 'ลาครึ่งวัน') badgeClass = 'badge-warning';
    else if (r.status === 'วันหยุด') badgeClass = 'badge-info';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${r.date}</td>
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

// Bootstrap
document.addEventListener('DOMContentLoaded', init);
