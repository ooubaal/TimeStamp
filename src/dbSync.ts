// Database structure
export interface RuleSettings {
  morningWorkStart: string; // "08:30"
  afternoonWorkEnd: string; // "16:30"
  lateAllowanceMinutes: number; // e.g. 0 minutes
  halfDayLateMinutes: number; // e.g. 120 minutes (sian for half day)
  earlyCheckoutAllowanceMinutes: number; // e.g. 0 minutes
  morningLeaveStart: string; // "10:30"
  morningLeaveEnd: string; // "13:00"
  afternoonLeaveStart: string; // "12:00"
  afternoonLeaveEnd: string; // "13:00"
  otWeekdayStart: string; // "19:30"
  otHolidayHours: number; // e.g. 8
}

export interface Holiday {
  date: string; // "YYYY-MM-DD"
  name: string; // e.g. "วันปิยมหาราช"
}

export interface EmployeeData {
  id: string;
  name: string;
  position: string;
  group: string;
  department: string;
  records: { [date: string]: string[] }; // date -> times
}

export interface TimeStampDB {
  rules: RuleSettings;
  holidays: Holiday[];
  employees: EmployeeData[];
  importedFiles: string[];
  fileDates?: { [filename: string]: string[] };
  version: string;
  updatedAt: string;
}

export const DEFAULT_RULES: RuleSettings = {
  morningWorkStart: "08:30",
  afternoonWorkEnd: "16:30",
  lateAllowanceMinutes: 0,
  halfDayLateMinutes: 240,
  earlyCheckoutAllowanceMinutes: 0,
  morningLeaveStart: "10:30",
  morningLeaveEnd: "13:00",
  afternoonLeaveStart: "12:00",
  afternoonLeaveEnd: "13:00",
  otWeekdayStart: "19:30",
  otHolidayHours: 8,
};

export const DEFAULT_HOLIDAYS_2026: Holiday[] = [
  { date: "2026-01-01", name: "วันขึ้นปีใหม่" },
  { date: "2026-03-03", name: "วันมาฆบูชา" },
  { date: "2026-04-06", name: "วันจักรี" },
  { date: "2026-04-13", name: "วันสงกรานต์" },
  { date: "2026-04-14", name: "วันสงกรานต์" },
  { date: "2026-04-15", name: "วันสงกรานต์" },
  { date: "2026-05-01", name: "วันแรงงานแห่งชาติ" },
  { date: "2026-05-04", name: "วันฉัตรมงคล" },
  { date: "2026-05-11", name: "วันพืชมงคล" },
  { date: "2026-05-31", name: "วันวิสาขบูชา" },
  { date: "2026-06-03", name: "วันเฉลิมพระชนมพรรษาสมเด็จพระนางเจ้าฯ พระบรมราชินี" },
  { date: "2026-07-28", name: "วันเฉลิมพระชนมพรรษาพระบาทสมเด็จพระเจ้าอยู่หัว" },
  { date: "2026-07-29", name: "วันอาสาฬหบูชา" },
  { date: "2026-07-30", name: "วันเข้าพรรษา" },
  { date: "2026-08-12", name: "วันแม่แห่งชาติ" },
  { date: "2026-10-13", name: "วันนวมินทรมหาราช" },
  { date: "2026-10-23", name: "วันปิยมหาราช" },
  { date: "2026-12-05", name: "วันคล้ายวันพระบรมราชสมภพ ร.9 / วันพ่อแห่งชาติ" },
  { date: "2026-12-10", name: "วันรัฐธรรมนูญ" },
  { date: "2026-12-31", name: "วันสิ้นปี" }
];

// Initial empty DB structure
export const initialDB = (): TimeStampDB => ({
  rules: { ...DEFAULT_RULES },
  holidays: [...DEFAULT_HOLIDAYS_2026],
  employees: [],
  importedFiles: [],
  fileDates: {},
  version: "1.0.0",
  updatedAt: new Date().toISOString(),
});

// Load DB from localStorage if exists
export function loadFromLocalStorage(): TimeStampDB {
  const data = localStorage.getItem("TimeStampDB");
  if (data) {
    try {
      return sanitizeDB(JSON.parse(data));
    } catch (e) {
      console.error("Failed to parse TimeStampDB from localStorage", e);
    }
  }
  return initialDB();
}

// Save DB to localStorage
export function saveToLocalStorage(db: TimeStampDB) {
  db.updatedAt = new Date().toISOString();
  localStorage.setItem("TimeStampDB", JSON.stringify(db));
}

// Export database as JSON file for OneDrive storage
export function exportToJSONFile(db: TimeStampDB) {
  db.updatedAt = new Date().toISOString();
  const dataStr = JSON.stringify(db, null, 2);
  const blob = new Blob([dataStr], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  
  const link = document.createElement("a");
  link.href = url;
  link.download = "TimeStampDB.json";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

// Import database from JSON file
export function importFromJSONFile(file: File): Promise<TimeStampDB> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const json = JSON.parse(e.target?.result as string);
        resolve(sanitizeDB(json));
      } catch (err) {
        reject(new Error("Failed to parse JSON file."));
      }
    };
    reader.onerror = () => reject(new Error("Failed to read file."));
    reader.readAsText(file);
  });
}

// Sanitize and guarantee a complete structure even if empty {} was fetched
export function sanitizeDB(json: any): TimeStampDB {
  const base = initialDB();
  if (!json || typeof json !== "object") return base;
  
  const rules = {
    ...base.rules,
    ...(json.rules || {})
  };
  
  const holidays = Array.isArray(json.holidays) ? json.holidays : base.holidays;
  const employees = Array.isArray(json.employees) ? json.employees : base.employees;
  const importedFiles = Array.isArray(json.importedFiles) ? json.importedFiles : base.importedFiles;
  const fileDates = json.fileDates && typeof json.fileDates === 'object' ? json.fileDates : base.fileDates;
  
  return {
    rules,
    holidays,
    employees,
    importedFiles,
    fileDates,
    version: json.version || base.version,
    updatedAt: json.updatedAt || base.updatedAt
  };
}
