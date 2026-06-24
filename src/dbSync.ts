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

// Initial empty DB structure
export const initialDB = (): TimeStampDB => ({
  rules: { ...DEFAULT_RULES },
  holidays: [],
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
