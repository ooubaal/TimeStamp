// Database structure
export interface RuleSettings {
  morningWorkStart: string; // "08:30"
  afternoonWorkEnd: string; // "16:30"
  lateAllowanceMinutes: number; // e.g. 0 minutes
  halfDayLateMinutes: number; // e.g. 120 minutes (sian for half day)
  earlyCheckoutAllowanceMinutes: number; // e.g. 0 minutes
}

export interface Holiday {
  date: string; // "YYYY-MM-DD"
  name: string; // e.g. "วันปิยมหาราช"
}

export interface TimeStampDB {
  rules: RuleSettings;
  holidays: Holiday[];
  version: string;
  updatedAt: string;
}

export const DEFAULT_RULES: RuleSettings = {
  morningWorkStart: "08:30",
  afternoonWorkEnd: "16:30",
  lateAllowanceMinutes: 0,
  halfDayLateMinutes: 240,
  earlyCheckoutAllowanceMinutes: 0,
};

// Initial empty DB structure
export const initialDB = (): TimeStampDB => ({
  rules: { ...DEFAULT_RULES },
  holidays: [],
  version: "1.0.0",
  updatedAt: new Date().toISOString(),
});

// Load DB from localStorage if exists
export function loadFromLocalStorage(): TimeStampDB {
  const data = localStorage.getItem("TimeStampDB");
  if (data) {
    try {
      return JSON.parse(data);
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
        if (json && json.rules && Array.isArray(json.holidays)) {
          resolve(json as TimeStampDB);
        } else {
          reject(new Error("Invalid database format. Must contain rules and holidays arrays."));
        }
      } catch (err) {
        reject(new Error("Failed to parse JSON file."));
      }
    };
    reader.onerror = () => reject(new Error("Failed to read file."));
    reader.readAsText(file);
  });
}
