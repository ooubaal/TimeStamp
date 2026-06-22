import * as XLSX from 'xlsx';

export interface ScanRecord {
  employeeId: string;
  employeeName: string;
  position: string;
  group: string;
  department: string;
  date: string; // "YYYY-MM-DD"
  times: string[]; // ["07:26", "16:32"]
}

export interface EmployeeData {
  id: string;
  name: string;
  position: string;
  group: string;
  department: string;
  records: { [date: string]: string[] }; // date -> times
}

/**
 * Parses Thai Buddhist Era date string "DD/MM/YYYY" to ISO "YYYY-MM-DD"
 * e.g., "03/10/2565" -> "2022-10-03"
 */
export function parseThaiDate(thaiDateStr: string): string {
  const parts = thaiDateStr.trim().split('/');
  if (parts.length === 3) {
    const day = parts[0].padStart(2, '0');
    const month = parts[1].padStart(2, '0');
    let year = parseInt(parts[2], 10);
    // Convert BE to AD if year is > 2400
    if (year > 2400) {
      year = year - 543;
    }
    return `${year}-${month}-${day}`;
  }
  return '';
}

/**
 * Parsed Excel data in columns and rows
 */
export function parseExcelFile(file: File): Promise<EmployeeData[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        
        // Convert sheet to 2D array
        const rows: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        
        const employeesMap: { [id: string]: EmployeeData } = {};
        
        let currentEmployee: Partial<EmployeeData> | null = null;
        let currentDate: string = '';

        for (let i = 0; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;

          // Trim cell values to clean spaces
          const cells = row.map(cell => String(cell || '').trim());

          // 1. Detect Employee metadata row
          // Based on diagnosis, Employee ID is in Col 0 (e.g. "3630189"), Card ID in Col 3, Name in Col 6 (or next index), 
          // Position in Col 12, Group in Col 14, Dept in Col 15
          // We can check if cells[0] is a valid employee ID (numeric e.g. "3630189") and cells[4]/cells[5]/cells[6] contains a name (e.g. "นางสาว ปุณยนุช แต่งผิว")
          const isNumericId = /^\d{5,10}$/.test(cells[0] || '');
          const hasName = cells[4] || cells[5] || cells[6] || '';
          const hasNoSlash = !cells[0].includes('/'); // Ensure it's not a date in column 0

          if (isNumericId && hasName && hasNoSlash && !cells.includes('รหัสพนักงาน') && !cells.includes('เวลา')) {
            const empId = cells[0];
            const empName = String(cells[4] || cells[5] || cells[6] || '').trim();
            currentEmployee = {
              id: empId,
              name: empName,
              position: String(cells[12] || cells[10] || cells[11] || '').trim(),
              group: String(cells[14] || cells[13] || '').trim(),
              department: String(cells[15] || cells[16] || '').trim(),
              records: {}
            };
            employeesMap[empId] = currentEmployee as EmployeeData;
            currentDate = '';
            continue;
          }

          // 2. Parse Date and Scan Times
          // Based on diagnosis:
          // Row with date: Col 1 has date (e.g., '03/10/2565'), Col 5 or Col 6 has time (e.g., '07:26')
          // Row without date: Col 6 has time (e.g., '16:32')
          if (currentEmployee) {
            let foundDate = '';
            let foundTimes: string[] = [];

            // Look for Date in Col 1 or Col 2 (dd/mm/yyyy format)
            const dateCandidate1 = cells[1] || '';
            const dateCandidate2 = cells[2] || '';
            const dateCandidate0 = cells[0] || '';

            if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(dateCandidate1)) {
              foundDate = parseThaiDate(dateCandidate1);
            } else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(dateCandidate2)) {
              foundDate = parseThaiDate(dateCandidate2);
            } else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(dateCandidate0)) {
              foundDate = parseThaiDate(dateCandidate0);
            }

            // Look for Scan Times (hh:mm format) in all columns of this row
            for (let c = 0; c < cells.length; c++) {
              const cellVal = cells[c];
              if (/^\d{2}:\d{2}$/.test(cellVal)) {
                foundTimes.push(cellVal);
              }
            }

            if (foundDate) {
              currentDate = foundDate;
              if (!currentEmployee.records![currentDate]) {
                currentEmployee.records![currentDate] = [];
              }
              if (foundTimes.length > 0) {
                currentEmployee.records![currentDate].push(...foundTimes);
              }
            } else if (currentDate && foundTimes.length > 0) {
              // Time only row, append to the active date
              currentEmployee.records![currentDate].push(...foundTimes);
            }
          }
        }
        
        resolve(Object.values(employeesMap));
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("Failed to read Excel file."));
    reader.readAsArrayBuffer(file);
  });
}
