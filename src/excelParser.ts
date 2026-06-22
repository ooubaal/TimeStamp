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

          // Convert all cells in row to trimmed strings, keeping array length intact
          const cells: string[] = [];
          for (let c = 0; c < row.length; c++) {
            cells.push(String(row[c] || '').trim());
          }

          // 1. Detect Employee metadata row
          // Col 0: Employee ID (e.g. "3630189")
          // Col 7: Employee Name (e.g. "นางสาว ปุณยนุช แต่งผิว")
          // Col 14: Job Position (e.g. "นักวิทยาศาสตร์การแพทย์")
          // Col 16: Employee Group (e.g. "จ้างชั่วคราว")
          // Col 17: Department (e.g. "ศูนย์บริการโลหิตฯ ...")
          const isNumericId = /^\d{5,10}$/.test(cells[0] || '');
          const hasName = cells[7] || '';
          const hasNoSlash = cells[0] ? !cells[0].includes('/') : true;

          if (isNumericId && hasName && hasNoSlash && !cells.includes('รหัสพนักงาน') && !cells.includes('เวลา')) {
            const empId = cells[0];
            const empName = cells[7];
            currentEmployee = {
              id: empId,
              name: empName,
              position: cells[14] || '',
              group: cells[16] || '',
              department: cells[17] || '',
              records: {}
            };
            employeesMap[empId] = currentEmployee as EmployeeData;
            currentDate = '';
            continue;
          }

          // 2. Parse Date and Scan Times
          // Date is in Col 1 (e.g. '03/10/2565')
          // Times are HH:MM in any column (usually Col 6)
          if (currentEmployee) {
            let foundDate = '';
            let foundTimes: string[] = [];

            const dateCandidate = cells[1] || '';

            if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(dateCandidate)) {
              foundDate = parseThaiDate(dateCandidate);
            }

            // Inspect all columns for HH:MM format
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
              // Time only row (like checkout on row below checkin), append to active date
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
