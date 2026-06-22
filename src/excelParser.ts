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

          const nameIndex = row.findIndex(cell => String(cell || '').includes('ชื่อพนักงาน'));
          const idIndex = row.findIndex(cell => String(cell || '').includes('รหัสพนักงาน'));

          if (idIndex !== -1 && nameIndex !== -1) {
            // This is the header row, e.g. "รหัสพนักงาน   รหัสบัตรพนักงาน   ชื่อพนักงาน"
            // The actual employee details are usually on the row IMMEDIATELY below it!
            const nextRow = rows[i + 1];
            if (nextRow) {
              const empId = String(nextRow[idIndex] || '').trim();
              const empName = String(nextRow[nameIndex] || nextRow[nameIndex + 1] || '').trim();
              
              if (empId && empName && !empId.includes('รหัส')) {
                currentEmployee = {
                  id: empId,
                  name: empName,
                  position: String(nextRow[idIndex + 3] || nextRow[idIndex + 4] || '').trim(),
                  group: String(nextRow[idIndex + 5] || '').trim(),
                  department: String(nextRow[idIndex + 6] || nextRow[idIndex + 7] || '').trim(),
                  records: {}
                };
                employeesMap[empId] = currentEmployee as EmployeeData;
                currentDate = '';
                i++; // Skip the next row since we processed it
                continue;
              }
            }
          }

          // Fallback parser if employee ID is in column 0 and name is in column 2 on the row directly
          const col0Str = String(row[0] || '').trim();
          const col2Str = String(row[2] || '').trim();
          if (/^\d{5,10}$/.test(col0Str) && col2Str && !col2Str.includes('ชื่อพนักงาน') && !col2Str.includes('เวลา')) {
            const empId = col0Str;
            // Ensure this is not a date (just in case)
            if (!empId.includes('/')) {
              currentEmployee = {
                id: empId,
                name: col2Str,
                position: String(row[4] || '').trim(),
                group: String(row[5] || '').trim(),
                department: String(row[6] || '').trim(),
                records: {}
              };
              employeesMap[empId] = currentEmployee as EmployeeData;
              currentDate = '';
              continue;
            }
          }

          // If we have an active employee, check for date and time values
          // The Excel file can have empty cells or shifted columns due to merging.
          // Let's inspect the entire row.
          if (currentEmployee) {
            // Find any cell matching "DD/MM/YYYY" format
            let foundDate = '';
            let foundTimes: string[] = [];

            for (let c = 0; c < row.length; c++) {
              const cellVal = String(row[c] || '').trim();
              if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(cellVal)) {
                foundDate = parseThaiDate(cellVal);
              } else if (/^\d{2}:\d{2}$/.test(cellVal)) {
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
              // Append times to active date
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
