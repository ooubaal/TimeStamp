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

          // Check if this row is an Employee Header
          // Format from image:
          // Col 0: รหัสพนักงาน (contains numeric ID e.g. 3630189)
          // Col 2: ชื่อพนักงาน (contains name e.g. นางสาว ปุณยนุช แต่งผิว)
          // Col 4: ตำแหน่งงาน
          // Col 5: กลุ่มพนักงาน
          // Col 6: หน่วยงาน (or nearby column)
          
          const col0Str = String(row[0] || '').trim();
          const col2Str = String(row[2] || '').trim();
          
          // Identify header indicator or just a new employee entry
          // Usually, employee ID is numeric, and header row is "รหัสพนักงาน"
          if (col0Str && col0Str !== 'รหัสพนักงาน' && col0Str !== 'รหัสบัตรพนักงาน' && !col0Str.includes('วันที่') && !col0Str.includes('สภากาชาดไทย') && !col0Str.includes('เวลา')) {
            // Check if we hit an employee row
            // If col0Str is a number, it's likely an employee ID!
            const isEmployeeId = /^\d+$/.test(col0Str);
            if (isEmployeeId && col2Str) {
              const empId = col0Str;
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
          // Col 0 or Col 1 may contain Date like "03/10/2565"
          // Col 2 or Col 3 may contain Time like "07:26"
          if (currentEmployee) {
            const potentialDate = String(row[0] || '').trim();
            const potentialTime = String(row[2] || '').trim();

            // Date matches format DD/MM/YYYY (Buddhist era or standard)
            const isDate = /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(potentialDate);
            const isTime = /^\d{2}:\d{2}$/.test(potentialTime);

            if (isDate) {
              currentDate = parseThaiDate(potentialDate);
              if (currentDate) {
                if (!currentEmployee.records![currentDate]) {
                  currentEmployee.records![currentDate] = [];
                }
                if (isTime) {
                  currentEmployee.records![currentDate].push(potentialTime);
                }
              }
            } else if (currentDate && isTime) {
              // Time only row, append to the last active date
              currentEmployee.records![currentDate].push(potentialTime);
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
