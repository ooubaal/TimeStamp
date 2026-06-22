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

          // Check if this row contains employee metadata anywhere
          // We look for cells that contain "รหัสพนักงาน" or "ชื่อพนักงาน" or "ตำแหน่งงาน"
          let foundEmpId = '';
          let foundEmpName = '';
          let foundPos = '';
          let foundGroup = '';
          let foundDept = '';

          // Let's inspect each cell in the row
          for (let c = 0; c < row.length; c++) {
            const val = String(row[c] || '').trim();
            
            // Check if cell is "รหัสพนักงาน" and the next cells have the ID
            if (val.includes('รหัสพนักงาน') || val.includes('รหัสบัตร')) {
              // ID is usually in this cell or subsequent non-empty cells
              for (let k = c + 1; k < Math.min(row.length, c + 5); k++) {
                const nextVal = String(row[k] || '').trim();
                if (/^\d+$/.test(nextVal)) {
                  foundEmpId = nextVal;
                  break;
                }
              }
            }

            if (val.includes('ชื่อพนักงาน') || val.includes('ชื่อ-สกุล')) {
              for (let k = c + 1; k < Math.min(row.length, c + 5); k++) {
                const nextVal = String(row[k] || '').trim();
                if (nextVal && !nextVal.includes('ชื่อ') && !nextVal.includes('รหัส') && !nextVal.includes('ตำแหน่ง')) {
                  foundEmpName = nextVal;
                  break;
                }
              }
            }

            if (val.includes('ตำแหน่ง')) {
              for (let k = c + 1; k < Math.min(row.length, c + 5); k++) {
                const nextVal = String(row[k] || '').trim();
                if (nextVal && !nextVal.includes('ตำแหน่ง') && !nextVal.includes('กลุ่ม') && !nextVal.includes('หน่วยงาน')) {
                  foundPos = nextVal;
                  break;
                }
              }
            }

            if (val.includes('กลุ่มพนักงาน')) {
              for (let k = c + 1; k < Math.min(row.length, c + 5); k++) {
                const nextVal = String(row[k] || '').trim();
                if (nextVal && !nextVal.includes('กลุ่ม')) {
                  foundGroup = nextVal;
                  break;
                }
              }
            }

            if (val.includes('หน่วยงาน')) {
              for (let k = c + 1; k < Math.min(row.length, c + 5); k++) {
                const nextVal = String(row[k] || '').trim();
                if (nextVal && !nextVal.includes('หน่วยงาน')) {
                  foundDept = nextVal;
                  break;
                }
              }
            }
          }

          // If we found an ID and a Name anywhere in this row, treat as new employee!
          if (foundEmpId && foundEmpName) {
            currentEmployee = {
              id: foundEmpId,
              name: foundEmpName,
              position: foundPos,
              group: foundGroup,
              department: foundDept,
              records: {}
            };
            employeesMap[foundEmpId] = currentEmployee as EmployeeData;
            currentDate = '';
            continue;
          }

          // If no keywords found, check if this is the details row directly
          // For example, some files place ID in Col 0, Name in Col 2.
          const col0Str = String(row[0] || '').trim();
          const col2Str = String(row[2] || '').trim();
          if (/^\d{5,10}$/.test(col0Str) && col2Str && !col2Str.includes('ชื่อพนักงาน') && !col2Str.includes('เวลา') && !col2Str.includes('/')) {
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
