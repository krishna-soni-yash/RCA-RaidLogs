import * as XLSX from 'xlsx';

export interface ExportRowsOptions {
  rows: Record<string, any>[];
  sheetName: string;
  fileName: string;
  headers?: string[];
}

const sanitizeSheetName = (name: string): string => {
  const cleaned = (name || 'Sheet1').replace(/[\[\]\*\/\\\?\:]/g, '').trim();
  return cleaned.substring(0, 31) || 'Sheet1';
};

const buildFileName = (base: string): string => {
  const withoutExtension = base ? base.replace(/\.xlsx$/i, '') : 'export';
  const safeBase = withoutExtension.replace(/[^a-zA-Z0-9 _-]/g, '').trim().replace(/\s+/g, '_') || 'export';
  return `${safeBase}.xlsx`;
};

const createWorksheet = (rows: Record<string, any>[], headers?: string[]): XLSX.WorkSheet => {
  if (!rows || rows.length === 0) {
    const headerRow = headers && headers.length > 0 ? headers : [''];
    return XLSX.utils.aoa_to_sheet([headerRow]);
  }
  return XLSX.utils.json_to_sheet(rows, {
    header: headers && headers.length > 0 ? headers : undefined
  });
};

export const exportRowsToExcel = ({ rows, sheetName, fileName, headers }: ExportRowsOptions): void => {
  const workbook = XLSX.utils.book_new();
  const worksheet = createWorksheet(rows, headers);
  XLSX.utils.book_append_sheet(workbook, worksheet, sanitizeSheetName(sheetName));
  XLSX.writeFile(workbook, buildFileName(fileName));
};
