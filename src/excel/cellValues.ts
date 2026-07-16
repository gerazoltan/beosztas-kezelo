import type ExcelJS from 'exceljs';

export function displayedCellText(cell: ExcelJS.Cell): string {
  try {
    return cell.text ?? '';
  } catch {
    const value = cell.value;
    if (value === null || value === undefined) return '';
    if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
      return String(value);
    }
    if ('result' in value) {
      const result = value.result;
      return typeof result === 'string' || typeof result === 'number' || typeof result === 'boolean'
        ? String(result)
        : '';
    }
    if ('richText' in value) return value.richText.map((part) => part.text).join('');
    if ('text' in value && typeof value.text === 'string') return value.text;
    return '';
  }
}
