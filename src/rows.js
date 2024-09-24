import { formatCell } from './cells';

export function formatRow(row, index) {
  // To ensure the row number starts as in excel.
  const rowIndex = index + 1;
  const rowCells = row
    .map((cell, cellIndex) => formatCell(cell, cellIndex, rowIndex))
    .join('');

  return `<row r="${rowIndex}">${rowCells}</row>`;
}

export function generatorRows(rows) {
  return rows
    .map((row, index) => formatRow(row, index))
    .join('');
}
