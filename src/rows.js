import { formatCell } from './cells';

export const formatRow = (row, index) => {
  // To ensure the row number starts as in excel.
  const rowIndex = index + 1;
  const rowCells = row
    .map((cell, cellIndex) => formatCell(cell, cellIndex, rowIndex))
    .join('');

  return `<row r="${rowIndex}">${rowCells}</row>`;
};

export const generatorRows = (rows) => (
  rows
    .map((row, index) => formatRow(row, index))
    .join('')
);
