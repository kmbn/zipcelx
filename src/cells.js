const CELL_TYPE_STRING = 'string';

const CELL_TYPE_NUMBER = 'number';

const CELL_TYPE_DATETIME = 'datetime';

export const WARNING_INVALID_TYPE = 'Invalid type supplied in cell config, falling back to "string"';

function escape(str) {
  const escapeChars = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;',
  };
  return str.replace(/[&<>"']/g, (m) => escapeChars[m]);
}

function generateColumnLetter(colIndex) {
  if (typeof colIndex !== 'number') {
    return '';
  }

  const prefix = Math.floor(colIndex / 26);
  const letter = String.fromCharCode(97 + (colIndex % 26)).toUpperCase();
  if (prefix === 0) {
    return letter;
  }
  return generateColumnLetter(prefix - 1) + letter;
}

export function generatorCellNumber(index, rowNumber) {
  return `${generateColumnLetter(index)}${rowNumber}`;
}

export function generatorNumberCell(index, value, rowIndex) {
  return `<c r="${generatorCellNumber(index, rowIndex)}"><v>${value}</v></c>`;
}

export function generatorStringCell(index, value, rowIndex) {
  return `<c r="${generatorCellNumber(index, rowIndex)}" t="inlineStr"><is><t>${escape(value)}</t></is></c>`;
}

export function generatorDatetimeCell(index, value, rowIndex) {
  return `<c r="${generatorCellNumber(index, rowIndex)}" s="0"><v>${value}</v></c>`;
}

export function formatCell(cell, index, rowIndex) {
  let formatter;
  switch (cell.type) {
    case CELL_TYPE_STRING:
      formatter = generatorStringCell;
      break;
    case CELL_TYPE_NUMBER:
      formatter = generatorNumberCell;
      break;
    case CELL_TYPE_DATETIME:
      formatter = generatorDatetimeCell;
      break;
    default:
      console.warn(WARNING_INVALID_TYPE);
      formatter = generatorStringCell;
  }
  return formatter(index, cell.value, rowIndex);
}
