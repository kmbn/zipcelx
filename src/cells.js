import escape from 'lodash.escape';

const CELL_TYPE_STRING = 'string';

const CELL_TYPE_NUMBER = 'number';

const validTypes = [CELL_TYPE_STRING, CELL_TYPE_NUMBER];

export const WARNING_INVALID_TYPE = 'Invalid type supplied in cell config, falling back to "string"';

const generateColumnLetter = (colIndex) => {
  if (typeof colIndex !== 'number') {
    return '';
  }

  const prefix = Math.floor(colIndex / 26);
  const letter = String.fromCharCode(97 + (colIndex % 26)).toUpperCase();
  if (prefix === 0) {
    return letter;
  }
  return generateColumnLetter(prefix - 1) + letter;
};

export const generatorCellNumber = (index, rowNumber) => (
  `${generateColumnLetter(index)}${rowNumber}`
);

export const generatorNumberCell = (index, value, rowIndex) => (`<c r="${generatorCellNumber(index, rowIndex)}"><v>${value}</v></c>`);

export const generatorStringCell = (index, value, rowIndex) => (`<c r="${generatorCellNumber(index, rowIndex)}" t="inlineStr"><is><t>${escape(value)}</t></is></c>`);

export const formatCell = (cell, index, rowIndex) => {
  if (validTypes.indexOf(cell.type) === -1) {
    console.warn(WARNING_INVALID_TYPE);
    cell.type = CELL_TYPE_STRING;
  }

  return (
    cell.type === CELL_TYPE_STRING
      ? generatorStringCell(index, cell.value, rowIndex)
      : generatorNumberCell(index, cell.value, rowIndex)
  );
};
