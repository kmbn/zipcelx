import { WARNING_INVALID_TYPE } from '../../src/cells';
import { formatCell } from '../../src/cells';
import baseConfig from '../baseConfig';

const cells = baseConfig.sheet.data[0];

console.warn = jest.fn();

describe('Format Cell', () => {
  describe('Create a cell of type sting', () => {
    const expectedXML = '<c r="A1" t="inlineStr"><is><t>Test</t></is></c>';

    it('Should fallback to string if invalid type was supplied', () => {
      const cell = Object.assign({}, cells[0], { type: 'date' });
      formatCell(cell);
      expect(console.warn).toBeCalledWith(WARNING_INVALID_TYPE);
      expect(formatCell(cell, 0, 1)).toBe(expectedXML);
    });

    it('Create cell', () => {
      expect(formatCell(cells[0], 0, 1)).toBe(expectedXML);
    });
  });

  describe('Create a cell of type number', () => {
    const expectedXML = '<c r="B1"><v>1000</v></c>';
    it('Create cell', () => {
      expect(formatCell(cells[1], 1, 1)).toBe(expectedXML);
    });
  });

  describe('Create a cell of type datetime', () => {
    const expectedXML = '<c r="C1" s="0"><v>2024-11-01T13:30</v></c>';
    it('Create cell', () => {
      expect(formatCell(cells[2], 2, 1)).toBe(expectedXML);
    });
  });  
});
