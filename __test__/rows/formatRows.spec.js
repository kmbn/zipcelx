import { formatRow } from '../../src/rows';
import baseConfig from '../baseConfig';

const expectedXML = '<row r="1"><c r="A1" t="inlineStr"><is><t>Test</t></is></c><c r="B1"><v>1000</v></c><c r="C1" s="0"><v>2024-11-01T13:30</v></c></row>';

describe('Format Row', () => {
  it('Should create one row from given data', () => {
    expect(formatRow(baseConfig.sheet.data[0], 0)).toBe(expectedXML);
  });
});
