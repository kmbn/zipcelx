import { generatorDatetimeCell } from '../../src/cells';

describe('Cell of type datetime', () => {
  const expectedXML = '<c r="A1" s="0"><v>2024-11-01T13:30</v></c>';
  it('Should create a new xml markup cell', () => {
    expect(generatorDatetimeCell(0, "2024-11-01T13:30", 1)).toBe(expectedXML);
  });
});
