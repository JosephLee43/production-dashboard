const test = require('node:test');
const assert = require('node:assert/strict');

const { normalizeCellValue, hasMeaningfulRow } = require('./excelData');

test('normalizeCellValue strips whitespace and resolves formula objects', () => {
  assert.equal(normalizeCellValue('  DP5-BASE UNIT '), 'DP5-BASE UNIT');
  assert.equal(normalizeCellValue({ result: 42 }), 42);
  assert.equal(normalizeCellValue({ richText: [{ text: 'AB' }, { text: 'C' }] }), 'ABC');
  assert.equal(normalizeCellValue({ text: 'QC WAIT' }), 'QC WAIT');
  assert.equal(normalizeCellValue(null), '');
});

test('hasMeaningfulRow ignores empty rows with whitespace-only values', () => {
  assert.equal(hasMeaningfulRow({ Item: ' ', 'Work Order': '', Workstation: '' }), false);
  assert.equal(hasMeaningfulRow({ Item: 'DP5-BASE UNIT', Workstation: 'IXCWS1' }), true);
});
