import { expect, it } from 'vitest';
import { parseXml } from '../xml';
import { independentTextFieldFormula } from './text-field-independent';
const cell = (name: string, extra = '') =>
	parseXml(
		`<Shape><Section N="Field"><Row IX="0" ${extra}><Cell N="${name}"/></Row></Section></Shape>`,
	).getElementsByTagName('Cell')[0]!;
it.each(['2', '37', '204'])('proves only measured independent Field picture %s', (code) =>
	expect(independentTextFieldFormula(cell('Format'), `FIELDPICTURE(${code})`)).toBe(true),
);
it.each([
	'FIELDPICTURE(999)',
	'FIELDPICTURE(21)',
	'FIELDPICTURE(Width)',
	'FIELDPICTURE(2)+Width',
	'INDIRECT("Width")',
])('refuses unmeasured/dynamic picture dependency: %s', (formula) =>
	expect(independentTextFieldFormula(cell('Format'), formula)).toBe(false),
);
it('keeps this proof confined to live canonical Field Value/Format cells', () => {
	for (const formula of ['PAGENAME()', 'DOCCREATION()', 'DATE(2026,1,1)'])
		expect(independentTextFieldFormula(cell('Value'), formula)).toBe(true);
	for (const formula of ['DATE(Width,1,1)', 'NOW()+Width', 'UNKNOWN()'])
		expect(independentTextFieldFormula(cell('Value'), formula)).toBe(false);
	for (const extra of ['Del="0"', 'N="named"', 'T="Rel"'])
		expect(independentTextFieldFormula(cell('Format', extra), 'FIELDPICTURE(2)')).toBe(false);
	expect(independentTextFieldFormula(cell('Other'), 'FIELDPICTURE(2)')).toBe(false);
	expect(
		independentTextFieldFormula(
			parseXml('<Shape><Cell N="Format"/></Shape>').getElementsByTagName('Cell')[0]!,
			'FIELDPICTURE(2)',
		),
	).toBe(false);
});
