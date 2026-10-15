import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index';
import { editVsdx } from './edit';
import { analyzeVisioFormula, evaluateVisioFormula, type VisioFormulaValue } from './formula';
import { fixture, cell, shape, section, rectangle } from './test-fixtures';

const resolve = (reference: { cell: string }): VisioFormulaValue => {
	if (reference.cell === 'Width') return { value: 3.25, unit: 'length' };
	if (reference.cell === 'Height') return { value: 2, unit: 'length' };
	throw new Error(`Unknown cell ${reference.cell}`);
};
const run = (source: string) => evaluateVisioFormula(source, resolve);
const scalar = (value: number): VisioFormulaValue => ({ value, unit: 'scalar' });

describe('ShapeSheet rounding, power, logarithm and inverse trigonometric functions', () => {
	it('rounds half away from zero to a number of digits and keeps the unit', () => {
		expect(run('ROUND(2.5,0)')).toEqual(scalar(3));
		expect(run('ROUND(-2.5)')).toEqual(scalar(-3));
		expect(run('ROUND(1.2345,2)').value).toBeCloseTo(1.23, 12);
		expect(run('ROUND(1250,-2)')).toEqual(scalar(1300));
		expect(run('ROUND(Width,1)')).toEqual({ value: 3.3, unit: 'length' });
		expect(run('ROUND(1.005,2)').value).toBeCloseTo(1.01, 12);
		expect(() => run('ROUND(1.5,0.5)')).toThrow(/digits/);
		expect(() => run('ROUND(1,2,3)')).toThrow(/argument count/);
	});

	it('takes the floor with INT and the integer part with TRUNC', () => {
		expect(run('INT(2.9)')).toEqual(scalar(2));
		expect(run('INT(-2.1)')).toEqual(scalar(-3));
		expect(run('TRUNC(-2.9)')).toEqual(scalar(-2));
		expect(run('INT(Width/Height*8)')).toEqual(scalar(13));
		expect(run('INT(Width)')).toEqual({ value: 3, unit: 'length' });
		expect(run('INT(2.9999999999)')).toEqual(scalar(3));
	});

	it('raises to a power with the dimension of the result', () => {
		expect(run('POW(2,10)')).toEqual(scalar(1024));
		expect(run('POW(Width,2)')).toEqual({ value: 3.25 ** 2, unit: 'area' });
		expect(run('POW(Width,0)')).toEqual(scalar(1));
		expect(() => run('POW(Width,Height)')).toThrow(/scalar/);
	});

	it('evaluates logarithms and exponentials of scalars only', () => {
		expect(run('EXP(0)')).toEqual(scalar(1));
		expect(run('LN(1)')).toEqual(scalar(0));
		expect(run('LOG10(1000)').value).toBeCloseTo(3, 12);
		expect(() => run('LN(0)')).toThrow(/positive/);
		expect(() => run('LOG10(-1)')).toThrow(/positive/);
		expect(() => run('EXP(Width)')).toThrow(/scalar/);
	});

	it('returns angles from ASIN, ACOS and ATAN', () => {
		expect(run('ASIN(1)')).toEqual({ value: Math.PI / 2, unit: 'angle' });
		expect(run('ACOS(1)')).toEqual({ value: 0, unit: 'angle' });
		expect(run('ATAN(Height/Width)').unit).toBe('angle');
		expect(run('ATAN(Height/Width)').value).toBeCloseTo(Math.atan(2 / 3.25), 12);
		expect(() => run('ASIN(2)')).toThrow(/-1 to 1/);
	});

	it('no longer reports them as unsupported, and still reports SUM', () => {
		expect(analyzeVisioFormula('ROUND(INT(Width),POW(2,1))').unsupportedFunctions).toEqual([]);
		expect(analyzeVisioFormula('SUM(Width,Height)').unsupportedFunctions).toEqual(['SUM']);
	});
});

describe('a resize recalculates cells that use them', () => {
	const dimensions =
		cell('PinX', 2) +
		cell('PinY', 3) +
		cell('Width', 3) +
		cell('Height', 4) +
		cell('LocPinX', 1.5, 'Width*0.5') +
		cell('LocPinY', 2, 'Height*0.5');
	const user = (name: string, value: number, formula: string) =>
		`<Row N="${name}">${cell('Value', value, formula)}</Row>`;
	const source = () =>
		fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape(
						'1',
						dimensions +
							rectangle +
							section(
								'User',
								user('Ratio', 0.75, 'ROUND(Width/Height,2)') +
									user('Cells', 3, 'INT(Width/Height*4)') +
									user('Square', 0.5625, 'POW(Width/Height,2)') +
									user('Level', 1, 'IF(ROUND(Width/Height,1)>1,2,1)') +
									user('Flag', 0.75, 'IF(FALSE,0,Width/Height)'),
							),
					)}</Shapes>`,
				},
			],
		});

	it('writes the new caches and leaves every other part alone', async () => {
		const original = await source();
		const result = await editVsdx(original, [
			{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 6.35, height: 2 },
		]);
		const zip = await JSZip.loadAsync(result.bytes);
		const root = parseXml(await zip.file('visio/pages/page1.xml')!.async('string'));
		const caches = new Map<string, number>();
		for (const node of Array.from(root.documentElement.getElementsByTagName('*')))
			if (node.hasAttribute('F'))
				caches.set(node.getAttribute('F')!, Number(node.getAttribute('V')));
		expect(caches.get('ROUND(Width/Height,2)')).toBeCloseTo(3.18, 12);
		expect(caches.get('INT(Width/Height*4)')).toBe(12);
		expect(caches.get('POW(Width/Height,2)')).toBeCloseTo((6.35 / 2) ** 2, 12);
		expect(caches.get('IF(ROUND(Width/Height,1)>1,2,1)')).toBe(2);
		// TRUE and FALSE are the constants 1 and 0, not cells.
		expect(caches.get('IF(FALSE,0,Width/Height)')).toBeCloseTo(3.175, 12);
		expect(result.changedParts).toEqual(['visio/pages/page1.xml']);
	});
});
