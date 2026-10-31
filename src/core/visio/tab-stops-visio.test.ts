import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { visioTabStarts } from './tab-stops';
import { cell, fixture, shape } from './test-fixtures';

/**
 * Positions Visio 16 gave the same text (`scripts/record-visio-tab-stops.ps1`: shapes whose Tabs
 * rows were injected, page exported to SVG, Calibri 12 pt, left margin 0, checked 2026-10-31).
 * The SVG gives each part as a `dx` in em from the end of the previous one; the widths below are
 * the Calibri advances in points, and positions are compared in inches (72 pt).
 */
const pt = (value: number) => value / 72;
const A = pt(6.948);
const B = pt(6.528);
const C = pt(6.4);
const num = { width: pt(27.36), decimal: pt(12.168) }; // "12.50"
const part = (width: number) => ({ width, decimal: width });
const near = (actual: number[], expected: number[]) => {
	expect(actual).toHaveLength(expected.length);
	actual.forEach((value, index) => expect(value).toBeCloseTo(pt(expected[index]!), 2));
};

describe('tab stops against Visio 16 text layout', () => {
	it('reads alignment codes 0 to 3 as left, centre, right and decimal', async () => {
		const cells = [0, 1, 2, 3]
			.map(
				(code) =>
					`<Cell N="Position${code}" V="${code + 1}"/><Cell N="Alignment${code}" V="${code}"/>`,
			)
			.join('');
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape(
						'1',
						cell('Width', 4) +
							cell('Height', 1) +
							`<Section N="Tabs"><Row IX="0">${cells}</Row></Section><Text><tp IX="0"/>a\tb</Text>`,
					)}</Shapes>`,
				},
			],
		});
		const tabs = (await parseVsdx(bytes)).pages[0]!.shapes[0]!.text.paragraphs![0]!.tabs;
		expect(tabs?.map((stop) => stop.align)).toEqual(['left', 'center', 'right', 'decimal']);
	});

	it('centres, right-aligns and decimal-aligns text on a 1 in stop', () => {
		const at = (align: 'center' | 'right' | 'decimal') =>
			visioTabStarts([part(A), num], [{ position: 1, align }])[1]!;
		expect(at('center')).toBeCloseTo(pt(58.32), 2);
		expect(at('right')).toBeCloseTo(pt(44.652), 2);
		expect(at('decimal')).toBeCloseTo(pt(59.844), 2);
	});

	it('right-aligns text without a decimal point on a decimal stop, then uses the default', () => {
		const stops = [
			{ position: 1, align: 'left' as const },
			{ position: 2.5, align: 'decimal' as const },
		];
		near(visioTabStarts([part(A), num, part(B), part(pt(7))], stops), [0, 72, 173.472, 216]);
	});

	it('uses 0.5 in steps with no stops and after the last stop', () => {
		near(visioTabStarts([part(A), part(B), part(C), part(C)]), [0, 36, 72, 108]);
		const last = [
			{ position: 1, align: 'left' as const },
			{ position: 2, align: 'left' as const },
		];
		near(visioTabStarts([part(A), part(B), part(C), part(C)], last), [0, 72, 144, 180]);
	});

	it('lets text that does not fit before a right stop follow the text before it', () => {
		const starts = visioTabStarts([part(A), num], [{ position: 0.1, align: 'right' }]);
		expect(starts[1]).toBeCloseTo(A, 3);
	});
});
