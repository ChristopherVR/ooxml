import { describe, expect, it, vi } from 'vitest';
import { linePattern } from './line-pattern';
import { parseVsdx } from './parser';
import type { Cells } from './sheet';
import { cell, fixture, shape } from './test-fixtures';
import { generatedTheme, themeFixture } from './theme-fixtures';

const cells = (value: string): Cells => new Map([['LinePattern', { value }]]);
const expected: [number, number[]][] = [
	[2, [7, 5]],
	[3, [0, 5]],
	[4, [7, 5, 0, 5]],
	[5, [7, 5, 0, 5, 0, 5]],
	[6, [7, 5, 7, 5, 0, 5]],
	[7, [19, 5, 7, 5]],
	[8, [19, 5, 7, 5, 7, 5]],
	[9, [3, 3]],
	[10, [0, 3]],
	[11, [3, 3, 0, 3]],
	[12, [3, 3, 0, 3, 0, 3]],
	[13, [3, 3, 3, 3, 0, 3]],
	[14, [9, 3, 3, 3]],
	[15, [9, 3, 3, 3, 3, 3]],
	[16, [15, 9]],
	[17, [0, 9]],
	[18, [15, 9, 0, 9]],
	[19, [15, 9, 0, 9, 0, 9]],
	[20, [15, 9, 15, 9, 0, 9]],
	[21, [39, 9, 15, 9]],
	[22, [39, 9, 15, 9, 15, 9]],
	[23, [1, 2]],
];

describe('Visio cached line-pattern normalization', () => {
	// Expected sequences are stated independently of the implementation's three-band formula.
	it.each(expected)('matches native built-in %i with round caps', (id, dash) => {
		const report = vi.fn();
		expect(linePattern(cells(String(id)), report)).toEqual({ linePattern: id, lineDash: dash });
		expect(report).not.toHaveBeenCalled();
	});
	it.each(['0', '1'])('keeps %s undashed without an inference warning', (value) => {
		const report = vi.fn();
		expect(linePattern(cells(value), report)).toEqual({ linePattern: Number(value) });
		expect(report).not.toHaveBeenCalled();
	});
	it('defaults absent cells to solid without a warning', () => {
		const report = vi.fn();
		expect(linePattern(new Map(), report)).toEqual({ linePattern: 1 });
		expect(report).not.toHaveBeenCalled();
	});
	it('preserves custom sentinel 254 and diagnoses the solid fallback without evaluating USE', () => {
		const report = vi.fn();
		const pattern: Cells = new Map([
			['LinePattern', { value: '254', formula: 'USE("CustomPattern")' }],
		]);
		expect(linePattern(pattern, report)).toEqual({ linePattern: 254 });
		expect(report.mock.calls.map(([code]) => code)).toEqual(['unsupported-custom-line-pattern']);
	});
	it.each(['-1', '1.5', '24', '253', '255', '1000000000'])(
		'normalizes invalid cached %s to a solid fallback',
		(value) => {
			const report = vi.fn();
			expect(linePattern(cells(value), report)).toEqual({ linePattern: 1 });
			expect(report.mock.calls.map(([code]) => code)).toEqual(['invalid-line-pattern']);
		},
	);
	it.each([
		'',
		'NaN',
		'Infinity',
		'1e999',
		'1000000001',
		'__proto__',
		'0'.repeat(129),
		' '.repeat(129) + '23',
		'23' + ' '.repeat(129),
	])('bounds invalid or oversized numeric cache %j', (value) => {
		const report = vi.fn();
		expect(linePattern(cells(value), report)).toEqual({ linePattern: 1 });
		expect(report.mock.calls.map(([code]) => code)).toEqual(['missing-cached-value']);
	});
	it('uses numeric saved results without evaluating a formula', () => {
		const pattern: Cells = new Map([['LinePattern', { value: ' 2.3e1 ', formula: 'THEMEVAL()' }]]);
		expect(linePattern(pattern, vi.fn())).toEqual({ linePattern: 23, lineDash: [1, 2] });
		const report = vi.fn();
		expect(linePattern(new Map([['LinePattern', { formula: '2+2' }]]), report)).toEqual({
			linePattern: 1,
		});
		expect(report).toHaveBeenCalledWith('missing-cached-value', expect.any(String));
	});
	it('does not share mutable sequence arrays between shapes', () => {
		const first = linePattern(cells('23'), vi.fn());
		(first.lineDash as number[])[0] = 100;
		expect(linePattern(cells('23'), vi.fn()).lineDash).toEqual([1, 2]);
	});
	it.each([0, 0.001, 0.02])(
		'keeps stroke-width multipliers independent of line width %s',
		async (width) => {
			const doc = await parseVsdx(
				await fixture({
					pages: [
						{
							id: '0',
							contents: `<Shapes>${shape('1', cell('LinePattern', 23) + cell('LineWeight', width))}</Shapes>`,
						},
					],
				}),
			);
			expect(doc.pages[0]!.shapes[0]!.style).toMatchObject({
				lineWidth: width,
				linePattern: 23,
				lineDash: [1, 2],
			});
		},
	);
	it('honors style, master and local caches including F=Inh', async () => {
		const doc = await parseVsdx(
			await fixture({
				document: `<StyleSheets><StyleSheet ID="0">${cell('LinePattern', 9)}</StyleSheet><StyleSheet ID="4" LineStyle="0">${cell('LinePattern', 23)}</StyleSheet></StyleSheets>`,
				masters: [{ id: '5', shapes: shape('10', cell('LinePattern', 6), 'LineStyle="4"') }],
				pages: [
					{
						id: '0',
						contents: `<Shapes>${shape('1', '', 'LineStyle="4"')}${shape('2', '', 'Master="5"')}${shape('3', cell('LinePattern', 0, 'Inh'), 'Master="5"')}${shape('4', '<Cell N="LinePattern" F="Inh"/>', 'Master="5"')}${shape('5', cell('LinePattern', 1), 'Master="5"')}</Shapes>`,
					},
				],
			}),
		);
		expect(doc.pages[0]!.shapes.map(({ style }) => [style.linePattern, style.lineDash])).toEqual([
			[23, [1, 2]],
			[6, [7, 5, 7, 5, 0, 5]],
			[0, undefined],
			[6, [7, 5, 7, 5, 0, 5]],
			[1, undefined],
		]);
	});
	it.each(['solid', 'sysDash'])(
		'resolves saved solid theme dashes but leaves %s spacing unresolved',
		async (dash) => {
			const theme = generatedTheme().replaceAll(
				'<a:ln>',
				`<a:ln w="12700"><a:prstDash val="${dash}"/>`,
			);
			for (const width of [cell('LineWeight', 0.02), cell('LineWeight', 'Themed')]) {
				const doc = await parseVsdx(
					await themeFixture({ theme, contents: width + cell('LinePattern', 'Themed') }),
				);
				expect(doc.pages[0]!.shapes[0]!.style.lineDash).toBeUndefined();
				expect(doc.pages[0]!.shapes[0]!.style.linePattern).toBe(1);
				expect(doc.diagnostics.some(({ code }) => code === 'unresolved-line-pattern')).toBe(
					dash !== 'solid',
				);
			}
		},
	);
	it.each([undefined, 0, 1, 23])(
		'does not diagnose unused theme dash when cached pattern is %s',
		async (pattern) => {
			const theme = generatedTheme().replaceAll(
				'<a:ln>',
				'<a:ln w="12700"><a:prstDash val="sysDash"/>',
			);
			const doc = await parseVsdx(
				await themeFixture({
					theme,
					contents:
						cell('LineWeight', 'Themed') +
						(pattern === undefined ? '' : cell('LinePattern', pattern, 'THEMEVAL()')),
				}),
			);
			const codes = doc.diagnostics.map(({ code }) => code);
			expect(codes).not.toContain('unsupported-theme-line-dash');
			expect(codes).not.toContain('unresolved-line-pattern');
			expect(doc.pages[0]!.shapes[0]!.style.linePattern).toBe(pattern ?? 1);
		},
	);
});
