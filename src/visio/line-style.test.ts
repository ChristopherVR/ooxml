import { describe, expect, it, vi } from 'vitest';
import { lineCap } from './line-style.js';
import { parseVsdx } from './parser.js';
import type { Cells } from './sheet.js';
import type { ThemeResources } from './theme-resolve.js';
import { generatedTheme, themeFixture } from './theme-fixtures.js';
import { cell, fixture, shape } from './test-fixtures.js';

const caps = (value: string): Cells => new Map([['LineCap', { value }]]);
const resources: ThemeResources = {};
async function themed(cap: string | undefined, contents = cell('LineCap', 'Themed')) {
	const theme = generatedTheme().replaceAll(
		'<a:ln>',
		`<a:ln${cap === undefined ? '' : ` cap="${cap}"`}>`,
	);
	const doc = await parseVsdx(await themeFixture({ theme, contents }));
	return { style: doc.pages[0]!.shapes[0]!.style, codes: doc.diagnostics.map((item) => item.code) };
}

describe('Visio normalized line caps', () => {
	it.each([
		['0', 'round'],
		['1', 'butt'],
		['2', 'square'],
	] as const)(
		'normalizes cached %s to %s without hiding compatibility inferences',
		(value, expected) => {
			const report = vi.fn();
			expect(lineCap(caps(value), resources, report)).toBe(expected);
			expect(report.mock.calls.map(([code]) => code)).toEqual(
				value === '0' ? [] : ['inferred-line-cap'],
			);
		},
	);
	it('leaves absent caps undefined instead of manufacturing a ShapeSheet default', () => {
		const report = vi.fn();
		expect(lineCap(new Map(), resources, report)).toBeUndefined();
		expect(report).not.toHaveBeenCalled();
	});
	it.each(['3', '-1', '0.5', '255'])('diagnoses out-of-range cached %s', (value) => {
		const report = vi.fn();
		expect(lineCap(caps(value), resources, report)).toBeUndefined();
		expect(report).toHaveBeenCalledWith('invalid-line-cap', expect.any(String));
	});
	it.each([
		'',
		'bad',
		'NaN',
		'Infinity',
		'1e999',
		'0'.repeat(129),
		' '.repeat(129) + '0',
		'0' + ' '.repeat(129),
	])('rejects unusable cached %s', (value) => {
		const report = vi.fn();
		expect(lineCap(caps(value), resources, report)).toBeUndefined();
		expect(report).toHaveBeenCalledWith('missing-cached-value', expect.any(String));
	});
	it('does not evaluate formula-only cells', () => {
		const report = vi.fn();
		expect(lineCap(new Map([['LineCap', { formula: '1+1' }]]), resources, report)).toBeUndefined();
		expect(report).toHaveBeenCalledWith('missing-cached-value', expect.any(String));
	});
	it.each([
		['rnd', 'round'],
		['flat', 'butt'],
		['sq', 'square'],
		[undefined, 'butt'],
	] as const)(
		'resolves selected themed cap %s to %s independently of width',
		async (cap, expected) => {
			const result = await themed(cap, cell('LineCap', 'Themed') + cell('LineWeight', 0.02));
			expect(result.style.lineCap).toBe(expected);
			expect(result.style.lineWidth).toBe(0.02);
			expect(result.codes).not.toContain('missing-cached-value');
			expect(result.codes).not.toContain('inferred-line-cap');
		},
	);
	it('does not require a cached or themed line width for themed caps', async () => {
		expect((await themed('sq')).style.lineCap).toBe('square');
		const result = await themed('sq', cell('LineCap', 'Themed') + cell('LineWeight', 'bad'));
		expect(result.style.lineCap).toBe('square');
		expect(result.codes).toContain('missing-cached-value');
	});
	it.each(['', 'round', 'SQ', 'injected'])(
		'diagnoses unsupported DrawingML cap token %j',
		async (cap) => {
			const result = await themed(cap);
			expect(result.style.lineCap).toBeUndefined();
			expect(result.codes).toContain('unsupported-theme-line-cap');
		},
	);
	it('leaves unresolved themed cap absent without selecting a guessed theme', async () => {
		const doc = await parseVsdx(
			await themeFixture({ omitTheme: true, contents: cell('LineCap', 'Themed') }),
		);
		expect(doc.pages[0]!.shapes[0]!.style.lineCap).toBeUndefined();
		expect(doc.diagnostics.map((item) => item.code)).toContain('unresolved-line-cap');
	});
	it('preserves a cached cap override without diagnosing unused theme cap geometry', async () => {
		const result = await themed(
			'unknown',
			cell('LineCap', 0, 'THEMEVAL()') + cell('LineWeight', 'Themed'),
		);
		expect(result.style.lineCap).toBe('round');
		expect(result.codes).not.toContain('unsupported-theme-line-cap');
	});
	it('does not import theme caps when only line width is themed', async () => {
		const result = await themed('unknown', cell('LineWeight', 'Themed'));
		expect(result.style.lineCap).toBeUndefined();
		expect(result.codes).not.toContain('unsupported-theme-line-cap');
	});
	it('honors style, master and local cached inheritance including F=Inh', async () => {
		const bytes = await fixture({
			document: `<StyleSheets><StyleSheet ID="0">${cell('LineCap', 0)}</StyleSheet><StyleSheet ID="4" LineStyle="0">${cell('LineCap', 1)}</StyleSheet></StyleSheets>`,
			masters: [{ id: '5', shapes: shape('10', cell('LineCap', 2), 'LineStyle="4"') }],
			pages: [
				{
					id: '1',
					contents: `<Shapes>${shape('1', '', 'LineStyle="4"')}${shape('2', '', 'Master="5"')}${shape('3', cell('LineCap', 0, 'Inh'), 'Master="5"')}${shape('4', '<Cell N="LineCap" F="Inh"/>', 'Master="5"')}</Shapes>`,
				},
			],
		});
		const doc = await parseVsdx(bytes);
		expect(doc.pages[0]!.shapes.map((item) => item.style.lineCap)).toEqual([
			'butt',
			'square',
			'round',
			'square',
		]);
	});
	it('selects connector theme caps using saved QuickStyleType', async () => {
		const theme = generatedTheme()
			.replaceAll('<a:ln>', '<a:ln cap="flat">')
			.replace(
				/(<v:fmtConnectorScheme>[\s\S]*?<a:lnStyleLst>)([\s\S]*?)(<\/a:lnStyleLst>)/,
				(_, start: string, content: string, end: string) =>
					start + content.replaceAll('cap="flat"', 'cap="rnd"') + end,
			);
		const doc = await parseVsdx(
			await themeFixture({
				theme,
				contents: cell('LineCap', 'Themed') + cell('QuickStyleType', 3),
			}),
		);
		expect(doc.pages[0]!.shapes[0]!.style.lineCap).toBe('round');
	});
});
