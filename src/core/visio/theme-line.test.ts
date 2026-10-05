import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser.js';
import { generatedTheme, themeFixture } from './theme-fixtures.js';
import { cell } from './test-fixtures.js';
async function parse(
	options: { width?: string; line?: string; effects?: string; contents?: string } = {},
) {
	const theme = generatedTheme()
		.replaceAll('<a:ln>', `<a:ln w="${options.width ?? '12700'}" ${options.line ?? ''}>`)
		.replace(
			'</a:fmtScheme>',
			`<a:effectStyleLst><a:effectStyle><a:effectLst>${options.effects ?? ''}</a:effectLst></a:effectStyle></a:effectStyleLst></a:fmtScheme>`,
		);
	const document = await parseVsdx(
		await themeFixture({
			theme,
			contents: cell('LineWeight', 'Themed') + (options.contents ?? ''),
		}),
	);
	return { style: document.pages[0]!.shapes[0]!.style, diagnostics: document.diagnostics };
}
describe('Visio theme line widths and unresolved effects', () => {
	it('converts theme EMUs and preserves literal cached overrides', async () => {
		expect((await parse()).style.lineWidth).toBeCloseTo(1 / 72, 14);
		expect((await parse({ contents: cell('LineWeight', 0.02) })).style.lineWidth).toBe(0.02);
	});
	it('resolves connector theme widths using the saved Quick Style type', async () => {
		expect((await parse({ contents: cell('QuickStyleType', 3) })).style.lineWidth).toBeCloseTo(
			1 / 72,
			14,
		);
	});
	it.each(['bad', '-1', '20116801', '12345.67'])(
		'retains cached-value diagnostics for invalid line width %s',
		async (width) => {
			const result = await parse({ width });
			expect(result.style.lineWidth).toBe(0.01);
			expect(
				result.diagnostics.some(
					(item) => item.code === 'missing-cached-value' && item.message.includes('LineWeight'),
				),
			).toBe(true);
		},
	);
	it('resolves supported caps and reports unsupported compound strokes separately', async () => {
		const result = await parse({
			line: 'cap="rnd" cmpd="dbl"',
			contents: cell('LineCap', 'Themed'),
		});
		expect(result.style.lineCap).toBe('round');
		expect(result.diagnostics.map((item) => item.code)).not.toContain('unsupported-theme-line-cap');
		expect(result.diagnostics.map((item) => item.code)).toContain(
			'unsupported-theme-line-compound',
		);
	});
	it('reports actual selected shadow effects without fabricating empty effects', async () => {
		const contents = cell('QuickStyleEffectsMatrix', 1) + cell('ShdwPattern', 'Themed');
		const shadow = await parse({
			effects: '<a:outerShdw><a:srgbClr val="000000"/></a:outerShdw>',
			contents,
		});
		expect(shadow.diagnostics.map((item) => item.code)).toContain('unsupported-theme-effects');
		for (const effects of ['', '<a:blur rad="1"/>'])
			expect(
				(await parse({ effects, contents })).diagnostics.map((item) => item.code),
			).not.toContain('unsupported-theme-effects');
	});
});

describe('bounded selected solid theme lines', () => {
	it.each([
		'',
		'<a:prstDash/>',
		'<a:prstDash val="other"/>',
		'<a:prstDash val="solid"/><a:prstDash val="solid"/>',
		'<a:prstDash val="solid"/><a:custDash/>',
		'<v:prstDash val="solid"/>',
	])('keeps missing, malformed or ambiguous dash definitions unresolved (%s)', async (dash) => {
		const document = await parseVsdx(
			await themeFixture({
				theme: generatedTheme().replaceAll('<a:ln>', `<a:ln>${dash}`),
				contents: cell('LinePattern', 'Themed'),
			}),
		);
		expect(document.diagnostics.map((item) => item.code)).toContain('unresolved-line-pattern');
	});
});
