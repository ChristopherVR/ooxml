import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { parseVsdx } from './parser.js';
import { parseXml } from '../xml/index.js';
import { drawingColor } from './theme-color.js';
import {
	drawingNamespace,
	generatedTheme,
	themeFixture,
	themeRelationship,
} from './theme-fixtures.js';
import { cell } from './test-fixtures.js';

async function parse(options: Parameters<typeof themeFixture>[0] = {}) {
	const document = await parseVsdx(await themeFixture(options));
	return { shape: document.pages[0]!.shapes[0]!, diagnostics: document.diagnostics };
}
describe('Visio theme colors', () => {
	it('resolves inherited theme selectors and matrix variants through the page', async () => {
		const result = await parse({
			pageCells: cell('VariationColorIndex', 3) + cell('VariationStyleIndex', 3),
		});
		expect(result.shape.style.fill).toBe('#789abc');
		expect(result.shape.style.lineColor).toBe('#789abc');
		expect(result.shape.text.color).toBe('#789abc');
		expect(result.diagnostics.some((item) => item.code === 'unsupported-color')).toBe(false);
	});
	it('uses explicit solid matrix colors before the placeholder color', async () => {
		const result = await parse({
			contents: cell('QuickStyleFillColor', 106) + cell('QuickStyleFillMatrix', 1),
		});
		expect(result.shape.style.fill).toBe('#ffffff');
	});
	it.each([
		[0, '#102030'],
		[1, '#f0e0d0'],
		[2, '#214365'],
		[7, '#cbed0f'],
		[8, '#ffeedd'],
		[100, '#123456'],
		[106, '#a9cbed'],
	])('resolves Quick Style color %i', async (index, expected) => {
		const result = await parse({ contents: cell('QuickStyleFillColor', index) });
		expect(result.shape.style.fill).toBe(expected);
	});
	it('reports the observed 200-series variant-color extension', async () => {
		const result = await parse({ contents: cell('QuickStyleFillColor', 202) });
		expect(result.shape.style.fill).toBe('#214365');
		expect(result.diagnostics.some((item) => item.code === 'theme-color-extension')).toBe(true);
	});
	it('keeps saved literal colors authoritative without evaluating formulas', async () => {
		const result = await parse({
			contents: cell('FillForegnd', '#fedcba', 'THEMEVAL()+RUNADDON(&quot;bad&quot;)'),
		});
		expect(result.shape.style.fill).toBe('#fedcba');
	});
	it('uses a diagnostic and the base color for gradients', async () => {
		const result = await parse({
			theme: generatedTheme({
				fill: '<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:gs></a:gsLst></a:gradFill>',
			}),
		});
		expect(result.shape.style.fill).toBe('#123456');
		expect(result.diagnostics.some((item) => item.code === 'unsupported-theme-gradient')).toBe(
			true,
		);
	});
	it('supports bounded tint/shade matrix colors', async () => {
		expect(
			(
				await parse({
					theme: generatedTheme({
						fill: '<a:solidFill><a:srgbClr val="FFFFFF"><a:shade val="50000"/></a:srgbClr></a:solidFill>',
					}),
				})
			).shape.style.fill,
		).toBe('#bcbcbc');
	});
	it.each([
		{ omitTheme: true },
		{ theme: generatedTheme({ fill: '<a:noFill/>' }) },
		{ external: true },
		{ contents: cell('QuickStyleFillColor', 999) },
		{ contents: cell('ColorSchemeIndex', 65535) },
		{ contents: cell('EffectSchemeIndex', 0) },
		{ contents: cell('VariationColorIndex', 5) },
		{ contents: cell('QuickStyleFillMatrix', 99) },
		{ contents: cell('QuickStyleFillMatrix', 'Themed', '1+3') },
		{ contents: cell('QuickStyleVariation', 8) },
		{ contents: cell('QuickStyleVariation', 'Themed') },
		{ contents: cell('QuickStyleType', 'Themed') },
		{ contents: cell('QuickStyleType', 4) },
		{ contents: cell('ColorSchemeIndex', 'oops') },
		{ pageCells: cell('VariationColorIndex', 'oops') },
		{ pageCells: cell('ColorSchemeIndex', 65534) },
		{
			theme: generatedTheme({
				fill: '<a:solidFill><a:schemeClr val="phClr"><a:alpha val="50000"/></a:schemeClr></a:solidFill>',
			}),
		},
	])('retains unresolved diagnostics for unsupported theme cases (%j)', async (options) => {
		const result = await parse(options);
		expect(result.shape.style.fill).toBe('#ffffff');
		expect(
			result.diagnostics.some(
				(item) => item.code === 'unsupported-color' && item.message.includes('FillForegnd'),
			),
		).toBe(true);
	});
	it('does not let unrelated visibility flags block fill resolution', async () => {
		expect((await parse({ contents: cell('QuickStyleVariation', 2) })).shape.style.fill).toBe(
			'#123456',
		);
	});
	it('uses explicit local variant choices before page values', async () => {
		const result = await parse({
			pageCells: cell('VariationColorIndex', 3),
			contents: cell('VariationColorIndex', 1),
		});
		expect(result.shape.style.fill).toBe('#345678');
	});
	it('selects the effect component independently from the color component', async () => {
		const result = await parse({
			theme: generatedTheme({ effectId: 73 }),
			pageCells: cell('EffectSchemeIndex', 73),
		});
		expect(result.shape.style.fill).toBe('#123456');
	});
	it('fails closed when a component ID is ambiguous between theme parts', async () => {
		const zip = await JSZip.loadAsync(await themeFixture());
		const rels = await zip.file('visio/_rels/document.xml.rels')!.async('string');
		zip.file(
			'visio/_rels/document.xml.rels',
			rels.replace(
				'</Relationships>',
				`<Relationship Id="theme2" Type="${themeRelationship}" Target="theme/theme2.xml"/></Relationships>`,
			),
		);
		zip.file('visio/theme/theme2.xml', generatedTheme());
		const document = await parseVsdx(await zip.generateAsync({ type: 'uint8array' }));
		expect(document.pages[0]!.shapes[0]!.style.fill).toBe('#ffffff');
		expect(document.diagnostics.some((item) => item.code === 'unsupported-color')).toBe(true);
	});
	it('charges theme XML against the package node budget', async () => {
		await expect(
			parseVsdx(await themeFixture(), { limits: { maxXmlNodes: 100 } }),
		).rejects.toThrow();
		await expect(
			parseVsdx(await themeFixture({ omitTheme: true }), { limits: { maxXmlNodes: 100 } }),
		).resolves.toHaveProperty('format', 'vsdx');
	});
	it('ignores namespace-spoofed theme parts', async () => {
		const result = await parse({
			theme: generatedTheme().replace(drawingNamespace, 'urn:untrusted'),
		});
		expect(result.shape.style.fill).toBe('#ffffff');
		expect(result.diagnostics.some((item) => item.code === 'unsupported-theme')).toBe(true);
	});
});
describe('bounded DrawingML color choices', () => {
	const color = (contents: string) =>
		parseXml(`<a:srgbClr xmlns:a="${drawingNamespace}" val="000000">${contents}</a:srgbClr>`)
			.documentElement;
	it('applies tint and shade in order', () => {
		expect(drawingColor(color('<a:tint val="50000"/><a:shade val="50000"/>'), new Map())).toBe(
			'#898989',
		);
	});
	it('rejects unknown, invalid and excessive transforms', () => {
		for (const contents of [
			'<a:shade val="50000oops"/>',
			'<a:tint val="100001"/>',
			'<a:invert/>',
			'<a:tint val="100000"/>'.repeat(33),
		])
			expect(drawingColor(color(contents), new Map())).toBeUndefined();
	});
	it('stops scheme cycles and accepts saved system-color fallbacks', () => {
		const cycle = parseXml(
			`<a:schemeClr xmlns:a="${drawingNamespace}" val="accent1"/>`,
		).documentElement;
		expect(drawingColor(cycle, new Map([['accent1', cycle]]))).toBeUndefined();
		const system = parseXml(
			`<a:sysClr xmlns:a="${drawingNamespace}" val="windowText" lastClr="123456"/>`,
		).documentElement;
		expect(drawingColor(system, new Map())).toBe('#123456');
	});
});
