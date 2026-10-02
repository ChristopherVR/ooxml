import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser.js';
import { cell, fixture, rectangle, shape, xml, relations, relation } from './test-fixtures.js';
import { generatedTheme, themeRelationship } from './theme-fixtures.js';

const names = ['LineWeight', 'LineCap', 'LinePattern', 'LineColor', 'FillForegnd'];
const character = (contents: string) =>
	`<Section N="Character"><Row IX="0">${contents}</Row></Section>`;
async function parse(
	rootCache: string | undefined,
	inherited = false,
	rootName = 'No Style',
	colorSelector = true,
) {
	const root = rootCache === undefined ? '' : names.map((name) => cell(name, rootCache)).join('');
	const requested =
		names.map((name) => cell(name, 'Themed')).join('') + character(cell('Color', 'Themed'));
	const bytes = await fixture({
		document: `<StyleSheets><StyleSheet ID="0" NameU="${rootName}">${root}${character(rootCache === undefined ? '' : cell('Color', rootCache))}</StyleSheet><StyleSheet ID="1">${requested}</StyleSheet></StyleSheets>`,
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', rectangle + (colorSelector ? cell('ColorSchemeIndex', inherited ? 65534 : 0) : '') + cell('EffectSchemeIndex', 42) + cell('QuickStyleLineMatrix', 1) + cell('QuickStyleFillMatrix', 1) + cell('QuickStyleFontMatrix', 1) + '<Text>Test</Text>', 'LineStyle="1" FillStyle="1" TextStyle="1"')}</Shapes>`,
			},
		],
		edit: (zip) => {
			zip.file(
				'visio/_rels/document.xml.rels',
				relations(
					relation('rId1', 'pages', 'pages/pages.xml') +
						`<Relationship Id="theme" Type="${themeRelationship}" Target="theme/theme1.xml"/>`,
				),
			);
			zip.file(
				'visio/theme/theme1.xml',
				generatedTheme().replaceAll('<a:ln>', '<a:ln w="91440" cap="sq"><a:prstDash val="solid"/>'),
			);
			zip.file(
				'visio/pages/pages.xml',
				xml(
					'Pages',
					`<Page ID="0"><PageSheet>${cell('PageWidth', 8)}${cell('PageHeight', 10)}${inherited ? cell('ColorSchemeIndex', 0) : ''}</PageSheet><Rel r:id="rId1"/></Page>`,
				),
			);
		},
	});
	const doc = await parseVsdx(bytes);
	return { shape: doc.pages[0]!.shapes[0]!, diagnostics: doc.diagnostics };
}

describe('root selection cannot fall through to a populated dynamic theme', () => {
	it.each([
		[undefined, false],
		['Themed', false],
		[undefined, true],
		['Themed', true],
	] as const)(
		'keeps unusable root cache %s diagnosed (page inheritance %s)',
		async (cache, inherited) => {
			const { shape, diagnostics } = await parse(cache, inherited);
			expect(shape.style).toMatchObject({
				lineWidth: 0.01,
				linePattern: 1,
				lineColor: '#000000',
				fill: '#ffffff',
			});
			expect(shape.style.lineCap).toBeUndefined();
			expect(shape.text.color).toBe('#000000');
			expect(diagnostics.map((d) => d.code)).toContain('unresolved-line-cap');
			expect(diagnostics.map((d) => d.code)).toContain('unresolved-line-pattern');
			expect(diagnostics).toContainEqual(
				expect.objectContaining({
					code: 'missing-cached-value',
					message: expect.stringContaining('LineWeight'),
				}),
			);
			for (const name of ['LineColor', 'FillForegnd', 'Color'])
				expect(diagnostics).toContainEqual(
					expect.objectContaining({
						code: 'unsupported-color',
						message: `Cell ${name} uses an unresolved color; a default was used.`,
					}),
				);
		},
	);
	it('does not bypass root selection when the root identity cannot be verified', async () => {
		const { shape, diagnostics } = await parse(undefined, false, 'Other Style');
		expect(shape.style.lineWidth).toBe(0.01);
		expect(shape.style.lineCap).toBeUndefined();
		expect(diagnostics.map((d) => d.code)).toContain('unresolved-line-pattern');
	});
	it('keeps valid numeric root caches authoritative despite a populated dynamic theme', async () => {
		const { shape } = await parse('0');
		expect(shape.style).toMatchObject({
			lineWidth: 0,
			lineCap: 'round',
			linePattern: 0,
			lineColor: '#000000',
			fill: '#000000',
		});
	});
	it('does not mistake entirely absent color selectors for explicit root selection', async () => {
		const { shape, diagnostics } = await parse(undefined, false, 'No Style', false);
		expect(shape.style).toMatchObject({ lineWidth: 0.1, lineCap: 'square', linePattern: 1 });
		expect(diagnostics.map((d) => d.code)).not.toContain('unresolved-line-pattern');
	});
});
