import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser.js';
import { rootThemeSheet } from './theme-root.js';
import { emptySheet } from './sheet.js';
import { cell, fixture, rectangle, relation, relations, shape, xml } from './test-fixtures.js';
import { generatedTheme, themeRelationship } from './theme-fixtures.js';

const properties = [
	'FillForegndTrans',
	'FillBkgndTrans',
	'FillPattern',
	'FillGradientDir',
	'FillGradientAngle',
	'FillGradientEnabled',
	'RotateGradientWithShape',
	'UseGroupGradient',
];
async function parse(
	options: { root?: string; local?: string; selectors?: string; page?: string } = {},
) {
	const data = await fixture({
		document: `<StyleSheets><StyleSheet ID="0" NameU="No Style">${options.root ?? cell('FillPattern', 0) + cell('FillForegndTrans', 0.4)}</StyleSheet><StyleSheet ID="1" FillStyle="0">${cell('FillPattern', 'Themed')}${cell('FillForegndTrans', 'Themed')}</StyleSheet></StyleSheets>`,
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', rectangle + cell('FillForegnd', '#abcdef') + (options.selectors ?? cell('ColorSchemeIndex', 42) + cell('QuickStyleFillMatrix', 0)) + (options.local ?? ''), 'FillStyle="1"')}</Shapes>`,
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
			zip.file('visio/theme/theme1.xml', generatedTheme());
			zip.file(
				'visio/pages/pages.xml',
				xml(
					'Pages',
					`<Page ID="0"><PageSheet>${cell('PageWidth', 8)}${cell('PageHeight', 10)}${options.page ?? ''}</PageSheet><Rel r:id="rId1"/></Page>`,
				),
			);
		},
	});
	const document = await parseVsdx(data);
	return { style: document.pages[0]!.shapes[0]!.style, diagnostics: document.diagnostics };
}

describe('explicit root fill format selection', () => {
	it('inherits root no-fill and opacity through a fill style independently of colors', async () => {
		const { style, diagnostics } = await parse();
		expect(style).toMatchObject({ fill: 'none', fillOpacity: 0.6 });
		expect(diagnostics.some((d) => d.message.includes('FillPattern'))).toBe(false);
		expect(diagnostics.some((d) => d.message.includes('FillForegndTrans'))).toBe(false);
	});
	it('does not replace a literal foreground when resolving root fill properties', async () => {
		const { style } = await parse({
			root:
				cell('FillPattern', 1) + cell('FillForegndTrans', 0.25) + cell('FillForegnd', '#ff0000'),
		});
		expect(style).toMatchObject({ fill: '#abcdef', fillOpacity: 0.75 });
	});
	it.each([false, true])(
		'uses root fill formats for root color selection (page inheritance %s)',
		async (page) => {
			const { style } = await parse({
				selectors: cell('ColorSchemeIndex', page ? 65534 : 0) + cell('QuickStyleFillMatrix', 1),
				page: cell('ColorSchemeIndex', 0),
			});
			expect(style).toMatchObject({ fill: 'none', fillOpacity: 0.6 });
		},
	);
	it('preserves local caches instead of evaluating formulas or replacing them', async () => {
		const { style } = await parse({
			local: cell('FillPattern', 1, 'THEMEVAL()') + cell('FillForegndTrans', 0.1, 'THEMEVAL()'),
		});
		expect(style).toMatchObject({ fill: '#abcdef', fillOpacity: 0.9 });
	});
	it.each(['', 'Themed', '65534', 'oops'])(
		'does not infer root fill selection from unavailable selector %s',
		async (selector) => {
			const { style, diagnostics } = await parse({
				selectors:
					cell('ColorSchemeIndex', 42) + (selector ? cell('QuickStyleFillMatrix', selector) : ''),
			});
			expect(style).toMatchObject({ fill: '#abcdef', fillOpacity: 1 });
			expect(diagnostics.some((d) => d.message.includes('FillPattern'))).toBe(true);
		},
	);
	it.each(['', cell('FillPattern', 'Themed'), '<Cell N="FillPattern" F="0"/>'])(
		'does not invent missing or unusable root caches: %s',
		async (root) => {
			const { style, diagnostics } = await parse({ root });
			expect(style.fill).toBe('#abcdef');
			expect(diagnostics.some((d) => d.message.includes('FillPattern'))).toBe(true);
		},
	);
	it('retains unsupported root hatch and gradient diagnostics', async () => {
		const { style, diagnostics } = await parse({
			root: cell('FillPattern', 2) + cell('FillForegndTrans', 0) + cell('FillGradientEnabled', 1),
			local: cell('FillGradientEnabled', 'Themed'),
		});
		expect(style.fillGradient).toBeUndefined();
		expect(diagnostics.map((d) => d.code)).toContain('unsupported-fill-pattern');
		expect(diagnostics.map((d) => d.code)).toContain('unsupported-gradient');
	});
	it('copies supported scalar fill caches immutably without adding absent cells', () => {
		const root = emptySheet(),
			sheet = emptySheet();
		sheet.cells.set('QuickStyleFillMatrix', { value: '0' });
		for (const name of properties) {
			sheet.cells.set(name, { value: 'Themed' });
			root.cells.set(name, { value: '0' });
		}
		root.cells.set('FillForegnd', { value: '#112233' });
		const result = rootThemeSheet(sheet, { rootSheet: root });
		for (const name of properties) {
			expect(result.cells.get(name)?.value).toBe('0');
			expect(sheet.cells.get(name)?.value).toBe('Themed');
		}
		expect(result.cells.has('FillForegnd')).toBe(false);
	});
	it('reports errors from substituted root fill caches', async () => {
		const { style, diagnostics } = await parse({
			root: '<Cell N="FillPattern" V="0" E="#REF!"/>' + cell('FillForegndTrans', 0),
		});
		expect(style.fill).toBe('none');
		expect(
			diagnostics.filter(
				(d) => d.code === 'cached-cell-error' && d.message.includes('FillPattern'),
			),
		).toHaveLength(1);
	});
});
