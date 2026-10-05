import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser.js';
import { rootThemeSheet } from './theme-root.js';
import { emptySheet } from './sheet.js';
import { cell, fixture, rectangle, shape, xml } from './test-fixtures.js';

const rootCells =
	cell('LineWeight', 0.035) +
	cell('LineCap', 2) +
	cell('LinePattern', 23) +
	cell('LineColor', '#123456') +
	cell('FillForegnd', '#abcdef') +
	cell('LineColorTrans', 0.25);
const requested = [
	'LineWeight',
	'LineCap',
	'LinePattern',
	'LineColor',
	'FillForegnd',
	'LineColorTrans',
]
	.map((name) => cell(name, 'Themed', 'THEMEVAL()'))
	.join('');
const character = (contents: string) =>
	`<Section N="Character"><Row IX="0">${contents}</Row></Section>`;
async function parse(
	options: {
		contents?: string;
		selectors?: string;
		root?: string;
		rootName?: string;
		page?: string;
		pageStyle?: string;
		styles?: string;
	} = {},
) {
	const data = await fixture({
		document: `<StyleSheets><StyleSheet ID="0" NameU="${options.rootName ?? 'No Style'}">${options.root ?? rootCells}${character(cell('Color', '#654321'))}</StyleSheet><StyleSheet ID="1" LineStyle="0" FillStyle="0" TextStyle="0">${requested}${character(cell('Color', 'Themed'))}</StyleSheet>${options.styles ?? ''}</StyleSheets>`,
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', rectangle + (options.selectors ?? cell('ColorSchemeIndex', 0)) + (options.contents ?? '') + '<Text>Hello</Text>', 'LineStyle="1" FillStyle="1" TextStyle="1"')}</Shapes>`,
			},
		],
		edit: (zip) =>
			zip.file(
				'visio/pages/pages.xml',
				xml(
					'Pages',
					`<Page ID="0"><PageSheet ${options.pageStyle ?? ''}>${cell('PageWidth', 8)}${cell('PageHeight', 10)}${options.page ?? ''}</PageSheet><Rel r:id="rId1"/></Page>`,
				),
			),
	});
	const document = await parseVsdx(data);
	return { shape: document.pages[0]!.shapes[0]!, codes: document.diagnostics.map((d) => d.code) };
}
describe('explicit root theme selection', () => {
	it('uses saved root colors, width, cap, pattern and transparency for Themed cells', async () => {
		const result = await parse();
		expect(result.shape.style).toMatchObject({
			lineColor: '#123456',
			fill: '#abcdef',
			lineWidth: 0.035,
			lineCap: 'square',
			linePattern: 23,
			lineDash: [2, 2],
			lineOpacity: 0.75,
		});
		expect(result.shape.text.color).toBe('#654321');
		expect(result.shape.text.runs[0]!.color).toBe('#654321');
		expect(result.codes).not.toContain('unsupported-color');
		expect(result.codes).not.toContain('unresolved-line-cap');
	});
	it('resolves 65534 from explicit page selector', async () => {
		const result = await parse({
			selectors: cell('ColorSchemeIndex', 65534),
			page: cell('ColorSchemeIndex', 0),
		});
		expect(result.shape.style.lineWidth).toBe(0.035);
	});
	it('resolves page selectors through explicitly assigned PageSheet styles', async () => {
		const result = await parse({
			selectors: cell('ColorSchemeIndex', 65534),
			pageStyle: 'LineStyle="2" FillStyle="2" TextStyle="2"',
			styles: `<StyleSheet ID="2">${cell('ColorSchemeIndex', 0)}</StyleSheet>`,
		});
		expect(result.shape.style.lineColor).toBe('#123456');
	});
	it('lets local page selectors override PageSheet style selectors', async () => {
		const result = await parse({
			selectors: cell('ColorSchemeIndex', 65534),
			page: cell('ColorSchemeIndex', 42),
			pageStyle: 'LineStyle="2"',
			styles: `<StyleSheet ID="2">${cell('ColorSchemeIndex', 0)}</StyleSheet>`,
		});
		expect(result.shape.style.lineColor).toBe('#000000');
		expect(result.codes).toContain('unsupported-color');
	});
	it('resolves explicit line matrix zero independently without masking color selection', async () => {
		const result = await parse({
			selectors: cell('ColorSchemeIndex', 42) + cell('QuickStyleLineMatrix', 0),
		});
		expect(result.shape.style.lineWidth).toBe(0.035);
		expect(result.shape.style.lineCap).toBe('square');
		expect(result.shape.style.lineColor).toBe('#000000');
		expect(result.codes).toContain('unsupported-color');
	});
	it.each([
		'',
		cell('ColorSchemeIndex', 'Themed'),
		cell('ColorSchemeIndex', 65534),
		cell('ColorSchemeIndex', 42),
		'<Cell N="ColorSchemeIndex" F="0"/>',
	])(
		'keeps absent, unavailable and malformed theme selections unresolved (%s)',
		async (selectors) => {
			const result = await parse({ selectors });
			expect(result.shape.style.lineWidth).toBe(0.01);
			expect(result.codes).toContain('unsupported-color');
			expect(result.codes).toContain('unresolved-line-cap');
		},
	);
	it.each(['Themed', 'oops', '65534'])(
		'does not treat malformed or unresolved PageSheet %s as root',
		async (value) => {
			const result = await parse({
				selectors: cell('ColorSchemeIndex', 65534),
				page: cell('ColorSchemeIndex', value),
			});
			expect(result.shape.style.lineWidth).toBe(0.01);
		},
	);
	it('does not treat a non-root style with ID zero as a verified root', async () => {
		expect((await parse({ rootName: 'Something else' })).shape.style.lineWidth).toBe(0.01);
	});
	it('preserves numeric/literal caches and never evaluates their formulas', async () => {
		const result = await parse({
			contents:
				cell('LineWeight', 0.07, 'THEMEVAL()') +
				cell('LineColor', '#fedcba') +
				cell('LineCap', 0) +
				cell('LinePattern', 0),
		});
		expect(result.shape.style).toMatchObject({
			lineWidth: 0.07,
			lineColor: '#fedcba',
			lineCap: 'round',
			linePattern: 0,
		});
	});
	it('does not recursively evaluate a themed root or fill a missing root cache', async () => {
		const result = await parse({
			root: cell('LineWeight', 'Themed') + '<Cell N="LineCap" F="0"/>',
		});
		expect(result.shape.style.lineWidth).toBe(0.01);
		expect(result.shape.style.lineCap).toBeUndefined();
		expect(result.codes).toContain('unresolved-line-cap');
	});
	it('does not fill missing shape cells and does not mutate source maps', () => {
		const sheet = emptySheet(),
			root = emptySheet();
		sheet.cells.set('ColorSchemeIndex', { value: '0' });
		sheet.cells.set('LineColor', { value: 'Themed' });
		root.cells.set('LineColor', { value: '#123456' });
		root.cells.set('LineWeight', { value: '0.035' });
		const result = rootThemeSheet(sheet, { rootSheet: root });
		expect(result.cells.get('LineColor')?.value).toBe('#123456');
		expect(sheet.cells.get('LineColor')?.value).toBe('Themed');
		expect(result.cells.has('LineWeight')).toBe(false);
	});
});

describe('root cache errors and immutable Character lookup', () => {
	it('reports errors from substituted root caches without duplicating existing cell errors', async () => {
		const bytes = await fixture({
			document: `<StyleSheets><StyleSheet ID="0" NameU="No Style"><Cell N="LineWeight" V="0.035" E="#REF!"/>${character('<Cell N="Color" V="#123456" E="#REF!"/>')}</StyleSheet><StyleSheet ID="1" LineStyle="0" TextStyle="0">${cell('LineWeight', 'Themed')}${character(cell('Color', 'Themed'))}</StyleSheet></StyleSheets>`,
			pages: [
				{
					id: '0',
					contents: `<Shapes>${shape('1', rectangle + cell('ColorSchemeIndex', 0) + '<Cell N="LineCap" V="0" E="#REF!"/><Text>Test</Text>', 'LineStyle="1" TextStyle="1"')}</Shapes>`,
				},
			],
		});
		const doc = await parseVsdx(bytes);
		expect(doc.pages[0]!.shapes[0]!.style.lineWidth).toBe(0.035);
		expect(doc.pages[0]!.shapes[0]!.text.color).toBe('#123456');
		const errors = doc.diagnostics.filter((d) => d.code === 'cached-cell-error');
		expect(errors).toHaveLength(3);
		for (const name of ['LineWeight', 'LineCap', 'Color'])
			expect(errors.filter((d) => d.message.startsWith(`Cell ${name} `))).toHaveLength(1);
	});
	it('looks up the immutable root Character row once while respecting deleted sections and rows', () => {
		const root = emptySheet(),
			sheet = emptySheet();
		const row = (deleted: boolean, color: string) => ({
			index: '0',
			type: '',
			deleted,
			cells: new Map([['Color', { value: color }]]),
		});
		root.sections.set('deleted-section', {
			name: 'Character',
			index: '0',
			deleted: true,
			cells: new Map(),
			rows: new Map([['0', row(false, '#ff0000')]]),
		});
		root.sections.set('deleted-row', {
			name: 'Character',
			index: '1',
			deleted: false,
			cells: new Map(),
			rows: new Map([['0', row(true, '#00ff00')]]),
		});
		root.sections.set('live', {
			name: 'Character',
			index: '2',
			deleted: false,
			cells: new Map(),
			rows: new Map([['0', row(false, '#123456')]]),
		});
		sheet.cells.set('ColorSchemeIndex', { value: '0' });
		sheet.sections.set('live', {
			name: 'Character',
			index: '0',
			deleted: false,
			cells: new Map(),
			rows: new Map([['0', row(false, 'Themed')]]),
		});
		let scans = 0;
		const values = root.sections.values.bind(root.sections);
		root.sections.values = () => {
			scans++;
			return values();
		};
		for (let index = 0; index < 3; index++) {
			const resolved = rootThemeSheet(sheet, { rootSheet: root });
			expect(resolved.sections.get('live')!.rows.get('0')!.cells.get('Color')?.value).toBe(
				'#123456',
			);
		}
		expect(scans).toBe(1);
		expect(sheet.sections.get('live')!.rows.get('0')!.cells.get('Color')?.value).toBe('Themed');
	});
});
