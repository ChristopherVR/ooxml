import { describe, expect, it } from 'vitest';
import { parseVsdx } from './index.js';
import { cell, fixture, rectangle, section, shape } from './test-fixtures.js';

// Generated here from format rules. No third-party fixture bytes are redistributed.
const textShape = (data = '', attributes = '') =>
	shape('1', data + rectangle + '<Text>Saved text</Text>', attributes);
const parse = async (data: string, document = '') =>
	parseVsdx(
		await fixture({ document, pages: [{ id: '0', contents: `<Shapes>${data}</Shapes>` }] }),
	);

describe('format cases identified by external Visio corpus inspection', () => {
	it('diagnoses saved formula errors while preserving last valid cached geometry', async () => {
		const document = await parse(
			shape('1', '<Cell N="Width" V="3" E="Formula error"/>' + rectangle),
		);
		expect(document.pages[0]!.shapes[0]!.width).toBe(3);
		expect(document.diagnostics).toContainEqual(
			expect.objectContaining({ code: 'cached-cell-error' }),
		);
	});
	it('normalizes cached rectangle-corner rounding without a fallback warning', async () => {
		const document = await parse(shape('1', cell('Rounding', 0.2) + rectangle));
		expect(document.pages[0]!.shapes[0]!.geometry[0]!.path).toContain('A 0.2 0.2');
		expect(document.diagnostics.some((d) => d.code === 'unsupported-corner-rounding')).toBe(false);
	});
	it('does not flag geometry-free text-only master subshapes as missing geometry', async () => {
		const document = await parseVsdx(
			await fixture({
				masters: [
					{
						id: '7',
						shapes: shape(
							'10',
							`<Shapes>${shape('11', cell('HideText', 1) + '<Text>Label</Text>')}</Shapes>`,
							'Type="Group"',
						),
					},
				],
				pages: [{ id: '0', contents: `<Shapes>${shape('1', '', 'Master="7"')}</Shapes>` }],
			}),
		);
		expect(document.pages[0]!.shapes[0]!.children[0]!.text.plainText).toBe('');
		expect(document.diagnostics.some((d) => d.code === 'missing-geometry')).toBe(false);
	});
	it('preserves saved RGB text backgrounds and fractional transparency', async () => {
		const document = await parse(
			textShape(cell('TextBkgnd', '#9dbb61') + cell('TextBkgndTrans', 0.6)),
		);
		expect(document.pages[0]!.shapes[0]!.text).toMatchObject({
			backgroundColor: '#9dbb61',
			backgroundOpacity: 0.4,
		});
	});
	it.each([
		[1, '#000000'],
		[2, '#ffffff'],
		[3, '#ff0000'],
		[24, '#1a1a1a'],
	])('normalizes one-based text background palette entry %i', async (index, color) => {
		const document = await parse(textShape(cell('TextBkgnd', index)));
		expect(document.pages[0]!.shapes[0]!.text.backgroundColor).toBe(color);
	});
	it.each([0, 255])('keeps sentinel %i transparent', async (value) => {
		const document = await parse(textShape(cell('TextBkgnd', value)));
		expect(document.pages[0]!.shapes[0]!.text.backgroundColor).toBeUndefined();
	});
	it('resolves document palette overrides before built-in colors', async () => {
		const document = await parse(
			textShape(cell('TextBkgnd', 3)),
			'<Colors><ColorEntry IX="2" RGB="#aabbcc"/></Colors>',
		);
		expect(document.pages[0]!.shapes[0]!.text.backgroundColor).toBe('#aabbcc');
	});
	it('inherits text backdrop through the text style and allows a local transparent override', async () => {
		const styles = `<StyleSheets><StyleSheet ID="1">${cell('TextBkgnd', '#abcdef')}${cell('TextBkgndTrans', 0.25)}</StyleSheet><StyleSheet ID="2" TextStyle="1"/></StyleSheets>`;
		const document = await parse(textShape('', 'TextStyle="2"'), styles);
		expect(document.pages[0]!.shapes[0]!.text).toMatchObject({
			backgroundColor: '#abcdef',
			backgroundOpacity: 0.75,
		});
		const cleared = await parse(textShape(cell('TextBkgnd', 255), 'TextStyle="2"'), styles);
		expect(cleared.pages[0]!.shapes[0]!.text.backgroundColor).toBeUndefined();
	});
	it('inherits backdrop from a master without evaluating its formula', async () => {
		const document = await parseVsdx(
			await fixture({
				masters: [
					{
						id: '7',
						shapes: shape('10', cell('TextBkgnd', '#123456', 'RGB(18,52,86)+1') + rectangle),
					},
				],
				pages: [{ id: '0', contents: `<Shapes>${textShape('', 'Master="7"')}</Shapes>` }],
			}),
		);
		expect(document.pages[0]!.shapes[0]!.text.backgroundColor).toBe('#123456');
	});
	it('warns about unresolved text backdrop instead of inventing an opaque color', async () => {
		const document = await parse(textShape(cell('TextBkgnd', 'Themed')));
		expect(document.pages[0]!.shapes[0]!.text.backgroundColor).toBeUndefined();
		expect(document.diagnostics).toContainEqual(
			expect.objectContaining({ code: 'unsupported-color' }),
		);
	});
	it('does not misreport intentionally hidden master geometry as missing', async () => {
		const hidden = section('Geometry', cell('NoShow', 1));
		const document = await parseVsdx(
			await fixture({
				masters: [{ id: '7', shapes: shape('10', hidden) }],
				pages: [{ id: '0', contents: `<Shapes>${shape('1', '', 'Master="7"')}</Shapes>` }],
			}),
		);
		expect(document.pages[0]!.shapes[0]!.geometry).toEqual([]);
		expect(document.diagnostics.some((d) => d.code === 'missing-geometry')).toBe(false);
		const trulyMissing = await parse(shape('1', ''));
		expect(trulyMissing.diagnostics.some((d) => d.code === 'missing-geometry')).toBe(true);
	});
});
