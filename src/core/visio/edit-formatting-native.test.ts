import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { cell, fixture, shape, rectangle, section } from './test-fixtures';

const target = { pageId: '0', shapeId: '1' };
const textEdit = { ...target, type: 'format-text' as const };
const paintEdit = { ...target, type: 'format-shape' as const };
const text = '<Text><cp IX="0"/>Hello native text</Text>';
const local = cell('Width', 4) + cell('Height', 2) + rectangle;
const characters = (contents: string) => section('Character', `<Row IX="0">${contents}</Row>`);
const nativeFonts = '<FaceNames><FaceName NameU="Calibri"/><FaceName NameU="Arial"/></FaceNames>';
const make = (document: string, cells = '', attributes = '') =>
	fixture({
		document,
		pages: [
			{ id: '0', contents: `<Shapes>${shape('1', local + cells + text, attributes)}</Shapes>` },
		],
	});

describe('native font table and theme formatting admission', () => {
	it('supports ID-free FaceNames and name-valued FONT caches without guessing font IDs', async () => {
		const bytes = await make(
			nativeFonts,
			characters(cell('Font', 'Arial', 'FONT(&quot;Arial&quot;)')),
		);
		const saved = await editVsdx(bytes, [{ ...textEdit, fontFamily: 'Calibri' }]);
		expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.text.fontFamily).toBe('Calibri');
		const xml = await (
			await JSZip.loadAsync(saved.bytes)
		)
			.file('visio/pages/page1.xml')!
			.async('string');
		expect(xml).toContain('N="Font" V="Calibri" F="FONT(&quot;Calibri&quot;)"');
		const unchanged = await editVsdx(saved.bytes, [{ ...textEdit, fontFamily: 'Calibri' }]);
		expect(unchanged.bytes).toEqual(saved.bytes);
		expect(unchanged.changedParts).toEqual([]);
	});

	it('rejects ambiguous explicit IDs and duplicate names in ID-free tables', async () => {
		for (const fonts of [
			'<FaceNames><FaceName ID="0" Name="Arial"/><FaceName ID="0" Name="Calibri"/></FaceNames>',
			'<FaceNames><FaceName NameU="Arial"/><FaceName NameU="Arial"/></FaceNames>',
		])
			await expect(
				editVsdx(await make(fonts), [{ ...textEdit, fontFamily: 'Arial' }]),
			).rejects.toThrow(/unique|match one/i);
	});

	it.each(['THEMEVAL()', 'THEMEGUARD(THEMEVAL())', 'THEMEGUARD(THEME(&quot;AccentColor2&quot;))'])(
		'overrides inherited pure %s paints with local formatting and preserves original styles',
		async (formula) => {
			const document = `<StyleSheets><StyleSheet ID="0">${cell('LayerMember', '')}${cell('LockFormat', 0)}</StyleSheet><StyleSheet ID="6" FillStyle="0" LineStyle="0">${cell('FillPattern', 'Themed', formula)}${cell('FillForegnd', 'Themed', formula)}${cell('FillForegndTrans', 'Themed', formula)}${cell('LineColor', 'Themed', formula)}${cell('LineWeight', 'Themed', formula)}${cell('LineColorTrans', 'Themed', formula)}${cell('FillGradientEnabled', 'Themed', formula)}${cell('LineGradientEnabled', 'Themed', formula)}</StyleSheet><StyleSheet ID="3" FillStyle="6" LineStyle="6"/></StyleSheets>`;
			const bytes = await make(document, '', 'LineStyle="3" FillStyle="3"');
			const saved = await editVsdx(bytes, [
				{ ...paintEdit, fillColor: '#abcdef', lineColor: '#123456', lineWeight: 2.25 },
			]);
			expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.style).toMatchObject({
				fill: '#abcdef',
				lineColor: '#123456',
				lineWidth: 2.25 / 72,
				fillOpacity: 1,
				lineOpacity: 1,
			});
			const before = await JSZip.loadAsync(bytes),
				after = await JSZip.loadAsync(saved.bytes);
			expect(await after.file('visio/document.xml')!.async('uint8array')).toEqual(
				await before.file('visio/document.xml')!.async('uint8array'),
			);
		},
	);

	it.each([
		'GUARD(THEMEVAL())',
		'THEMEGUARD(GUARD(0))',
		'THEMEGUARD(SETATREF(Width))',
		'THEMEGUARD(Width)',
		'THEMEGUARD(THEMEVAL(Width))',
	])('keeps genuine guards and reference-bearing theme formula %s protected', async (formula) => {
		const bytes = await make('', cell('FillForegnd', 'Themed', formula));
		const before = bytes.slice();
		await expect(editVsdx(bytes, [{ ...paintEdit, fillColor: '#abcdef' }])).rejects.toMatchObject({
			code: 'EDIT_PROTECTED_CELL',
		});
		expect(bytes).toEqual(before);
	});

	it('retains explicit shape Inh and error cache refusals for theme lookups', async () => {
		for (const node of [
			'<Cell N="FillForegnd" V="Themed" F="Inh"/>',
			'<Cell N="FillForegnd" V="Themed" F="THEMEVAL()" E="error"/>',
		])
			await expect(
				editVsdx(await make('', node), [{ ...paintEdit, fillColor: '#abcdef' }]),
			).rejects.toMatchObject({ code: 'EDIT_PROTECTED_CELL' });
	});
});
