import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { snapshotFormatting, type VisioShapeFormatEdit } from './edit-formatting-commands';
import { rootThemeSheet } from './theme-root';
import { readSheet } from './sheet';
import { cell, fixture, shape, rectangle, xml } from './test-fixtures';
import { themeFixture, generatedTheme } from './theme-fixtures';

const target = { type: 'format-shape' as const, pageId: '0', shapeId: '1' };
const source = (extra = '', document = '', attributes = '') =>
	fixture({
		document,
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', rectangle + extra + '<Text>Paint</Text>', attributes)}</Shapes>`,
			},
		],
	});

describe('paint formatting admission and atomic refusal', () => {
	it.each([
		{ linePattern: -1 },
		{ linePattern: 24 },
		{ linePattern: 1.5 },
		{ linePattern: NaN },
		{ fillPattern: -1 },
		{ fillPattern: 25 },
		{ fillPattern: Infinity },
		{ fillPattern: 1.5 },
		{ lineTransparency: -1 },
		{ lineTransparency: 101 },
		{ lineTransparency: '25' },
		{ fillTransparency: -1 },
		{ fillTransparency: 101 },
		{ fillTransparency: Infinity },
		{ fillBackgroundColor: 'none' },
		{ fillBackgroundColor: '#fff' },
		{ fillColor: 'none', fillPattern: 1 },
	])('rejects malformed or contradictory patch %j', async (patch) => {
		const command = { ...target, ...patch } as VisioShapeFormatEdit;
		expect(() => snapshotFormatting(command)).toThrow();
		await expect(editVsdx(await source(), [command])).rejects.toMatchObject({
			code: 'INVALID_EDIT',
		});
	});
	it.each([
		'LinePattern',
		'FillPattern',
		'FillBkgnd',
		'FillForegndTrans',
		'FillBkgndTrans',
		'LineColorTrans',
	])('refuses guarded affected %s cells atomically', async (name) => {
		const bytes = await source(cell(name, 0, 'GUARD(0)'));
		const original = bytes.slice();
		const command = {
			...target,
			linePattern: 2,
			fillPattern: 3,
			fillBackgroundColor: '#0000ff',
			fillTransparency: 25,
			lineTransparency: 50,
		};
		await expect(editVsdx(bytes, [{ ...target, lineWeight: 3 }, command])).rejects.toMatchObject({
			code: 'EDIT_PROTECTED_CELL',
		});
		expect(bytes).toEqual(original);
	});
	it.each(['LinePattern', 'FillPattern', 'FillForegndTrans', 'FillBkgndTrans', 'LineColorTrans'])(
		'refuses wrong scalar units in affected %s',
		async (name) => {
			const bytes = await source(`<Cell N="${name}" V="0" U="PT"/>`);
			await expect(
				editVsdx(bytes, [
					{ ...target, linePattern: 2, fillPattern: 3, fillTransparency: 25, lineTransparency: 50 },
				]),
			).rejects.toMatchObject({ code: 'EDIT_FORMULA_UNIT' });
		},
	);
	it.each([
		cell('LockFormat', 1),
		cell('LayerMember', '2'),
		cell('LinePattern', 1, 'Inh'),
		cell('LinePattern', 2, '1'),
		cell('linepattern', 1),
	])(
		'retains lock, layer, inheritance, stale cache and ambiguous-name restrictions: %s',
		async (extra) => {
			const bytes = await source(extra);
			await expect(editVsdx(bytes, [{ ...target, linePattern: 3 }])).rejects.toThrow();
		},
	);
	it.each(['LinePattern', 'LineColorTrans', 'FillBkgnd', 'FillForegndTrans', 'FillBkgndTrans'])(
		'refuses retained dependencies on changed %s',
		async (name) => {
			const bytes = await source(
				`<Section N="User"><Row N="dependency"><Cell N="Value" V="0" F="${name}"/></Row></Section>`,
			);
			await expect(
				editVsdx(bytes, [
					{
						...target,
						linePattern: 2,
						fillPattern: 3,
						fillTransparency: 25,
						lineTransparency: 50,
						fillBackgroundColor: '#0000ff',
					},
				]),
			).rejects.toMatchObject({ code: 'EDIT_UNSUPPORTED_FORMAT_DEPENDENCY' });
		},
	);
	it('admits a narrowly scoped line-pattern change without overwriting protected unrelated fill', async () => {
		const saved = await editVsdx(await source(cell('FillForegndTrans', 0.25, 'GUARD(0.25)')), [
			{ ...target, linePattern: 2 },
		]);
		expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.style).toMatchObject({
			linePattern: 2,
			fillOpacity: 0.75,
		});
	});
	it.each([
		{ extra: cell('FillGradientEnabled', 1), patch: { fillTransparency: 25 } },
		{ extra: cell('LineGradientEnabled', 1), patch: { lineTransparency: 25 } },
		{ extra: cell('FillPattern', 25), patch: { fillTransparency: 25 } },
		{ extra: cell('FillGradientEnabled', 1), patch: { fillBackgroundColor: '#0000ff' } },
		{ extra: cell('FillGradientEnabled', 'Themed', 'THEMEVAL()'), patch: { fillTransparency: 25 } },
	])(
		'refuses active or unresolved gradients without explicit paint replacement',
		async ({ extra, patch }) => {
			await expect(editVsdx(await source(extra), [{ ...target, ...patch }])).rejects.toMatchObject({
				code: 'UNSUPPORTED_FORMAT_EDIT',
			});
		},
	);
	it('preserves exact no-op bytes even when an unchanged transparency cache belongs to a gradient', async () => {
		const bytes = await source(cell('LineGradientEnabled', 1) + cell('LineColorTrans', 0.25));
		expect((await editVsdx(bytes, [{ ...target, lineTransparency: 25 }])).bytes).toEqual(bytes);
	});
	it('allows an explicit solid-color replacement plus transparency, retaining inactive stop payloads', async () => {
		const saved = await editVsdx(
			await source(cell('LineGradientEnabled', 1) + cell('FillGradientEnabled', 1)),
			[
				{
					...target,
					fillPattern: 3,
					fillTransparency: 25,
					lineColor: '#123456',
					lineTransparency: 50,
				},
			],
		);
		expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.style).toMatchObject({
			fillPatternIndex: 3,
			lineOpacity: 0.5,
		});
	});
	it('resolves inherited solid theme formats using the shared resource resolver', async () => {
		const bytes = await themeFixture({
			contents:
				cell('FillGradientEnabled', 'Themed', 'THEMEVAL()') +
				cell('LineGradientEnabled', 'Themed', 'THEMEVAL()') +
				cell('QuickStyleFillMatrix', 1) +
				cell('QuickStyleLineMatrix', 1),
		});
		const saved = await editVsdx(bytes, [
			{ ...target, fillTransparency: 25, lineTransparency: 50 },
		]);
		expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.style).toMatchObject({
			fillForegroundOpacity: 0.75,
			lineColorOpacity: 0.5,
		});
	});
	it('rejects a selected theme gradient rather than treating missing cached numeric state as solid', async () => {
		const bytes = await themeFixture({
			contents:
				cell('FillGradientEnabled', 'Themed', 'THEMEVAL()') + cell('QuickStyleFillMatrix', 4),
			theme: generatedTheme({ fill: '<a:gradFill/>' }),
		});
		await expect(editVsdx(bytes, [{ ...target, fillTransparency: 25 }])).rejects.toMatchObject({
			code: 'UNSUPPORTED_FORMAT_EDIT',
		});
	});
});

describe('native root line-gradient inheritance', () => {
	it.each([0, 1])(
		'resolves the No Style root LineGradientEnabled=%i through existing root selection',
		(value) => {
			const read = (contents: string) =>
				readSheet(parseXml(xml('Sheet', contents)).documentElement);
			const root = read(cell('LineGradientEnabled', value));
			const sheet = read(
				cell('ColorSchemeIndex', 0) + cell('LineGradientEnabled', 'Themed', 'THEMEVAL()'),
			);
			expect(
				rootThemeSheet(sheet, { rootSheet: root }).cells.get('LineGradientEnabled')?.value,
			).toBe(String(value));
		},
	);
	it('does not trust a stale root gradient cache during transparency proof', async () => {
		const document = `<StyleSheets><StyleSheet ID="0" NameU="No Style">${cell('LineGradientEnabled', 0, '1')}</StyleSheet><StyleSheet ID="2" LineStyle="0">${cell('LineGradientEnabled', 'Themed', 'THEMEVAL()')}</StyleSheet></StyleSheets>`;
		await expect(
			editVsdx(await source(cell('ColorSchemeIndex', 0), document, 'LineStyle="2"'), [
				{ ...target, lineTransparency: 25 },
			]),
		).rejects.toMatchObject({ code: 'UNSUPPORTED_FORMAT_EDIT' });
	});
});
