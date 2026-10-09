import { describe, expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { createVsdx } from './create-document';
import { snapshotFormatting } from './edit-formatting-commands';
import { visioFallbackQuickStyle, visioShadowPresetGeometry } from './edit-formatting-effects';
import { copySnapshotScene } from './ui/snapshot-scene';
import { assertViewableDocument } from './ui/scene-validation';
import { visioShadowPreset, visioShapeFormattingState } from './ui/shape-formatting';
import { attribute, children } from './sheet';
import { cell, fixture, shape, rectangle } from './test-fixtures';
import { themeFixture } from './theme-fixtures';

const target = { type: 'format-shape' as const, pageId: '0', shapeId: '1' };
const plain = (extra = '', attributes = '') =>
	fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${shape('1', cell('Width', 2) + cell('Height', 1) + rectangle + extra + '<Text>Styled</Text>', attributes)}</Shapes>`,
			},
		],
	});
const read = async (bytes: Uint8Array) => (await parseVsdx(bytes)).pages[0]!.shapes[0]!;
async function cells(bytes: Uint8Array): Promise<Map<string, Element>> {
	const root = await (await VisioPackage.open(bytes)).readXml('visio/pages/page1.xml');
	const node = children(children(root, 'Shapes')[0], 'Shape')[0]!;
	return new Map(children(node, 'Cell').map((item) => [attribute(item, 'N')!, item]));
}

describe('Quick Style edits', () => {
	it('selects theme colour and matrix cells and resolves them from the document theme', async () => {
		const bytes = await themeFixture({
			contents: cell('FillForegnd', '#ff0000', 'RGB(255,0,0)') + cell('LineColor', '#00ff00'),
		});
		const saved = await editVsdx(bytes, [{ ...target, quickStyle: { color: 2, matrix: 2 } }]);
		const written = await cells(saved.bytes);
		for (const name of ['QuickStyleFillColor', 'QuickStyleLineColor', 'QuickStyleFontColor'])
			expect(attribute(written.get(name), 'V')).toBe('2');
		for (const name of ['QuickStyleFillMatrix', 'QuickStyleLineMatrix', 'QuickStyleEffectsMatrix'])
			expect(attribute(written.get(name), 'V')).toBe('2');
		for (const name of ['FillForegnd', 'LineColor']) {
			expect(attribute(written.get(name), 'V')).toBe('Themed');
			expect(attribute(written.get(name), 'F')).toBe('THEMEVAL()');
		}
		const actual = await read(saved.bytes);
		// Generated theme: accent1 is #214365 and matrix 2 uses the placeholder colour.
		expect(actual.style).toMatchObject({ fill: '#214365', lineColor: '#214365' });
		expect(actual.text.color).toBe('#214365');
		const subtle = await read(
			(await editVsdx(saved.bytes, [{ ...target, quickStyle: { color: 2, matrix: 1 } }])).bytes,
		);
		expect(subtle.style.fill).toBe('#ffffff');
		expect(
			(await editVsdx(saved.bytes, [{ ...target, quickStyle: { color: 2, matrix: 2 } }])).bytes,
		).toEqual(saved.bytes);
	});
	it('uses fixed Office colours for a drawing without a theme', async () => {
		const saved = await editVsdx(await plain(cell('FillPattern', 0)), [
			{ ...target, quickStyle: { color: 2, matrix: 4 } },
		]);
		const expected = visioFallbackQuickStyle({ color: 2, matrix: 4 });
		const actual = await read(saved.bytes);
		expect(actual.style).toMatchObject({ fill: expected.fill, lineColor: expected.line });
		expect(actual.text.runs.every((run) => run.color === expected.font)).toBe(true);
		const written = await cells(saved.bytes);
		expect(attribute(written.get('FillForegnd'), 'F')).toBe('RGB(91,155,213)');
		expect(attribute(written.get('QuickStyleFillMatrix'), 'V')).toBe('4');
	});
	it('styles shapes drawn in a new drawing', async () => {
		const blank = await createVsdx();
		const drawn = await editVsdx(blank, [
			{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 1, y: 1, width: 2, height: 1 },
		] as never);
		const saved = await editVsdx(drawn.bytes, [
			{ ...target, quickStyle: { color: 7, matrix: 6 }, shadow: 'bottom-right' },
		]);
		const actual = await read(saved.bytes);
		expect(actual.style.fill).toBe(visioFallbackQuickStyle({ color: 7, matrix: 6 }).fill);
		expect(visioShadowPreset(actual)).toBe('bottom-right');
	});
	it('rejects locked shapes, master instances and guarded selectors', async () => {
		const edit = { ...target, quickStyle: { color: 3, matrix: 3 } } as const;
		await expect(editVsdx(await plain(cell('LockFormat', 1)), [edit])).rejects.toMatchObject({
			code: 'EDIT_PROTECTED_CELL',
		});
		await expect(
			editVsdx(await plain(cell('QuickStyleFillColor', 2, 'GUARD(2)')), [edit]),
		).rejects.toMatchObject({ code: 'EDIT_PROTECTED_CELL' });
		await expect(editVsdx(await plain('', 'Master="2"'), [edit])).rejects.toMatchObject({
			code: 'UNSUPPORTED_FORMAT_EDIT',
		});
		await expect(
			editVsdx(await plain(cell('LineWeight', 0.01, 'THEMEVAL()')), [edit]),
		).rejects.toMatchObject({ code: 'EDIT_UNSUPPORTED_FORMAT_DEPENDENCY' });
	});
	it('validates the command', () => {
		for (const quickStyle of [
			{ color: 1, matrix: 1 },
			{ color: 2, matrix: 0 },
			{ color: 2, matrix: 7 },
			{ color: 2, matrix: 1.5 },
		])
			expect(() => snapshotFormatting({ ...target, quickStyle } as never)).toThrow();
		expect(() =>
			snapshotFormatting({ ...target, quickStyle: { color: 2, matrix: 1 }, fillColor: '#ffffff' }),
		).toThrow(/Quick Style/);
		expect(() => snapshotFormatting({ ...target, shadow: 'inner' } as never)).toThrow();
		expect(snapshotFormatting({ ...target, quickStyle: { color: 0, matrix: 6 } })).toEqual({
			...target,
			quickStyle: { color: 0, matrix: 6 },
		});
	});
});

describe('shadow edits', () => {
	it.each(['bottom-right', 'top-left', 'center'] as const)(
		'writes and renders the %s outer shadow',
		async (preset) => {
			const saved = await editVsdx(await plain(cell('FillForegnd', '#abcdef')), [
				{ ...target, shadow: preset },
			]);
			const written = await cells(saved.bytes);
			const geometry = visioShadowPresetGeometry(preset);
			expect(attribute(written.get('ShdwPattern'), 'V')).toBe('1');
			expect(attribute(written.get('ShapeShdwOffsetX'), 'U')).toBe('PT');
			const document = await parseVsdx(saved.bytes);
			const actual = document.pages[0]!.shapes[0]!;
			expect(actual.style.fill).toBe('#abcdef');
			expect(actual.style.shadow).toMatchObject({ color: '#000000' });
			expect(actual.style.shadow!.opacity).toBeCloseTo(0.4);
			expect(actual.style.shadow!.offsetX).toBeCloseTo(geometry.x / 72);
			expect(actual.style.shadow!.offsetY).toBeCloseTo(geometry.y / 72);
			expect(visioShapeFormattingState([actual]).shadowPreset).toBe(preset);
			const copy = copySnapshotScene(document);
			expect(() => assertViewableDocument(copy)).not.toThrow();
			expect(copy.pages[0]!.shapes[0]!.style.shadow).toEqual(actual.style.shadow);
			const cleared = await read(
				(await editVsdx(saved.bytes, [{ ...target, shadow: 'none' }])).bytes,
			);
			expect(cleared.style.shadow).toBeUndefined();
			expect(visioShadowPreset(cleared)).toBe('none');
		},
	);
	it('rejects a shadow on a format-locked shape', async () => {
		await expect(
			editVsdx(await plain(cell('LockFormat', 1)), [{ ...target, shadow: 'bottom' }]),
		).rejects.toMatchObject({ code: 'EDIT_PROTECTED_CELL' });
	});
	it('rejects an invalid scene shadow', async () => {
		const document = await parseVsdx(
			(await editVsdx(await plain(), [{ ...target, shadow: 'bottom' }])).bytes,
		);
		document.pages[0]!.shapes[0]!.style.shadow!.color = 'url(#x)';
		expect(() => assertViewableDocument(document)).toThrow(/shadow color/);
	});
});
