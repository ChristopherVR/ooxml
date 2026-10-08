import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index';
import { captureVisioClipboard } from './clipboard';
import { deleteVisioShapes } from './edit-delete';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { children } from './sheet';
import { cell, fixture, rectangle, section, shape, xml } from './test-fixtures';
import { visioPasteCommand } from './ui/shape-clipboard';

const remove = (shapeId: string, pageId = '0') => ({
	type: 'delete-shape' as const,
	pageId,
	shapeId,
});
const leaf = (id: string, extra = '', attributes = '') =>
	shape(
		id,
		cell('Width', 2) + cell('Height', 1) + cell('PinX', 2) + cell('PinY', 2) + rectangle + extra,
		attributes,
	);
const link = (target: string) =>
	section('User', `<Row N="Link">${cell('Value', 2, `Sheet.${target}!PinX`)}</Row>`);
const closed = () => leaf('1', link('2')) + leaf('2', link('1'));
const control = () =>
	leaf('3', link('3') + '<Text>Retained<cp IX="0"/> text</Text><Unknown data="preserved"/>');
const source = (shapes = closed() + control(), trailing = '', document = '') =>
	fixture({
		document,
		pages: [{ id: '0', contents: `<Shapes>${shapes}</Shapes>${trailing}` }],
		edit: (zip) => zip.file('unknown/control.bin', new Uint8Array([1, 5, 255])),
	});
const root = (contents: string) =>
	parseXml(xml('PageContents', `<Shapes>${contents}</Shapes>`)).documentElement;
const document = () => parseXml(xml('VisioDocument', '')).documentElement;

describe('atomic reference-closed deletion', () => {
	it.each([
		['1', '2'],
		['2', '1'],
	])('removes a closed pair in either selection order: %s,%s', async (first, second) => {
		const bytes = await source();
		const before = await JSZip.loadAsync(bytes);
		const result = await editVsdx(bytes, [remove(first), remove(second)]);
		const after = await JSZip.loadAsync(result.bytes);
		expect((await parseVsdx(result.bytes)).pages[0]!.shapes.map((node) => node.id)).toEqual(['3']);
		expect(result.changedParts).toEqual(['visio/pages/page1.xml']);
		expect(result.diagnostics.map((item) => item.code)).toContain('edit-delete-experimental');
		const savedRoot = parseXml(
			await after.file('visio/pages/page1.xml')!.async('string'),
		).documentElement;
		const originalRoot = parseXml(
			await before.file('visio/pages/page1.xml')!.async('string'),
		).documentElement;
		expect(children(children(savedRoot, 'Shapes')[0], 'Shape')[0]!.toString()).toEqual(
			children(children(originalRoot, 'Shapes')[0], 'Shape')[2]!.toString(),
		);
		for (const [path, entry] of Object.entries(before.files))
			if (!entry.dir && !result.changedParts.includes(path))
				expect(await after.file(path)!.async('uint8array'), path).toEqual(
					await entry.async('uint8array'),
				);
	});
	it('removes selected owners even if their internal numeric formulas form a cycle', async () => {
		const cycle = (id: string, target: string) =>
			leaf(
				id,
				section('User', `<Row N="Link">${cell('Value', 0, `Sheet.${target}!User.Link`)}</Row>`),
			);
		const result = await editVsdx(await source(cycle('1', '2') + cycle('2', '1') + control()), [
			remove('1'),
			remove('2'),
		]);
		expect((await parseVsdx(result.bytes)).pages[0]!.shapes.map((node) => node.id)).toEqual(['3']);
	});
	it('supports Cut then Paste of the captured closed pair after both originals disappear', async () => {
		const bytes = await source();
		const clipboard = await captureVisioClipboard(bytes, '0', ['2', '1']);
		const deleted = await editVsdx(bytes, [remove('2'), remove('1')]);
		const page = (await parseVsdx(deleted.bytes)).pages[0]!;
		const edit = visioPasteCommand(page, clipboard, { x: 0, y: 0 })!;
		expect(edit.copies).toEqual([
			{ shapeId: '2', newShapeId: '5' },
			{ shapeId: '1', newShapeId: '4' },
		]);
		const saved = await editVsdx(deleted.bytes, [edit]);
		expect((await parseVsdx(saved.bytes)).pages[0]!.shapes.map((node) => node.id)).toEqual([
			'3',
			'4',
			'5',
		]);
		const zip = await JSZip.loadAsync(saved.bytes);
		const content = await zip.file('visio/pages/page1.xml')!.async('string');
		expect(content).toContain('F="Sheet.5!PinX"');
		expect(content).toContain('F="Sheet.4!PinX"');
	});
	it('keeps numeric Sheet namespaces page-local across a multi-page delete array', async () => {
		const bytes = await fixture({
			pages: [
				{ id: '0', contents: `<Shapes>${closed()}${control()}</Shapes>` },
				{ id: '7', contents: `<Shapes>${leaf('1')}${leaf('2', link('1'))}</Shapes>` },
				{ id: '9', contents: `<Shapes>${closed()}${control()}</Shapes>` },
			],
		});
		const result = await editVsdx(bytes, [
			remove('2', '9'),
			remove('1'),
			remove('1', '9'),
			remove('2'),
		]);
		expect(
			(await parseVsdx(result.bytes)).pages.map((page) => page.shapes.map((node) => node.id)),
		).toEqual([['3'], ['1', '2'], ['3']]);
		expect(result.changedParts).toEqual(['visio/pages/page3.xml', 'visio/pages/page1.xml']);
	});
	it.each([
		leaf('3', link('1')),
		leaf('3', section('User', `<Row N="Link">${cell('Value', 2, 'Sheet.2!User.Link')}</Row>`)),
	])('refuses a retained incoming edge into the set', async (outside) => {
		const bytes = await source(closed() + outside),
			original = bytes.slice();
		await expect(editVsdx(bytes, [remove('2'), remove('1')])).rejects.toThrow(
			'referenced by a ShapeSheet formula',
		);
		expect(bytes).toEqual(original);
	});
	it.each(['<Connect FromSheet="1" ToSheet="2"/>', '<Connect FromSheet="3" ToSheet="1"/>'])(
		'refuses internal and external glue rather than guessing healing',
		async (connection) => {
			await expect(
				editVsdx(await source(undefined, `<Connects>${connection}</Connects>`), [
					remove('1'),
					remove('2'),
				]),
			).rejects.toThrow('Connect record');
		},
	);
	it.each([
		cell('LockDelete', 1),
		cell('LockDelete', 0, 'GUARD(1)'),
		'<Cell N="LockDelete" V="0" E="1"/>',
	])('preflights a late protected target without mutating earlier roots', (protection) => {
		const page = root(leaf('1') + leaf('2', protection));
		const before = page.toString();
		expect(() =>
			deleteVisioShapes(new Map([['0', page]]), document(), [remove('1'), remove('2')], () => {}),
		).toThrow();
		expect(page.toString()).toEqual(before);
	});
	it.each([remove('99'), remove('2', '99'), remove('1')])(
		'rejects missing/repeated targets before any removal',
		(last) => {
			const page = root(leaf('1') + leaf('2')),
				before = page.toString();
			expect(() =>
				deleteVisioShapes(new Map([['0', page]]), document(), [remove('1'), last], () => {}),
			).toThrow();
			expect(page.toString()).toEqual(before);
		},
	);
	it('keeps inherited protection admission and independent local movement locks', async () => {
		const styles = `<StyleSheets><StyleSheet ID="0">${cell('LockDelete', 1)}</StyleSheet></StyleSheets>`;
		await expect(
			editVsdx(await source(closed(), '', styles), [remove('1'), remove('2')]),
		).rejects.toThrow('Inherited protection');
		const result = await editVsdx(
			await source(leaf('1', cell('LockMoveX', 1) + cell('LockMoveY', 1)) + leaf('2')),
			[remove('1'), remove('2')],
		);
		expect((await parseVsdx(result.bytes)).pages[0]!.shapes).toHaveLength(0);
	});
	it.each([
		leaf('2', '', 'Master="4"'),
		leaf('2', `<Shapes>${leaf('4')}</Shapes>`, 'Type="Group"'),
		leaf('2', '<ForeignData/>'),
	])('retains conservative inherited/group/foreign restrictions', async (unsupported) => {
		await expect(
			editVsdx(await source(leaf('1') + unsupported), [remove('1'), remove('2')]),
		).rejects.toThrow();
	});
	it('rejects malformed dynamic selected formulas before removing their XML', async () => {
		await expect(
			editVsdx(
				await source(leaf('1', cell('PinX', 2, 'INDIRECT(&quot;Sheet.2!PinX&quot;)')) + leaf('2')),
				[remove('1'), remove('2')],
			),
		).rejects.toThrow('dependency');
	});
	it('refuses concealed nested identities that could leave retained references dangling', async () => {
		const nested = leaf('1', `<Unknown>${leaf('4')}</Unknown>`);
		await expect(
			editVsdx(await source(nested + leaf('2') + leaf('3', link('4'))), [remove('1'), remove('2')]),
		).rejects.toThrow('Nested shape identities');
	});
	it('refuses a retained page-sheet formula and global metadata reference', async () => {
		await expect(
			editVsdx(
				await source(undefined, `<PageSheet>${cell('PageWidth', 2, 'Sheet.1!Width')}</PageSheet>`),
				[remove('1'), remove('2')],
			),
		).rejects.toThrow('referenced');
		await expect(
			editVsdx(
				await source(
					undefined,
					'',
					`<DocumentSheet>${cell('User.Control', 2, 'Sheet.1!PinX')}</DocumentSheet>`,
				),
				[remove('1'), remove('2')],
			),
		).rejects.toThrow('dependencies');
	});
	it.each([
		{ type: 'replace-plain-text', pageId: '0', shapeId: '3', text: 'Changed' },
		{ type: 'move-shape', pageId: '0', shapeId: '3', x: 3, y: 3 },
	] satisfies VisioEdit[])('retains scalar refusal for mixed transactions', async (other) => {
		const bytes = await source(closed() + leaf('3', '<Text>Plain text</Text>')),
			before = bytes.slice();
		await expect(editVsdx(bytes, [other, remove('1'), remove('2')])).rejects.toThrow('referenced');
		expect(bytes).toEqual(before);
	});
});
