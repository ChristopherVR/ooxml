import { expect, it } from 'vitest';
import { editVsdx, type VisioEdit } from './edit';
import { parseVsdx } from './parser';
import { VisioPackage } from './package';
import { fixture, cell, shape, rectangle } from './test-fixtures';
import { attribute, children } from './sheet';
import { snapshotEdits } from './ui/edit-commands';
import { visioPageEditToDrawing } from './ui/page-edit';
import { visioOrderingShape } from './ui/shape-order';

const order = (value: 'front' | 'back' | 'forward' | 'backward', shapeId = '2'): VisioEdit => ({
	type: 'reorder-shape',
	pageId: '0',
	shapeId,
	order: value,
});
const box = (id: string, extra = '', attrs = '') =>
	shape(
		id,
		cell('Width', 2) +
			cell('Height', 1) +
			cell('PinX', 2) +
			cell('PinY', 3) +
			cell('FillForegnd', '#336699') +
			rectangle +
			extra +
			`<Text>Shape ${id}</Text>`,
		attrs,
	);
const source = (extra = '', document = '', attrs = '') =>
	fixture({
		document,
		pages: [
			{
				id: '0',
				contents: `<Shapes>${box('1')}<!--keep-->${box('2', extra, attrs)}${box('3')}${box('4')}</Shapes>`,
			},
		],
	});
const ids = async (bytes: Uint8Array) =>
	(await parseVsdx(bytes)).pages[0]!.shapes.map((shape) => shape.id);

it.each([
	['front', ['1', '3', '4', '2']],
	['back', ['2', '1', '3', '4']],
	['forward', ['1', '3', '2', '4']],
	['backward', ['2', '1', '3', '4']],
] as const)(
	'applies %s to source order and preserves every shape and untouched payload',
	async (command, expected) => {
		const original = await source(cell('UserValue', '9', 'Width/1in'));
		const saved = await editVsdx(original, [order(command)]);
		expect(await ids(saved.bytes)).toEqual(expected);
		expect(saved.changedParts).toEqual(['visio/pages/page1.xml']);
		const before = await VisioPackage.open(original),
			after = await VisioPackage.open(saved.bytes);
		for (const path of before.paths())
			if (!saved.changedParts.includes(path))
				expect(await after.readBytes(path)).toEqual(await before.readBytes(path));
		const root = await after.readXml('visio/pages/page1.xml', 'PageContents');
		expect(root.toString()).toContain('<!--keep-->');
		const target = children(children(root, 'Shapes')[0], 'Shape').find(
			(node) => attribute(node, 'ID') === '2',
		)!;
		expect(children(target, 'Text')[0]!.textContent).toBe('Shape 2');
		expect(
			attribute(
				children(target, 'Cell').find((node) => attribute(node, 'N') === 'UserValue'),
				'F',
			),
		).toBe('Width/1in');
	},
);

it('treats already front/back shapes as no-ops and preserves independent source bytes', async () => {
	const original = await source();
	for (const [command, id] of [
		['front', '4'],
		['forward', '4'],
		['back', '1'],
		['backward', '1'],
	] as const) {
		const saved = await editVsdx(original, [order(command, id)]);
		expect(saved.bytes).toEqual(original);
		expect(saved.bytes).not.toBe(original);
		expect(saved.changedParts).toEqual([]);
		expect(saved.diagnostics).toEqual([]);
	}
});

it('snapshots commands before awaiting and combines ordering with source-backed text edits', async () => {
	const original = await source();
	const command = {
		type: 'reorder-shape' as const,
		pageId: '0',
		shapeId: '2',
		order: 'front' as const,
	};
	const pending = editVsdx(original, [
		command,
		{ type: 'replace-plain-text', pageId: '0', shapeId: '2', text: 'New' },
	]);
	command.shapeId = '1';
	const saved = await pending;
	expect(await ids(saved.bytes)).toEqual(['1', '3', '4', '2']);
	expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[3]!.text.plainText).toBe('New');
	expect(snapshotEdits([{ ...command, discarded: new Map() } as VisioEdit])).toEqual([command]);
	const page = (await parseVsdx(original)).pages[0]!;
	expect(visioPageEditToDrawing({ ...page, drawingToPageScale: 3 }, command)).toEqual(command);
});

it.each([
	['nonzero display band', cell('DisplayLevel', 1), '', ''],
	['stale display band', cell('DisplayLevel', 0, '1'), '', ''],
	['dependent display band', cell('DisplayLevel', 0, 'Width/1in'), '', ''],
	['local selection protection', cell('LockSelect', 1), '', ''],
	['format protection', cell('LockFormat', 1), '', ''],
	['stale protection cache', cell('LockSelect', 0, '1'), '', ''],
	['layer membership', cell('LayerMember', 0), '', ''],
	[
		'inherited display band',
		'',
		`<StyleSheets><StyleSheet ID="0">${cell('DisplayLevel', 1)}</StyleSheet></StyleSheets>`,
		'',
	],
	[
		'inherited lock',
		'',
		`<StyleSheets><StyleSheet ID="0">${cell('LockSelect', 1)}</StyleSheet></StyleSheets>`,
		'',
	],
	[
		'inherited lock despite local zero',
		cell('LockSelect', 0),
		`<StyleSheets><StyleSheet ID="0">${cell('LockSelect', 1)}</StyleSheet></StyleSheets>`,
		'',
	],
	['ambiguous cell', cell('DisplayLevel', 0) + cell('DisplayLevel', 1), '', ''],
	['foreign shape', '<ForeignData/>', '', ''],
	['group', `<Shapes>${box('9')}</Shapes>`, '', 'Type="Group"'],
	['deleted shape', '', '', 'Del="1"'],
] as const)('refuses %s atomically', async (_name, extra, document, attrs) => {
	const original = await source(extra, document, attrs),
		copy = new Uint8Array(original);
	await expect(editVsdx(original, [order('front')])).rejects.toThrow();
	expect(original).toEqual(copy);
});

it('rejects invalid commands, missing targets and duplicate IDs', async () => {
	const original = await source();
	for (const command of [{ ...order('front'), order: 'up' }, order('front', '9')])
		await expect(editVsdx(original, [command as VisioEdit])).rejects.toThrow();
	await expect(
		editVsdx(
			await fixture({ pages: [{ id: '0', contents: `<Shapes>${box('2')}${box('2')}</Shapes>` }] }),
			[order('front')],
		),
	).rejects.toThrow();
});

it('refuses order-sensitive container and dynamic formulas elsewhere in the package', async () => {
	for (const formula of ['CONTAINERSHEETREF(1)!Height', 'INDIRECT("Sheet.2!Width")']) {
		const original = await fixture({
			pages: [
				{ id: '0', contents: `<Shapes>${box('1')}${box('2')}</Shapes>` },
				{
					id: '1',
					contents: `<Shapes>${box('8', cell('UserValue', 1, formula.replaceAll('"', '&quot;')))}</Shapes>`,
				},
			],
		});
		await expect(editVsdx(original, [order('back')])).rejects.toMatchObject({
			code: 'UNSUPPORTED_SHAPE_ORDER',
		});
	}
});

it('keeps connection identities and literal color formulas intact while reordering', async () => {
	const original = await fixture({
		pages: [
			{
				id: '0',
				contents: `<Shapes>${box('1', cell('LineColor', '#336699', 'RGB(51,102,153)'))}${box('2')}</Shapes><Connects><Connect FromSheet="1" ToSheet="2" FromCell="BeginX" ToCell="PinX"/></Connects>`,
			},
		],
	});
	const saved = await editVsdx(original, [order('back')]);
	expect((await parseVsdx(saved.bytes)).pages[0]!.connectors).toEqual(
		(await parseVsdx(original)).pages[0]!.connectors,
	);
	expect(await ids(saved.bytes)).toEqual(['2', '1']);
});

it('exposes scene admission for top-level ordinary shapes without promising source protections', async () => {
	const page = (await parseVsdx(await source())).pages[0]!;
	expect(visioOrderingShape(page, '2')).toBe(page.shapes[1]);
	expect(visioOrderingShape(page, 'missing')).toBeUndefined();
	page.shapes[1]!.layerIds = ['0'];
	expect(visioOrderingShape(page, '2')).toBeUndefined();
	page.shapes[1]!.layerIds = [];
	page.shapes[2]!.masterId = '1';
	expect(visioOrderingShape(page, '2')).toBeUndefined();
});
