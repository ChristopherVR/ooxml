import { describe, it, expect } from 'vitest';
import { composeAffine, IDENTITY_AFFINE, type AffineMatrix } from '../geometry/affine';
import { parseVsdx } from './parser';
import { editVsdx, type VisioEdit } from './edit';
import { VisioPackage } from './package';
import { fixture, cell, shape, rectangle } from './test-fixtures';
import type { VisioShape } from './model';

const box = (
	id: string,
	x: number,
	y: number,
	w: number,
	h: number,
	extra = '',
	attributes = 'Type="Shape"',
) =>
	shape(
		id,
		cell('PinX', x) +
			cell('PinY', y) +
			cell('Width', w) +
			cell('Height', h) +
			cell('LocPinX', w / 2, 'Width*0.5') +
			cell('LocPinY', h / 2, 'Height*0.5') +
			(extra.includes('N="Angle"') ? '' : cell('Angle', 0)) +
			extra +
			(attributes.includes('Group') ? '' : rectangle),
		attributes,
	);
const groupCells =
	cell('PinX', 2) +
	cell('PinY', 2) +
	cell('Width', 2) +
	cell('Height', 2) +
	cell('LocPinX', 1) +
	cell('LocPinY', 1) +
	cell('Angle', 0);
const page = (contents: string) => fixture({ pages: [{ id: '0', contents }] });
const three = () =>
	page(
		`<Shapes>${box('1', 1, 1, 1, 0.5)}${box('2', 3, 2, 2, 1, cell('Angle', 0.5) + cell('FlipX', 1))}${box('3', 5, 5, 1, 1)}</Shapes>`,
	);

/** Page-space matrices of every leaf sheet, keyed by ID. */
function leaves(shapes: readonly VisioShape[], parent: AffineMatrix = IDENTITY_AFFINE) {
	const result = new Map<string, AffineMatrix>();
	for (const item of shapes) {
		const world = composeAffine(parent, item.transform);
		if (item.children.length)
			for (const [id, matrix] of leaves(item.children, world)) result.set(id, matrix);
		else result.set(item.id, world);
	}
	return result;
}
async function world(bytes: Uint8Array) {
	return leaves((await parseVsdx(bytes)).pages[0]!.shapes);
}
function expectSame(
	actual: Map<string, AffineMatrix>,
	expected: Map<string, AffineMatrix>,
	shift: [number, number] = [0, 0],
	moved: ReadonlySet<string> = new Set(expected.keys()),
) {
	expect([...actual.keys()].sort()).toEqual([...expected.keys()].sort());
	for (const [id, matrix] of expected) {
		const other = actual.get(id)!;
		for (let i = 0; i < 6; i++)
			expect(other[i]).toBeCloseTo(
				matrix[i]! + (!moved.has(id) ? 0 : i === 4 ? shift[0] : i === 5 ? shift[1] : 0),
				9,
			);
	}
}
const group: VisioEdit = { type: 'group-shapes', pageId: '0', shapeId: '9', memberIds: ['2', '1'] };
const ungroup: VisioEdit = { type: 'ungroup-shape', pageId: '0', shapeId: '9' };

describe('group-shapes and ungroup-shape', () => {
	it('groups members into a native Type="Group" sheet without moving them', async () => {
		const source = await three();
		const before = await world(source);
		const saved = await editVsdx(source, [group]);
		expect(saved.changedParts).toEqual(['visio/pages/page1.xml']);
		expect(saved.diagnostics.map((item) => item.code)).toContain('edit-group-experimental');
		const parsed = (await parseVsdx(saved.bytes)).pages[0]!;
		expect(parsed.shapes.map((item) => item.id)).toEqual(['9', '3']);
		const created = parsed.shapes[0]!;
		expect(created.kind).toBe('group');
		expect(created.children.map((item) => item.id)).toEqual(['1', '2']);
		expectSame(await world(saved.bytes), before);
		const xml = new TextDecoder().decode(
			await (await VisioPackage.open(saved.bytes)).readBytes('visio/pages/page1.xml'),
		);
		expect(xml).toMatch(/<Shape ID="9" Type="Group">.*<Shapes><Shape ID="1"/);
		// The bounds are the union of the members' (rotated) alignment boxes.
		expect(created.rotation!.pinX).toBeGreaterThan(0.5);
		expect(created.width).toBeGreaterThan(3);
	});

	it('moves the group, then ungroups to the moved page positions', async () => {
		const source = await three();
		const before = await world(source);
		const grouped = await editVsdx(source, [group]);
		const created = (await parseVsdx(grouped.bytes)).pages[0]!.shapes[0]!;
		const moved = await editVsdx(grouped.bytes, [
			{
				type: 'move-shape',
				pageId: '0',
				shapeId: '9',
				x: created.rotation!.pinX + 1.25,
				y: created.rotation!.pinY - 0.5,
			},
		]);
		expectSame(await world(moved.bytes), before, [1.25, -0.5], new Set(['1', '2']));
		const ungrouped = await editVsdx(moved.bytes, [ungroup]);
		const parsed = (await parseVsdx(ungrouped.bytes)).pages[0]!;
		expect(parsed.shapes.map((item) => item.id)).toEqual(['1', '2', '3']);
		expect(parsed.shapes.every((item) => !item.children.length)).toBe(true);
		expectSame(await world(ungrouped.bytes), before, [1.25, -0.5], new Set(['1', '2']));
	});

	it('ungroups a rotated group, preserving every member transform', async () => {
		const grouped = await editVsdx(await three(), [group]);
		const rotated = await editVsdx(grouped.bytes, [
			{ type: 'rotate-shape', pageId: '0', shapeId: '9', angle: 2.2 },
		]);
		const before = await world(rotated.bytes);
		const ungrouped = await editVsdx(rotated.bytes, [ungroup]);
		expectSame(await world(ungrouped.bytes), before);
		const member = (await parseVsdx(ungrouped.bytes)).pages[0]!.shapes[1]!;
		expect(member.rotation!.angle).toBeCloseTo(2.2 + 0.5, 9);
	});

	it.each([
		['horizontal flip', cell('FlipX', 1)],
		['vertical flip', cell('FlipY', 1)],
		['both flips', cell('FlipX', 1) + cell('FlipY', 1)],
	])('ungroups a rotated group with a %s', async (_name, flips) => {
		const members =
			box('2', 1, 1, 1, 0.5, cell('Angle', 0.3)) + box('3', 2.5, 1.5, 1, 1, cell('FlipY', 1));
		const source = await page(
			`<Shapes>${box('1', 4, 4, 4, 2, cell('Angle', -0.7) + flips + `<Shapes>${members}</Shapes>`, 'Type="Group"')}</Shapes>`,
		);
		const before = await world(source);
		const saved = await editVsdx(source, [{ ...ungroup, shapeId: '1' }]);
		expectSame(await world(saved.bytes), before);
	});

	it('converts native group-scaling member formulas to values on ungroup', async () => {
		const member = shape(
			'2',
			cell('PinX', 1, 'Sheet.1!Width*0.25') +
				cell('PinY', 1, 'Sheet.1!Height*0.5') +
				cell('Width', 2, 'Sheet.1!Width*0.5') +
				cell('Height', 2, 'Sheet.1!Height*1') +
				cell('LocPinX', 1, 'Width*0.5') +
				cell('LocPinY', 1, 'Height*0.5') +
				cell('Angle', 0) +
				rectangle,
			'Type="Shape"',
		);
		const source = await page(
			`<Shapes>${box('1', 3, 3, 4, 2, cell('Angle', 0) + `<Shapes>${member}</Shapes>`, 'Type="Group"')}</Shapes>`,
		);
		const before = await world(source);
		const saved = await editVsdx(source, [{ ...ungroup, shapeId: '1' }]);
		expectSame(await world(saved.bytes), before);
		const xml = new TextDecoder().decode(
			await (await VisioPackage.open(saved.bytes)).readBytes('visio/pages/page1.xml'),
		);
		expect(xml).not.toContain('Sheet.1!');
		expect(xml).toContain('<Cell N="Width" V="2"/>');
	});

	it('round-trips group then ungroup to the original member caches', async () => {
		const source = await three();
		const ungrouped = await editVsdx((await editVsdx(source, [group])).bytes, [ungroup]);
		expectSame(await world(ungrouped.bytes), await world(source));
		expect((await parseVsdx(ungrouped.bytes)).pages[0]!.shapes.map((item) => item.id)).toEqual([
			'1',
			'2',
			'3',
		]);
	});

	it('keeps group resizing and flips unsupported', async () => {
		const grouped = await editVsdx(await three(), [group]);
		for (const command of [
			{ type: 'resize-shape', width: 6, height: 6 },
			{ type: 'flip-shape', axis: 'horizontal' },
		] as const)
			await expect(
				editVsdx(grouped.bytes, [{ ...command, pageId: '0', shapeId: '9' }]),
			).rejects.toThrow();
	});

	const rejected: [string, string, VisioEdit?][] = [
		[
			'glued member',
			`<Shapes>${box('1', 1, 1, 1, 1)}${box('2', 3, 3, 1, 1)}</Shapes><Connects><Connect FromSheet="2" ToSheet="1"/></Connects>`,
		],
		[
			'LockGroup',
			`<Shapes>${box('1', 1, 1, 1, 1, cell('LockGroup', 1))}${box('2', 3, 3, 1, 1)}</Shapes>`,
		],
		[
			'LockGroup on the group',
			`<Shapes>${shape('9', groupCells + cell('LockGroup', 1) + `<Shapes>${box('1', 1, 1, 1, 1)}</Shapes>`, 'Type="Group"')}</Shapes>`,
			ungroup,
		],
		[
			'layer member',
			`<Shapes>${box('1', 1, 1, 1, 1, cell('LayerMember', 0))}${box('2', 3, 3, 1, 1)}</Shapes>`,
		],
		[
			'master instance',
			`<Shapes>${box('1', 1, 1, 1, 1, '', 'Type="Shape" Master="4"')}${box('2', 3, 3, 1, 1)}</Shapes>`,
		],
		[
			'line member',
			`<Shapes>${box('1', 1, 1, 1, 1, cell('BeginX', 0) + cell('BeginY', 0) + cell('EndX', 1) + cell('EndY', 1))}${box('2', 3, 3, 1, 1)}</Shapes>`,
		],
		[
			'pin formula',
			`<Shapes>${box('1', 1, 1, 1, 1).replace('<Cell N="PinX" V="1"/>', '<Cell N="PinX" V="1" F="ThePage!PageWidth*0.1"/>')}${box('2', 3, 3, 1, 1)}</Shapes>`,
		],
		[
			'external reference to a member pin',
			`<Shapes>${box('1', 1, 1, 1, 1)}${box('2', 3, 3, 1, 1)}${box('3', 5, 5, 1, 1, cell('User.x', 0, 'Sheet.1!PinX'))}</Shapes>`,
		],
		[
			'member page reference',
			`<Shapes>${box('1', 1, 1, 1, 1, cell('TxtWidth', 1, 'ThePage!PageWidth*0.1'))}${box('2', 3, 3, 1, 1)}</Shapes>`,
		],
		[
			'existing group ID',
			`<Shapes>${box('1', 1, 1, 1, 1)}${box('2', 3, 3, 1, 1)}${box('9', 5, 5, 1, 1)}</Shapes>`,
		],
		[
			'nested member',
			`<Shapes>${box('1', 1, 1, 1, 1)}${box('5', 3, 3, 2, 2, `<Shapes>${box('2', 1, 1, 1, 1)}</Shapes>`, 'Type="Group"')}</Shapes>`,
		],
		[
			'group with its own text',
			`<Shapes>${shape('9', groupCells + `<Shapes>${box('1', 1, 1, 1, 1)}</Shapes><Text>Label</Text>`, 'Type="Group"')}</Shapes>`,
			ungroup,
		],
		[
			'reference to the removed group',
			`<Shapes>${shape('9', groupCells + `<Shapes>${box('1', 1, 1, 1, 1)}</Shapes>`, 'Type="Group"')}${box('3', 5, 5, 1, 1, cell('TxtWidth', 2, 'Sheet.9!Width'))}</Shapes>`,
			ungroup,
		],
		[
			'member angle formula in a group',
			`<Shapes>${shape('9', groupCells + `<Shapes>${box('1', 1, 1, 1, 1, cell('Angle', 0.1, '0.1+0'))}</Shapes>`, 'Type="Group"')}</Shapes>`,
			ungroup,
		],
		['ungrouping a plain shape', `<Shapes>${box('9', 1, 1, 1, 1)}</Shapes>`, ungroup],
	];
	it.each(rejected)('refuses %s without changing the source', async (_name, contents, edit) => {
		const source = await page(contents);
		const original = source.slice();
		await expect(
			editVsdx(
				source,
				[edit ?? group].map((item) =>
					item.type === 'group-shapes' ? { ...item, memberIds: ['1', '2'] } : item,
				),
			),
		).rejects.toThrow();
		expect(source).toEqual(original);
	});

	it('validates command shape', async () => {
		const source = await three();
		for (const memberIds of [['1'], ['1', '1'], ['1', '9'], ['01', '2']])
			await expect(
				editVsdx(source, [{ type: 'group-shapes', pageId: '0', shapeId: '9', memberIds }]),
			).rejects.toThrow(/INVALID_EDIT|Group/);
	});
});
