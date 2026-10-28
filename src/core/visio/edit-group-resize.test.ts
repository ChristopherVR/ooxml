import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';
import { cell, fixture, rectangle, shape, section } from './test-fixtures';
import type { VisioEdit } from './edit-commands';
import { VISIO_RESIZE_HANDLES } from './ui/shape-resize';

const box = (id: string, x: number, y: number, w: number, h: number, extra = '') =>
	shape(
		id,
		cell('Width', w) +
			cell('Height', h) +
			cell('PinX', x) +
			cell('PinY', y) +
			cell('LocPinX', w / 4, 'Width*0.25') +
			cell('LocPinY', h / 4, 'Height*0.25') +
			cell('Angle', 0) +
			rectangle +
			extra,
		'Type="Shape"',
	);
const resize = (width = 8, height = 6): Extract<VisioEdit, { type: 'resize-shape' }> => ({
	type: 'resize-shape',
	pageId: '0',
	shapeId: '9',
	width,
	height,
	anchor: { x: 0, y: 0 },
});
async function source(memberExtra = '', groupExtra = '', pageExtra = '', scale = 1) {
	const members = box('1', 1, 1, 1, 1, memberExtra) + box('2', 3, 1.5, 0.5, 0.5, cell('FlipX', 1));
	return fixture({
		pages: [
			{
				id: '0',
				pageCells: cell('DrawingScale', 1) + cell('PageScale', scale),
				contents: `<Shapes>${shape(
					'9',
					cell('Width', 4) +
						cell('Height', 3) +
						cell('PinX', 5) +
						cell('PinY', 6) +
						cell('LocPinX', 1, 'Width*0.25') +
						cell('LocPinY', 0.75, 'Height*0.25') +
						cell('Angle', 0) +
						groupExtra +
						`<Shapes>${members}</Shapes>`,
					'Type="Group"',
				)}</Shapes>${pageExtra}`,
			},
		],
		edit: (zip) => zip.file('unknown/preserved.bin', new Uint8Array([3, 8, 9])),
	});
}
async function xml(bytes: Uint8Array) {
	return (await JSZip.loadAsync(bytes)).file('visio/pages/page1.xml')!.async('string');
}

it.each([0.5, 1, 2])(
	'resizes flat local group members and preserves the opposite anchor at scale %s',
	async (scale) => {
		const bytes = await source('', '', '', scale);
		const before = (await parseVsdx(bytes)).pages[0]!.shapes[0]!;
		const result = await editVsdx(bytes, [resize()]);
		const after = (await parseVsdx(result.bytes)).pages[0]!.shapes[0]!;
		expect(after.width).toBe(8 * scale);
		expect(after.height).toBe(6 * scale);
		expect(after.transform).toEqual(before.transform);
		for (let i = 0; i < before.children.length; i++) {
			const old = before.children[i]!,
				next = after.children[i]!;
			expect(next.width).toBeCloseTo(old.width * 2, 10);
			expect(next.height).toBeCloseTo(old.height * 2, 10);
			expect(next.rotation!.pinX).toBeCloseTo(old.rotation!.pinX * 2, 10);
			expect(next.rotation!.pinY).toBeCloseTo(old.rotation!.pinY * 2, 10);
			expect(next.transform.slice(0, 4)).toEqual(old.transform.slice(0, 4));
			expect(next.transform[4]).toBeCloseTo(old.transform[4] * 2, 10);
			expect(next.transform[5]).toBeCloseTo(old.transform[5] * 2, 10);
		}
		expect(await xml(result.bytes)).toContain('F="Sheet.9!Width*0.25"');
		expect(
			await (
				await JSZip.loadAsync(result.bytes)
			)
				.file('unknown/preserved.bin')!
				.async('uint8array'),
		).toEqual(new Uint8Array([3, 8, 9]));
		// The saved parent dependencies must remain usable on the next resize, not just the first.
		const again = await editVsdx(result.bytes, [resize(12, 9)]);
		const third = (await parseVsdx(again.bytes)).pages[0]!.shapes[0]!;
		expect(third.children[0]!.width).toBeCloseTo(3 * scale, 10);
		expect(third.children[0]!.rotation!.pinX).toBeCloseTo(3 * scale, 10);
	},
);

it.each([0, Math.PI / 6, Math.PI / 2])(
	'keeps a rotated/reflected parent anchor fixed at angle %s',
	async (angle) => {
		const zip = await JSZip.loadAsync(await source('', cell('FlipY', 1)));
		zip.file(
			'visio/pages/page1.xml',
			(await xml(await source('', cell('FlipY', 1)))).replace(
				'N="Angle" V="0"',
				`N="Angle" V="${angle}"`,
			),
		);
		const bytes = await zip.generateAsync({ type: 'uint8array' });
		const before = (await parseVsdx(bytes)).pages[0]!.shapes[0]!;
		const result = await editVsdx(bytes, [resize(6, 9)]);
		const after = (await parseVsdx(result.bytes)).pages[0]!.shapes[0]!;
		after.transform.forEach((value, i) => expect(value).toBeCloseTo(before.transform[i]!, 10));
		expect(after.children[0]).toMatchObject({
			width: 1.5,
			height: 3,
			rotation: { pinX: 1.5, pinY: 3 },
		});
	},
);

it.each([
	[cell('ResizeMode', 1), '', 1],
	['', cell('ResizeMode', 1), 1],
	[cell('ResizeMode', 2), cell('ResizeMode', 1), 2],
] as const)('honors the member and group resize modes', async (member, group, factor) => {
	const saved = await editVsdx(await source(member, group), [resize()]);
	const child = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.children[0]!;
	expect(child.width).toBe(factor);
	expect(child.height).toBe(factor);
	expect(child.rotation).toMatchObject({ pinX: 2, pinY: 2 });
});

it('retains authored and guarded parent formulas when they follow the requested size', async () => {
	const zip = await JSZip.loadAsync(await source());
	zip.file(
		'visio/pages/page1.xml',
		(await xml(await source())).replace(
			'<Cell N="Width" V="1"/>',
			'<Cell N="Width" V="1" F="GUARD(Sheet.9!Width/4)"/>',
		),
	);
	const saved = await editVsdx(await zip.generateAsync({ type: 'uint8array' }), [resize()]);
	expect(await xml(saved.bytes)).toContain('F="GUARD(Sheet.9!Width/4)"');
	expect((await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.children[0]!.width).toBe(2);
});

it('resizes a newly created group and then ungroups without changing world geometry', async () => {
	const input = await fixture({
		pages: [
			{ id: '0', contents: `<Shapes>${box('1', 2, 3, 2, 1)}${box('2', 5, 5, 1, 1)}</Shapes>` },
		],
	});
	const grouped = await editVsdx(input, [
		{ type: 'group-shapes', pageId: '0', shapeId: '9', memberIds: ['1', '2'] },
	]);
	const saved = await editVsdx(grouped.bytes, [resize()]);
	const group = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
	const combined = await editVsdx(input, [
		{ type: 'group-shapes', pageId: '0', shapeId: '9', memberIds: ['1', '2'] },
		resize(),
	]);
	expect((await parseVsdx(combined.bytes)).pages[0]!.shapes[0]).toEqual(group);
	const ungrouped = await editVsdx(saved.bytes, [
		{ type: 'ungroup-shape', pageId: '0', shapeId: '9' },
	]);
	const shapes = (await parseVsdx(ungrouped.bytes)).pages[0]!.shapes;
	shapes.forEach((shape, i) => {
		expect(shape.width).toBe(group.children[i]!.width);
		expect(shape.height).toBe(group.children[i]!.height);
		expect(shape.transform[4]).toBeCloseTo(
			group.transform[4] + group.children[i]!.transform[4],
			10,
		);
		expect(shape.transform[5]).toBeCloseTo(
			group.transform[5] + group.children[i]!.transform[5],
			10,
		);
	});
});

it.each([
	['locked width', cell('LockWidth', 1), ''],
	['locked position', cell('LockMoveX', 1), ''],
	['invalid mode', cell('ResizeMode', 3), ''],
	['layer', cell('LayerMember', 0), ''],
	['member glue', '', '<Connects><Connect FromSheet="2" ToSheet="1"/></Connects>'],
	['group glue', '', '<Connects><Connect FromSheet="2" ToSheet="9"/></Connects>'],
] as const)('rejects %s without modifying source bytes', async (_name, extra, connections) => {
	const bytes = await source(extra, '', connections),
		before = bytes.slice();
	await expect(editVsdx(bytes, [resize()])).rejects.toThrow();
	expect(bytes).toEqual(before);
});

it.each([
	['rotated member', '<Cell N="Angle" V="0"/>', '<Cell N="Angle" V="0.3"/>'],
	['guarded constant', '<Cell N="Width" V="1"/>', '<Cell N="Width" V="1" F="GUARD(1 in)"/>'],
	[
		'non-scaling formula',
		'<Cell N="Width" V="1"/>',
		'<Cell N="Width" V="1" F="Sheet.9!Height/3"/>',
	],
] as const)('refuses %s instead of breaking its pose or expression', async (_name, from, to) => {
	const zip = await JSZip.loadAsync(await source());
	let text = await zip.file('visio/pages/page1.xml')!.async('string');
	// Restrict replacements to the members, leaving the group's own rotation alone.
	const split = text.indexOf('<Shapes>', text.indexOf('ID="9"'));
	text = text.slice(0, split) + text.slice(split).replace(from, to);
	zip.file('visio/pages/page1.xml', text);
	await expect(
		editVsdx(await zip.generateAsync({ type: 'uint8array' }), [resize(8, 3)]),
	).rejects.toThrow();
});

it.each(['direct', 'indirect', 'new group'] as const)(
	'refuses %s package dependencies on member writes',
	async (kind) => {
		const original =
			kind === 'new group'
				? await fixture({
						pages: [
							{
								id: '0',
								contents: `<Shapes>${box('1', 2, 3, 2, 1)}${box('2', 5, 5, 1, 1)}</Shapes>`,
							},
						],
					})
				: await source();
		const zip = await JSZip.loadAsync(original);
		if (kind === 'indirect') {
			const text = await zip.file('visio/pages/page1.xml')!.async('string');
			zip.file(
				'visio/pages/page1.xml',
				text
					.replace('</PageContents>', '')
					.replace(
						/<\/Shapes>\s*$/,
						`${box('3', 2, 3, 1, 1, section('User', '<Row N="memberWidth">' + cell('Value', 1, 'Sheet.1!Width') + '</Row>'))}</Shapes></PageContents>`,
					),
			);
		}
		const document = await zip.file('visio/document.xml')!.async('string');
		zip.file(
			'visio/document.xml',
			document.replace(
				'</VisioDocument>',
				`${cell('Outside', 1, kind === 'indirect' ? 'Sheet.3!User.memberWidth' : 'Sheet.1!Width')}</VisioDocument>`,
			),
		);
		const commands: VisioEdit[] =
			kind === 'new group'
				? [{ type: 'group-shapes', pageId: '0', shapeId: '9', memberIds: ['1', '2'] }, resize()]
				: [resize()];
		await expect(
			editVsdx(await zip.generateAsync({ type: 'uint8array' }), commands).then(() => undefined),
		).rejects.toThrow(/outside editable page|affected|dependencies/i);
	},
);

it.each([cell('ResizeMode', 1), '<Cell N="ResizeMode" V="2" U="IN"/>'])(
	'rejects nondefault or dimensioned page-level modes',
	async (mode) => {
		const zip = await JSZip.loadAsync(await source());
		const pages = await zip.file('visio/pages/pages.xml')!.async('string');
		zip.file('visio/pages/pages.xml', pages.replace('</PageSheet>', `${mode}</PageSheet>`));
		await expect(
			editVsdx(await zip.generateAsync({ type: 'uint8array' }), [resize()]).then(() => undefined),
		).rejects.toThrow('page-level ResizeMode');
	},
);

it.each(VISIO_RESIZE_HANDLES)(
	'preserves the opposite $id group anchor with nonuniform scaling',
	async (handle) => {
		const bytes = await source();
		const before = (await parseVsdx(bytes)).pages[0]!.shapes[0]!;
		const anchor = { x: (1 - handle.x) as 0 | 0.5 | 1, y: (1 - handle.y) as 0 | 0.5 | 1 };
		const width = handle.x === 0.5 ? before.width : 6;
		const height = handle.y === 0.5 ? before.height : 6;
		const result = await editVsdx(bytes, [{ ...resize(width, height), anchor }]);
		const after = (await parseVsdx(result.bytes)).pages[0]!.shapes[0]!;
		expect(after.transform[4] + anchor.x * width).toBeCloseTo(
			before.transform[4] + anchor.x * before.width,
			10,
		);
		expect(after.transform[5] + anchor.y * height).toBeCloseTo(
			before.transform[5] + anchor.y * before.height,
			10,
		);
		expect(after.children[0]!.width).toBeCloseTo(width / before.width, 10);
		expect(after.children[0]!.height).toBeCloseTo(height / before.height, 10);
	},
);

it('does not grant other group members general recalculation permission', async () => {
	const zip = await JSZip.loadAsync(await source());
	const text = await zip.file('visio/pages/page1.xml')!.async('string');
	const other = shape(
		'8',
		cell('Width', 1) +
			cell('Height', 1) +
			cell('PinX', 1) +
			cell('PinY', 1) +
			cell('LocPinX', 0.5) +
			cell('LocPinY', 0.5) +
			cell('Angle', 0) +
			`<Shapes>${box('7', 1, 1, 1, 1, section('User', '<Row N="siblingWidth">' + cell('Value', 1, 'Sheet.1!Width') + '</Row>'))}</Shapes>`,
		'Type="Group"',
	);
	zip.file(
		'visio/pages/page1.xml',
		text
			.replace('</PageContents>', '')
			.replace(/<\/Shapes>\s*$/, `${other}</Shapes></PageContents>`),
	);
	await expect(
		editVsdx(await zip.generateAsync({ type: 'uint8array' }), [resize()]).then(() => undefined),
	).rejects.toThrow(/grouped cells/i);
});
