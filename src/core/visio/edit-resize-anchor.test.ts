import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index';
import { editVsdx, type VisioEdit } from './edit';
import { snapshotVisioEdits, geometryChangedCells } from './edit-commands';
import { resizeVisioShapeAtAnchor } from './edit-resize-anchor';
import { transform } from './geometry';
import { parseVsdx } from './parser';
import { attribute, children } from './sheet';
import { snapshotEdits } from './ui/edit-commands';
import { visioPageEditToDrawing } from './ui/page-edit';
import { cell, fixture, rectangle, shape, section, xml } from './test-fixtures';

type Resize = Extract<VisioEdit, { type: 'resize-shape' }>;
const resize = (anchor: NonNullable<Resize['anchor']> = { x: 0, y: 0.5 }): Resize => ({
	type: 'resize-shape',
	pageId: '0',
	shapeId: '1',
	width: 3,
	height: 1,
	anchor,
});
const target = (
	extra = '',
	localX = cell('LocPinX', 1, 'Width*0.5'),
	localY = cell('LocPinY', 0.5, 'Height*0.5'),
) =>
	shape(
		'1',
		cell('Width', 2) +
			cell('Height', 1) +
			cell('PinX', 2) +
			cell('PinY', 3) +
			localX +
			localY +
			rectangle +
			extra,
	);
const source = (contents = target(), document = '') =>
	fixture({
		document,
		pages: [{ id: '0', contents: `<Shapes>${contents}</Shapes>` }],
		edit: (zip) => zip.file('unknown/preserved.bin', new Uint8Array([2, 7, 255])),
	});
async function local(bytes: Uint8Array) {
	const zip = await JSZip.loadAsync(bytes);
	const root = parseXml(await zip.file('visio/pages/page1.xml')!.async('string')).documentElement;
	const node = children(children(root, 'Shapes')[0], 'Shape')[0]!;
	return new Map(children(node, 'Cell').map((cell) => [attribute(cell, 'N')!, cell]));
}

describe('source-backed anchored resize', () => {
	it.each([0, 0.5, 1] as const)(
		'fixes normalized x anchor %s and preserves retained payloads',
		async (x) => {
			const text =
				'<Text>A<cp IX="0"/> B<fld IX="0"/>2</Text>' +
				section('Field', `<Row IX="0">${cell('Value', 2, '1+1')}</Row>`) +
				'<Unknown retained="yes"/>';
			const bytes = await source(target(text)),
				saved = await editVsdx(bytes, [resize({ x, y: 0.5 })]);
			const parsed = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
			expect(parsed.width).toBe(3);
			expect(parsed.rotation!.pinX).toBe(2 + (0.5 - x));
			expect(parsed.transform[4] + parsed.width * x).toBe(1 + 2 * x);
			const zip = await JSZip.loadAsync(saved.bytes),
				before = await JSZip.loadAsync(bytes);
			expect(await zip.file('visio/pages/page1.xml')!.async('string')).toContain(text);
			for (const [path, entry] of Object.entries(before.files))
				if (!entry.dir && !saved.changedParts.includes(path))
					expect(await zip.file(path)!.async('uint8array'), path).toEqual(
						await entry.async('uint8array'),
					);
		},
	);
	it.each([
		[cell('LocPinX', 0.25), 0.375, 'Width*0.125'],
		[cell('LocPinX', 0.5, 'Width*0.25'), 0.75, 'Width*0.25'],
		[cell('LocPinX', 0.6, 'Width*0.25+0.1 in'), 0.9, 'Width*0.3'],
		[cell('LocPinX', 0.5, 'PinX*0.25'), 0.75, 'Width*0.25'],
	])(
		'normalizes admissible changed-axis local pins like native handle resize',
		async (x, value, formula) => {
			const saved = await editVsdx(await source(target('', x as string, cell('LocPinY', 0.75))), [
				resize(),
			]);
			const cells = await local(saved.bytes);
			expect(Number(attribute(cells.get('LocPinX'), 'V'))).toBeCloseTo(value as number, 12);
			expect(attribute(cells.get('LocPinX'), 'F')).toBe(formula);
			expect(attribute(cells.get('LocPinY'), 'V')).toBe('0.75');
			expect(attribute(cells.get('LocPinY'), 'F')).toBeUndefined();
		},
	);
	it('retains proportional guarded local-pin formulas and unused guarded dimensions', async () => {
		const input = target('', cell('LocPinX', 0.5, 'GUARD(Width*0.25)')).replace(
			'N="Height" V="1"',
			'N="Height" V="1" F="GUARD(1 in)"',
		);
		const saved = await editVsdx(await source(input), [resize()]),
			cells = await local(saved.bytes);
		expect(attribute(cells.get('LocPinX'), 'F')).toBe('GUARD(Width*0.25)');
		expect(attribute(cells.get('LocPinX'), 'V')).toBe('0.75');
		expect(attribute(cells.get('Height'), 'F')).toBe('GUARD(1 in)');
	});
	it('projects both local pins together before validating guarded cross-axis expressions', async () => {
		const input = target('', cell('LocPinX', 0.5, 'GUARD(LocPinY)'), cell('LocPinY', 0.5));
		const saved = await editVsdx(await source(input), [
			{ ...resize({ x: 0, y: 0 }), width: 4, height: 2 },
		]);
		const cells = await local(saved.bytes);
		expect(attribute(cells.get('LocPinX'), 'V')).toBe('1');
		expect(attribute(cells.get('LocPinX'), 'F')).toBe('GUARD(LocPinY)');
	});
	it('materializes implicit pins for a center anchor rather than accepting changed defaults', async () => {
		const input = target()
			.replace('<Cell N="PinX" V="2"/>', '')
			.replace('<Cell N="PinY" V="3"/>', '');
		const saved = await editVsdx(await source(input), [
			{ ...resize({ x: 0.5, y: 0.5 }), height: 2 },
		]);
		const parsed = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
		expect(parsed.rotation).toEqual({ pinX: 1, pinY: 0.5, angle: 0 });
	});
	it.each([
		[Math.PI / 6, 0, 0],
		[0, 1, 0],
		[0, 0, 1],
	] as const)(
		'anchors rotated and flipped local bounds without guessing the pin',
		async (angle, flipX, flipY) => {
			const saved = await editVsdx(
				await source(target(cell('Angle', angle) + cell('FlipX', flipX) + cell('FlipY', flipY))),
				[resize({ x: 0, y: 0.5 })],
			);
			const parsed = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!;
			const before = transform(2, 3, 1, 0.5, angle, flipX === 1, flipY === 1);
			const at = (matrix: readonly number[], x: number, y: number) => [
				matrix[0]! * x + matrix[2]! * y + matrix[4]!,
				matrix[1]! * x + matrix[3]! * y + matrix[5]!,
			];
			at(parsed.transform, 0, 0.5).forEach((value, index) =>
				expect(value).toBeCloseTo(at(before, 0, 0.5)[index]!, 12),
			);
		},
	);
	it('computes one bounded dependent closure and preserves unrelated properties', async () => {
		const user = section('User', `<Row N="Pinned">${cell('Value', 2, 'PinX')}</Row>`).replace(
			'N="Value" V="2"',
			'N="Value" V="2" U="IN"',
		);
		const saved = await editVsdx(
			await source(
				target(user) +
					shape(
						'2',
						cell('Width', 1) + cell('Height', 1) + cell('PinX', 2, 'Sheet.1!PinX') + rectangle,
					),
			),
			[resize()],
		);
		const parsed = (await parseVsdx(saved.bytes)).pages[0]!;
		expect(parsed.shapes[1]!.rotation!.pinX).toBe(2.5);
		const zip = await JSZip.loadAsync(saved.bytes);
		expect(await zip.file('visio/pages/page1.xml')!.async('string')).toContain(
			'N="Value" V="2.5" U="IN" F="PinX"',
		);
	});
	it('preserves fixed-pin semantics when anchor is omitted and makes repeated anchored edits no-ops', async () => {
		const bytes = await source(target('', cell('LocPinX', 0.25)));
		const fixed = await editVsdx(bytes, [
			{ type: 'resize-shape', pageId: '0', shapeId: '1', width: 3, height: 1 },
		]);
		expect(attribute((await local(fixed.bytes)).get('LocPinX'), 'V')).toBe('0.25');
		expect(attribute((await local(fixed.bytes)).get('PinX'), 'V')).toBe('2');
		const anchored = await editVsdx(bytes, [resize()]),
			repeated = await editVsdx(anchored.bytes, [resize()]);
		expect(repeated.bytes).toEqual(anchored.bytes);
		expect(repeated.changedParts).toEqual([]);
	});
});

describe('anchored resize admission and snapshots', () => {
	it.each([
		['guarded pin', target().replace('N="PinX" V="2"', 'N="PinX" V="2" F="GUARD(2 in)"')],
		['guarded width', target().replace('N="Width" V="2"', 'N="Width" V="2" F="GUARD(2 in)"')],
		['constant guarded local pin', target('', cell('LocPinX', 0.25, 'GUARD(0.25 in)'))],
		['movement lock', target(cell('LockMoveX', 1))],
		['dimension lock', target(cell('LockWidth', 1))],
		['stale source cache', target('', cell('LocPinX', 0.7, 'Width*0.25'))],
		[
			'cycle',
			target('', cell('LocPinX', 1, 'PinX')).replace(
				'N="PinX" V="2"',
				'N="PinX" V="1" F="LocPinX"',
			),
		],
		['angle dependency', target(cell('Angle', 0, '(Width-2 in)/(1 in)*15 deg'))],
		['flip dependency', target(cell('FlipX', 0, 'IF(Width&gt;2 in,1,0)'))],
		[
			'field display',
			target(
				'<Text><fld IX="0"/>2</Text>' +
					section('Field', `<Row IX="0">${cell('Value', 2, 'Width/1 in')}</Row>`),
			),
		],
		[
			'nested identity',
			target(`<Unknown>${shape('4', cell('Width', 1) + cell('Height', 1))}</Unknown>`),
		],
		['noncanonical guarded local pin', target('', cell('locpinx', 0.25, 'GUARD(0.25 in)'))],
		['noncanonical movement lock', target(cell('lockmovex', 1))],
	])('refuses %s before changing even an internal source root', (_, contents) => {
		const root = parseXml(xml('PageContents', `<Shapes>${contents}</Shapes>`)).documentElement,
			before = root.toString();
		expect(() =>
			resizeVisioShapeAtAnchor(
				new Map([['0', root]]),
				parseXml(xml('VisioDocument', '')).documentElement,
				resize(),
				() => {},
			),
		).toThrow();
		expect(root.toString()).toBe(before);
	});
	it('refuses an affected stale cache on another retained shape atomically', async () => {
		const bytes = await source(
				target() +
					shape(
						'2',
						cell('Width', 1) + cell('Height', 1) + cell('PinX', 99, 'Sheet.1!PinX') + rectangle,
					),
			),
			original = bytes.slice();
		await expect(editVsdx(bytes, [resize()])).rejects.toThrow('cache disagree');
		expect(bytes).toEqual(original);
	});
	it.each([null, {}, { x: 0 }, { x: 0.2, y: 0 }, { x: '0', y: 0 }, { x: NaN, y: 0 }])(
		'rejects malformed anchor %j at both snapshot boundaries',
		(anchor) => {
			const command = { ...resize(), anchor } as unknown as VisioEdit;
			expect(() => snapshotVisioEdits([command], 1000, 1000)).toThrow('anchor');
			expect(() => snapshotEdits([command])).toThrow('anchor');
		},
	);
	it('owns anchor before await, strips host fields and preserves ratios during unit conversion', async () => {
		const command = resize(),
			bytes = await source(),
			pending = editVsdx(bytes, [command]);
		command.anchor!.x = 1;
		expect(attribute((await local((await pending).bytes)).get('PinX'), 'V')).toBe('2.5');
		const owned = snapshotEdits([
			{ ...resize(), anchor: { x: 0, y: 0.5, arbitrary: 'discard' } } as Resize,
		])[0]!;
		expect(owned.type === 'resize-shape' && owned.anchor).toEqual({ x: 0, y: 0.5 });
		const converted = visioPageEditToDrawing(
			{ ...(await parseVsdx(bytes)).pages[0]!, drawingToPageScale: 0.5 },
			resize(),
		);
		expect(
			converted.type === 'resize-shape' && [converted.width, converted.height, converted.anchor],
		).toEqual([6, 2, { x: 0, y: 0.5 }]);
		expect(geometryChangedCells(resize())).toEqual([
			'Width',
			'Height',
			'PinX',
			'PinY',
			'LocPinX',
			'LocPinY',
		]);
	});
});
