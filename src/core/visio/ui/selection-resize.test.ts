import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { editVsdx } from '../edit';
import { parseVsdx } from '../parser';
import { cell, fixture, rectangle, shape } from '../test-fixtures';
import { VISIO_RESIZE_HANDLES, visioResizeDrag } from './shape-resize';
import { visioResizeSelectionDrag, visioResizeSelectionFrame } from './selection-resize';
import { assertVisioGeometryExpectation } from './geometry-expectation';

async function source(scale = 1, extra = '') {
	const contents = [
		{ id: '1', width: 2, height: 1, x: 3, y: 5, lx: 0.25, ly: 0.75, flipX: 0, flipY: 0 },
		{ id: '2', width: 1, height: 2, x: 7, y: 3, lx: 0.75, ly: 0.25, flipX: 1, flipY: 1 },
	]
		.map((s) =>
			shape(
				s.id,
				cell('Width', s.width / scale) +
					cell('Height', s.height / scale) +
					cell('PinX', s.x / scale) +
					cell('PinY', s.y / scale) +
					cell('LocPinX', s.lx / scale) +
					cell('LocPinY', s.ly / scale) +
					cell('FlipX', s.flipX) +
					cell('FlipY', s.flipY) +
					rectangle +
					(s.id === '2' ? extra : '') +
					'<Unknown keep="yes"/>',
			),
		)
		.join('');
	return fixture({
		pages: [
			{
				id: '0',
				width: 8.5 / scale,
				height: 11 / scale,
				pageCells: cell('PageScale', scale) + cell('DrawingScale', 1),
				contents: `<Shapes>${contents}</Shapes>`,
			},
		],
		edit: (zip) => zip.file('unknown/retained.bin', new Uint8Array([3, 5, 7])),
	});
}

it.each(VISIO_RESIZE_HANDLES)(
	'scales members and gaps from $id, with saved reflected/off-center pins',
	async (handle) => {
		for (const scale of [0.5, 1, 2]) {
			const bytes = await source(scale),
				page = (await parseVsdx(bytes)).pages[0]!;
			const snapshot = structuredClone(page);
			const initial = visioResizeSelectionFrame(page, ['1', '2'])!;
			const result = visioResizeSelectionDrag(
				page,
				['1', '2'],
				handle.id,
				{ x: 0, y: 0 },
				{
					x: handle.x === 0.5 ? 0 : handle.x === 1 ? 1 : -1,
					y: handle.y === 0.5 ? 0 : handle.y === 1 ? -0.5 : 0.5,
				},
			)!;
			const saved = await editVsdx(bytes, result.commands);
			const reopenedDocument = await parseVsdx(saved.bytes),
				reopened = reopenedDocument.pages[0]!;
			expect(() =>
				assertVisioGeometryExpectation(reopenedDocument, {
					pageId: page.id,
					members: result.members,
				}),
			).not.toThrow();
			const sx = result.frame.width / initial.width,
				sy = result.frame.height / initial.height;
			for (let i = 0; i < page.shapes.length; ++i) {
				const before = page.shapes[i]!,
					after = reopened.shapes[i]!;
				expect(after.width).toBeCloseTo(before.width * sx, 10);
				expect(after.height).toBeCloseTo(before.height * sy, 10);
				expect(after.transform.slice(0, 4)).toEqual(before.transform.slice(0, 4));
				for (const [u, v] of [
					[0, 0],
					[1, 0],
					[0, 1],
					[1, 1],
				]) {
					const x = before.transform[4] + before.transform[0] * before.width * u!;
					const y = before.transform[5] + before.transform[3] * before.height * v!;
					expect(after.transform[4] + after.transform[0] * after.width * u!).toBeCloseTo(
						result.frame.transform[4] + (x - initial.transform[4]) * sx,
						9,
					);
					expect(after.transform[5] + after.transform[3] * after.height * v!).toBeCloseTo(
						result.frame.transform[5] + (y - initial.transform[5]) * sy,
						9,
					);
				}
			}
			const actual = visioResizeSelectionFrame(reopened, ['1', '2'])!;
			expect(actual.width).toBeCloseTo(result.frame.width, 10);
			expect(actual.height).toBeCloseTo(result.frame.height, 10);
			expect(actual.transform[4]).toBeCloseTo(result.frame.transform[4], 10);
			expect(actual.transform[5]).toBeCloseTo(result.frame.transform[5], 10);
			expect(initial.transform[4] + (1 - handle.x) * initial.width).toBeCloseTo(
				actual.transform[4] + (1 - handle.x) * actual.width,
				10,
			);
			expect(initial.transform[5] + (1 - handle.y) * initial.height).toBeCloseTo(
				actual.transform[5] + (1 - handle.y) * actual.height,
				10,
			);
			expect(page).toEqual(snapshot);
			const zip = await JSZip.loadAsync(saved.bytes);
			expect(await zip.file('unknown/retained.bin')!.async('uint8array')).toEqual(
				new Uint8Array([3, 5, 7]),
			);
			expect(await zip.file('visio/pages/page1.xml')!.async('string')).toContain(
				'<Unknown keep="yes"/>',
			);
		}
	},
);

it('retains single-shape local-axis behavior and leaves no-op selections untouched', async () => {
	const page = (await parseVsdx(await source())).pages[0]!;
	const start = { x: 1, y: 2 },
		end = { x: 1.5, y: 1.75 };
	const single = visioResizeDrag(page, '2', 'ne', start, end)!;
	const result = visioResizeSelectionDrag(page, ['2'], 'ne', start, end)!;
	expect(result).toMatchObject({ frame: single.frame, commands: [single.command] });
	expect(visioResizeSelectionDrag(page, ['1', '2'], 'ne', start, start)!.commands).toEqual([]);
});

it('clamps by the smallest member without enlarging already tiny untouched dimensions', async () => {
	const page = (await parseVsdx(await source())).pages[0]!;
	page.shapes[0]!.width = 0.01;
	const original = visioResizeSelectionFrame(page, ['1', '2'])!;
	const result = visioResizeSelectionDrag(
		page,
		['1', '2'],
		'ne',
		{ x: 0, y: 0 },
		{ x: -100, y: 100 },
	)!;
	expect(result.frame.width).toBe(original.width);
	expect(result.frame.height / original.height).toBeCloseTo(1 / 16);
});

it('declines mixed, rotated, nested, attached and invalid selections as a whole', async () => {
	const page = (await parseVsdx(await source())).pages[0]!;
	for (const ids of [
		[],
		['1', 'missing'],
		['1', '1'],
		Array.from({ length: 501 }, (_, i) => String(i)),
	])
		expect(visioResizeSelectionFrame(page, ids)).toBeUndefined();
	expect(
		visioResizeSelectionDrag(page, ['1', '2'], 'ne', { x: 0, y: 0 }, { x: Infinity, y: 0 }),
	).toBeUndefined();
	const other = page.shapes[1]!;
	for (const change of [
		() => {
			const [a, , c, d, e, f] = other.transform;
			other.transform = [a, 0.2, c, d, e, f];
		},
		() => {
			other.children = [structuredClone(page.shapes[0]!)];
		},
		() => {
			other.masterId = '3';
		},
		() => {
			other.hidden = true;
		},
	]) {
		const original = structuredClone(other);
		change();
		expect(visioResizeSelectionFrame(page, ['1', '2'])).toBeUndefined();
		Object.assign(other, original);
		delete other.masterId;
	}
	page.connectors = [{ fromShapeId: '3', toShapeId: '2' } as (typeof page.connectors)[number]];
	expect(visioResizeSelectionFrame(page, ['1', '2'])).toBeUndefined();
});

it.each([cell('LockWidth', 1), cell('LockMoveX', 1), cell('LockHeight', 1)])(
	'rejects an entire source batch when a later member cannot resize',
	async (protection) => {
		const bytes = await source(1, protection),
			untouched = bytes.slice();
		const page = (await parseVsdx(bytes)).pages[0]!;
		const result = visioResizeSelectionDrag(
			page,
			['1', '2'],
			'ne',
			{ x: 0, y: 0 },
			{ x: 2, y: -1 },
		)!;
		await expect(editVsdx(bytes, result.commands)).rejects.toThrow();
		expect(bytes).toEqual(untouched);
	},
);

it('detects a later resize recalculating an earlier guarded dimension before accepting the plan', async () => {
	const zip = await JSZip.loadAsync(await source());
	const xml = await zip.file('visio/pages/page1.xml')!.async('string');
	zip.file(
		'visio/pages/page1.xml',
		xml.replace('<Cell N="Width" V="2"/>', '<Cell N="Width" V="2" F="GUARD(Sheet.2!Height)"/>'),
	);
	const bytes = await zip.generateAsync({ type: 'uint8array' });
	const page = (await parseVsdx(bytes)).pages[0]!;
	const result = visioResizeSelectionDrag(page, ['1', '2'], 'n', { x: 0, y: 0 }, { x: 0, y: -1 })!;
	const saved = await editVsdx(bytes, result.commands);
	const document = await parseVsdx(saved.bytes);
	expect(document.pages[0]!.shapes[0]!.width).not.toBe(page.shapes[0]!.width);
	expect(() =>
		assertVisioGeometryExpectation(document, { pageId: page.id, members: result.members }),
	).toThrow('Dependent formulas');
});
