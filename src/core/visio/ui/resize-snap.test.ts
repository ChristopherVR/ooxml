import { expect, it } from 'vitest';
import { parseVsdx } from '../parser';
import { cell, fixture, rectangle, shape } from '../test-fixtures';
import type { VisioPage } from '../model';
import { visioResizeSelectionDrag } from './selection-resize';
import { visioResizeSizeMatches, visioSnapResizeDrag } from './resize-snap';
import { visioResizeDrag } from './shape-resize';

const SNAP = { shapes: true, guides: true, threshold: 0.1 };

/** Shape 1 spans x 2..4 (y 4.5..5.5); shape 2 spans x 6..7; shape 3 is 4 wide at x 6.5..10.5. */
async function page(): Promise<VisioPage> {
	const box = (id: string, width: number, height: number, x: number, y: number) =>
		shape(
			id,
			cell('Width', width) +
				cell('Height', height) +
				cell('PinX', x) +
				cell('PinY', y) +
				cell('LocPinX', width / 2) +
				cell('LocPinY', height / 2) +
				cell('FlipX', 0) +
				cell('FlipY', 0) +
				rectangle,
		);
	const contents = [box('1', 2, 1, 3, 5), box('2', 1, 1, 6.5, 5), box('3', 4, 1, 8.5, 5)].join('');
	const bytes = await fixture({
		pages: [
			{
				id: '0',
				width: 12,
				height: 11,
				pageCells: cell('PageScale', 1) + cell('DrawingScale', 1),
				contents: `<Shapes>${contents}</Shapes>`,
			},
		],
	});
	return (await parseVsdx(bytes)).pages[0]!;
}

it('snaps a moving right edge onto the nearest shape edge and reports the line', async () => {
	const scene = await page();
	const result = visioSnapResizeDrag(
		scene,
		['1'],
		'e',
		{ x: 4, y: 5 },
		{ x: 5.97, y: 5 },
		{ snap: SNAP },
	)!;
	// Shape 2's left edge at 6 is 0.03 away; the right edge lands on it, the left stays at 2.
	expect(result.frame.width).toBeCloseTo(4, 6);
	expect(result.frame.pinX).toBeCloseTo(4, 6);
	expect(result.lines).toEqual([expect.objectContaining({ axis: 'x', pos: 6 })]);
});

it('keeps the plain drag when nothing is in the threshold, or when snapping is off', async () => {
	const scene = await page();
	const far = visioSnapResizeDrag(
		scene,
		['1'],
		'e',
		{ x: 4, y: 5 },
		{ x: 4.5, y: 5 },
		{
			snap: SNAP,
		},
	)!;
	expect(far.lines).toEqual([]);
	expect(far.frame.width).toBeCloseTo(2.5, 6);
	const off = visioSnapResizeDrag(scene, ['1'], 'e', { x: 4, y: 5 }, { x: 5.97, y: 5 })!;
	expect(off.lines).toEqual([]);
	expect(off.frame.width).toBeCloseTo(3.97, 6);
});

it('snaps only the axes the handle moves: a side handle ignores the y axis', async () => {
	const scene = await page();
	// The top-right corner also snaps in y; the east handle must not, although y is off a line.
	const east = visioSnapResizeDrag(
		scene,
		['1'],
		'e',
		{ x: 4, y: 5 },
		{ x: 4, y: 5.04 },
		{ snap: SNAP },
	)!;
	expect(east.lines.every((line) => line.axis === 'x')).toBe(true);
	expect(east.frame.height).toBeCloseTo(1, 6);
});

it('resizes about the centre with Ctrl (symmetric): the centre stays and each edge moves twice as far', async () => {
	const scene = await page();
	const drag = visioResizeDrag(
		scene,
		'1',
		'e',
		{ x: 4, y: 5 },
		{ x: 4.5, y: 5 },
		{ symmetric: true },
	);
	expect(drag?.frame.width).toBeCloseTo(3, 6);
	expect(drag?.frame.pinX).toBeCloseTo(3, 6);
	expect(drag?.frame.pinY).toBeCloseTo(5, 6);
	const selection = visioResizeSelectionDrag(
		scene,
		['1', '2'],
		'se',
		{ x: 7, y: 4.5 },
		{ x: 7.5, y: 4.5 },
		{ symmetric: true },
	)!;
	// The selection spans x 2..7 (centre 4.5); pulling its east corner 0.5 grows it 1 in total.
	expect(selection.frame.width).toBeCloseTo(6, 6);
	expect(selection.frame.pinX).toBeCloseTo(4.5, 6);
	// A corner keeps proportions, so the height follows the 6/5 scale too.
	expect(selection.frame.height).toBeCloseTo(1.2, 6);
});

it('hints a size match for a nearby shape of equal width', async () => {
	const scene = await page();
	const result = visioSnapResizeDrag(
		scene,
		['1'],
		'e',
		{ x: 4, y: 5 },
		{ x: 5.97, y: 5 },
		{ snap: SNAP },
	)!;
	// Shape 3 is 4 wide, as the 4-wide result is; shape 2 is 1 high, as the result is too.
	expect(result.sizeMatches).toEqual(
		expect.arrayContaining([
			expect.objectContaining({ axis: 'x', shapeId: '3' }),
			expect.objectContaining({ axis: 'y', shapeId: '2' }),
		]),
	);
	expect(
		visioResizeSizeMatches(scene, ['1'], result.frame).some(
			(match) => match.axis === 'x' && match.shapeId === '3',
		),
	).toBe(true);
});
