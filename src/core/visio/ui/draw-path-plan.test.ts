import { expect, it } from 'vitest';
import { parseVsdx } from '../parser';
import { editVsdx } from '../edit';
import { fixture } from '../test-fixtures';
import { visioPathDrawPlan, visioPathPreview } from './draw-path-plan';

const range = (count: number) => Array.from({ length: count }, (_, index) => index);
const page = async () => (await parseVsdx(await fixture())).pages[0]!;

it('plans a quarter arc in drawing coordinates and renders it after creation', async () => {
	const target = await page();
	const command = visioPathDrawPlan(target, 'arc', [
		{ x: 1, y: 1 },
		{ x: 2, y: 1.5 },
		{ x: 3, y: 2 },
	])!;
	expect(command).toMatchObject({
		type: 'create-path',
		pageId: target.id,
		x: 1,
		y: target.height - 1,
		segments: [{ kind: 'arc', x: 3, y: target.height - 2, ratio: 2 }],
	});
	const saved = await editVsdx(await fixture(), [command]);
	const shape = (await parseVsdx(saved.bytes)).pages[0]!.shapes.find(
		(item) => item.id === (command as { shapeId: string }).shapeId,
	)!;
	expect([shape.width, shape.height]).toEqual([2, 1]);
	expect(
		visioPathPreview('arc', [
			{ x: 1, y: 1 },
			{ x: 3, y: 2 },
		]),
	).toMatch(/^M 1 1 L .* L 3 2$/);
});

it('closes freeform strokes that return to their start and keeps open ones unfilled', async () => {
	const target = await page();
	const ring = range(60).map((index) => ({
		x: 3 + Math.cos((index / 60) * 2 * Math.PI),
		y: 3 + Math.sin((index / 60) * 2 * Math.PI),
	}));
	ring.push({ x: 4.05, y: 3.02 });
	const closed = visioPathDrawPlan(target, 'freeform', ring)!;
	expect(closed).toMatchObject({ type: 'create-path', closed: true });
	const open = visioPathDrawPlan(target, 'freeform', ring.slice(0, 40))!;
	expect(open.type).toBe('create-path');
	expect('closed' in open).toBe(false);
	const saved = await editVsdx(await fixture(), [closed]);
	const shape = (await parseVsdx(saved.bytes)).pages[0]!.shapes.at(-1)!;
	expect(shape.geometry[0]!.fill).toBe(true);
	expect(shape.width).toBeCloseTo(2, 1);
});

it('turns a straight stroke into a line and chains pencil lines and arcs into one path', async () => {
	const target = await page();
	const straight = range(20).map((index) => ({ x: 1 + index / 10, y: 2 }));
	expect(visioPathDrawPlan(target, 'pencil', straight)?.type).toBe('create-line');
	expect(visioPathDrawPlan(target, 'freeform', straight)?.type).toBe('create-line');
	const stroke = [
		...range(21).map((index) => ({ x: 1 + index / 20, y: 4 })),
		...range(20).map((index) => {
			const angle = Math.PI - ((index + 1) / 20) * Math.PI;
			return { x: 2.5 + 0.5 * Math.cos(angle), y: 4 - 0.5 * Math.sin(angle) };
		}),
	];
	const chained = visioPathDrawPlan(target, 'pencil', stroke)!;
	expect(chained.type).toBe('create-path');
	const kinds = (chained as { segments: { kind: string }[] }).segments.map((item) => item.kind);
	expect(kinds).toEqual(['line', 'arc']);
	expect(visioPathDrawPlan(target, 'pencil', [{ x: 1, y: 1 }])).toBeUndefined();
	expect(() => visioPathDrawPlan(target, 'bogus' as never, stroke)).toThrow();
});
