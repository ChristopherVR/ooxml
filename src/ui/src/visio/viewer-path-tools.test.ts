import { afterEach, expect, it, vi } from 'vitest';
import { pointerViewer } from './__fixtures__/pointer-viewer';
import { ShapeDrawTool } from './viewer-draw-tool';
import type { VisioPathTool } from 'ooxml-core/visio/ui';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

async function draw(tool: VisioPathTool, points: [number, number][]) {
	const ui = await pointerViewer();
	ui.inactive();
	const announce = vi.fn();
	const drawer = new ShapeDrawTool(ui.viewport, ui.controller, { tool: () => tool, announce });
	const dispose = drawer.wire();
	const [first, ...rest] = points;
	ui.pointer('pointerdown', ui.svg, first![0], first![1]);
	expect(ui.svg.querySelector('path.draw-preview')).not.toBeNull();
	for (const [x, y] of rest) ui.pointer('pointermove', ui.svg, x, y);
	const last = points[points.length - 1]!;
	ui.pointer('pointerup', ui.svg, last[0], last[1]);
	await ui.done();
	expect(ui.svg.querySelector('.draw-preview')).toBeNull();
	return { ui, announce, dispose };
}
const circle = (cx: number, cy: number, r: number, count: number, sweep = 2 * Math.PI) =>
	Array.from({ length: count + 1 }, (_, index): [number, number] => [
		cx + r * Math.cos((index / count) * sweep),
		cy + r * Math.sin((index / count) * sweep),
	]);

it('draws a closed freeform shape that undo removes and redo restores', async () => {
	const { ui, announce, dispose } = await draw('freeform', circle(40, 40, 15, 48));
	try {
		const [command] = ui.edits.at(-1)!;
		expect(command).toMatchObject({ type: 'create-path', closed: true });
		expect(announce).toHaveBeenCalledWith(
			`Shape ${(command as { shapeId: string }).shapeId} added.`,
		);
		const id = (command as { shapeId: string }).shapeId;
		const shape = ui.controller.state.document!.pages[0]!.shapes.find((item) => item.id === id)!;
		expect(shape.geometry[0]!.fill).toBe(true);
		expect(shape.geometry[0]!.path).toContain('C ');
		await ui.controller.undo();
		expect(ui.controller.state.document!.pages[0]!.shapes.some((item) => item.id === id)).toBe(
			false,
		);
		await ui.controller.redo();
		expect(ui.controller.state.document!.pages[0]!.shapes.some((item) => item.id === id)).toBe(
			true,
		);
	} finally {
		dispose();
		ui.dispose();
	}
});

it('draws a quarter arc and an open pencil chain of a line and an arc', async () => {
	const arc = await draw('arc', [
		[20, 20],
		[30, 25],
		[40, 30],
	]);
	try {
		expect(arc.ui.edits.at(-1)![0]).toMatchObject({
			type: 'create-path',
			segments: [{ kind: 'arc', ratio: 2 }],
		});
	} finally {
		arc.dispose();
		arc.ui.dispose();
	}
	const stroke: [number, number][] = [
		...Array.from({ length: 21 }, (_, index): [number, number] => [10 + index, 50]),
		...circle(40, 50, 10, 20, Math.PI)
			.reverse()
			.map(([x, y]): [number, number] => [x, 100 - y])
			.slice(1),
	];
	const pencil = await draw('pencil', stroke);
	try {
		const [command] = pencil.ui.edits.at(-1)!;
		expect(command).toMatchObject({ type: 'create-path' });
		expect('closed' in command!).toBe(false);
		expect((command as { segments: { kind: string }[] }).segments.map((item) => item.kind)).toEqual(
			['line', 'arc'],
		);
	} finally {
		pencil.dispose();
		pencil.ui.dispose();
	}
});

it('explains a click without a drag instead of creating a path', async () => {
	const { ui, announce, dispose } = await draw('freeform', [
		[20, 20],
		[20.1, 20],
	]);
	try {
		expect(ui.edits).toEqual([]);
		expect(announce).toHaveBeenCalledWith('Drag on the page to draw a freeform.');
	} finally {
		dispose();
		ui.dispose();
	}
});
