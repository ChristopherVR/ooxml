import { afterEach, expect, it, vi } from 'vitest';
import { parseVsdx } from 'ooxml-core/visio';
import { pointerViewer } from './__fixtures__/pointer-viewer';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

it('switches snapping off with Alt during a drag and commits the release position once', async () => {
	const snap = vi.fn(() => ({ x: 2, y: 1 }));
	const clearSnap = vi.fn();
	const ui = await pointerViewer(false, true, false, { snap, clearSnap });
	try {
		ui.select(['1', '2']);
		ui.pointer('pointerdown', ui.group());
		ui.pointer('pointermove', ui.svg, 52.3, 34.7);
		expect(snap).toHaveBeenCalledTimes(1);
		const snappedPreview = ui.svg
			.querySelector('[data-movement-preview]')!
			.getAttribute('transform');
		ui.pointer('pointermove', ui.svg, 52.3, 34.7, { altKey: true });
		expect(snap).toHaveBeenCalledTimes(1);
		expect(clearSnap).toHaveBeenCalled();
		expect(ui.svg.querySelector('[data-movement-preview]')!.getAttribute('transform')).not.toBe(
			snappedPreview,
		);
		expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
		ui.pointer('pointerup', ui.svg, 52.3, 34.7, { altKey: true });
		await ui.done();
		expect(ui.edits).toHaveLength(1);
		expect(ui.edits[0]).toHaveLength(2);
		const saved = await parseVsdx(ui.controller.exportVsdx().bytes);
		expect(saved.pages[0]!.shapes[0]!.rotation).toMatchObject({
			pinX: expect.closeTo(5.23),
			pinY: expect.closeTo(7.53),
		});
		expect(saved.pages[0]!.shapes[1]!.rotation).toMatchObject({
			pinX: expect.closeTo(8.23),
			pinY: expect.closeTo(7.53),
		});
		await ui.controller.undo();
		expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
		await ui.controller.redo();
		expect(ui.controller.state.selectedShapes.map((shape) => shape.id)).toEqual(['1', '2']);
	} finally {
		ui.dispose();
	}
});

it.each([
	{ x: 52.3, y: 34.7, axis: 'x', pinX: 6, pinY: 7 },
	{ x: 43, y: 24, axis: 'y', pinX: 4, pinY: 8 },
])(
	'constrains Shift movement on $axis despite snapping on both axes',
	async ({ x, y, axis, pinX, pinY }) => {
		const snap = vi.fn(() => ({ x: 2, y: 1 }));
		const ui = await pointerViewer(false, true, false, { snap });
		try {
			ui.select();
			ui.pointer('pointerdown', ui.group(), 40, 40, { shiftKey: true });
			ui.pointer('pointermove', ui.svg, x, y, { shiftKey: true });
			expect(snap).toHaveBeenLastCalledWith(expect.anything(), ['1'], expect.anything(), axis);
			ui.pointer('pointerup', ui.svg, x, y, { shiftKey: true });
			await ui.done();
			expect(ui.edits[0]).toEqual([
				{ type: 'move-shape', pageId: '1', shapeId: '1', x: pinX, y: pinY },
			]);
		} finally {
			ui.dispose();
		}
	},
);

it('honors Alt+Shift on release and resumes snapping when Alt is released', async () => {
	const snap = vi.fn(() => ({ x: 2, y: 1 }));
	const ui = await pointerViewer(false, true, false, { snap });
	try {
		ui.select();
		ui.pointer('pointerdown', ui.group());
		ui.pointer('pointermove', ui.svg, 52.3, 34.7, { altKey: true });
		expect(snap).not.toHaveBeenCalled();
		ui.pointer('pointermove', ui.svg, 52.3, 34.7);
		expect(snap).toHaveBeenCalledTimes(1);
		ui.pointer('pointerup', ui.svg, 52.3, 34.7, { altKey: true, shiftKey: true });
		await ui.done();
		expect(snap).toHaveBeenCalledTimes(1);
		expect(ui.edits[0]).toEqual([
			{ type: 'move-shape', pageId: '1', shapeId: '1', x: expect.closeTo(5.23), y: 7 },
		]);
	} finally {
		ui.dispose();
	}
});
