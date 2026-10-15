import { afterEach, expect, it, vi } from 'vitest';
import { pointerViewer } from './__fixtures__/pointer-viewer';
import { nearbyShape } from './viewer-touch-hit';

afterEach(() => {
	Reflect.deleteProperty(document, 'elementsFromPoint');
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});
const touch = (id: number) => ({ pointerType: 'touch', pointerId: id });
/** A 5px-wide vertical connector at x, spanning every y, for shape `id`, plus the paper. */
function thinLines(ui: Awaited<ReturnType<typeof pointerViewer>>, lines: Record<string, number>) {
	Object.defineProperty(document, 'elementsFromPoint', {
		configurable: true,
		value: (x: number) => {
			const hit = Object.entries(lines).filter(([, at]) => Math.abs(x - at) <= 2.5);
			return [...hit.map(([id]) => ui.group(id)), ui.svg];
		},
	});
}
const event = (x: number, type = 'touch') =>
	({ clientX: x, clientY: 40, pointerType: type }) as PointerEvent;

it('finds a thin shape within the touch radius and not beyond it', async () => {
	const ui = await pointerViewer();
	thinLines(ui, { '1': 100 });
	expect(nearbyShape(ui.viewport, ui.svg, event(108), [])).toBe(ui.group('1'));
	expect(nearbyShape(ui.viewport, ui.svg, event(86), [])).toBe(ui.group('1'));
	expect(nearbyShape(ui.viewport, ui.svg, event(120), [])).toBeNull();
	ui.dispose();
});
it('prefers the nearest of two shapes and ignores a mouse', async () => {
	const ui = await pointerViewer();
	thinLines(ui, { '1': 100, '2': 116 });
	expect(nearbyShape(ui.viewport, ui.svg, event(112), [])).toBe(ui.group('2'));
	expect(nearbyShape(ui.viewport, ui.svg, event(104), [])).toBe(ui.group('1'));
	expect(nearbyShape(ui.viewport, ui.svg, event(108, 'mouse'), [])).toBeNull();
	ui.dispose();
});
it('selects a thin shape on a tap beside it', async () => {
	const ui = await pointerViewer();
	thinLines(ui, { '2': 100 });
	ui.pointer('pointerdown', ui.svg, 108, 40, touch(1));
	ui.pointer('pointerup', ui.svg, 108, 40, touch(1));
	expect(ui.controller.state.selectedShapes.map((shape) => shape.id)).toEqual(['2']);
	expect(ui.edits).toHaveLength(0);
	ui.dispose();
});
it('a tap far from any shape selects nothing', async () => {
	const ui = await pointerViewer();
	thinLines(ui, { '2': 100 });
	ui.pointer('pointerdown', ui.svg, 160, 40, touch(1));
	ui.pointer('pointerup', ui.svg, 160, 40, touch(1));
	expect(ui.controller.state.selectedShapes).toHaveLength(0);
	ui.dispose();
});
