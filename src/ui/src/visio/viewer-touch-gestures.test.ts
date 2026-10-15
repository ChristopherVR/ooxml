import { afterEach, expect, it, vi } from 'vitest';
import { pointerViewer } from './__fixtures__/pointer-viewer';

afterEach(() => {
	vi.useRealTimers();
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});
const touch = (id: number) => ({ pointerType: 'touch', pointerId: id });

it('pinch zooms by the finger distance ratio and never edits', async () => {
	const ui = await pointerViewer();
	ui.pointer('pointerdown', ui.svg, 100, 100, touch(1));
	ui.pointer('pointerdown', ui.svg, 200, 100, touch(2));
	ui.pointer('pointermove', ui.viewport, 300, 100, touch(2));
	expect(ui.controller.state.zoom).toBeCloseTo(2);
	ui.pointer('pointermove', ui.viewport, 500, 100, touch(2));
	expect(ui.controller.state.zoom).toBeCloseTo(4);
	ui.pointer('pointermove', ui.viewport, 5000, 100, touch(2));
	expect(ui.controller.state.zoom).toBe(8);
	ui.pointer('pointerup', ui.viewport, 5000, 100, touch(2));
	ui.pointer('pointerup', ui.viewport, 100, 100, touch(1));
	expect(ui.edits).toHaveLength(0);
	ui.dispose();
});
it('a second finger cancels a shape drag in progress', async () => {
	const ui = await pointerViewer();
	ui.select();
	ui.pointer('pointerdown', ui.group('1'), 40, 40, touch(1));
	ui.pointer('pointermove', ui.svg, 60, 60, touch(1));
	expect(ui.svg.querySelectorAll('[data-movement-preview]')).toHaveLength(1);
	ui.pointer('pointerdown', ui.svg, 140, 40, touch(2));
	expect(ui.svg.querySelectorAll('[data-movement-preview]')).toHaveLength(0);
	ui.pointer('pointerup', ui.viewport, 60, 60, touch(1));
	ui.pointer('pointerup', ui.viewport, 140, 40, touch(2));
	expect(ui.edits).toHaveLength(0);
	ui.dispose();
});
it('does not start a pan with the finger left after a pinch', async () => {
	const ui = await pointerViewer();
	ui.viewport.scrollLeft = 50;
	ui.pointer('pointerdown', ui.svg, 100, 100, touch(1));
	ui.pointer('pointerdown', ui.svg, 200, 100, touch(2));
	ui.pointer('pointerup', ui.viewport, 200, 100, touch(2));
	ui.pointer('pointermove', ui.viewport, 150, 100, touch(1));
	ui.pointer('pointerup', ui.viewport, 150, 100, touch(1));
	expect(ui.viewport.scrollLeft).toBe(50);
	ui.dispose();
});
it('a held finger on a shape opens the context menu once on release', async () => {
	const ui = await pointerViewer();
	vi.useFakeTimers();
	const menu = vi.fn();
	ui.viewport.addEventListener('contextmenu', menu);
	ui.pointer('pointerdown', ui.group('1'), 40, 40, touch(1));
	ui.pointer('pointermove', ui.svg, 43, 41, touch(1));
	vi.advanceTimersByTime(499);
	expect(menu).not.toHaveBeenCalled();
	vi.advanceTimersByTime(2);
	ui.pointer('pointerup', ui.svg, 43, 41, touch(1));
	vi.advanceTimersByTime(101);
	expect(menu).toHaveBeenCalledTimes(1);
	expect(ui.edits).toHaveLength(0);
	ui.dispose();
});
it('moving a finger cancels the long press', async () => {
	const ui = await pointerViewer();
	vi.useFakeTimers();
	const menu = vi.fn();
	ui.viewport.addEventListener('contextmenu', menu);
	ui.pointer('pointerdown', ui.group('1'), 40, 40, touch(1));
	ui.pointer('pointermove', ui.svg, 80, 80, touch(1));
	vi.advanceTimersByTime(800);
	expect(menu).not.toHaveBeenCalled();
	ui.pointer('pointerup', ui.svg, 80, 80, touch(1));
	ui.dispose();
});
it('a mouse drag on empty paper still draws a marquee', async () => {
	const ui = await pointerViewer();
	ui.pointer('pointerdown', ui.svg, 300, 300, { pointerType: 'mouse' });
	ui.pointer('pointermove', ui.svg, 340, 340, { pointerType: 'mouse' });
	expect(ui.svg.querySelector('[data-marquee-preview]')).not.toBeNull();
	ui.pointer('pointerup', ui.svg, 340, 340, { pointerType: 'mouse' });
	ui.dispose();
});
