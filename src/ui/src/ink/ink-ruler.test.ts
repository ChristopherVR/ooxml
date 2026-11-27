// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { DrawState } from './draw-state';
import { defineInkCanvas, type InkStrokeDetail, type OfficeUiInkCanvas } from './ink-canvas';
import { InkRuler, RULER_SNAP } from './ink-ruler';

beforeAll(() => defineInkCanvas());
afterEach(() => document.body.replaceChildren());

function pointer(type: string, x: number, y: number) {
	// jsdom has no PointerEvent: a MouseEvent carries the same coordinates.
	return Object.assign(
		new MouseEvent(type, { clientX: x, clientY: y, bubbles: true, composed: true, button: 0 }),
		{ pointerId: 1, pointerType: 'mouse', pressure: 0.5 },
	);
}

describe('the ink ruler', () => {
	it('snaps a point near its edge onto it, within its ends, at any angle', () => {
		const ruler = new InkRuler();
		ruler.place({ x: 0, y: 0, width: 500, height: 200 });
		expect(ruler.center).toEqual({ x: 250, y: 100 });
		expect(ruler.length).toBe(400);
		expect(ruler.near({ x: 100, y: 100 - RULER_SNAP })).toBe(true);
		expect(ruler.near({ x: 100, y: 100 - RULER_SNAP - 1 })).toBe(false);
		expect(ruler.near({ x: 20, y: 100 })).toBe(false);
		expect(ruler.snap({ x: 120, y: 90 })).toEqual({ x: 120, y: 100 });
		expect(ruler.snap({ x: 0, y: 90 })).toEqual({ x: 50, y: 100 });
		ruler.rotate(90);
		const snapped = ruler.snap({ x: 245, y: 130 });
		expect(snapped.x).toBeCloseTo(250);
		expect(snapped.y).toBeCloseTo(130);
	});

	it('turns with the keyboard and the wheel and moves with the arrow keys', () => {
		const ruler = new InkRuler();
		ruler.moveTo({ x: 10, y: 10 });
		const key = (key: string, init: KeyboardEventInit = {}) =>
			ruler.key(new KeyboardEvent('keydown', { key, ...init }));
		expect(key('ArrowRight')).toBe(true);
		expect(ruler.center).toEqual({ x: 20, y: 10 });
		key('ArrowUp', { ctrlKey: true });
		expect(ruler.center).toEqual({ x: 20, y: 9 });
		key('ArrowRight', { shiftKey: true });
		expect(ruler.angle).toBe(1);
		key('ArrowLeft', { shiftKey: true, ctrlKey: true });
		expect(ruler.angle).toBe(-14);
		ruler.wheel(new WheelEvent('wheel', { deltaY: 100, shiftKey: true }));
		expect(ruler.angle).toBe(1);
		ruler.rotate(190);
		expect(ruler.angle).toBe(-169);
		expect(key('a')).toBe(false);
		expect(ruler.label('de')).toBe('Lineal, -169 Grad');
	});

	it('lays the ruler on the canvas, where a pen stroke begun at its edge follows it', () => {
		const canvas = document.createElement('office-ui-ink-canvas') as OfficeUiInkCanvas;
		document.body.append(canvas);
		canvas.ruler = true;
		canvas.locale = 'fr';
		const ruler = canvas.shadowRoot!.querySelector<SVGGElement>('.ruler')!;
		expect(ruler).not.toBeNull();
		expect(ruler.getAttribute('tabindex')).toBe('0');
		expect(ruler.getAttribute('aria-label')).toBe('Règle, 0 degrés');
		// Without a tool only the ruler takes the pointer.
		expect(canvas.getAttribute('tool')).toBe('');
		canvas.tool = 'pen';
		const strokes: InkStrokeDetail[] = [];
		canvas.addEventListener('office-ink-stroke', (event) =>
			strokes.push((event as CustomEvent<InkStrokeDetail>).detail),
		);
		// The jsdom canvas has no size: the ruler lies across the middle of 600 x 300 at y = 150.
		canvas.dispatchEvent(pointer('pointerdown', 100, 140));
		canvas.dispatchEvent(pointer('pointermove', 200, 160));
		canvas.dispatchEvent(pointer('pointerup', 300, 141));
		expect(strokes[0]?.samples.map(({ x, y }) => [x, y])).toEqual([
			[100, 150],
			[200, 150],
			[300, 150],
		]);
		// One begun away from the edge is drawn as it is.
		canvas.dispatchEvent(pointer('pointerdown', 100, 40));
		canvas.dispatchEvent(pointer('pointerup', 120, 60));
		expect(strokes[1]?.samples.at(-1)).toMatchObject({ x: 120, y: 60 });
		// Shift and an arrow key turn it.
		ruler.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'ArrowDown', shiftKey: true, bubbles: true }),
		);
		expect(canvas.rulerAngle).toBe(1);
	});

	it('is part of the Draw tab state', () => {
		const state = new DrawState();
		let changes = 0;
		state.onChange(() => changes++);
		expect(state.snapshot.ruler).toBe(false);
		state.setRuler(true);
		state.setRuler(true);
		expect(state.snapshot.ruler).toBe(true);
		expect(changes).toBe(1);
	});
});
