// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it } from 'vitest';
import { DrawState, inkPensLabels } from './draw-state';
import { defineInkCanvas, type InkStrokeDetail, type OfficeUiInkCanvas } from './ink-canvas';
import { defineInkPens, type OfficeUiInkPens } from './ink-pens';

beforeAll(() => {
	defineInkCanvas();
	defineInkPens();
});
afterEach(() => document.body.replaceChildren());

function pointer(type: string, x: number, y: number, extra: Partial<PointerEventInit> = {}) {
	// jsdom has no PointerEvent: a MouseEvent carries the same coordinates.
	return Object.assign(new MouseEvent(type, { clientX: x, clientY: y, bubbles: true, button: 0 }), {
		pointerId: 1,
		pointerType: 'mouse',
		pressure: 0.5,
		...extra,
	});
}

describe('office-ui-ink-canvas', () => {
	it('lets the pointer through until a tool is set, then reports a finished stroke', () => {
		const canvas = document.createElement('office-ui-ink-canvas') as OfficeUiInkCanvas;
		document.body.append(canvas);
		expect(canvas.hasAttribute('role')).toBe(false);
		canvas.tool = 'pen';
		canvas.label = 'Drawing canvas';
		expect(canvas.getAttribute('tool')).toBe('pen');
		expect(canvas.getAttribute('role')).toBe('application');
		const strokes: InkStrokeDetail[] = [];
		canvas.addEventListener('office-ink-stroke', (event) =>
			strokes.push((event as CustomEvent<InkStrokeDetail>).detail),
		);
		canvas.dispatchEvent(pointer('pointerdown', 10, 10));
		canvas.dispatchEvent(pointer('pointermove', 20, 15));
		expect(canvas.shadowRoot!.querySelector('polyline.live')).not.toBeNull();
		canvas.dispatchEvent(pointer('pointerup', 30, 20));
		expect(strokes).toHaveLength(1);
		expect(strokes[0]!.tool).toBe('pen');
		expect(strokes[0]!.samples.map(({ x, y }) => [x, y])).toEqual([
			[10, 10],
			[20, 15],
		]);
		expect(canvas.shadowRoot!.querySelector('polyline.live')).toBeNull();
	});

	it('leaves touch to scrolling unless Draw with Touch is on, and Escape abandons a stroke', () => {
		const canvas = document.createElement('office-ui-ink-canvas') as OfficeUiInkCanvas;
		canvas.tool = 'highlighter';
		document.body.append(canvas);
		let count = 0;
		canvas.addEventListener('office-ink-stroke', () => count++);
		canvas.dispatchEvent(pointer('pointerdown', 1, 1, { pointerType: 'touch' }));
		canvas.dispatchEvent(pointer('pointerup', 9, 9, { pointerType: 'touch' }));
		expect(count).toBe(0);
		canvas.dispatchEvent(pointer('pointerdown', 1, 1));
		document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		canvas.dispatchEvent(pointer('pointerup', 9, 9));
		expect(count).toBe(0);
	});
});

describe('office-ui-ink-pens', () => {
	function strip() {
		const state = new DrawState();
		const pens = document.createElement('office-ui-ink-pens') as OfficeUiInkPens;
		pens.labels = inkPensLabels((text) => text);
		pens.pens = state.snapshot.pens;
		document.body.append(pens);
		return { state, pens };
	}

	it('names each pen by kind, thickness and colour and presses the pen in use', () => {
		const { pens } = strip();
		const buttons = [...pens.shadowRoot!.querySelectorAll<HTMLButtonElement>('.pen')];
		expect(buttons.map((button) => button.getAttribute('aria-label'))).toEqual([
			'Pen: 0.5 mm, Black',
			'Pen: 0.5 mm, Red',
			'Pen: 0.5 mm, Blue',
			'Pencil: 1 mm, #404040',
			'Highlighter: 6 mm, Highlight Yellow',
		]);
		pens.selected = 1;
		expect(buttons[1]!.getAttribute('aria-pressed')).toBe('true');
		// One tab stop: the pen in use.
		expect(buttons.map((button) => button.tabIndex)).toEqual([-1, 0, -1, -1, -1]);
	});

	it('opens the pressed pen’s colour and thickness menu and reports changes', () => {
		const { pens } = strip();
		pens.selected = 0;
		const changes: unknown[] = [];
		pens.addEventListener('office-ink-pen-change', (event) =>
			changes.push((event as CustomEvent).detail),
		);
		pens.shadowRoot!.querySelector<HTMLButtonElement>('.pen')!.click();
		const menu = pens.shadowRoot!.querySelector('.menu');
		expect(menu?.getAttribute('role')).toBe('dialog');
		menu!.querySelector<HTMLButtonElement>('.swatch[aria-label="Red"]')!.click();
		menu!.querySelector<HTMLButtonElement>('.thickness[aria-label="2 mm"]')!.click();
		expect(changes).toEqual([
			{ index: 0, color: '#E81123' },
			{ index: 0, width: 7.56 },
		]);
		document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
		expect(pens.shadowRoot!.querySelector('.menu')).toBeNull();
	});
});

describe('DrawState', () => {
	it('tracks the tool, pens and Draw with Touch', () => {
		const state = new DrawState();
		expect(state.canvasTool).toBe('');
		state.selectPen(4);
		expect(state.canvasTool).toBe('highlighter');
		state.updatePen({ color: '#00FF00', width: 15 });
		expect(state.pen).toEqual({ kind: 'highlighter', color: '#00FF00', width: 15 });
		expect(state.addPen('pencil')).toBe(true);
		expect(state.canvasTool).toBe('pencil');
		state.removePen(5);
		expect(state.snapshot.pens).toHaveLength(5);
		state.setTool('eraser');
		expect(state.canvasTool).toBe('eraser');
	});
});
