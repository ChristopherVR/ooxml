import { afterEach, expect, it, vi } from 'vitest';
import { pointerViewer } from './__fixtures__/pointer-viewer';
import { ViewerDrawingGesture } from './viewer-drawing-gesture';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});
it.each([false, true])(
	'owns a rejected completion promise and keeps cancellation quiet (%s)',
	async (cancelled) => {
		const ui = await pointerViewer();
		ui.inactive();
		const announce = vi.fn();
		const gesture = new ViewerDrawingGesture(ui.viewport, ui.controller, {
			tool: () => 'rectangle',
			announce,
			finish: async () => {
				throw cancelled
					? new DOMException('Cancelled', 'AbortError')
					: new Error('Creation refused');
			},
		});
		const dispose = gesture.wire();
		try {
			ui.pointer('pointerdown', ui.svg, 10, 10);
			ui.pointer('pointermove', ui.svg, 30, 20);
			ui.pointer('pointerup', ui.svg, 30, 20);
			await Promise.resolve();
			await Promise.resolve();
			if (cancelled) expect(announce).not.toHaveBeenCalled();
			else expect(announce).toHaveBeenCalledWith('Creation refused');
			expect(ui.controller.exportVsdx().bytes).toEqual(ui.bytes);
			expect(ui.svg.querySelector('.draw-preview')).toBeNull();
		} finally {
			dispose();
			ui.dispose();
		}
	},
);
it.each(['selection', 'source', 'destroy'])(
	'does not promote a reentrant %s change into a newer drawing token',
	async (change) => {
		const ui = await pointerViewer();
		ui.inactive();
		const finish = vi.fn(async () => {});
		const gesture = new ViewerDrawingGesture(ui.viewport, ui.controller, {
			tool: () => 'rectangle',
			announce: vi.fn(),
			finish,
		});
		const dispose = gesture.wire();
		Object.defineProperty(ui.svg, 'getScreenCTM', {
			configurable: true,
			value: () => {
				if (change === 'selection') ui.select();
				else if (change === 'source') ui.controller.setDocument(null);
				else ui.controller.destroy();
				return { inverse: () => ({}) };
			},
		});
		try {
			ui.pointer('pointerdown', ui.svg, 10, 10);
			expect(ui.svg.querySelector('.draw-preview')).toBeNull();
			ui.pointer('pointerup', ui.svg, 30, 20);
			expect(finish).not.toHaveBeenCalled();
			expect(ui.edits).toEqual([]);
		} finally {
			dispose();
			ui.dispose();
		}
	},
);
