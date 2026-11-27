// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { openInkEquationDialog } from './ink-equation-dialog';
import {
	inkMathRecognizer,
	onInkMathRecognizerChange,
	setInkMathRecognizer,
	type InkMathRecognizer,
} from './math-recognizer';

afterEach(() => {
	vi.useRealTimers();
	document.body.replaceChildren();
});

const stroke = (canvas: Element, points: [number, number][]) =>
	canvas.dispatchEvent(
		new CustomEvent('office-ink-stroke', {
			detail: {
				tool: 'pen',
				samples: points.map(([x, y], time) => ({ x, y, pressure: 0.5, time })),
			},
		}),
	);

describe('the shared ink math recognizer', () => {
	it('is one recognizer for every product, removed by the function that set it', () => {
		const heard = vi.fn();
		const off = onInkMathRecognizerChange(heard);
		const recognizer: InkMathRecognizer = async () => 'x';
		const remove = setInkMathRecognizer(recognizer);
		try {
			expect(inkMathRecognizer()).toBe(recognizer);
			expect(heard).toHaveBeenCalledTimes(1);
		} finally {
			remove();
			off();
		}
		expect(inkMathRecognizer()).toBeUndefined();
		expect(heard).toHaveBeenCalledTimes(2);
	});
});

describe('the Ink Equation dialog', () => {
	it('sends only the strokes to the recognizer and inserts the equation it recognized', async () => {
		vi.useFakeTimers();
		const recognizer = vi.fn<InkMathRecognizer>(async () => '\frac{1}{2}');
		const insert = vi.fn(() => true);
		const closed = vi.fn();
		openInkEquationDialog({
			host: document.body,
			recognizer,
			locale: 'de',
			heading: 'Ink to Math',
			insert,
			closed,
		});
		const dialog = document.querySelector('office-ui-dialog')!;
		expect(dialog.getAttribute('heading') ?? (dialog as { heading?: string }).heading).toBe(
			'Freihand in Formel',
		);
		const canvas = dialog.querySelector('office-ui-ink-canvas')!;
		stroke(canvas, [
			[1, 2],
			[3, 4],
		]);
		stroke(canvas, [[5, 6]]);
		// Recognized once the writing settles, with all strokes and nothing but their points.
		expect(recognizer).not.toHaveBeenCalled();
		await vi.advanceTimersByTimeAsync(500);
		expect(recognizer).toHaveBeenCalledTimes(1);
		expect(recognizer.mock.calls[0]?.[0]).toEqual([
			[
				{ x: 1, y: 2 },
				{ x: 3, y: 4 },
			],
			[{ x: 5, y: 6 }],
		]);
		const source = dialog.querySelector('textarea')!;
		expect(source.value).toBe('\frac{1}{2}');
		dialog
			.querySelector('office-ui-dialog-footer')!
			.dispatchEvent(new CustomEvent('office-dialog-footer-request', { detail: { id: 'insert' } }));
		expect(insert).toHaveBeenCalledWith(expect.objectContaining({ latex: '\frac{1}{2}' }));
	});

	it('aborts a pending recognition when the pad is cleared and says when nothing was recognized', async () => {
		vi.useFakeTimers();
		let signal: AbortSignal | undefined;
		const recognizer = vi.fn<InkMathRecognizer>(
			(_strokes, abort) =>
				new Promise((resolve) => {
					signal = abort;
					setTimeout(() => resolve(undefined), 1000);
				}),
		);
		openInkEquationDialog({
			host: document.body,
			recognizer,
			locale: 'en',
			heading: 'Ink Equation',
			insert: () => true,
		});
		const canvas = document.querySelector('office-ui-ink-canvas')!;
		stroke(canvas, [[1, 1]]);
		await vi.advanceTimersByTimeAsync(500);
		const status = document.querySelector('[role="status"]')!;
		expect(status.textContent).toBe('Recognizing...');
		document.querySelector<HTMLButtonElement>('.office-ink-equation-clear')!.click();
		expect(signal?.aborted).toBe(true);
		stroke(canvas, [[2, 2]]);
		await vi.advanceTimersByTimeAsync(1600);
		expect(status.textContent).toContain('could not be recognized');
	});
});
