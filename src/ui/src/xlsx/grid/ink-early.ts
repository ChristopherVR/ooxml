// Pen strokes drawn while the Draw tab's surface is still loading. The first stroke right after a
// pen is picked can land before the surface is over the cells: it is recorded here (the press never
// selects cells) and handed to the surface once it has loaded, in order, so a quick second stroke
// joins it instead of racing it.
import type { GridInkTool } from 'ooxml-core/xlsx/ui';
import type { InkStrokeDetail } from '../../ink/ink-canvas';
import type { GridView } from './grid-view';

const PENS = new Set(['pen', 'pencil', 'highlighter']);

export class EarlyInk {
	readonly #view: GridView;
	readonly #strokes: InkStrokeDetail[] = [];
	/** The loaded surface's stroke handler: a stroke finished after it loaded goes straight to it. */
	deliver: ((detail: InkStrokeDetail) => void) | undefined;

	constructor(view: GridView) {
		this.#view = view;
	}

	/** Records the stroke a press on the cells starts with `tool`, until the pointer is released. */
	press(event: PointerEvent, tool: GridInkTool | undefined): void {
		if (!tool || !PENS.has(tool.tool)) return;
		const el = event.currentTarget as HTMLElement;
		const start = event.timeStamp;
		const samples: InkStrokeDetail['samples'] = [];
		const add = (sample: PointerEvent) => {
			const { x, y } = this.#view.clientToView(sample.clientX, sample.clientY);
			const pressure = sample.pointerType === 'mouse' ? 0.5 : sample.pressure || 0.5;
			samples.push({ x, y, pressure, time: Math.max(0, sample.timeStamp - start) });
		};
		const move = (next: PointerEvent) => {
			if (next.pointerId === event.pointerId) add(next);
		};
		const stop = () => {
			el.removeEventListener('pointermove', move);
			el.removeEventListener('pointerup', up);
			el.removeEventListener('pointercancel', cancel);
		};
		const cancel = (next: PointerEvent) => {
			if (next.pointerId === event.pointerId) stop();
		};
		const up = (next: PointerEvent) => {
			if (next.pointerId !== event.pointerId) return;
			add(next);
			stop();
			const stroke = { tool: tool.tool as InkStrokeDetail['tool'], samples };
			if (this.deliver) this.deliver(stroke);
			else this.#strokes.push(stroke);
		};
		el.setPointerCapture?.(event.pointerId);
		add(event);
		el.addEventListener('pointermove', move);
		el.addEventListener('pointerup', up);
		el.addEventListener('pointercancel', cancel);
	}

	/** The strokes recorded so far, emptied. */
	take(): InkStrokeDetail[] {
		return this.#strokes.splice(0);
	}
}
