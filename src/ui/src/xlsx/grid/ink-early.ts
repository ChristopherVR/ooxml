// Pen strokes drawn while the Draw tab's surface is still loading. The first stroke right after a
// pen is picked can land before the surface is over the cells: it is recorded here (the press never
// selects cells), drawn as it goes, and handed to the surface once it has loaded, in order and with
// the pen it was drawn with, so a quick second stroke joins it instead of racing it.
import type { GridInkTool } from 'ooxml-core/xlsx/ui';
import type { InkStrokeDetail } from '../../ink/ink-canvas';
import type { GridView } from './grid-view';

const PENS = new Set(['pen', 'pencil', 'highlighter']);
const SVG = 'http://www.w3.org/2000/svg';

/** A recorded stroke and the pen that drew it. */
export type EarlyStroke = InkStrokeDetail & { pen: { color: string; width: number } };

export class EarlyInk {
	readonly #view: GridView;
	readonly #strokes: EarlyStroke[] = [];
	readonly #listeners = new AbortController();
	#preview: SVGSVGElement | undefined;
	/** The loaded surface's stroke handler: a stroke finished after it loaded goes straight to it. */
	deliver: ((detail: EarlyStroke) => void) | undefined;

	constructor(view: GridView) {
		this.#view = view;
	}

	/** Records the stroke a press on the cells starts with `tool`, until the pointer is released. */
	press(event: PointerEvent, tool: GridInkTool | undefined): void {
		if (!tool || !PENS.has(tool.tool)) return;
		const el = event.currentTarget as HTMLElement;
		const start = event.timeStamp;
		const samples: InkStrokeDetail['samples'] = [];
		const line = this.#line(tool);
		const add = (sample: PointerEvent) => {
			const { x, y } = this.#view.clientToView(sample.clientX, sample.clientY);
			const pressure = sample.pointerType === 'mouse' ? 0.5 : sample.pressure || 0.5;
			samples.push({ x, y, pressure, time: Math.max(0, sample.timeStamp - start) });
			line.setAttribute('points', samples.map((point) => `${point.x},${point.y}`).join(' '));
		};
		const gesture = new AbortController();
		const signal = AbortSignal.any([gesture.signal, this.#listeners.signal]);
		const mine = (next: PointerEvent) => next.pointerId === event.pointerId;
		el.addEventListener('pointermove', (next) => mine(next) && add(next), { signal });
		el.addEventListener(
			'pointercancel',
			(next) => {
				if (!mine(next)) return;
				gesture.abort();
				line.remove();
			},
			{ signal },
		);
		el.addEventListener(
			'pointerup',
			(next) => {
				if (!mine(next)) return;
				add(next);
				gesture.abort();
				const stroke: EarlyStroke = {
					tool: tool.tool as InkStrokeDetail['tool'],
					samples,
					pen: { color: tool.color, width: tool.width },
				};
				if (this.deliver) {
					line.remove();
					this.deliver(stroke);
				} else this.#strokes.push(stroke);
			},
			{ signal },
		);
		el.setPointerCapture?.(event.pointerId);
		add(event);
	}

	/** A line drawing the stroke as it is made, until the surface takes it over. */
	#line(tool: GridInkTool): SVGPolylineElement {
		const doc = this.#view.doc;
		if (!this.#preview) {
			this.#preview = doc.createElementNS(SVG, 'svg');
			this.#preview.setAttribute('aria-hidden', 'true');
			this.#preview.style.cssText =
				'position:absolute;left:0;top:0;width:100%;height:100%;pointer-events:none;overflow:visible';
			this.#view.view.append(this.#preview);
		}
		const line = doc.createElementNS(SVG, 'polyline');
		line.setAttribute('fill', 'none');
		line.setAttribute('stroke', tool.color);
		line.setAttribute('stroke-width', String(tool.width * (this.#view.zoom() / 100)));
		line.setAttribute('stroke-linecap', 'round');
		line.setAttribute('stroke-linejoin', 'round');
		this.#preview.append(line);
		return line;
	}

	/** The strokes recorded so far, emptied (their preview goes with them). */
	take(): EarlyStroke[] {
		this.#preview?.remove();
		this.#preview = undefined;
		return this.#strokes.splice(0);
	}

	destroy(): void {
		this.#listeners.abort();
		this.take();
	}
}
