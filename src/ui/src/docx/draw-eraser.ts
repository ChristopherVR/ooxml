// The eraser on Word's pages: its moves are applied once a frame (however fast the pointer
// reports them), each pass removes the strokes the eraser touched, and the whole gesture is one
// undo step. While the pointer moves the pages still show the boxes it erases against, so the
// boxes are fitted to the strokes left only when the gesture ends.
import type { InkCoordinate } from 'ooxml-core/drawingml';
import type { EditorView } from 'prosemirror-view';
import type { InkEraseDetail } from '../ink/ink-canvas';
import { inkTools } from '../ink/tools';
import { localPoint, type InkTarget } from './draw-targets';
import { applyInkEdits, fitInk, type InkEdit } from './ink-commands';

export interface EraserHost {
	view(): EditorView | undefined;
	active(): boolean;
	client(detail: { x: number; y: number }): InkCoordinate;
	targets(test: (rect: DOMRect) => boolean): InkTarget[];
}

export class InkEraser {
	readonly #host: EraserHost;
	#points: InkEraseDetail[] = [];
	#frame: number | undefined;
	/** The drawings this gesture changed, and when it last changed them. */
	#touched = new Set<number>();
	#time: number | undefined;

	constructor(host: EraserHost) {
		this.#host = host;
	}

	push(detail: InkEraseDetail): void {
		this.#points.push(detail);
		if (!detail.done) {
			this.#frame ??= requestAnimationFrame(() => {
				this.#frame = undefined;
				this.#flush();
			});
			return;
		}
		this.cancel();
		this.#flush();
		const view = this.#host.view();
		if (view && this.#time !== undefined) fitInk(view, this.#touched, this.#time);
		this.#touched.clear();
		this.#time = undefined;
	}

	cancel(): void {
		if (this.#frame !== undefined) cancelAnimationFrame(this.#frame);
		this.#frame = undefined;
	}

	#flush(): void {
		const view = this.#host.view();
		const points = this.#points.splice(0);
		if (!view || !points.length || !this.#host.active()) return;
		const edits = new Map<number, InkEdit>();
		for (const detail of points) {
			const point = this.#host.client(detail);
			const reach = detail.radius;
			for (const target of this.#host.targets(
				(rect) =>
					point.x >= rect.left - reach &&
					point.x <= rect.right + reach &&
					point.y >= rect.top - reach &&
					point.y <= rect.bottom + reach,
			)) {
				const strokes = edits.get(target.pos)?.strokes ?? target.ink.strokes;
				const local = localPoint(target, point);
				const kept = inkTools().eraseInkStrokes(strokes, local.point, reach * local.factor);
				if (kept.length !== strokes.length)
					edits.set(target.pos, { pos: target.pos, strokes: kept });
			}
		}
		if (!edits.size) return;
		const tr = applyInkEdits(view, [...edits.values()], {
			group: this.#time !== undefined,
			fit: false,
		});
		if (!tr) return;
		this.#time = tr.time;
		this.#touched = new Set(
			// A drawing the eraser deleted maps to a position fitInk passes over (no ink there).
			[...this.#touched, ...edits.keys()].map((pos) => tr.mapping.map(pos, -1)),
		);
	}
}
