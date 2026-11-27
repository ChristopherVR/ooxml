// Drawing over Excel's cells: the shared ink canvas over the grid turns pen strokes into ink anchored
// to the cells (strokes drawn one after another join the same drawing, as in Excel), the eraser
// removes the strokes it touches (once a frame, one undo step a gesture), the lasso selects strokes
// for Delete or Ink to Shape, and Ink Replay redraws the sheet's ink in the order it was drawn.
// Coordinates of ink are pixels at 100%.
import type { InkCoordinate, InkStroke } from 'ooxml-core/drawingml';
import type { GridInk, GridInkTool } from 'ooxml-core/xlsx/ui';
import {
	defineInkCanvas,
	type InkEraseDetail,
	type InkLassoDetail,
	type InkStrokeDetail,
	type OfficeUiInkCanvas,
} from '../../ink/ink-canvas';
import type { InkReplay } from '../../ink/ink-replay';
import { inkTools } from '../../ink/tools';
import type { GridView } from './grid-view';
import { startInkReplay } from './ink-replay-run';
import { InkSelection } from './ink-selection';
import { anchorOf, fittedInk, sheetInks, zoomOf } from './ink-sheet';

/** Whether a key goes to something that takes Delete itself (a box, a list, a dialog). */
function keepsOwnDelete(event: KeyboardEvent): boolean {
	const path = event.composedPath();
	const target = path[0];
	if (
		target instanceof HTMLInputElement ||
		target instanceof HTMLTextAreaElement ||
		target instanceof HTMLSelectElement ||
		(target instanceof HTMLElement && target.isContentEditable)
	)
		return true;
	return path.some((node) => node instanceof HTMLElement && node.matches('[role="dialog"]'));
}

export class InkLayer implements GridInk {
	readonly #view: GridView;
	readonly #canvas: OfficeUiInkCanvas;
	readonly #selection: InkSelection;
	#tool: GridInkTool | undefined;
	#last: { sheet: number; index: number; ink: object; at: number } | undefined;
	#replay: InkReplay | undefined;
	/** Eraser points not yet applied, and the gesture they belong to (one undo step). */
	#erase: { points: InkEraseDetail[]; frame: number | undefined; gesture: string | undefined } = {
		points: [],
		frame: undefined,
		gesture: undefined,
	};
	#gestures = 0;
	readonly #listeners = new AbortController();

	constructor(view: GridView) {
		this.#view = view;
		this.#selection = new InkSelection(view);
		defineInkCanvas(view.doc.defaultView?.customElements);
		this.#canvas = view.doc.createElement('office-ui-ink-canvas') as OfficeUiInkCanvas;
		this.#canvas.classList.add('xg-ink-surface');
		const { signal } = this.#listeners;
		const on = <T>(type: string, run: (detail: T) => void) =>
			this.#canvas.addEventListener(type, (event) => run((event as CustomEvent<T>).detail), {
				signal,
			});
		on<InkStrokeDetail>('office-ink-stroke', (detail) => this.stroke(detail));
		on<InkEraseDetail>('office-ink-erase', (detail) => this.#erasing(detail));
		on<InkLassoDetail>('office-ink-lasso', (detail) =>
			this.#selection.select(
				detail.points.flatMap((point) => {
					const at = this.#sheetPoint(point);
					return at ? [at] : [];
				}),
			),
		);
		this.#canvas.addEventListener(
			'wheel',
			(event) => this.#view.scroller.scrollBy(event.deltaX, event.deltaY),
			{ signal, passive: true },
		);
		view.doc.addEventListener('keydown', (event) => this.#key(event), { capture: true, signal });
	}

	setTool(tool: GridInkTool | undefined): void {
		this.#tool = tool;
		if (tool?.tool !== 'lasso') this.#selection.clear();
		const canvas = this.#canvas;
		if (!tool || this.#view.ctx.readOnly()) {
			canvas.tool = '';
			canvas.remove();
			return;
		}
		if (canvas.parentElement !== this.#view.view) this.#view.view.append(canvas);
		canvas.label = this.#view.ctx.t('Drawing canvas');
		canvas.tool = tool.tool;
		canvas.color = tool.color;
		canvas.width = tool.width * zoomOf(this.#view);
		canvas.touch = tool.touch;
	}

	/** A canvas point (view coordinates) on the sheet at 100%, or undefined over the headers. */
	#sheetPoint(point: { x: number; y: number }): InkCoordinate | undefined {
		const g = this.#view.geometry;
		if (g.hit(point.x, point.y).area !== 'cell') return undefined;
		const plane = g.planeAt(point.x, point.y);
		const f = zoomOf(this.#view);
		return { x: plane.x / f, y: plane.y / f };
	}

	/** Adds a pen stroke (canvas coordinates), drawn on the canvas or before it had loaded. */
	stroke(detail: InkStrokeDetail, pen?: { color: string; width: number }): void {
		const tool = this.#tool;
		if (!this.#view.ctx.session() || !tool || !this.#view.sheet() || this.#view.ctx.readOnly())
			return;
		const end = detail.samples.at(-1)?.time ?? 0;
		const now = Date.now();
		const samples = detail.samples.flatMap((sample) => {
			const point = this.#sheetPoint(sample);
			return point ? [{ ...point, pressure: sample.pressure, time: now - end + sample.time }] : [];
		});
		const tools = inkTools();
		const stroke = tools.inkStrokeFromSamples(samples, {
			kind: detail.tool,
			color: pen?.color ?? tool.color,
			width: pen?.width ?? tool.width,
		});
		const bounds = stroke && tools.inkStrokesBounds([stroke]);
		if (!stroke || !bounds) return;
		const sheetIndex = this.#view.sheetIndex();
		const last = this.#last;
		const target = last && sheetInks(this.#view).find((entry) => entry.index === last.index);
		if (
			last &&
			target &&
			last.sheet === sheetIndex &&
			target.ink === last.ink &&
			tools.inkJoins(target.box, bounds, now - last.at) &&
			// A drawing collects strokes up to a size its part is always read back whole at.
			tools.canJoinInk(target.ink.strokes, [stroke])
		) {
			const { ink, box } = target;
			const kept = tools.transformInkStrokes(
				ink.strokes,
				0,
				0,
				box.width / (ink.width || 1),
				box.height / (ink.height || 1),
			);
			this.#place(
				sheetIndex,
				[...tools.transformInkStrokes(kept, box.x, box.y), stroke],
				target.index,
			);
			return;
		}
		this.#place(sheetIndex, [stroke]);
	}

	/** Puts strokes (sheet coordinates at 100%) in a new ink drawing, or in the one at `index`. */
	#place(sheetIndex: number, strokes: InkStroke[], index?: number): void {
		const session = this.#view.ctx.session();
		const bounds = inkTools().inkStrokesBounds(strokes);
		if (!session || !bounds) return;
		const box = {
			x: Math.max(0, bounds.x),
			y: Math.max(0, bounds.y),
			width: bounds.width,
			height: bounds.height,
		};
		const ink = {
			strokes: inkTools().transformInkStrokes(strokes, -box.x, -box.y),
			width: box.width,
			height: box.height,
		};
		const anchor = anchorOf(this.#view, box);
		const placed =
			index === undefined
				? session.addInk(sheetIndex, anchor, { ...ink, name: this.#newName() })
				: (session.setInk(sheetIndex, index, anchor, ink), index);
		const drawing = this.#view.sheet()?.drawings[placed];
		this.#last =
			drawing?.kind === 'ink'
				? { sheet: sheetIndex, index: placed, ink: drawing, at: Date.now() }
				: undefined;
	}

	/** A new drawing's name, as Excel numbers them: Ink 1, Ink 2 (in the editor's language). */
	#newName(): string {
		const word = this.#view.ctx.t('Ink');
		const numbers = (this.#view.sheet()?.drawings ?? []).flatMap((drawing) => {
			const match =
				drawing.kind === 'ink' &&
				drawing.name?.startsWith(`${word} `) &&
				/\d+$/u.exec(drawing.name);
			return match ? [Number(match[0])] : [];
		});
		return `${word} ${Math.max(0, ...numbers) + 1}`;
	}

	/** Eraser moves are applied once a frame; the gesture is one undo step. */
	#erasing(detail: InkEraseDetail): void {
		const erase = this.#erase;
		erase.gesture ??= `ink-erase-${++this.#gestures}`;
		erase.points.push(detail);
		if (!detail.done) {
			erase.frame ??= this.#view.doc.defaultView?.requestAnimationFrame(() => {
				erase.frame = undefined;
				this.#flushErase();
			});
			return;
		}
		if (erase.frame !== undefined) this.#view.doc.defaultView?.cancelAnimationFrame(erase.frame);
		erase.frame = undefined;
		this.#flushErase();
		erase.gesture = undefined;
	}

	#flushErase(): void {
		const session = this.#view.ctx.session();
		const points = this.#erase.points.splice(0);
		const gesture = this.#erase.gesture;
		if (!session || this.#view.ctx.readOnly()) return;
		const sheetIndex = this.#view.sheetIndex();
		for (const detail of points) {
			const point = this.#sheetPoint(detail);
			if (!point) continue;
			const reach = detail.radius / zoomOf(this.#view);
			// Last first, so deleting an emptied drawing keeps the indexes before it.
			for (const entry of sheetInks(this.#view).reverse()) {
				const { ink, box } = entry;
				if (!inkTools().inkBoxReaches(box, point, reach)) continue;
				const local = inkTools().inkLocalPoint(box, ink, point);
				const kept = inkTools().eraseInkStrokes(ink.strokes, local.point, reach * local.factor);
				if (kept.length === ink.strokes.length) continue;
				const next = fittedInk(this.#view, entry, kept);
				session.setInk(sheetIndex, entry.index, next.anchor, next.ink, gesture);
			}
		}
	}

	#key(event: KeyboardEvent): void {
		if (event.key === 'Escape' && this.#replay) {
			this.#stopReplay();
			event.preventDefault();
			return;
		}
		if (event.key !== 'Delete' && event.key !== 'Backspace') return;
		if (!this.#selection.size || keepsOwnDelete(event)) return;
		event.preventDefault();
		event.stopPropagation();
		this.#selection.remove(this.#selection.held());
		this.#selection.clear();
	}

	hasSelection(): boolean {
		return this.#selection.held().length > 0;
	}

	async inkToShape(): Promise<boolean> {
		return this.#selection.inkToShape();
	}

	replaying(): boolean {
		return this.#replay !== undefined;
	}

	replay(): void {
		if (this.#replay) {
			this.#stopReplay();
			return;
		}
		this.#replay = startInkReplay(this.#view, () => this.#replay && this.#stopReplay());
		if (this.#replay) this.#view.ctx.requestRender();
	}

	#stopReplay(): void {
		const replay = this.#replay;
		if (!replay) return;
		this.#replay = undefined;
		replay.dispose();
		this.#view.root.classList.remove('xg-ink-replaying');
		this.#view.ctx.requestRender();
	}

	destroy(): void {
		this.#stopReplay();
		this.#selection.clear();
		if (this.#erase.frame !== undefined)
			this.#view.doc.defaultView?.cancelAnimationFrame(this.#erase.frame);
		this.#listeners.abort();
		this.#canvas.remove();
	}
}
