// Drawing on Word's Print Layout pages: the shared ink canvas over the pages turns pen strokes into
// floating ink (strokes drawn one after another join the same drawing, as in Word), the eraser
// removes the strokes it touches (one undo step a gesture), the lasso selects strokes for Delete or
// Ink to Shape, and Ink Replay redraws the ink on the pages in the order it was drawn. Floating ink
// and ink in line with text both take the eraser, the lasso and Ink to Shape.
import type { InkCoordinate } from 'ooxml-core/drawingml';
import { NodeSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { DrawState } from '../ink/draw-state';
import {
	defineInkCanvas,
	type InkEraseDetail,
	type InkLassoDetail,
	type InkStrokeDetail,
	type OfficeUiInkCanvas,
} from '../ink/ink-canvas';
import { inkTools } from '../ink/tools';
import { InkEraser } from './draw-eraser';
import { InkLasso } from './draw-lasso';
import { InkStrokes } from './draw-stroke';
import { injectDrawStyles, startInkReplay, type RunningReplay } from './draw-replay-run';
import { inkTargets, pageMapping, scaleOf, type InkTarget } from './draw-targets';
import { applyInkEdits, convertInkToShapes, type InkEdit } from './ink-commands';
import { drawHelpers } from './draw-helpers';
import { translate, type EditorLocale, type LocalizationKey } from './localization';
import { editsOwnText } from './ribbon-history-keys';

export interface DocxDrawHost {
	view(): EditorView | undefined;
	state: DrawState;
	/** The scrolling `.dve-canvas` that holds the pages. */
	scroller(): HTMLElement | undefined;
	pages(): HTMLElement | undefined;
	/** True while the pages can take ink (Print Layout, editable body). */
	active(): boolean;
	locale(): EditorLocale;
	reportError(error: Error): void;
	refresh(): void;
}

export class DocxDrawController {
	readonly #host: DocxDrawHost;
	readonly #canvas: OfficeUiInkCanvas;
	readonly #strokes: InkStrokes;
	readonly #eraser: InkEraser;
	readonly #lasso = new InkLasso();
	#replay: RunningReplay | undefined;
	readonly #listeners = new AbortController();

	constructor(host: DocxDrawHost) {
		this.#host = host;
		const client = (detail: { x: number; y: number }) => this.#client(detail);
		this.#strokes = new InkStrokes({
			view: () => host.view(),
			pages: () => host.pages(),
			pen: () => host.state.pen,
			client,
			name: () => this.#t('Ink'),
		});
		this.#eraser = new InkEraser({
			view: () => host.view(),
			active: () => host.active(),
			client,
			targets: (test) => this.#targets(test),
		});
		defineInkCanvas();
		this.#canvas = document.createElement('office-ui-ink-canvas') as OfficeUiInkCanvas;
		this.#canvas.classList.add('dve-ink-surface');
		const { signal } = this.#listeners;
		const on = <T>(type: string, run: (detail: T) => void) =>
			this.#canvas.addEventListener(type, (event) => run((event as CustomEvent<T>).detail), {
				signal,
			});
		on<InkStrokeDetail>('office-ink-stroke', (detail) => this.#stroke(detail));
		on<InkEraseDetail>('office-ink-erase', (detail) => this.#eraser.push(detail));
		on<InkLassoDetail>('office-ink-lasso', (detail) => this.#select(detail));
		this.#canvas.addEventListener(
			'wheel',
			(event) => this.#host.scroller()?.scrollBy(event.deltaX, event.deltaY),
			{ signal, passive: true },
		);
		document.addEventListener('keydown', (event) => this.#key(event), { capture: true, signal });
	}

	#t(text: string): string {
		return translate(this.#host.locale(), text as LocalizationKey);
	}

	/** Places the canvas over the pages and follows the Draw tab's tool and pen. */
	sync(): void {
		const scroller = this.#host.scroller();
		const tool = this.#host.active() ? this.#host.state.canvasTool : '';
		// The lasso holds strokes of one document: an edit or an undo drops it.
		if (tool !== 'lasso') this.clearLasso();
		else if (this.#lasso.follow(this.#host.view()?.state.doc)) this.#host.refresh();
		if (!scroller || !tool) {
			this.#canvas.tool = '';
			scroller?.classList.remove('dve-inking');
			this.#canvas.remove();
			return;
		}
		injectDrawStyles(scroller);
		// While drawing there is no caret, as in Word: the pages must not scroll to it as ink is added.
		const view = this.#host.view();
		if (view?.hasFocus()) (view.dom as HTMLElement).blur();
		scroller.classList.add('dve-inking');
		if (this.#canvas.parentElement !== scroller) scroller.append(this.#canvas);
		Object.assign(this.#canvas.style, {
			width: `${scroller.scrollWidth}px`,
			height: `${scroller.scrollHeight}px`,
		});
		const sheet = this.#host.pages()?.querySelector<HTMLElement>('.dve-print-page');
		const pen = this.#host.state.pen;
		this.#canvas.label = this.#t('Drawing canvas');
		this.#canvas.tool = tool;
		this.#canvas.color = pen.color;
		this.#canvas.width = pen.width * (sheet ? scaleOf(sheet) : 1);
		this.#canvas.touch = this.#host.state.snapshot.touch;
	}

	/** Lasso selection or the selected ink, for Ink to Shape. */
	hasSelection(): boolean {
		return this.#lasso.size > 0 || this.#selectedInk() !== undefined;
	}

	/** Whether any ink is on the laid-out pages, for Ink Replay. */
	hasInk(): boolean {
		return Boolean(this.#host.pages()?.querySelector('.dve-print-page [data-ink]'));
	}

	get replaying(): boolean {
		return this.#replay !== undefined;
	}

	#targets(test: (rect: DOMRect) => boolean, editable = true): InkTarget[] {
		const view = this.#host.view();
		const pages = this.#host.pages();
		if (!view || !pages) return [];
		return inkTargets(view, pages, test, editable);
	}

	#client(detail: { x: number; y: number }): InkCoordinate {
		const rect = this.#canvas.getBoundingClientRect();
		return { x: rect.left + detail.x, y: rect.top + detail.y };
	}

	#stroke(detail: InkStrokeDetail): void {
		if (!this.#host.active() || this.#strokes.add(detail)) return;
		this.#host.reportError(new Error(this.#t('Ink can only be drawn on a page with text.')));
	}

	#select(detail: InkLassoDetail): void {
		const view = this.#host.view();
		if (!view) return;
		const polygon = detail.points.map((point) => this.#client(point));
		const xs = polygon.map((point) => point.x);
		const ys = polygon.map((point) => point.y);
		const [left, right, top, bottom] = [
			Math.min(...xs),
			Math.max(...xs),
			Math.min(...ys),
			Math.max(...ys),
		];
		const candidates = this.#targets(
			(rect) =>
				rect.left <= right && rect.right >= left && rect.top <= bottom && rect.bottom >= top,
		);
		this.#lasso.select(view, candidates, polygon);
		this.#host.refresh();
	}

	clearLasso(): void {
		if (this.#lasso.clear()) this.#host.refresh();
	}

	#key(event: KeyboardEvent): void {
		if (event.key === 'Escape' && this.#replay) {
			this.stopReplay();
			event.preventDefault();
			return;
		}
		if (event.key !== 'Delete' && event.key !== 'Backspace') return;
		const view = this.#host.view();
		if (!view || !this.#lasso.size || this.#keepsOwnDelete(view, event)) return;
		const targets = this.#lasso.targets(view);
		this.clearLasso();
		if (!targets.length) return;
		event.preventDefault();
		event.stopPropagation();
		applyInkEdits(
			view,
			targets.map((target) => ({
				pos: target.pos,
				strokes: target.ink.strokes.filter((_, index) => !target.indexes.includes(index)),
			})),
		);
	}

	/** A box, list or dialog (anything but the body) takes Delete itself. */
	#keepsOwnDelete(view: EditorView, event: KeyboardEvent): boolean {
		const path = event.composedPath();
		const target = path[0];
		if (editsOwnText(event) && !(target instanceof Node && view.dom.contains(target))) return true;
		return path.some((node) => node instanceof HTMLElement && node.matches('[role="dialog"]'));
	}

	/** The selected ink drawing (a node selection), as a target holding all its strokes. */
	#selectedInk(): InkTarget | undefined {
		const view = this.#host.view();
		const selection = view?.state.selection;
		if (!view || !(selection instanceof NodeSelection)) return undefined;
		const ink = drawHelpers().inkOf(selection.node);
		if (!ink || ink.truncated) return undefined;
		const target = this.#targets(() => true).find((item) => item.pos === selection.from);
		return target && { ...target, indexes: ink.strokes.map((_, index) => index) };
	}

	/** Ink to Shape: each recognised selection becomes a shape outlined in the ink's colour. */
	inkToShape(): void {
		const view = this.#host.view();
		if (!view) return;
		const held = this.#lasso.targets(view);
		const targets = held.length
			? held
			: [this.#selectedInk()].filter((target) => target !== undefined);
		if (!targets.length) return;
		const { recognizeInkShape, inkShapeOf, inkPathPoints } = inkTools();
		const edits: InkEdit[] = [];
		const shapes: Parameters<typeof convertInkToShapes>[2][number][] = [];
		for (const target of targets) {
			const { ox, oy, kx, ky } = pageMapping(target);
			const strokes = target.indexes.map((index) => target.ink.strokes[index]!);
			const recognized = recognizeInkShape(
				strokes.map((stroke) =>
					inkPathPoints(stroke.path).map((point) => ({
						x: ox + point.x * kx,
						y: oy + point.y * ky,
					})),
				),
			);
			if (!recognized || !strokes[0]) continue;
			shapes.push({
				shape: inkShapeOf(recognized, strokes[0].color, strokes[0].width),
				blockId: target.blockId,
			});
			edits.push({
				pos: target.pos,
				strokes: target.ink.strokes.filter((_, index) => !target.indexes.includes(index)),
			});
		}
		const converted = convertInkToShapes(view, edits, shapes);
		this.clearLasso();
		if (!converted)
			this.#host.reportError(
				new Error(this.#t('Ink to Shape did not recognize a shape in the selected ink.')),
			);
	}

	/** Ink Replay: redraws the ink on the pages in the order it was drawn; again stops it. */
	replay(): void {
		if (this.#replay) {
			this.stopReplay();
			return;
		}
		const scroller = this.#host.scroller();
		if (!scroller) return;
		this.#replay = startInkReplay(
			scroller,
			this.#targets(() => true, false),
			() => this.#replay && this.stopReplay(),
			() => this.#canvas.isConnected,
		);
		if (this.#replay) this.#host.refresh();
	}

	stopReplay(): void {
		const current = this.#replay;
		if (!current) return;
		this.#replay = undefined;
		current.replay.dispose();
		current.dispose();
		this.#host.refresh();
	}

	reset(): void {
		this.#strokes.forget();
		this.#eraser.cancel();
		this.clearLasso();
		this.stopReplay();
	}

	destroy(): void {
		this.reset();
		this.#listeners.abort();
		this.#canvas.remove();
		this.#host.scroller()?.classList.remove('dve-inking');
	}
}
