// Drawing on Word's Print Layout pages: the shared ink canvas over the pages turns pen strokes into
// floating ink (strokes drawn one after another join the same drawing, as in Word), the eraser
// removes the strokes it touches, the lasso selects strokes for Delete or Ink to Shape, and Ink
// Replay redraws the ink on the pages in the order it was drawn.
import type { InkCoordinate, InkStroke } from 'ooxml-core/drawingml';
import { inkTools } from '../ink/tools';
import type { DocxInk } from 'ooxml-core/docx';
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
import { createInkReplay, type InkReplay } from '../ink/ink-replay';
import {
	appendInk,
	applyInkEdits,
	convertInkToShapes,
	insertInk,
	paragraphStart,
	type InkEdit,
} from './ink-commands';
import { inkOf } from './ink-node-view';
import { locateFloat } from './print-float-locate';
import surfaceCss from './draw-surface.css?raw';

export interface DocxDrawHost {
	view(): EditorView | undefined;
	state: DrawState;
	/** The scrolling `.dve-canvas` that holds the pages. */
	scroller(): HTMLElement | undefined;
	pages(): HTMLElement | undefined;
	/** True while the pages can take ink (Print Layout, editable body). */
	active(): boolean;
	reportError(error: Error): void;
	refresh(): void;
}

/** Ink drawn this close in time and space joins the previous drawing, as Word groups strokes. */
const JOIN_MS = 2000;
const JOIN_PX = 48;

/** One ink drawing on a page, with the strokes the lasso holds in it. */
interface InkTarget {
	pos: number;
	ink: DocxInk;
	element: HTMLElement;
	sheet: HTMLElement;
	blockId: string;
	indexes: number[];
}

const scaleOf = (element: HTMLElement): number =>
	(element.offsetWidth ? element.getBoundingClientRect().width / element.offsetWidth : 1) || 1;

export class DocxDrawController {
	readonly #host: DocxDrawHost;
	readonly #canvas: OfficeUiInkCanvas;
	/** The ink the last stroke made, with its page (sheets are rebuilt by every layout). */
	#last: { pos: number; ink: string; page: number; at: number } | undefined;
	#lassoed: InkTarget[] = [];
	#boxes: HTMLElement[] = [];
	#replay: { replay: InkReplay; dispose: () => void } | undefined;
	readonly #listeners = new AbortController();

	constructor(host: DocxDrawHost) {
		this.#host = host;
		defineInkCanvas();
		this.#canvas = document.createElement('office-ui-ink-canvas') as OfficeUiInkCanvas;
		this.#canvas.classList.add('dve-ink-surface');
		this.#canvas.label = 'Drawing canvas';
		const { signal } = this.#listeners;
		this.#canvas.addEventListener(
			'office-ink-stroke',
			(event) => this.#stroke((event as CustomEvent<InkStrokeDetail>).detail),
			{ signal },
		);
		this.#canvas.addEventListener(
			'office-ink-erase',
			(event) => this.#erase((event as CustomEvent<InkEraseDetail>).detail),
			{ signal },
		);
		this.#canvas.addEventListener(
			'office-ink-lasso',
			(event) => this.#lasso((event as CustomEvent<InkLassoDetail>).detail),
			{ signal },
		);
		this.#canvas.addEventListener(
			'wheel',
			(event) => this.#host.scroller()?.scrollBy(event.deltaX, event.deltaY),
			{ signal, passive: true },
		);
		document.addEventListener('keydown', (event) => this.#key(event), { capture: true, signal });
	}

	/** Places the canvas over the pages and follows the Draw tab's tool and pen. */
	sync(): void {
		const scroller = this.#host.scroller();
		const tool = this.#host.active() ? this.#host.state.canvasTool : '';
		if (tool !== 'lasso') this.clearLasso();
		if (!scroller || !tool) {
			this.#canvas.tool = '';
			scroller?.classList.remove('dve-inking');
			this.#canvas.remove();
			return;
		}
		this.#style(scroller);
		scroller.classList.add('dve-inking');
		if (this.#canvas.parentElement !== scroller) scroller.append(this.#canvas);
		Object.assign(this.#canvas.style, {
			width: `${scroller.scrollWidth}px`,
			height: `${scroller.scrollHeight}px`,
		});
		const sheet = this.#host.pages()?.querySelector<HTMLElement>('.dve-print-page');
		const pen = this.#host.state.pen;
		this.#canvas.tool = tool;
		this.#canvas.color = pen.color;
		this.#canvas.width = pen.width * (sheet ? scaleOf(sheet) : 1);
		this.#canvas.touch = this.#host.state.snapshot.touch;
	}

	/** The surface's styles, added to the editor's shadow root with the Draw tab. */
	#style(scroller: HTMLElement): void {
		const node = scroller.getRootNode();
		const root = node instanceof ShadowRoot ? node : scroller.ownerDocument.head;
		if (root.querySelector('style[data-draw-surface]')) return;
		const style = document.createElement('style');
		style.dataset.drawSurface = '';
		style.textContent = surfaceCss;
		root.append(style);
	}

	/** Lasso selection or the selected ink, for Ink to Shape. */
	hasSelection(): boolean {
		return this.#lassoed.length > 0 || this.#selectedInk() !== undefined;
	}

	/** Whether any ink is on the laid-out pages, for Ink Replay. */
	hasInk(): boolean {
		return Boolean(this.#host.pages()?.querySelector('.dve-print-float[data-ink]'));
	}

	get replaying(): boolean {
		return this.#replay !== undefined;
	}

	#client(detail: { x: number; y: number }): InkCoordinate {
		const rect = this.#canvas.getBoundingClientRect();
		return { x: rect.left + detail.x, y: rect.top + detail.y };
	}

	#sheetAt(point: InkCoordinate): HTMLElement | undefined {
		return [...(this.#host.pages()?.querySelectorAll<HTMLElement>('.dve-print-page') ?? [])].find(
			(sheet) => {
				const rect = sheet.getBoundingClientRect();
				return (
					point.x >= rect.left &&
					point.x <= rect.right &&
					point.y >= rect.top &&
					point.y <= rect.bottom
				);
			},
		);
	}

	/** The body paragraph a drawing at `top` (client) on `sheet` anchors to: the line above it. */
	#anchor(view: EditorView, sheet: HTMLElement, top: number): string | undefined {
		let chosen: string | undefined;
		let first: string | undefined;
		for (const line of sheet.querySelectorAll<HTMLElement>('.dve-print-line[data-block-id]')) {
			const id = line.dataset.blockId!;
			if (paragraphStart(view.state.doc, id) === undefined) continue;
			first ??= id;
			if (line.getBoundingClientRect().top <= top) chosen = id;
		}
		return chosen ?? first;
	}

	#stroke(detail: InkStrokeDetail): void {
		const view = this.#host.view();
		const first = detail.samples[0];
		if (!view || !first || !this.#host.active()) return;
		const sheet = this.#sheetAt(this.#client(first));
		if (!sheet) return;
		const rect = sheet.getBoundingClientRect();
		const scale = scaleOf(sheet);
		const samples = detail.samples.map((sample) => {
			const client = this.#client(sample);
			return {
				x: (client.x - rect.left) / scale,
				y: (client.y - rect.top) / scale,
				pressure: sample.pressure,
				time: Date.now() - (detail.samples.at(-1)?.time ?? 0) + sample.time,
			};
		});
		const stroke = inkTools().inkStrokeFromSamples(samples, {
			...this.#host.state.pen,
			kind: detail.tool,
		});
		if (!stroke) return;
		const now = Date.now();
		const last = this.#last;
		const node = last && view.state.doc.nodeAt(last.pos);
		const bounds = inkTools().inkStrokesBounds([stroke]);
		if (
			last &&
			node &&
			bounds &&
			now - last.at < JOIN_MS &&
			last.page === this.#pageOf(sheet) &&
			node.attrs.ink === last.ink
		) {
			const placement = JSON.parse(String(node.attrs.placement)) as {
				offsetXPx?: number;
				offsetYPx?: number;
			};
			const near =
				bounds.x < (placement.offsetXPx ?? 0) + Number(node.attrs.widthPx) + JOIN_PX &&
				bounds.x + bounds.width > (placement.offsetXPx ?? 0) - JOIN_PX &&
				bounds.y < (placement.offsetYPx ?? 0) + Number(node.attrs.heightPx) + JOIN_PX &&
				bounds.y + bounds.height > (placement.offsetYPx ?? 0) - JOIN_PX;
			if (near && appendInk(view, last.pos, [stroke])) {
				this.#remember(view, last.pos, sheet, now);
				return;
			}
		}
		const blockId = this.#anchor(view, sheet, this.#client(first).y);
		const pos = blockId ? insertInk(view, [stroke], { blockId }) : undefined;
		if (pos === undefined) {
			this.#host.reportError(new Error('Ink can only be drawn on a page with text.'));
			return;
		}
		this.#remember(view, pos, sheet, now);
	}

	#remember(view: EditorView, pos: number, sheet: HTMLElement, at: number): void {
		const ink = view.state.doc.nodeAt(pos)?.attrs.ink;
		this.#last = typeof ink === 'string' ? { pos, ink, page: this.#pageOf(sheet), at } : undefined;
	}

	#pageOf(sheet: HTMLElement): number {
		return Number(sheet.dataset.pageIndex ?? -1);
	}

	/** The ink drawings on the pages whose box `test` accepts, with their nodes. */
	#targets(test: (rect: DOMRect) => boolean): InkTarget[] {
		const view = this.#host.view();
		if (!view) return [];
		return [
			...(this.#host.pages()?.querySelectorAll<HTMLElement>('.dve-print-float[data-ink]') ?? []),
		].flatMap((element) => {
			const sheet = element.closest<HTMLElement>('.dve-print-page');
			const blockId = element.dataset.blockId;
			if (!sheet || !blockId || !test(element.getBoundingClientRect())) return [];
			const located = locateFloat(view.state.doc, blockId, Number(element.dataset.floatIndex));
			const ink = located && inkOf(located.node);
			return located && ink
				? [{ pos: located.pos, ink, element, sheet, blockId, indexes: [] }]
				: [];
		});
	}

	/** A client point in an ink drawing's own coordinates (and the factor for lengths). */
	#local(target: InkTarget, point: InkCoordinate): { point: InkCoordinate; factor: number } {
		const rect = target.element.getBoundingClientRect();
		const factor = target.ink.width / (rect.width || 1);
		return {
			point: {
				x: (point.x - rect.left) * factor,
				y: (point.y - rect.top) * (target.ink.height / (rect.height || 1)),
			},
			factor,
		};
	}

	#erase(detail: InkEraseDetail): void {
		const view = this.#host.view();
		if (!view || !this.#host.active()) return;
		const point = this.#client(detail);
		const reach = detail.radius;
		const edits = this.#targets(
			(rect) =>
				point.x >= rect.left - reach &&
				point.x <= rect.right + reach &&
				point.y >= rect.top - reach &&
				point.y <= rect.bottom + reach,
		).flatMap((target) => {
			const local = this.#local(target, point);
			const kept = inkTools().eraseInkStrokes(
				target.ink.strokes,
				local.point,
				reach * local.factor,
			);
			return kept.length === target.ink.strokes.length ? [] : [{ pos: target.pos, strokes: kept }];
		});
		if (edits.length) applyInkEdits(view, edits, true);
	}

	#lasso(detail: InkLassoDetail): void {
		this.clearLasso();
		const polygon = detail.points.map((point) => this.#client(point));
		const xs = polygon.map((point) => point.x);
		const ys = polygon.map((point) => point.y);
		const box = {
			left: Math.min(...xs),
			right: Math.max(...xs),
			top: Math.min(...ys),
			bottom: Math.max(...ys),
		};
		this.#lassoed = this.#targets(
			(rect) =>
				rect.left <= box.right &&
				rect.right >= box.left &&
				rect.top <= box.bottom &&
				rect.bottom >= box.top,
		).flatMap((target) => {
			const indexes = inkTools().lassoInkStrokes(
				target.ink.strokes,
				polygon.map((point) => this.#local(target, point).point),
			);
			return indexes.length ? [{ ...target, indexes }] : [];
		});
		for (const target of this.#lassoed) this.#paintBox(target);
		this.#host.refresh();
	}

	#paintBox(target: InkTarget): void {
		const bounds = inkTools().inkStrokesBounds(
			target.indexes.map((index) => target.ink.strokes[index]!),
		);
		if (!bounds) return;
		const rect = target.element.getBoundingClientRect();
		const sheetRect = target.sheet.getBoundingClientRect();
		const scale = scaleOf(target.sheet);
		const kx = rect.width / (target.ink.width || 1) / scale;
		const ky = rect.height / (target.ink.height || 1) / scale;
		const frame = document.createElement('div');
		frame.className = 'dve-ink-lasso-box';
		frame.setAttribute('aria-hidden', 'true');
		Object.assign(frame.style, {
			left: `${(rect.left - sheetRect.left) / scale + bounds.x * kx}px`,
			top: `${(rect.top - sheetRect.top) / scale + bounds.y * ky}px`,
			width: `${bounds.width * kx}px`,
			height: `${bounds.height * ky}px`,
		});
		target.sheet.append(frame);
		this.#boxes.push(frame);
	}

	clearLasso(): void {
		for (const box of this.#boxes.splice(0)) box.remove();
		if (this.#lassoed.length) {
			this.#lassoed = [];
			this.#host.refresh();
		}
	}

	#key(event: KeyboardEvent): void {
		if (event.key === 'Escape' && this.#replay) {
			this.stopReplay();
			event.preventDefault();
			return;
		}
		if ((event.key !== 'Delete' && event.key !== 'Backspace') || !this.#lassoed.length) return;
		const view = this.#host.view();
		if (!view) return;
		event.preventDefault();
		event.stopPropagation();
		applyInkEdits(
			view,
			this.#lassoed.map((target) => ({
				pos: target.pos,
				strokes: target.ink.strokes.filter((_, index) => !target.indexes.includes(index)),
			})),
		);
		this.clearLasso();
	}

	/** The selected ink drawing (a node selection), as a target holding all its strokes. */
	#selectedInk(): InkTarget | undefined {
		const view = this.#host.view();
		const selection = view?.state.selection;
		if (!view || !(selection instanceof NodeSelection)) return undefined;
		const ink = inkOf(selection.node);
		if (!ink) return undefined;
		const target = this.#targets(() => true).find((item) => item.pos === selection.from);
		return target && { ...target, indexes: ink.strokes.map((_, index) => index) };
	}

	/** Ink to Shape: each recognised selection becomes a shape outlined in the ink's colour. */
	async inkToShape(): Promise<void> {
		const view = this.#host.view();
		const targets = this.#lassoed.length
			? this.#lassoed
			: [this.#selectedInk()].filter((target) => target !== undefined);
		if (!view || !targets.length) return;
		const { recognizeInkShape, inkShapeOf } = inkTools();
		if (this.#host.view() !== view) return;
		const edits: InkEdit[] = [];
		const shapes: Parameters<typeof convertInkToShapes>[2][number][] = [];
		for (const target of targets) {
			const rect = target.element.getBoundingClientRect();
			const sheetRect = target.sheet.getBoundingClientRect();
			const scale = scaleOf(target.sheet);
			const kx = rect.width / (target.ink.width || 1) / scale;
			const ky = rect.height / (target.ink.height || 1) / scale;
			const ox = (rect.left - sheetRect.left) / scale;
			const oy = (rect.top - sheetRect.top) / scale;
			const strokes = target.indexes.map((index) => target.ink.strokes[index]!);
			const paths = strokes.map((stroke) =>
				inkTools()
					.inkPathPoints(stroke.path)
					.map((point) => ({ x: ox + point.x * kx, y: oy + point.y * ky })),
			);
			const recognized = recognizeInkShape(paths);
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
				new Error('Ink to Shape did not recognize a shape in the selected ink.'),
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
		const origin = scroller.getBoundingClientRect();
		const strokes = this.#targets(() => true).flatMap((target) => {
			const rect = target.element.getBoundingClientRect();
			const kx = rect.width / (target.ink.width || 1);
			const ky = rect.height / (target.ink.height || 1);
			return target.ink.strokes.map((stroke) => ({
				points: inkTools()
					.inkPathPoints(stroke.path)
					.map((point) => ({
						x: rect.left - origin.left + scroller.scrollLeft + point.x * kx,
						y: rect.top - origin.top + scroller.scrollTop + point.y * ky,
					})),
				style: { ...stroke, width: stroke.width * kx },
			}));
		});
		const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
		svg.classList.add('dve-ink-replay');
		svg.setAttribute('aria-hidden', 'true');
		svg.setAttribute('width', String(scroller.scrollWidth));
		svg.setAttribute('height', String(scroller.scrollHeight));
		this.#style(scroller);
		scroller.classList.add('dve-inking', 'dve-ink-replaying');
		scroller.append(svg);
		const dispose = () => {
			scroller.classList.remove('dve-ink-replaying');
			if (!this.#canvas.isConnected) scroller.classList.remove('dve-inking');
		};
		const replay = createInkReplay(
			svg,
			strokes,
			(time, duration, playing) => {
				if (!playing && time >= duration && this.#replay) setTimeout(() => this.stopReplay(), 400);
			},
			inkTools(),
		);
		if (!replay) {
			svg.remove();
			dispose();
			return;
		}
		this.#replay = { replay, dispose };
		replay.toggle();
		this.#host.refresh();
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
		this.#last = undefined;
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
