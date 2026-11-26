// Drawing over Excel's cells: the shared ink canvas over the grid turns pen strokes into ink anchored
// to the cells (strokes drawn one after another join the same drawing, as in Excel), the eraser
// removes the strokes it touches, the lasso selects strokes for Delete or Ink to Shape, and Ink
// Replay redraws the sheet's ink in the order it was drawn. Coordinates of ink are pixels at 100%.
import type { InkBox, InkCoordinate, InkStroke } from 'ooxml-core/drawingml';
import { inkTools } from '../../ink/tools';
import { anchorToPixelBox, pixelBoxToAnchor, type InkObject } from 'ooxml-core/xlsx';
import type { GridInk, GridInkTool } from 'ooxml-core/xlsx/ui';
import {
	defineInkCanvas,
	type InkEraseDetail,
	type InkLassoDetail,
	type InkStrokeDetail,
	type OfficeUiInkCanvas,
} from '../../ink/ink-canvas';
import { createInkReplay, type InkReplay } from '../../ink/ink-replay';
import type { GridView } from './grid-view';

const JOIN_MS = 2000;
const JOIN_PX = 48;

interface Lassoed {
	index: number;
	ink: InkObject;
	indexes: number[];
}

export class InkLayer implements GridInk {
	readonly #view: GridView;
	readonly #canvas: OfficeUiInkCanvas;
	#tool: GridInkTool | undefined;
	#last: { sheet: number; index: number; ink: InkObject; at: number } | undefined;
	#lassoed: Lassoed[] = [];
	#boxes: HTMLElement[] = [];
	#replay: InkReplay | undefined;
	readonly #listeners = new AbortController();

	constructor(view: GridView) {
		this.#view = view;
		defineInkCanvas(view.doc.defaultView?.customElements);
		this.#canvas = view.doc.createElement('office-ui-ink-canvas') as OfficeUiInkCanvas;
		this.#canvas.classList.add('xg-ink-surface');
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
			(event) => this.#view.scroller.scrollBy(event.deltaX, event.deltaY),
			{ signal, passive: true },
		);
		view.doc.addEventListener('keydown', (event) => this.#key(event), { capture: true, signal });
	}

	/** Zoom as a factor: ink is kept at 100%, the grid's plane is drawn at the sheet's zoom. */
	#factor(): number {
		return this.#view.zoom() / 100;
	}

	setTool(tool: GridInkTool | undefined): void {
		this.#tool = tool;
		if (tool?.tool !== 'lasso') this.#clearLasso();
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
		canvas.width = tool.width * this.#factor();
		canvas.touch = tool.touch;
	}

	/** A canvas point (view coordinates) on the sheet at 100%, or undefined over the headers. */
	#sheetPoint(point: { x: number; y: number }): InkCoordinate | undefined {
		const g = this.#view.geometry;
		if (g.hit(point.x, point.y).area !== 'cell') return undefined;
		const plane = g.planeAt(point.x, point.y);
		const f = this.#factor();
		return { x: plane.x / f, y: plane.y / f };
	}

	/** Every ink drawing of the shown sheet with its box at 100%. */
	#inks(): { index: number; ink: InkObject; box: InkBox }[] {
		const sheet = this.#view.sheet();
		if (!sheet) return [];
		const f = this.#factor();
		return sheet.drawings.flatMap((drawing, index) => {
			if (drawing.kind !== 'ink') return [];
			const box = anchorToPixelBox(sheet, drawing.anchor, this.#view.metrics);
			return [
				{
					index,
					ink: drawing,
					box: { x: box.x / f, y: box.y / f, width: box.w / f, height: box.h / f },
				},
			];
		});
	}

	#anchor(box: InkBox) {
		const sheet = this.#view.sheet()!;
		const f = this.#factor();
		return pixelBoxToAnchor(
			sheet,
			{ x: box.x * f, y: box.y * f, w: box.width * f, h: box.height * f },
			this.#view.metrics,
		);
	}

	#stroke(detail: InkStrokeDetail): void {
		const session = this.#view.ctx.session();
		const tool = this.#tool;
		const sheet = this.#view.sheet();
		if (!session || !tool || !sheet || this.#view.ctx.readOnly()) return;
		const end = detail.samples.at(-1)?.time ?? 0;
		const now = Date.now();
		const samples = detail.samples.flatMap((sample) => {
			const point = this.#sheetPoint(sample);
			return point ? [{ ...point, pressure: sample.pressure, time: now - end + sample.time }] : [];
		});
		const stroke = inkTools().inkStrokeFromSamples(samples, {
			kind: detail.tool,
			color: tool.color,
			width: tool.width,
		});
		const bounds = stroke && inkTools().inkStrokesBounds([stroke]);
		if (!stroke || !bounds) return;
		const sheetIndex = this.#view.sheetIndex();
		const last = this.#last;
		const target = last && this.#inks().find((entry) => entry.index === last.index);
		if (
			last &&
			target &&
			last.sheet === sheetIndex &&
			target.ink === last.ink &&
			now - last.at < JOIN_MS &&
			bounds.x < target.box.x + target.box.width + JOIN_PX &&
			bounds.x + bounds.width > target.box.x - JOIN_PX &&
			bounds.y < target.box.y + target.box.height + JOIN_PX &&
			bounds.y + bounds.height > target.box.y - JOIN_PX
		) {
			const kept = inkTools().transformInkStrokes(
				inkTools().transformInkStrokes(
					target.ink.strokes,
					0,
					0,
					target.box.width / (target.ink.width || 1),
					target.box.height / (target.ink.height || 1),
				),
				target.box.x,
				target.box.y,
			);
			this.#place(sheetIndex, [...kept, stroke], target.index);
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
		const anchor = this.#anchor(box);
		const placed =
			index === undefined
				? session.addInk(sheetIndex, anchor, { ...ink, name: this.#view.ctx.t('Ink') })
				: (session.setInk(sheetIndex, index, anchor, ink), index);
		const drawing = this.#view.sheet()?.drawings[placed];
		this.#last =
			drawing?.kind === 'ink'
				? { sheet: sheetIndex, index: placed, ink: drawing, at: Date.now() }
				: undefined;
	}

	#erase(detail: InkEraseDetail): void {
		const session = this.#view.ctx.session();
		const point = this.#sheetPoint(detail);
		if (!session || !point || this.#view.ctx.readOnly()) return;
		const reach = detail.radius / this.#factor();
		const sheetIndex = this.#view.sheetIndex();
		// Last first, so deleting an emptied drawing keeps the indexes before it.
		for (const { index, ink, box } of this.#inks().reverse()) {
			if (
				point.x < box.x - reach ||
				point.x > box.x + box.width + reach ||
				point.y < box.y - reach ||
				point.y > box.y + box.height + reach
			)
				continue;
			const kx = ink.width / (box.width || 1);
			const local = {
				x: (point.x - box.x) * kx,
				y: (point.y - box.y) * (ink.height / (box.height || 1)),
			};
			const kept = inkTools().eraseInkStrokes(ink.strokes, local, reach * kx);
			if (kept.length !== ink.strokes.length)
				session.setInk(sheetIndex, index, ink.anchor, {
					strokes: kept,
					width: ink.width,
					height: ink.height,
				});
		}
	}

	#lasso(detail: InkLassoDetail): void {
		this.#clearLasso();
		const polygon = detail.points.flatMap((point) => {
			const at = this.#sheetPoint(point);
			return at ? [at] : [];
		});
		if (polygon.length < 3) return;
		this.#lassoed = this.#inks().flatMap(({ index, ink, box }) => {
			const local = polygon.map((point) => ({
				x: ((point.x - box.x) * ink.width) / (box.width || 1),
				y: ((point.y - box.y) * ink.height) / (box.height || 1),
			}));
			const indexes = inkTools().lassoInkStrokes(ink.strokes, local);
			if (!indexes.length) return [];
			this.#paintBox(ink, box, indexes);
			return [{ index, ink, indexes }];
		});
		this.#view.ctx.requestRender();
	}

	#paintBox(ink: InkObject, box: InkBox, indexes: number[]): void {
		const bounds = inkTools().inkStrokesBounds(indexes.map((index) => ink.strokes[index]!));
		if (!bounds) return;
		const f = this.#factor();
		const kx = box.width / (ink.width || 1);
		const ky = box.height / (ink.height || 1);
		const frame = this.#view.doc.createElement('div');
		frame.className = 'xg-ink-lasso';
		frame.setAttribute('aria-hidden', 'true');
		Object.assign(frame.style, {
			left: `${(box.x + bounds.x * kx) * f}px`,
			top: `${(box.y + bounds.y * ky) * f}px`,
			width: `${bounds.width * kx * f}px`,
			height: `${bounds.height * ky * f}px`,
		});
		this.#view.quads.main.drawings.append(frame);
		this.#boxes.push(frame);
	}

	#clearLasso(): void {
		for (const box of this.#boxes.splice(0)) box.remove();
		if (!this.#lassoed.length) return;
		this.#lassoed = [];
		this.#view.ctx.requestRender();
	}

	#remove(selection: readonly Lassoed[]): void {
		const session = this.#view.ctx.session();
		if (!session) return;
		const sheetIndex = this.#view.sheetIndex();
		for (const { index, ink, indexes } of [...selection].sort((a, b) => b.index - a.index))
			session.setInk(sheetIndex, index, ink.anchor, {
				strokes: ink.strokes.filter((_, at) => !indexes.includes(at)),
				width: ink.width,
				height: ink.height,
			});
	}

	#key(event: KeyboardEvent): void {
		if (event.key === 'Escape' && this.#replay) {
			this.#stopReplay();
			event.preventDefault();
			return;
		}
		if ((event.key !== 'Delete' && event.key !== 'Backspace') || !this.#lassoed.length) return;
		event.preventDefault();
		event.stopPropagation();
		this.#remove(this.#lassoed);
		this.#clearLasso();
	}

	/** The lasso's strokes, or all strokes of the selected ink drawing. */
	#selection(): Lassoed[] {
		if (this.#lassoed.length) return this.#lassoed;
		const index = this.#view.ctx.selection.get().drawing;
		const drawing = index === undefined ? undefined : this.#view.sheet()?.drawings[index];
		return index !== undefined && drawing?.kind === 'ink'
			? [{ index, ink: drawing, indexes: drawing.strokes.map((_, at) => at) }]
			: [];
	}

	hasSelection(): boolean {
		return this.#selection().length > 0;
	}

	async inkToShape(): Promise<boolean> {
		const selection = this.#selection();
		const session = this.#view.ctx.session();
		if (!selection.length || !session) return false;
		const { recognizeInkShape, inkShapeOf } = inkTools();
		const sheetIndex = this.#view.sheetIndex();
		const boxes = new Map(this.#inks().map((entry) => [entry.index, entry.box]));
		const converted: Lassoed[] = [];
		for (const item of selection) {
			const box = boxes.get(item.index);
			const strokes = item.indexes.map((index) => item.ink.strokes[index]!);
			if (!box || !strokes[0]) continue;
			const kx = box.width / (item.ink.width || 1);
			const ky = box.height / (item.ink.height || 1);
			const recognized = recognizeInkShape(
				strokes.map((stroke) =>
					inkTools()
						.inkPathPoints(stroke.path)
						.map((point) => ({
							x: box.x + point.x * kx,
							y: box.y + point.y * ky,
						})),
				),
			);
			if (!recognized) continue;
			const shape = inkShapeOf(recognized, strokes[0].color, strokes[0].width);
			session.addShape(sheetIndex, this.#anchor(shape.box), {
				kind: 'outline',
				preset: shape.preset,
				name: this.#view.ctx.t('Shape'),
				line: shape.line,
				...(shape.rotation ? { rotation: shape.rotation } : {}),
				...(shape.flipH ? { flipH: true } : {}),
				...(shape.flipV ? { flipV: true } : {}),
			});
			converted.push(item);
		}
		this.#remove(converted);
		this.#clearLasso();
		return converted.length > 0;
	}

	replaying(): boolean {
		return this.#replay !== undefined;
	}

	replay(): void {
		if (this.#replay) {
			this.#stopReplay();
			return;
		}
		const g = this.#view.geometry;
		const f = this.#factor();
		const main = g.box('main');
		const { ox, oy } = g.origin('main');
		const strokes = this.#inks().flatMap(({ ink, box }) => {
			const kx = box.width / (ink.width || 1);
			const ky = box.height / (ink.height || 1);
			return ink.strokes.map((stroke) => ({
				points: inkTools()
					.inkPathPoints(stroke.path)
					.map((point) => ({
						x: main.x + (box.x + point.x * kx) * f - ox,
						y: main.y + (box.y + point.y * ky) * f - oy,
					})),
				style: { ...stroke, width: stroke.width * f },
			}));
		});
		const svg = this.#view.doc.createElementNS('http://www.w3.org/2000/svg', 'svg');
		svg.classList.add('xg-ink-replay');
		svg.setAttribute('aria-hidden', 'true');
		svg.setAttribute('width', String(this.#view.view.clientWidth));
		svg.setAttribute('height', String(this.#view.view.clientHeight));
		this.#view.view.append(svg);
		this.#view.root.classList.add('xg-ink-replaying');
		const replay = createInkReplay(
			svg,
			strokes,
			(time, duration, playing) => {
				if (!playing && time >= duration && this.#replay) setTimeout(() => this.#stopReplay(), 400);
			},
			inkTools(),
		);
		if (!replay) {
			svg.remove();
			this.#view.root.classList.remove('xg-ink-replaying');
			return;
		}
		this.#replay = replay;
		replay.toggle();
		this.#view.ctx.requestRender();
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
		this.#clearLasso();
		this.#listeners.abort();
		this.#canvas.remove();
	}
}
