// The lasso's selection over Excel's cells: the strokes it holds in each ink drawing, framed on
// the sheet, for Delete and Ink to Shape. It belongs to the drawings it framed: once one of them
// changes (an edit, an undo, a remote change) the selection is dropped.
import type { InkCoordinate } from 'ooxml-core/drawingml';
import type { InkObject } from 'ooxml-core/xlsx';
import { inkTools } from '../../ink/tools';
import type { GridView } from './grid-view';
import { anchorOf, fittedInk, sheetInks, zoomOf, type SheetInk } from './ink-sheet';

interface Held {
	index: number;
	ink: InkObject;
	/** The strokes it framed: an edit of the drawing replaces them. */
	strokes: InkObject['strokes'];
	indexes: number[];
}

export class InkSelection {
	readonly #view: GridView;
	#held: Held[] = [];
	#boxes: HTMLElement[] = [];
	#sheet = -1;

	constructor(view: GridView) {
		this.#view = view;
	}

	/** Holds the strokes inside `polygon` (sheet points at 100%) and frames them. */
	select(polygon: InkCoordinate[]): void {
		this.clear();
		if (polygon.length < 3) return;
		this.#sheet = this.#view.sheetIndex();
		this.#held = sheetInks(this.#view).flatMap(({ index, ink, box }) => {
			const local = polygon.map((point) => ({
				x: ((point.x - box.x) * ink.width) / (box.width || 1),
				y: ((point.y - box.y) * ink.height) / (box.height || 1),
			}));
			const indexes = inkTools().lassoInkStrokes(ink.strokes, local);
			if (!indexes.length) return [];
			this.#frame({ index, ink, box }, indexes);
			return [{ index, ink, strokes: ink.strokes, indexes }];
		});
		this.#view.ctx.requestRender();
	}

	/** The held strokes, or all strokes of the selected ink drawing; drops a stale selection. */
	held(): Held[] {
		if (this.#held.length && !this.#current()) this.clear();
		if (this.#held.length) return this.#held;
		const index = this.#view.ctx.selection.get().drawing;
		const drawing = index === undefined ? undefined : this.#view.sheet()?.drawings[index];
		return index !== undefined && drawing?.kind === 'ink' && !drawing.truncated
			? [
					{
						index,
						ink: drawing,
						strokes: drawing.strokes,
						indexes: drawing.strokes.map((_, at) => at),
					},
				]
			: [];
	}

	get size(): number {
		return this.#current() ? this.#held.length : 0;
	}

	/** Whether every held drawing is still the one the lasso framed. */
	#current(): boolean {
		const drawings = this.#view.sheet()?.drawings;
		return (
			this.#view.sheetIndex() === this.#sheet &&
			this.#held.every(
				({ index, ink, strokes }) => drawings?.[index] === ink && ink.strokes === strokes,
			)
		);
	}

	#frame(entry: SheetInk, indexes: number[]): void {
		const { ink, box } = entry;
		const bounds = inkTools().inkStrokesBounds(indexes.map((index) => ink.strokes[index]!));
		if (!bounds) return;
		const f = zoomOf(this.#view);
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

	clear(): void {
		for (const box of this.#boxes.splice(0)) box.remove();
		if (!this.#held.length) return;
		this.#held = [];
		this.#view.ctx.requestRender();
	}

	/** Removes the held strokes from their drawings, each box fitted to what it keeps. */
	remove(selection: readonly Held[]): void {
		const session = this.#view.ctx.session();
		if (!session) return;
		const sheetIndex = this.#view.sheetIndex();
		const entries = new Map(sheetInks(this.#view).map((entry) => [entry.index, entry]));
		// Last first, so deleting an emptied drawing keeps the indexes before it.
		for (const { index, ink, strokes, indexes } of [...selection].sort(
			(a, b) => b.index - a.index,
		)) {
			const entry = entries.get(index);
			// A drawing changed since it was selected keeps its strokes.
			if (!entry || entry.ink !== ink || ink.strokes !== strokes) continue;
			const kept = ink.strokes.filter((_, at) => !indexes.includes(at));
			const next = fittedInk(this.#view, entry, kept);
			session.setInk(sheetIndex, index, next.anchor, next.ink);
		}
	}

	/** Ink to Shape, one undo step: each recognised selection becomes a shape outlined in its colour. */
	inkToShape(): boolean {
		const selection = this.held();
		const session = this.#view.ctx.session();
		if (!selection.length || !session) return false;
		const { recognizeInkShape, inkShapeOf, inkPathPoints } = inkTools();
		const sheetIndex = this.#view.sheetIndex();
		const boxes = new Map(sheetInks(this.#view).map((entry) => [entry.index, entry.box]));
		const converted: Held[] = [];
		session.batch(this.#view.ctx.t('Ink to Shape'), () => {
			for (const item of selection) {
				const box = boxes.get(item.index);
				const strokes = item.indexes.map((index) => item.ink.strokes[index]!);
				if (!box || !strokes[0]) continue;
				const kx = box.width / (item.ink.width || 1);
				const ky = box.height / (item.ink.height || 1);
				const recognized = recognizeInkShape(
					strokes.map((stroke) =>
						inkPathPoints(stroke.path).map((point) => ({
							x: box.x + point.x * kx,
							y: box.y + point.y * ky,
						})),
					),
				);
				if (!recognized) continue;
				const shape = inkShapeOf(recognized, strokes[0].color, strokes[0].width);
				session.addShape(sheetIndex, anchorOf(this.#view, shape.box), {
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
			this.remove(converted);
		});
		this.clear();
		return converted.length > 0;
	}
}
