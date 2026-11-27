import { DEFAULT_INK_PENS, MAX_INK_PENS, type InkPen } from 'ooxml-core/drawingml';
import { INK_PEN_COLORS, type InkPensLabels } from './pen-labels';
import type { InkCanvasTool } from './ink-canvas';

/**
 * The Draw tab's state, shared by Word and Excel: the tool (Select, Lasso Select, Eraser or a pen
 * of the gallery), the pen gallery and Draw with Touch. DOM-free; each product's ribbon renders
 * it and its drawing surface reads `canvasTool`.
 */
export type DrawTool = 'select' | 'lasso' | 'eraser' | 'pen';

export interface DrawStateSnapshot {
	tool: DrawTool;
	pens: readonly InkPen[];
	/** Index of the pen in `pens` the next stroke draws with. */
	pen: number;
	touch: boolean;
}

export class DrawState {
	#tool: DrawTool = 'select';
	#pens: InkPen[] = DEFAULT_INK_PENS.map((pen) => ({ ...pen }));
	#pen = 0;
	#touch = false;
	readonly #listeners = new Set<() => void>();

	get snapshot(): DrawStateSnapshot {
		return { tool: this.#tool, pens: this.#pens, pen: this.#pen, touch: this.#touch };
	}

	/** The pen the next stroke draws with. */
	get pen(): InkPen {
		return this.#pens[this.#pen] ?? DEFAULT_INK_PENS[0]!;
	}

	/** The tool the drawing surface captures the pointer for ('' lets the pointer through). */
	get canvasTool(): InkCanvasTool {
		if (this.#tool === 'pen') {
			const kind = this.pen.kind;
			return kind === 'actionPen' ? 'lasso' : kind;
		}
		return this.#tool === 'select' ? '' : this.#tool;
	}

	onChange(listener: () => void): () => void {
		this.#listeners.add(listener);
		return () => this.#listeners.delete(listener);
	}

	#changed(): void {
		for (const listener of this.#listeners) listener();
	}

	setTool(tool: DrawTool): void {
		if (this.#tool === tool) return;
		this.#tool = tool;
		this.#changed();
	}

	/** Picks a pen of the gallery, which turns drawing on (a second pick of it keeps it on). */
	selectPen(index: number): void {
		if (!this.#pens[index]) return;
		this.#pen = index;
		this.#tool = 'pen';
		this.#changed();
	}

	/** Add Pen: a copy of the current pen at the end of the gallery, picked. */
	addPen(kind: InkPen['kind'] = this.pen.kind): boolean {
		if (this.#pens.length >= MAX_INK_PENS) return false;
		const like = this.#pens.find((pen) => pen.kind === kind) ?? this.pen;
		this.#pens = [...this.#pens, { ...like, kind }];
		this.selectPen(this.#pens.length - 1);
		return true;
	}

	removePen(index: number): void {
		if (this.#pens.length <= 1 || !this.#pens[index]) return;
		this.#pens = this.#pens.filter((_, at) => at !== index);
		this.#pen = Math.min(this.#pen, this.#pens.length - 1);
		this.#changed();
	}

	/** Recolours or resizes the current pen. */
	updatePen(change: Partial<Pick<InkPen, 'color' | 'width'>>): void {
		const pen = this.#pens[this.#pen];
		if (!pen) return;
		const color = change.color && /^#[\da-f]{6}$/iu.test(change.color) ? change.color : pen.color;
		const width = change.width && change.width > 0 && change.width <= 64 ? change.width : pen.width;
		this.#pens = this.#pens.map((item, at) =>
			at === this.#pen ? { ...item, color, width } : item,
		);
		this.#changed();
	}

	setTouch(on: boolean): void {
		if (this.#touch === on) return;
		this.#touch = on;
		this.#changed();
	}
}

/** The pen strip's texts in the product's language (`t` translates English). */
export function inkPensLabels(t: (text: string) => string, locale?: string): InkPensLabels {
	const colorNames = Object.fromEntries(INK_PEN_COLORS.map(({ name }) => [name, t(name)]));
	return {
		group: t('Pens'),
		pen: '{kind}: {size} mm, {color}',
		kinds: {
			pen: t('Pen'),
			pencil: t('Pencil'),
			highlighter: t('Highlighter'),
			actionPen: t('Action Pen'),
		},
		thickness: t('Thickness'),
		colors: t('Colors'),
		remove: t('Delete Pen'),
		moreColors: t('More Colors...'),
		colorNames,
		description: t('Draw ink with this pen. Select it again to change its color and thickness.'),
		standardColors: t('Standard Colors'),
		...(locale ? { locale } : {}),
	};
}
