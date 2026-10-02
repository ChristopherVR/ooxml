// Cached core cell views of the shown sheet, with the sheet's conditional-format evaluator. The
// cache is dropped after every model change; a blank, unstyled cell is answered without calling
// the core (most of a sheet is empty).
import {
	cellKey,
	cellView,
	createConditionalFormatEvaluator,
	effectiveStyleId,
	getCell,
	type CellView,
	type ConditionalFormatEvaluator,
	type EditSession,
	type Workbook,
	type Worksheet,
} from '@christophervr/xlsx-core';
import { isBlankView } from './cell-items.js';

export interface ViewSource {
	workbook(): Workbook | undefined;
	sheet(): Worksheet | undefined;
	sheetIndex(): number;
	session(): EditSession | undefined;
}

const cfCovers = (sheet: Worksheet, row: number, col: number): boolean =>
	sheet.conditionalFormats.some((cf) =>
		cf.ranges.some(
			(r) => row >= r.start.row && row <= r.end.row && col >= r.start.col && col <= r.end.col,
		),
	);

export class CellViewCache {
	#views = new Map<number, CellView | null>();
	#cf: ConditionalFormatEvaluator | undefined;
	#cfReady = false;
	#source: ViewSource;

	constructor(source: ViewSource) {
		this.#source = source;
	}

	clear(): void {
		this.#views.clear();
		this.#cf = undefined;
		this.#cfReady = false;
	}

	/** The core view of a cell; null when it paints nothing. */
	get(row: number, col: number): CellView | null {
		const key = cellKey(row, col);
		const hit = this.#views.get(key);
		if (hit !== undefined) return hit;
		const workbook = this.#source.workbook();
		const sheet = this.#source.sheet();
		let result: CellView | null = null;
		if (workbook && sheet) {
			const stored = getCell(sheet, row, col);
			const plain =
				!stored && effectiveStyleId(sheet, row, col) === 0 && !cfCovers(sheet, row, col);
			if (!plain) {
				const view = cellView(workbook, this.#source.sheetIndex(), row, col, this.#evaluator());
				result = isBlankView(view) ? null : view;
			}
		}
		this.#views.set(key, result);
		return result;
	}

	#evaluator(): ConditionalFormatEvaluator | undefined {
		if (this.#cfReady) return this.#cf;
		this.#cfReady = true;
		const workbook = this.#source.workbook();
		const session = this.#source.session();
		const sheet = this.#source.sheet();
		if (!workbook || !sheet?.conditionalFormats.length) return undefined;
		try {
			this.#cf = createConditionalFormatEvaluator(
				workbook,
				this.#source.sheetIndex(),
				(formula, at) => (session ? session.calc.evaluate(formula, at) : null),
			);
		} catch {
			this.#cf = undefined;
		}
		return this.#cf;
	}
}
