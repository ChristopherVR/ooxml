// Building the dependency graph in slices. Parsing and analysing every formula of a large workbook
// takes hundreds of milliseconds, too long for the first edit to wait on, so an editor can run the
// work while it is idle: a `GraphPreparation` walks the sheets cell by cell, resumable at any point,
// and indexes each formula's precedents as it goes. Edits that arrive first finish what is left.
import type { Cell, Workbook } from '../model';
import type { FormulaNode } from './graph';
import { ReverseIndex } from './graph-index';

/** How many cells are visited between two checks of the time left. */
const CHECK_EVERY = 32;

/** What the first recalculation after opening needs to decide whether stored values can stand. */
export interface PreparedFormulas {
	/** Formulas that may spill (neither legacy nor CSE arrays). */
	dynamic: FormulaNode[];
	/** Formulas saved without a value (computed on the first recalculation). */
	unvalued: FormulaNode[];
}

type CreateNode = (sheet: number, row: number, col: number, cell: Cell) => FormulaNode;

export class GraphPreparation {
	readonly reverse = new ReverseIndex();
	readonly formulas: PreparedFormulas = { dynamic: [], unvalued: [] };
	private sheet = 0;
	private rows: Iterator<[number, Map<number, Cell>]> | undefined;
	private row = 0;
	private cells: Iterator<[number, Cell]> | undefined;
	private finished = false;

	constructor(
		private readonly workbook: Workbook,
		private readonly create: CreateNode,
	) {}

	/**
	 * Adds formulas to the graph until `timeRemaining` (milliseconds, as an idle deadline reports
	 * them) reaches zero, or until every formula is in. Returns true when the graph is complete.
	 */
	run(timeRemaining?: () => number): boolean {
		let visited = 0;
		while (!this.finished) {
			const cells = this.cells ?? this.nextRow();
			if (!cells) continue;
			const next = cells.next();
			if (next.done) {
				this.cells = undefined;
				continue;
			}
			const [col, cell] = next.value;
			if (cell.formula) this.add(this.create(this.sheet, this.row, col, cell), cell);
			if (timeRemaining && ++visited % CHECK_EVERY === 0 && timeRemaining() <= 0) break;
		}
		return this.finished;
	}

	private add(node: FormulaNode, cell: Cell): void {
		this.reverse.addNode(node);
		if (!node.legacy && !node.arrayRange) this.formulas.dynamic.push(node);
		if (cell.value === null && !node.keepCached) this.formulas.unvalued.push(node);
	}

	/** Moves to the next stored row (and sheet), or marks the walk finished. */
	private nextRow(): Iterator<[number, Cell]> | undefined {
		for (;;) {
			const sheet = this.workbook.sheets[this.sheet];
			if (!sheet) {
				this.finished = true;
				return undefined;
			}
			this.rows ??= sheet.rows.entries();
			const next = this.rows.next();
			if (!next.done) {
				this.row = next.value[0];
				this.cells = next.value[1].entries();
				return this.cells;
			}
			this.rows = undefined;
			this.sheet++;
		}
	}
}
