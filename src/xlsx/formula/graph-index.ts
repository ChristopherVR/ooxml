// The indexes the calc engine uses to order and propagate work: formulas by position, and
// formulas by the areas they read.
import { AreaIndex } from './area-index.js';
import type { FormulaNode } from './graph.js';
import type { Area } from './values.js';

/** Formula nodes by sheet and column, rows sorted, for "which formulas are inside this range". */
export class ColumnIndex {
	private readonly sheets = new Map<number, Map<number, FormulaNode[]>>();

	constructor(nodes: Iterable<FormulaNode>) {
		for (const node of nodes) {
			let cols = this.sheets.get(node.sheet);
			if (!cols) this.sheets.set(node.sheet, (cols = new Map()));
			let list = cols.get(node.col);
			if (!list) cols.set(node.col, (list = []));
			list.push(node);
		}
		for (const cols of this.sheets.values())
			for (const list of cols.values()) list.sort((a, b) => a.row - b.row);
	}

	nodesIn(area: Area, out: FormulaNode[]): void {
		const cols = this.sheets.get(area.sheet);
		if (!cols) return;
		const { start, end } = area.range;
		const width = end.col - start.col + 1;
		const scan = (list: FormulaNode[]): void => {
			let lo = 0;
			let hi = list.length;
			while (lo < hi) {
				const mid = (lo + hi) >> 1;
				if ((list[mid] as FormulaNode).row < start.row) lo = mid + 1;
				else hi = mid;
			}
			for (let i = lo; i < list.length; i++) {
				const node = list[i] as FormulaNode;
				if (node.row > end.row) break;
				out.push(node);
			}
		};
		if (width <= cols.size) {
			for (let c = start.col; c <= end.col; c++) {
				const list = cols.get(c);
				if (list) scan(list);
			}
		} else {
			for (const [c, list] of cols) if (c >= start.col && c <= end.col) scan(list);
		}
	}
}

/** Which formulas read a cell: static and dynamic precedents of every node. */
export class ReverseIndex extends AreaIndex<FormulaNode> {
	constructor(nodes: Iterable<FormulaNode>) {
		super();
		for (const node of nodes) {
			for (const dep of node.deps) this.add(dep, node);
			for (const dep of node.dynamicDeps) this.add(dep, node);
		}
	}
}

export { AreaIndex };
