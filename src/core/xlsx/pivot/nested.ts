import type { CellValue, PivotAggregate, PivotTable } from '../model';
import type { PivotGrid, PivotRowEntry } from './aggregate';
import { pivotItemLabel, pivotItems, pivotRowFields, pivotValueKey } from './data';
import {
	aggregateValues,
	pivotHeaderDimensions,
	pivotHeaders,
	pivotMeasures,
	pivotNumbers,
} from './measures';
import { fieldSubtotalFunctions, PIVOT_SUBTOTALS } from './subtotals';

const label = pivotItemLabel;
const MAX_CELLS = 1_000_000;

interface Node {
	index: number;
	value: CellValue;
	records: CellValue[][];
	/** Records per column item key, when the report has a column field. */
	columns: Map<string, CellValue[][]>;
	children: Map<string, Node>;
}

const node = (index: number, value: CellValue): Node => ({
	index,
	value,
	records: [],
	columns: new Map(),
	children: new Map(),
});

/**
 * Values on columns with any number of row fields. Every group is a node holding its records, so
 * the group's row, its automatic subtotal and each custom subtotal read the same source values.
 */
export function nestedPivotGrid(pivot: PivotTable): PivotGrid {
	const names = pivotRowFields(pivot);
	const depth = names.length;
	const axisItems = names.map((name) => pivotItems(pivot, name));
	const itemIndices = axisItems.map(
		(items) => new Map(items.map((item, index) => [pivotValueKey(item), index])),
	);
	const columnItems = pivot.columnField ? pivotItems(pivot, pivot.columnField) : [];
	const fieldIndices = names.map((name) => pivot.fields.indexOf(name));
	const columnField = pivot.columnField ? pivot.fields.indexOf(pivot.columnField) : -1;
	const layout = pivot.layout ?? 'tabular';
	const compact = layout === 'compact';
	const tabular = layout === 'tabular';
	const above = !tabular && (pivot.subtotalPosition ?? 'top') === 'top';
	const labelColumns = compact ? 1 : depth;
	const measures = pivotMeasures(pivot);

	const root = node(0, null);
	for (const record of pivot.records) {
		const column = columnField >= 0 ? pivotValueKey(record[columnField] ?? null) : '';
		let current = root;
		for (let level = -1; level < depth; level++) {
			if (level >= 0) {
				const value = record[fieldIndices[level]!] ?? null;
				const key = pivotValueKey(value);
				let child = current.children.get(key);
				if (!child) {
					child = node(itemIndices[level]!.get(key)!, value);
					current.children.set(key, child);
				}
				current = child;
			}
			current.records.push(record);
			if (columnField >= 0) {
				const list = current.columns.get(column);
				if (list) list.push(record);
				else current.columns.set(column, [record]);
			}
		}
	}

	const dimensions = pivotHeaderDimensions(pivot, columnItems.length, labelColumns);
	const functions = names.map((name) => fieldSubtotalFunctions(pivot, name));
	let rows = 1;
	const count = (parent: Node, level: number) => {
		for (const child of parent.children.values()) {
			rows++;
			if (level < depth - 1) {
				const f = functions[level]!;
				rows += f === 'auto' ? 1 : f.length;
				count(child, level + 1);
			}
		}
	};
	count(root, 0);
	if ((rows + dimensions.rows) * dimensions.width > MAX_CELLS)
		throw new Error('The PivotTable result is too large.');

	const values = pivotHeaders(pivot, columnItems, compact ? ['Row Labels'] : names, label);
	const firstDataRow = values.length;
	const indents = values.map(() => 0);
	const rowEntries: PivotRowEntry[] = [];
	const blanks = (): CellValue[] => Array<CellValue>(dimensions.width - labelColumns).fill(null);
	const bands = (group: Node, run: (records: CellValue[][]) => CellValue[]): CellValue[] => [
		...columnItems.flatMap((column) => run(group.columns.get(pivotValueKey(column)) ?? [])),
		...run(group.records),
	];
	const numbers = (group: Node) => bands(group, (records) => pivotNumbers(pivot, records));
	const subtotalNumbers = (group: Node, fn: PivotAggregate) =>
		bands(group, (records) =>
			measures.map((measure) => aggregateValues(records, pivot.fields.indexOf(measure.field), fn)),
		);

	/** Label cells: the ancestors' labels (when shown) then this row's own at its column. */
	const labelCells = (path: CellValue[], shown: boolean[], text: CellValue): CellValue[] => {
		if (compact) return [text];
		const cells: CellValue[] = Array(depth).fill(null);
		path.forEach((value, level) => {
			if (shown[level]) cells[level] = label(value);
		});
		cells[path.length] = text;
		return cells;
	};
	const push = (
		cells: CellValue[],
		numbersOnRow: CellValue[],
		entry: PivotRowEntry,
		indent: number,
	) => {
		values.push([...cells, ...numbersOnRow]);
		rowEntries.push(entry);
		indents.push(compact ? indent : 0);
	};

	const emit = (
		parent: Node,
		level: number,
		path: CellValue[],
		indices: number[],
		firsts: boolean[],
	) => {
		const last = level === depth - 1;
		const own = functions[level]!;
		const children = [...parent.children.values()].sort((a, b) => a.index - b.index);
		children.forEach((child, position) => {
			const here = [...path, child.value];
			const at = [...indices, child.index];
			const firstHere = [...firsts, position === 0];
			// Only a repeat shows ancestors on header and subtotal rows; tabular detail rows show
			// an ancestor on the first row of its group.
			const repeated = Array<boolean>(level).fill(!!pivot.repeatLabels);
			if (last) {
				const shown = repeated.map(
					(repeat, k) => repeat || (tabular && firstHere.slice(k + 1).every(Boolean)),
				);
				push(labelCells(path, shown, label(child.value)), numbers(child), { indices: at }, level);
				return;
			}
			if (!tabular)
				push(
					labelCells(path, repeated, label(child.value)),
					own === 'auto' && above ? numbers(child) : blanks(),
					{ indices: at },
					level,
				);
			emit(child, level + 1, here, at, firstHere);
			const caption = String(label(child.value));
			if (own === 'auto') {
				if (!above)
					push(
						labelCells(path, repeated, `${caption} Total`),
						numbers(child),
						{ indices: at, type: 'default' },
						level,
					);
			} else
				for (const fn of own)
					push(
						labelCells(path, repeated, `${caption} ${PIVOT_SUBTOTALS[fn].label}`),
						subtotalNumbers(child, fn),
						{ indices: at, type: fn },
						level,
					);
		});
	};
	emit(root, 0, [], [], []);
	push(
		compact ? ['Grand Total'] : ['Grand Total', ...Array<CellValue>(depth - 1).fill(null)],
		numbers(root),
		{ indices: [0], type: 'grand' },
		0,
	);
	return {
		values,
		rowItems: axisItems[0]!,
		secondRowItems: axisItems[1]!,
		axisItems,
		columnItems,
		rowEntries,
		firstDataRow,
		rowLabelColumns: labelColumns,
		indents,
	};
}
