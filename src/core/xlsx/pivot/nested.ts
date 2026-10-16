import type { CellValue, PivotAggregate, PivotTable } from '../model';
import type { PivotGrid, PivotRowEntry } from './aggregate';
import { pivotItems, pivotRowFields, pivotValueKey } from './data';
import { pivotItemLabeler } from './item-labels';
import { pivotValuesIndex } from './axes';
import {
	columnKey,
	pivotColumnDimensions,
	pivotColumnEntries,
	pivotColumnHeaders,
	recordColumnKeys,
} from './columns';
import { aggregateValues, pivotMeasureNumber, pivotMeasures } from './measures';
import { fillCaption, pivotCaptions } from './captions';
import { fieldSubtotalFunctions } from './subtotals';

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

/** What the walk carries down: the labels, item indices and first-child flags of every level. */
interface Path {
	labels: CellValue[];
	indices: number[];
	firsts: boolean[];
}

/**
 * Any number of row and column fields, with Values on the columns or at any place among the row
 * fields. Every group is a node holding its records, so the group's row, its automatic subtotal
 * and each custom subtotal read the same source values. The axis is a list of levels: the row
 * fields and, with Values on rows, one level whose items are the measures.
 */
export function nestedPivotGrid(pivot: PivotTable): PivotGrid {
	const names = pivotRowFields(pivot);
	const depth = names.length;
	const valuesAt = pivotValuesIndex(pivot);
	const onRows = valuesAt >= 0;
	const levels = depth + (onRows ? 1 : 0);
	/** The row field of a level, or -1 for the Values level. */
	const fieldOf = (level: number) =>
		!onRows || level < valuesAt ? level : level === valuesAt ? -1 : level - 1;
	const axisItems = names.map((name) => pivotItems(pivot, name));
	const itemIndices = axisItems.map(
		(items) => new Map(items.map((item, index) => [pivotValueKey(item), index])),
	);
	const columns = pivotColumnEntries(pivot);
	const fieldIndices = names.map((name) => pivot.fields.indexOf(name));
	const columnFields = columns.names.map((name) => pivot.fields.indexOf(name));
	const layout = pivot.layout ?? 'tabular';
	const compact = layout === 'compact';
	const tabular = layout === 'tabular';
	const above = !tabular && (pivot.subtotalPosition ?? 'top') === 'top';
	const labelColumns = compact ? 1 : levels;
	const measures = pivotMeasures(pivot);
	const functions = names.map((name) => fieldSubtotalFunctions(pivot, name));
	const caps = pivotCaptions(pivot);
	const labelOf = pivotItemLabeler(pivot);

	const root = node(0, null);
	for (const record of pivot.records) {
		const keys = recordColumnKeys(record, columnFields);
		let current = root;
		for (let field = -1; field < depth; field++) {
			if (field >= 0) {
				const value = record[fieldIndices[field]!] ?? null;
				const key = pivotValueKey(value);
				let child = current.children.get(key);
				if (!child) {
					child = node(itemIndices[field]!.get(key)!, value);
					current.children.set(key, child);
				}
				current = child;
			}
			current.records.push(record);
			for (let width = 1; width <= keys.length; width++) {
				const key = columnKey(keys, width);
				const list = current.columns.get(key);
				if (list) list.push(record);
				else current.columns.set(key, [record]);
			}
		}
	}

	const dimensions = pivotColumnDimensions(pivot, columns, labelColumns, onRows);
	const rowTitles = names.map((name): CellValue => name);
	if (onRows) rowTitles.splice(valuesAt, 0, caps.values);
	let values: CellValue[][] = [];
	let indents: number[] = [];
	const rowEntries: PivotRowEntry[] = [];
	// A first pass only counts the rows, so a report that is too large is refused before any
	// header or cell is built.
	let counting = true;
	let rowCount = 0;
	const width = dimensions.width - labelColumns;
	const blanks = (): CellValue[] => Array<CellValue>(width).fill(null);

	/**
	 * The cells of a group across the column entries, for one measure (Values on rows) or all of
	 * them. A custom column subtotal summarizes with its own function; crossed with a custom row
	 * subtotal it shows a value only where both agree.
	 */
	const cells = (group: Node, rowFn?: PivotAggregate, measure?: number): CellValue[] =>
		columns.entries.flatMap((entry) => {
			const records = entry.kind === 'grand' ? group.records : (group.columns.get(entry.key) ?? []);
			const picked = measure === undefined ? measures : [measures[measure]!];
			if (rowFn && entry.fn && rowFn !== entry.fn) return picked.map((): CellValue => null);
			const fn = rowFn ?? entry.fn;
			return picked.map((one) =>
				fn
					? aggregateValues(records, pivot.fields.indexOf(one.field), fn)
					: pivotMeasureNumber(pivot, records, one),
			);
		});

	/** Label cells: the ancestors' labels (when shown) then this row's own at its column. */
	const labelCells = (path: Path, shown: boolean[], level: number, text: CellValue) => {
		if (compact) return [text];
		const row: CellValue[] = Array(levels).fill(null);
		path.labels.forEach((value, at) => {
			if (shown[at]) row[at] = value;
		});
		row[level] = text;
		return row;
	};
	type Row = [labels: CellValue[], numbers: CellValue[], entry: PivotRowEntry];
	const push = (indent: number, build: () => Row) => {
		if (counting) {
			if ((++rowCount + dimensions.rows) * dimensions.width > MAX_CELLS)
				throw new Error('The PivotTable result is too large.');
			return;
		}
		const [cells, numbers, entry] = build();
		values.push([...cells, ...numbers]);
		rowEntries.push(entry);
		indents.push(compact ? indent : 0);
	};
	const entryOf = (indices: number[], measure?: number, type?: PivotRowEntry['type']) => ({
		indices,
		...(measure !== undefined ? { measure } : {}),
		...(type ? { type } : {}),
	});

	/** Walks one level: the Values level lists the measures, a field level its items. */
	const walk = (level: number, group: Node, measure: number | undefined, path: Path) => {
		const field = fieldOf(level);
		const last = level === levels - 1;
		// Only a repeat shows ancestors on header and subtotal rows; tabular detail rows show
		// an ancestor on the first row of its group.
		// The Values level is not an item label and is never repeated.
		const repeated = Array.from(
			{ length: level },
			(_v, k) => !!pivot.repeatLabels && fieldOf(k) >= 0,
		);
		const leafShown = (firsts: boolean[]) =>
			repeated.map((repeat, k) => repeat || (tabular && firsts.slice(k + 1).every(Boolean)));
		if (field < 0) {
			measures.forEach((one, m) => {
				const here: Path = {
					labels: [...path.labels, one.caption],
					indices: [...path.indices, m],
					firsts: [...path.firsts, m === 0],
				};
				if (last)
					return push(level, () => [
						labelCells(path, leafShown(here.firsts), level, one.caption),
						cells(group, undefined, m),
						entryOf(here.indices, m),
					]);
				if (!tabular)
					push(level, () => [
						labelCells(path, repeated, level, one.caption),
						blanks(),
						entryOf(here.indices, m),
					]);
				walk(level + 1, group, m, here);
			});
			return;
		}
		const own = functions[field]!;
		// With Values above (or on columns) the measure is known, so a header row can show the
		// group's value; below it the group's subtotals expand into one row per measure.
		const valuesAbove = !onRows || valuesAt < level;
		const showsValue = above && own === 'auto' && valuesAbove;
		const children = [...group.children.values()].sort((a, b) => a.index - b.index);
		children.forEach((child, position) => {
			const here: Path = {
				labels: [...path.labels, labelOf(names[field]!, child.value)],
				indices: [...path.indices, child.index],
				firsts: [...path.firsts, position === 0],
			};
			const text = labelOf(names[field]!, child.value);
			if (last)
				return push(level, () => [
					labelCells(path, leafShown(here.firsts), level, text),
					cells(child, undefined, measure),
					entryOf(here.indices, measure),
				]);
			if (!tabular)
				push(level, () => [
					labelCells(path, repeated, level, text),
					showsValue ? cells(child, undefined, measure) : blanks(),
					entryOf(here.indices, measure),
				]);
			walk(level + 1, child, measure, here);
			if (field === depth - 1) return;
			const caption = String(text);
			const subtotal = (title: string, fn: PivotAggregate | undefined, one?: number) =>
				push(level, () => [
					labelCells(path, repeated, level, title),
					cells(child, fn, one),
					entryOf(here.indices, one, fn ?? 'default'),
				]);
			if (own === 'auto') {
				if (showsValue) return;
				if (valuesAbove) subtotal(fillCaption(caps.subtotal, caption), undefined, measure);
				else measures.forEach((one, m) => subtotal(`${caption} ${one.caption}`, undefined, m));
			} else
				for (const fn of own)
					if (valuesAbove) subtotal(`${caption} ${caps.functions[fn]}`, fn, measure);
					else
						measures.forEach((one, m) =>
							subtotal(`${caption} ${caps.functions[fn]} of ${one.field}`, fn, m),
						);
		});
	};
	const grand = (text: CellValue, measure?: number) =>
		push(0, () => [
			compact ? [text] : [text, ...Array<CellValue>(levels - 1).fill(null)],
			cells(root, undefined, measure),
			entryOf([0], measure, 'grand'),
		]);
	const build = () => {
		walk(0, root, undefined, { labels: [], indices: [], firsts: [] });
		if (onRows) measures.forEach((one, m) => grand(fillCaption(caps.measureTotal, one.caption), m));
		else grand(caps.grandTotal);
	};
	build();
	counting = false;
	values = pivotColumnHeaders(pivot, columns, compact ? [caps.rowLabels] : rowTitles, onRows);
	const firstDataRow = values.length;
	indents = values.map(() => 0);
	build();
	return {
		values,
		rowItems: axisItems[0]!,
		...(axisItems[1] ? { secondRowItems: axisItems[1] } : {}),
		axisItems,
		columnItems: columns.items[0] ?? [],
		columnAxisItems: columns.items,
		columnEntries: columns.entries,
		rowEntries,
		firstDataRow,
		rowLabelColumns: labelColumns,
		indents,
		...(onRows ? { firstHeaderRow: 1 } : {}),
	};
}
