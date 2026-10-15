import type { CellValue, PivotAggregate, PivotTable } from '../model';
import { attrs } from '../write/xml-out';
import { pivotColumnFields, pivotItemLabel, pivotItems, pivotValueKey } from './data';
import { pivotMeasures } from './measures';
import { fieldSubtotalFunctions, PIVOT_SUBTOTALS } from './subtotals';

const SEPARATOR = '\u0001';
const label = pivotItemLabel;

/** One column group of the report: a detail column, a subtotal of a group, or the grand total. */
export interface ColumnEntry {
	kind: 'leaf' | 'subtotal' | 'grand';
	/** Item indices per level, outermost first (empty for the grand total). */
	path: number[];
	/** Identifies the records of the group: its item keys joined (empty for the grand total). */
	key: string;
	/** The item value of each level of the path. */
	values: CellValue[];
	/** The function of a custom subtotal column; an automatic subtotal has none. */
	fn?: PivotAggregate;
}

export interface PivotColumns {
	names: string[];
	/** The visible items of every column field, outermost first. */
	items: CellValue[][];
	/** Display order: each group's columns, then its subtotal, and the grand total last. */
	entries: ColumnEntry[];
}

interface Node {
	index: number;
	value: CellValue;
	children: Map<string, Node>;
}

/** The key of the records under the first `depth` column items of a record. */
export const columnKey = (keys: readonly string[], depth: number): string =>
	keys.slice(0, depth).join(SEPARATOR);

/** The column keys of a record, one per column field. */
export const recordColumnKeys = (record: CellValue[], indices: readonly number[]): string[] =>
	indices.map((index) => pivotValueKey(record[index] ?? null));

/**
 * The column groups of a report: the combinations of column items present in the records, in
 * item order, with a subtotal column after every group that has a deeper level.
 */
export function pivotColumnEntries(pivot: PivotTable): PivotColumns {
	const names = pivotColumnFields(pivot);
	const items = names.map((name) => pivotItems(pivot, name));
	const entries: ColumnEntry[] = [];
	const functions = names.map((name) => fieldSubtotalFunctions(pivot, name));
	if (names.length) {
		const order = items.map(
			(list) => new Map(list.map((item, index) => [pivotValueKey(item), index])),
		);
		const indices = names.map((name) => pivot.fields.indexOf(name));
		const root: Node = { index: 0, value: null, children: new Map() };
		for (const record of pivot.records) {
			let current = root;
			indices.forEach((field, level) => {
				const value = record[field] ?? null;
				const key = pivotValueKey(value);
				let child = current.children.get(key);
				if (!child) {
					child = { index: order[level]!.get(key)!, value, children: new Map() };
					current.children.set(key, child);
				}
				current = child;
			});
		}
		const visit = (
			node: Node,
			level: number,
			path: number[],
			values: CellValue[],
			keys: string[],
		) => {
			for (const child of [...node.children.entries()].sort((a, b) => a[1].index - b[1].index)) {
				const here = {
					path: [...path, child[1].index],
					values: [...values, child[1].value],
					keys: [...keys, child[0]],
				};
				const entry = {
					path: here.path,
					key: columnKey(here.keys, here.keys.length),
					values: here.values,
				};
				if (level === names.length - 1) entries.push({ kind: 'leaf', ...entry });
				else {
					visit(child[1], level + 1, here.path, here.values, here.keys);
					const own = functions[level]!;
					if (own === 'auto') entries.push({ kind: 'subtotal', ...entry });
					else for (const fn of own) entries.push({ kind: 'subtotal', ...entry, fn });
				}
			}
		};
		visit(root, 0, [], [], []);
	}
	entries.push({ kind: 'grand', path: [], key: '', values: [] });
	return { names, items, entries };
}

/** The shape of the header block and the full width of the report. */
export function pivotColumnDimensions(
	pivot: Pick<
		PivotTable,
		'valueField' | 'aggregate' | 'valueCaption' | 'additionalValues' | 'pageFilter'
	>,
	columns: PivotColumns,
	labelColumns: number,
): { rows: number; width: number } {
	const measures = 1 + (pivot.additionalValues?.length ?? 0);
	const levels = columns.names.length;
	return {
		rows: (levels ? levels + 1 + Number(measures > 1) : 1) + (pivot.pageFilter ? 2 : 0),
		width: labelColumns + (levels ? columns.entries.length : 1) * measures,
	};
}

/**
 * The header rows for Values on columns. The first holds the value caption (one measure) and the
 * column field names; then one row per column level labels each group where it starts, subtotal
 * columns in the row of their level and the grand total in the first; the last holds the row
 * labels and, with several measures, the measure captions.
 */
export function pivotColumnHeaders(
	pivot: PivotTable,
	columns: PivotColumns,
	rowLabels: CellValue[],
): CellValue[][] {
	const measures = pivotMeasures(pivot);
	const count = measures.length;
	const levels = columns.names.length;
	if (!levels) return [[...rowLabels, ...measures.map((measure) => measure.caption)]];
	const compact = pivot.layout === 'compact';
	const labelColumns = rowLabels.length;
	const width = labelColumns + columns.entries.length * count;
	const rows = levels + 1 + Number(count > 1);
	const grid = Array.from({ length: rows }, () => Array<CellValue>(width).fill(null));
	if (count === 1) grid[0]![0] = pivot.valueCaption;
	const titles = compact ? ['Column Labels'] : [...columns.names, ...(count > 1 ? ['Values'] : [])];
	titles.forEach((title, i) => {
		grid[0]![labelColumns + i] = title;
	});
	rowLabels.forEach((text, i) => {
		grid[rows - 1]![i] = text;
	});
	let previous: number[] = [];
	columns.entries.forEach((entry, j) => {
		const at = labelColumns + j * count;
		if (entry.kind === 'leaf') {
			// A group is labelled where its first column starts.
			entry.path.forEach((_index, level) => {
				const continued = entry.path.slice(0, level + 1).every((p, i) => p === previous[i]);
				if (!continued) grid[1 + level]![at] = label(entry.values[level]!);
			});
			previous = entry.path;
			if (count > 1) measures.forEach((measure, m) => (grid[rows - 1]![at + m] = measure.caption));
			return;
		}
		const row = entry.kind === 'grand' ? 1 : entry.path.length;
		const name = entry.kind === 'grand' ? '' : String(label(entry.values.at(-1)!));
		measures.forEach((measure, m) => {
			// A custom subtotal names its function, and with several measures its source field.
			const custom = entry.fn ? PIVOT_SUBTOTALS[entry.fn].label : undefined;
			grid[row]![at + m] = custom
				? count === 1
					? `${name} ${custom}`
					: `${name} ${custom} of ${measure.field}`
				: count === 1
					? entry.kind === 'grand'
						? 'Grand Total'
						: `${name} Total`
					: entry.kind === 'grand'
						? `Total ${measure.caption}`
						: `${name} ${measure.caption}`;
		});
	});
	return grid;
}

/** How many leading items an entry repeats from the previous one (never its last item). */
export function sharedPrefix(path: readonly number[], previous: readonly number[]): number {
	let prefix = 0;
	while (prefix < path.length - 1 && path[prefix] === previous[prefix]) prefix++;
	return prefix;
}

/**
 * `colItems` for column fields: one item per group and measure, each repeating the leading items
 * it shares with the previous one. A detail column also lists its measure; a subtotal or grand
 * total names it with `i`. `remap` turns a visible item index into the one in the field's list.
 */
export function pivotColumnItemsXml(
	entries: readonly ColumnEntry[],
	measures: number,
	remap: (level: number, index: number) => number,
): string {
	const out: string[] = [];
	let previous: number[] = [];
	for (const entry of entries)
		for (let m = 0; m < measures; m++) {
			if (entry.kind === 'grand') {
				out.push(`<i${attrs({ t: 'grand', i: m || undefined })}><x/></i>`);
				previous = [];
				continue;
			}
			const group = entry.path.map((index, level) => remap(level, index));
			const path = entry.kind === 'leaf' && measures > 1 ? [...group, m] : group;
			const prefix = sharedPrefix(path, previous);
			const xs = path.slice(prefix).map((index) => `<x${attrs({ v: index || undefined })}/>`);
			out.push(
				`<i${attrs({ t: entry.kind === 'subtotal' ? (entry.fn ? PIVOT_SUBTOTALS[entry.fn].item : 'default') : undefined, r: prefix || undefined, i: m || undefined })}>${xs.join('')}</i>`,
			);
			previous = path;
		}
	return `<colItems count="${out.length}">${out.join('')}</colItems>`;
}
