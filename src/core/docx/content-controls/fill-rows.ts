// Filling the controls of a table: a repeating section or a control around rows repeats its rows
// per record, a plain row holding a control named in `repeat` is copied per record with the rows
// after it whose controls are per-item templates (sub-rows such as a discount, kept for a record
// only when it fills one of their controls), cells are filled as block lists, and a row an
// optional control emptied goes. Copied rows are recorded for the writer, which copies the source
// row whole (its borders, shading, height and fonts) and writes only its cells' paragraphs.
import { rowEditKey } from '../block-copies';
import { documentParagraphs } from '../cell-blocks';
import type { ContentControl } from '../content-control-model';
import type { Table, TableCell } from '../model';
import type { TableRowProperties } from '../table-model';
import { chainOf, groupAt } from './chains';
import { anyControlShowsValue, blocksText, forEachControl } from './controls';
import {
	copyKey,
	isOptional,
	lookup,
	repeatsOf,
	report,
	type FillEnvironment,
	type Scope,
} from './environment';
import { fillCell } from './fill-cells';
import { copyUnder, retagControls } from './repeat';

interface Row {
	cells: TableCell[];
	props: TableRowProperties | undefined;
}

interface TableFill {
	table: Table;
	/** A table nested in a cell: its rows are not repeated or removed. */
	nested: boolean;
	/** Rows were copied or removed. */
	structural: boolean;
}

const rowChain = (row: Row) => chainOf(row.props);
const asTable = (rows: readonly Row[]): Table => ({
	type: 'table',
	id: '',
	rows: rows.map((row) => row.cells),
	rowProperties: rows.map((row) => row.props ?? {}),
});
const fromTable = (table: Table): Row[] =>
	table.rows.map((cells, index) => ({ cells, props: table.rowProperties?.[index] }));

/** The tags of the controls in the cells of rows (not of the controls around the rows). */
function rowTags(rows: readonly Row[]): string[] {
	const tags: string[] = [];
	const cells: Table = { type: 'table', id: '', rows: rows.map((row) => row.cells) };
	forEachControl([cells], (control) => {
		if (control.tag !== undefined) tags.push(control.tag);
	});
	return tags;
}

const TEMPLATE = /(^|[_.-])template([_.-]|$)/i;

/** Whether a control in a row makes it a row of repeat `key`: tagged `key`, or `key_template_…`. */
const belongs = (tag: string, key: string) =>
	tag === key ||
	(/^[_.-]/.test(tag.slice(key.length)) && tag.startsWith(key) && TEMPLATE.test(tag));

/** The repeat a plain row belongs to: the first key one of its controls belongs to. */
function rowKey(row: Row, env: FillEnvironment): string | undefined {
	const keys = Object.keys(env.data.repeat ?? {});
	const tags = rowTags([row]);
	return keys.find((key) => tags.some((tag) => belongs(tag, key)));
}

/** Copies of rows under new ids; the writer copies the source rows. */
function copyRows(
	rows: readonly Row[],
	env: FillEnvironment,
	fill: TableFill,
	keep: (control: ContentControl) => boolean,
): Row[] {
	const rekey = copyKey(env);
	const source = asTable(rows);
	const [table] = copyUnder([source], env, keep, rekey) as [Table];
	if (!fill.nested)
		for (const paragraph of documentParagraphs([source]))
			env.copies.set(rekey(paragraph.id), { sourceId: paragraph.id, rekey });
	return fromTable(table);
}

/** Fills the cells of a plain row; none when an optional control emptied it. */
function fillRow(row: Row, scope: Scope, env: FillEnvironment, fill: TableFill): Row[] {
	const removed = env.removed;
	const cells = row.cells.map((cell) => fillCell(cell, scope, env));
	const next = { ...row, cells };
	if (env.removed === removed || fill.nested) return [next];
	const table = asTable([next]);
	const empty = !blocksText([table]).trim() && !rowTags([next]).length;
	if (!empty) return [next];
	fill.structural = true;
	return [];
}

/** Whether a row is a sub-row of repeat `key`: each of its tagged controls is `key_template_…`. */
const isSubRow = (row: Row, key: string) => {
	const tags = rowTags([row]);
	return tags.length > 0 && tags.every((tag) => tag !== key && belongs(tag, key));
};

/** The copies of `unit` (rows repeated together), one per record. */
function repeatRows(
	unit: readonly Row[],
	fillCopy: (rows: Row[], scope: Scope) => Row[],
	records: ReadonlyArray<Readonly<Record<string, string>>>,
	keep: (control: ContentControl) => boolean,
	keepTag: (control: ContentControl) => boolean,
	subRows: boolean,
	env: FillEnvironment,
	fill: TableFill,
): Row[] {
	fill.structural = true;
	return records.flatMap((record, index) => {
		const item = index + 1;
		const scope = { record, item };
		const rows = unit.filter(
			(row, at) =>
				!subRows ||
				at === 0 ||
				rowTags([row]).some((tag) => (lookup(tag, scope, env) ?? '').trim() !== ''),
		);
		const copy = index ? copyRows(rows, env, fill, keep) : rows;
		const filledRows = fillCopy(copy, scope);
		return fromTable(retagControls([asTable(filledRows)], item, keepTag)[0] as Table);
	});
}

/** The rows of one depth of row controls, filled. */
function fillRows(
	rows: readonly Row[],
	depth: number,
	scope: Scope,
	env: FillEnvironment,
	fill: TableFill,
): Row[] {
	const output: Row[] = [];
	for (const group of groupAt(rows, depth, rowChain)) {
		const control = group.control;
		if (control) {
			const tag = control.tag;
			if (tag !== undefined) env.seen.add(tag);
			const records = control.kind === 'repeatingSectionItem' ? undefined : repeatsOf(tag, env);
			if (records && fill.nested)
				report(env, tag!, 'skipped', scope, {
					reason: 'the rows of a table nested in a cell are not repeated',
				});
			else if (records) {
				report(env, tag!, records.length ? 'repeated' : 'removed', scope, {
					items: records.length,
				});
				// The copies' rows are filled below the repeated control, so they do not repeat again.
				const inside = (copy: Row[], item: Scope) => fillRows(copy, depth + 1, item, env, fill);
				if (control.kind === 'repeatingSection') {
					const items = groupAt(group.items, depth + 1, rowChain);
					const first = items.find((item) => item.control?.kind === 'repeatingSectionItem');
					const same = (item: ContentControl) => item.id === control.id;
					const unit = first?.items ?? group.items;
					output.push(...repeatRows(unit, inside, records, same, same, false, env, fill));
				} else {
					const own = (item: ContentControl) => item.tag === tag;
					output.push(
						...repeatRows(group.items, inside, records, () => false, own, false, env, fill),
					);
				}
				continue;
			}
			if (tag !== undefined && lookup(tag, scope, env) !== undefined)
				report(env, tag, 'skipped', scope, {
					reason: 'a control around table rows holds cells: name the controls inside it',
				});
			const inner = fillRows(group.items, depth + 1, scope, env, fill);
			if (tag !== undefined && isOptional(tag, env) && !fill.nested) {
				const shows = anyControlShowsValue([asTable(inner)], control.id);
				report(env, tag, shows ? 'kept' : 'removed', scope);
				if (!shows) {
					env.removed++;
					fill.structural = true;
					continue;
				}
			}
			output.push(...inner);
			continue;
		}
		for (let at = 0; at < group.items.length; at++) {
			const row = group.items[at]!;
			const key = fill.nested ? undefined : rowKey(row, env);
			if (key === undefined) {
				output.push(...fillRow(row, scope, env, fill));
				continue;
			}
			let end = at + 1;
			while (end < group.items.length && isSubRow(group.items[end]!, key)) end++;
			const records = repeatsOf(key, env)!;
			env.seen.add(key);
			report(env, key, records.length ? 'repeated' : 'removed', scope, { items: records.length });
			const plain = (copy: Row[], item: Scope) =>
				copy.flatMap((each) => fillRow(each, item, env, fill));
			const own = (item: ContentControl) => item.tag === key;
			const unit = group.items.slice(at, end);
			output.push(...repeatRows(unit, plain, records, () => false, own, true, env, fill));
			at = end - 1;
		}
	}
	return output;
}

/** Fills the controls of a table; undefined when no row is left. */
export function fillTable(
	table: Table,
	scope: Scope,
	env: FillEnvironment,
	nested: boolean,
): Table | undefined {
	const fill: TableFill = { table, nested, structural: false };
	const rows = fillRows(fromTable(table), 0, scope, env, fill);
	if (!rows.length) {
		env.removed++;
		return undefined;
	}
	if (fill.structural && !nested)
		env.copies.set(rowEditKey(table.id), { sourceId: '', rekey: (id) => id });
	const properties = rows.some((row) => row.props && Object.keys(row.props).length);
	const { rowProperties: _properties, ...rest } = table;
	return {
		...rest,
		rows: rows.map((row) => row.cells),
		...((properties || table.rowProperties) && {
			rowProperties: rows.map((row) => row.props ?? {}),
		}),
	};
}
