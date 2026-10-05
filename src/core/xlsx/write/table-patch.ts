import { NS, buildXml, elements, parseXml, type XmlElement } from '../../xml/index.js';
import { formatRange, parseRange, type CellRange } from '../address.js';
import type { Table, Worksheet } from '../model.js';
import { addFuturePrefixes } from '../read/formula-text.js';
import { num } from './xml-out.js';

const xKids = (parent: XmlElement, local: string): XmlElement[] =>
	elements(parent).filter((e) => e.localName === local);
const xKid = (parent: XmlElement, local: string): XmlElement | undefined => xKids(parent, local)[0];

function setOrRemove(element: XmlElement, name: string, value: string | undefined): void {
	if (value === undefined) element.removeAttribute(name);
	else element.setAttribute(name, value);
}

const flag = (value: boolean | undefined, fallback: boolean): string =>
	(value ?? fallback) ? '1' : '0';

/** Inserts `child` before the first sibling whose local name is in `before` (else appends). */
function insertBefore(parent: XmlElement, child: XmlElement, before: readonly string[]): void {
	const next = elements(parent).find((e) => before.includes(e.localName));
	parent.insertBefore(child, next ?? null);
}

function setChildText(
	doc: Document,
	parent: XmlElement,
	local: string,
	text: string | undefined,
	before: readonly string[],
): void {
	const existing = xKid(parent, local);
	if (text === undefined) {
		if (existing) parent.removeChild(existing);
		return;
	}
	const node = existing ?? doc.createElementNS(NS.x, local);
	node.textContent = text;
	if (!existing) insertBefore(parent, node, before);
}

/** The autoFilter range of a table: the whole table less its totals row. */
export function tableFilterRange(table: Table): CellRange {
	return {
		start: table.range.start,
		end: { ...table.range.end, row: table.range.end.row - (table.totalsRow ? 1 : 0) },
	};
}

function patchAutoFilter(doc: Document, root: XmlElement, table: Table, before: CellRange): void {
	const existing = xKid(root, 'autoFilter');
	if (!table.headerRow) {
		if (existing) root.removeChild(existing);
		return;
	}
	const filter = existing ?? doc.createElementNS(NS.x, 'autoFilter');
	const range = tableFilterRange(table);
	filter.setAttribute('ref', formatRange(range));
	const width = range.end.col - range.start.col + 1;
	for (const column of xKids(filter, 'filterColumn')) {
		const colId = Number(column.getAttribute('colId') ?? '0');
		if (!(colId < width)) filter.removeChild(column);
	}
	if (!existing) insertBefore(root, filter, ['sortState', 'tableColumns']);
	const sort = xKid(filter, 'sortState') ?? xKid(root, 'sortState');
	if (!sort) return;
	const moved =
		before.start.row !== table.range.start.row ||
		before.start.col !== table.range.start.col ||
		before.end.col !== table.range.end.col;
	if (moved) {
		sort.parentNode?.removeChild(sort);
		return;
	}
	// Only the row extent changed: stretch the sort range and its conditions to the data rows.
	const firstData = range.start.row + 1;
	const stretch = (ref: string | null): string | undefined => {
		const parsed = ref ? parseRange(ref) : undefined;
		if (!parsed) return undefined;
		return formatRange({
			start: { row: firstData, col: parsed.start.col },
			end: { row: Math.max(firstData, range.end.row), col: parsed.end.col },
		});
	};
	const sortRef = stretch(sort.getAttribute('ref'));
	if (sortRef) sort.setAttribute('ref', sortRef);
	for (const condition of xKids(sort, 'sortCondition')) {
		const ref = stretch(condition.getAttribute('ref'));
		if (ref) condition.setAttribute('ref', ref);
	}
}

const COLUMN_CHILD_ORDER = ['calculatedColumnFormula', 'totalsRowFormula', 'xmlColumnPr', 'extLst'];

function patchColumns(
	doc: Document,
	root: XmlElement,
	sheet: Worksheet,
	table: Table,
	names: readonly string[],
): void {
	let list = xKid(root, 'tableColumns');
	if (!list) {
		list = doc.createElementNS(NS.x, 'tableColumns');
		insertBefore(root, list, ['tableStyleInfo', 'extLst']);
	}
	const sourceColumns = xKids(list, 'tableColumn');
	const byName = new Map(
		sourceColumns.map((c) => [(c.getAttribute('name') ?? '').toLowerCase(), c]),
	);
	const sameCount = sourceColumns.length === names.length;
	const used = new Set<XmlElement>();
	const ids = new Set<number>();
	const picked = names.map((name, index) => {
		const match = sameCount ? sourceColumns[index] : byName.get(name.toLowerCase());
		if (!match || used.has(match)) return undefined;
		used.add(match);
		const id = Number(match.getAttribute('id'));
		if (Number.isInteger(id) && id > 0 && !ids.has(id)) ids.add(id);
		else match.removeAttribute('id');
		return match;
	});
	for (const column of sourceColumns) list.removeChild(column);
	let nextId = Math.max(0, ...ids) + 1;
	const totalsRow = table.range.end.row;
	names.forEach((name, index) => {
		const column = picked[index] ?? doc.createElementNS(NS.x, 'tableColumn');
		if (!column.getAttribute('id')) column.setAttribute('id', num(nextId++));
		column.setAttribute('name', name);
		const model = table.columns[index];
		const fn = table.totalsRow ? model?.totalsRowFunction : undefined;
		setOrRemove(column, 'totalsRowFunction', fn);
		setOrRemove(column, 'totalsRowLabel', table.totalsRow ? model?.totalsRowLabel : undefined);
		const calculated = model?.calculatedColumnFormula;
		setChildText(
			doc,
			column,
			'calculatedColumnFormula',
			calculated ? addFuturePrefixes(calculated) : undefined,
			COLUMN_CHILD_ORDER.slice(1),
		);
		if (fn !== 'custom') setChildText(doc, column, 'totalsRowFormula', undefined, []);
		else {
			const cellFormula = sheet.rows.get(totalsRow)?.get(table.range.start.col + index)?.formula;
			if (cellFormula)
				setChildText(
					doc,
					column,
					'totalsRowFormula',
					addFuturePrefixes(cellFormula),
					COLUMN_CHILD_ORDER.slice(2),
				);
		}
		list.appendChild(column);
	});
	list.setAttribute('count', num(names.length));
}

function patchStyleInfo(doc: Document, root: XmlElement, table: Table): void {
	let info = xKid(root, 'tableStyleInfo');
	if (!info) {
		info = doc.createElementNS(NS.x, 'tableStyleInfo');
		insertBefore(root, info, ['extLst']);
	}
	setOrRemove(info, 'name', table.styleName);
	info.setAttribute('showFirstColumn', flag(table.showFirstColumn, false));
	info.setAttribute('showLastColumn', flag(table.showLastColumn, false));
	info.setAttribute('showRowStripes', flag(table.showRowStripes, true));
	info.setAttribute('showColumnStripes', flag(table.showColumnStripes, false));
}

/**
 * Rewrites a table part read from the source package, changing only the modelled attributes
 * (id, names, range, header/totals rows, columns, style) and keeping everything else: the
 * autoFilter criteria and sort state, dxf ids, `totalsRowFormula`, `xr:uid`s and `extLst`.
 * Returns `undefined` when the source XML cannot be read (the caller then writes a fresh part).
 */
export function patchTableXml(
	sourceXml: string,
	sheet: Worksheet,
	table: Table,
	id: number,
	names: readonly string[],
): string | undefined {
	let doc: Document;
	try {
		doc = parseXml(sourceXml, { label: 'XLSX table' });
	} catch {
		return undefined;
	}
	const root = doc.documentElement;
	if (root.localName !== 'table') return undefined;
	const before = parseRange(root.getAttribute('ref') ?? '') ?? table.range;
	root.setAttribute('id', num(id));
	root.setAttribute('name', table.name);
	root.setAttribute('displayName', table.displayName || table.name);
	root.setAttribute('ref', formatRange(table.range));
	setOrRemove(root, 'headerRowCount', table.headerRow ? undefined : '0');
	if (table.totalsRow) {
		root.setAttribute('totalsRowCount', '1');
		root.removeAttribute('totalsRowShown');
	} else {
		root.removeAttribute('totalsRowCount');
		if (!root.hasAttribute('totalsRowShown')) root.setAttribute('totalsRowShown', '0');
	}
	patchAutoFilter(doc, root, table, before);
	patchColumns(doc, root, sheet, table, names);
	patchStyleInfo(doc, root, table);
	return buildXml(doc);
}
