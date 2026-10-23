/**
 * Keeps the Excel tables of a worksheet consistent with cells the chart write-back changed. Word
 * and PowerPoint put a chart's data in a table (`Table1` over `A1:D5`); Excel repairs a table
 * whose column names differ from its header cells, so a renamed series renames its column, and a
 * write that extends the data right or down past the table's edge (an added series or category)
 * grows the table over it, as Excel's own data window does.
 *
 * @module chart/workbook/tables
 */
import type JSZip from 'jszip';
import { parseRelationships, relationshipsPartFor, resolvePartPath } from '../../opc/relationships';
import { buildXml, parseXml, type XmlDocument, type XmlElement } from '../../xml/index';
import { formatCellAddress, parseCellAddress, type CellAddress } from './cellref';
import { childNamed, childrenNamed, worksheetData, type ChartCellWrite } from './sheet-cells';

interface Box {
	start: CellAddress;
	end: CellAddress;
}

function parseBox(ref: string | null): Box | undefined {
	const [from, to] = (ref ?? '').split(':');
	const start = from ? parseCellAddress(from) : undefined;
	const end = to ? parseCellAddress(to) : start;
	return start && end ? { start, end } : undefined;
}

const formatBox = (box: Box) => `${formatCellAddress(box.start)}:${formatCellAddress(box.end)}`;

/** The furthest index past `end` reached by consecutive written indexes `end + 1, end + 2...`. */
function grown(end: number, written: ReadonlySet<number>): number {
	let next = end;
	while (written.has(next + 1)) next++;
	return next;
}

/** The text a header cell shows. */
function cellText(doc: XmlDocument, ref: string, shared: readonly string[]): string {
	for (const row of childrenNamed(worksheetData(doc), 'row'))
		for (const cell of childrenNamed(row, 'c')) {
			if (cell.getAttribute('r') !== ref) continue;
			const type = cell.getAttribute('t');
			const value = childNamed(cell, 'v')?.textContent ?? '';
			if (type === 's') return shared[Number(value)] ?? '';
			if (type === 'inlineStr') return childNamed(childNamed(cell, 'is'), 't')?.textContent ?? '';
			return value;
		}
	return '';
}

/** Applies the header and size changes to one table part; returns whether it changed. */
function syncTable(
	table: XmlElement,
	doc: XmlDocument,
	writes: ReadonlyMap<string, string>,
	shared: readonly string[],
): boolean {
	const box = parseBox(table.getAttribute('ref'));
	const columns = childNamed(table, 'tableColumns');
	if (!box || !columns) return false;
	const cells = [...writes.keys()].flatMap((ref) => parseCellAddress(ref) ?? []);
	const cols = new Set(
		cells
			.filter((cell) => cell.row >= box.start.row && cell.row <= box.end.row + 1)
			.map((cell) => cell.col),
	);
	const rows = new Set(
		cells
			.filter((cell) => cell.col >= box.start.col && cell.col <= box.end.col + 1)
			.map((cell) => cell.row),
	);
	const next: Box = {
		start: box.start,
		end: { col: grown(box.end.col, cols), row: grown(box.end.row, rows) },
	};
	let changed = false;
	const list = childrenNamed(columns, 'tableColumn');
	const used = new Set<string>();
	let id = list.reduce((max, column) => Math.max(max, Number(column.getAttribute('id')) || 0), 0);
	for (let col = next.start.col; col <= next.end.col; col++) {
		const ref = formatCellAddress({ col, row: next.start.row });
		let column = list[col - next.start.col];
		const header = writes.get(ref) ?? column?.getAttribute('name') ?? cellText(doc, ref, shared);
		// An existing column keeps its name (Word names the category column ' ') unless its header
		// was written; a new column with an empty header gets Excel's ColumnN.
		let name =
			header.trim() || (column && !writes.has(ref)) ? header : `Column${col - next.start.col + 1}`;
		for (let n = 2; used.has(name.toLowerCase()); n++) name = `${header}${n}`;
		used.add(name.toLowerCase());
		if (!column) {
			column = table.ownerDocument!.createElementNS(
				table.namespaceURI,
				columns.prefix ? `${columns.prefix}:tableColumn` : 'tableColumn',
			);
			column.setAttribute('id', String(++id));
			columns.appendChild(column);
			changed = true;
		}
		if (column.getAttribute('name') !== name) {
			column.setAttribute('name', name);
			changed = true;
		}
	}
	if (next.end.col !== box.end.col || next.end.row !== box.end.row) {
		const ref = formatBox(next);
		table.setAttribute('ref', ref);
		childNamed(table, 'autoFilter')?.setAttribute('ref', ref);
		columns.setAttribute('count', String(childrenNamed(columns, 'tableColumn').length));
		changed = true;
	}
	return changed;
}

/**
 * Brings every table of the worksheet `sheetPath` (already holding `writes`) in line with them,
 * and widens the worksheet's `dimension` over the written cells.
 */
export async function syncWorksheetTables(
	zip: JSZip,
	sheetPath: string,
	doc: XmlDocument,
	writes: readonly ChartCellWrite[],
	shared: readonly string[],
): Promise<void> {
	const values = new Map(writes.map((write) => [write.ref, write.value]));
	const dimension = childNamed(doc.documentElement ?? undefined, 'dimension');
	const area = parseBox(dimension?.getAttribute('ref') ?? null);
	if (dimension && area) {
		for (const cell of [...values.keys()].flatMap((ref) => parseCellAddress(ref) ?? [])) {
			area.end = { col: Math.max(area.end.col, cell.col), row: Math.max(area.end.row, cell.row) };
		}
		dimension.setAttribute('ref', formatBox(area));
	}
	const rels = parseRelationships(await zip.file(relationshipsPartFor(sheetPath))?.async('string'));
	for (const rel of rels.values()) {
		if (rel.mode === 'External' || !rel.type.endsWith('/table')) continue;
		const path = resolvePartPath(sheetPath, rel.target);
		const xml = await zip.file(path)?.async('string');
		if (!xml) continue;
		const tableDoc = parseXml(xml, { label: 'Table' });
		const table = tableDoc.documentElement;
		if (table && syncTable(table, doc, values, shared)) zip.file(path, buildXml(tableDoc));
	}
}

/** The texts of the workbook's shared strings, in index order (rich runs joined). */
export async function sharedStringTexts(zip: JSZip): Promise<string[]> {
	const xml = await zip.file('xl/sharedStrings.xml')?.async('string');
	if (!xml) return [];
	return childrenNamed(parseXml(xml).documentElement ?? undefined, 'si').map((item) => {
		const direct = childNamed(item, 't');
		if (direct) return direct.textContent ?? '';
		return childrenNamed(item, 'r')
			.map((run) => childNamed(run, 't')?.textContent ?? '')
			.join('');
	});
}
