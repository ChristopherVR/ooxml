import { NS } from '../../xml/index.js';
import { formatRange } from '../address.js';
import type { Table, Worksheet } from '../model.js';
import { addFuturePrefixes } from '../read/formula-text.js';
import { XML_HEADER, attrs, escapeText } from './xml-out.js';

/** Excel's own text for a header cell value. */
const headerTextOf = (value: unknown): string | undefined =>
	typeof value === 'string'
		? value
		: typeof value === 'number' || typeof value === 'boolean'
			? String(value).toUpperCase()
			: undefined;

/**
 * Column names Excel will accept: they must equal the header cells' text and be unique (case
 * insensitive). Header cells whose value is not text are written as the column name.
 */
export function resolveTableColumns(
	sheet: Worksheet,
	table: Table,
): { names: string[]; headerText: Map<string, string> } {
	const headerText = new Map<string, string>();
	const used = new Set<string>();
	const width = table.range.end.col - table.range.start.col + 1;
	const names: string[] = [];
	for (let i = 0; i < width; i++) {
		const col = table.range.start.col + i;
		const row = table.range.start.row;
		const fromCell = table.headerRow
			? headerTextOf(sheet.rows.get(row)?.get(col)?.value)
			: undefined;
		let name =
			(fromCell?.trim() ? fromCell : undefined) ?? table.columns[i]?.name ?? `Column${i + 1}`;
		if (!name.trim()) name = `Column${i + 1}`;
		let unique = name;
		for (let n = 2; used.has(unique.toLowerCase()); n++) unique = `${name}${n}`;
		used.add(unique.toLowerCase());
		names.push(unique);
		const cell = sheet.rows.get(row)?.get(col);
		if (table.headerRow && (typeof cell?.value !== 'string' || cell.value !== unique))
			headerText.set(`${row}:${col}`, unique);
	}
	return { names, headerText };
}

export function tableXml(table: Table, id: number, names: readonly string[]): string {
	const range = formatRange(table.range);
	const filterEnd = { ...table.range.end, row: table.range.end.row - (table.totalsRow ? 1 : 0) };
	const filter =
		table.headerRow && !table.totalsRow
			? `<autoFilter ref="${range}"/>`
			: table.headerRow
				? `<autoFilter ref="${formatRange({ start: table.range.start, end: filterEnd })}"/>`
				: '';
	const columns = names
		.map((name, index) => {
			const column = table.columns[index];
			const calculated = column?.calculatedColumnFormula
				? `<calculatedColumnFormula>${escapeText(addFuturePrefixes(column.calculatedColumnFormula))}</calculatedColumnFormula>`
				: '';
			const values = attrs({
				id: index + 1,
				name,
				totalsRowFunction: table.totalsRow ? column?.totalsRowFunction : undefined,
				totalsRowLabel: table.totalsRow ? column?.totalsRowLabel : undefined,
			});
			return calculated
				? `<tableColumn${values}>${calculated}</tableColumn>`
				: `<tableColumn${values}/>`;
		})
		.join('');
	const style = attrs({
		name: table.styleName,
		showFirstColumn: table.showFirstColumn ?? false,
		showLastColumn: table.showLastColumn ?? false,
		showRowStripes: table.showRowStripes ?? true,
		showColumnStripes: table.showColumnStripes ?? false,
	});
	const head = attrs({
		id,
		name: table.name,
		displayName: table.displayName || table.name,
		ref: range,
		headerRowCount: table.headerRow ? undefined : 0,
		totalsRowCount: table.totalsRow ? 1 : undefined,
		totalsRowShown: table.totalsRow ? undefined : false,
	});
	return `${XML_HEADER}<table xmlns="${NS.x}"${head}>${filter}<tableColumns count="${names.length}">${columns}</tableColumns><tableStyleInfo${style}/></table>`;
}

/** `SUBTOTAL` function numbers Excel writes for totals-row functions. */
const SUBTOTAL_CODES: Record<string, number> = {
	average: 101,
	countNums: 102,
	count: 103,
	max: 104,
	min: 105,
	stdDev: 107,
	sum: 109,
	var: 110,
};

/**
 * Totals-row formulas Excel requires to match the column's `totalsRowFunction`
 * (`SUBTOTAL(109,Table[Col])`); keyed `row:col`, only where the cell differs.
 */
export function totalsRowFormulas(
	sheet: Worksheet,
	table: Table,
	names: readonly string[],
): Map<string, string> {
	const out = new Map<string, string>();
	if (!table.totalsRow) return out;
	const row = table.range.end.row;
	names.forEach((name, index) => {
		const code = SUBTOTAL_CODES[table.columns[index]?.totalsRowFunction ?? ''];
		if (code === undefined) return;
		const col = table.range.start.col + index;
		const escaped = name.replace(/['#[\]]/g, "'$&");
		const formula = `SUBTOTAL(${code},${table.displayName || table.name}[${escaped}])`;
		if (sheet.rows.get(row)?.get(col)?.formula !== formula) out.set(`${row}:${col}`, formula);
	});
	return out;
}

/**
 * The `name` and `displayName` to write. Names read from the source and left unedited are kept
 * exactly (Excel itself writes `name="T1" displayName="T1_"`); only names the user created or
 * changed are made safe, since `displayName` is what formulas refer to.
 */
export function resolveTableNames(
	table: Table,
	before: Pick<Table, 'name' | 'displayName'> | undefined,
	index: number,
): { name: string; displayName: string } {
	const display = table.displayName || table.name;
	const displayName =
		before && before.displayName === display ? display : safeTableName(display, index);
	const name =
		before && before.name === table.name
			? table.name
			: table.name === display
				? displayName
				: safeTableName(table.name, index);
	return { name, displayName };
}

/** Excel rejects table names that are not identifiers or that read as cell references. */
export function safeTableName(name: string, index: number): string {
	const valid =
		/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(name) &&
		!/^[A-Za-z]{1,3}\d+$/.test(name) &&
		!/^[RrCc]$/.test(name) &&
		!/^[Rr]\d*[Cc]\d*$/.test(name) &&
		name.length <= 255;
	if (valid) return name;
	const cleaned = name.replace(/[^A-Za-z0-9_.]/g, '_');
	return `Table_${cleaned || index}`;
}
