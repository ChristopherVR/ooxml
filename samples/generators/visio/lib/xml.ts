/** Small string builders for the ShapeSheet XML of a VSDX part ([MS-VSDX] Cell, Row, Section). */
import { xmlEscape } from '../../lib/package';

export const VISIO_NS = 'http://schemas.microsoft.com/office/visio/2012/main';
export const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** A number as the file stores it: six decimals at most, so the bytes never depend on float noise. */
export function num(value: number): string {
	const rounded = Math.round(value * 1e6) / 1e6;
	return Object.is(rounded, -0) ? '0' : String(rounded);
}

/** `#RRGGBB` as the RGB() formula Visio keeps beside a literal colour. */
export function rgb(hex: string): string {
	const part = (at: number) => parseInt(hex.slice(at, at + 2), 16);
	return `RGB(${part(1)},${part(3)},${part(5)})`;
}

export function attrs(values: Record<string, string | number | undefined>): string {
	let out = '';
	for (const [name, value] of Object.entries(values)) {
		if (value === undefined) continue;
		out += ` ${name}="${xmlEscape(typeof value === 'number' ? num(value) : value)}"`;
	}
	return out;
}

/** One ShapeSheet cell: its cached value, and the formula and unit when it has them. */
export function cell(
	name: string,
	value: string | number,
	formula?: string,
	unit?: string,
): string {
	return `<Cell${attrs({ N: name, V: value, U: unit, F: formula })}/>`;
}

/** A colour cell: the literal and its RGB() formula. */
export function colorCell(name: string, hex: string): string {
	return cell(name, hex.toLowerCase(), rgb(hex));
}

export function row(
	values: { N?: string; IX?: number; T?: string },
	cells: readonly string[],
): string {
	return `<Row${attrs(values)}>${cells.join('')}</Row>`;
}

export function section(
	name: string,
	rows: readonly string[],
	options: { ix?: number; cells?: readonly string[] } = {},
): string {
	if (rows.length === 0 && !options.cells?.length) return '';
	const head = (options.cells ?? []).join('');
	return `<Section${attrs({ N: name, IX: options.ix })}>${head}${rows.join('')}</Section>`;
}

/** A string literal inside a ShapeSheet formula. */
export const quoted = (text: string): string => `"${text.replaceAll('"', '""')}"`;

/** Points to inches, the unit of every length cell. */
export const pt = (points: number): number => points / 72;

/** Cells from several sources as one list: a later cell replaces an earlier one of the same name. */
export function mergeCells(...lists: readonly (readonly string[] | undefined)[]): string[] {
	const byName = new Map<string, string>();
	for (const text of lists.flatMap((list) => list ?? [])) {
		const name = /^<Cell N="([^"]+)"/u.exec(text)?.[1];
		if (!name) throw new Error(`Not a cell: ${text}`);
		byName.delete(name);
		byName.set(name, text);
	}
	return [...byName.values()];
}
