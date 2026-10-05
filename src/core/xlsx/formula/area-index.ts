// A spatial index of sheet areas: which items (dependent formulas, spill footprints) cover a cell
// or range, without scanning every item per lookup.
import { cellKey, type CellRange, keyToAddress, rangesIntersect } from '../address.js';
import type { Area } from './values.js';

interface Entry<T> {
	range: CellRange;
	node: T;
}

/** Ranges at most this wide are indexed per column; wider ones are checked one by one. */
const BUCKET_WIDTH = 64;
/** Rows per chunk inside a column; ranges spanning at most `CHUNKS` chunks are chunked. */
const CHUNK_SHIFT = 6;
const CHUNKS = 64;

function push<K, V>(map: Map<K, V[]>, key: K, value: V): void {
	const list = map.get(key);
	if (list) list.push(value);
	else map.set(key, [value]);
}

function inner<K, V>(map: Map<number, Map<K, V>>, sheet: number): Map<K, V> {
	let found = map.get(sheet);
	if (!found) map.set(sheet, (found = new Map()));
	return found;
}

/**
 * Items keyed by the areas they cover. Short single-row ranges are found by cell key; other
 * narrow ranges sit in per-column row chunks, tall ranges in per-column lists and very wide
 * ranges in one list per sheet.
 */
export class AreaIndex<T> {
	private readonly cells = new Map<number, Map<number, T[]>>();
	private readonly chunked = new Map<number, Map<number, Map<number, Entry<T>[]>>>();
	private readonly tall = new Map<number, Map<number, Entry<T>[]>>();
	private readonly wide = new Map<number, Entry<T>[]>();

	add(area: Area, node: T): void {
		const { start, end } = area.range;
		if (start.row === end.row && end.col - start.col + 1 <= BUCKET_WIDTH) {
			const cells = inner(this.cells, area.sheet);
			for (let col = start.col; col <= end.col; col++) push(cells, cellKey(start.row, col), node);
			return;
		}
		const entry = { range: area.range, node };
		if (end.col - start.col + 1 > BUCKET_WIDTH) {
			const list = this.wide.get(area.sheet);
			if (list) list.push(entry);
			else this.wide.set(area.sheet, [entry]);
			return;
		}
		const first = start.row >> CHUNK_SHIFT;
		const last = end.row >> CHUNK_SHIFT;
		if (last - first + 1 > CHUNKS) {
			const cols = inner(this.tall, area.sheet);
			for (let c = start.col; c <= end.col; c++) push(cols, c, entry);
			return;
		}
		const cols = inner(this.chunked, area.sheet);
		for (let c = start.col; c <= end.col; c++) {
			let chunks = cols.get(c);
			if (!chunks) cols.set(c, (chunks = new Map()));
			for (let k = first; k <= last; k++) push(chunks, k, entry);
		}
	}

	/** Items whose areas intersect `range` on `sheet` (each once). */
	dependents(sheet: number, range: CellRange, out: Set<T> | T[]): void {
		const found = out instanceof Set ? out : new Set<T>();
		this.collect(sheet, range, found);
		if (Array.isArray(out)) for (const node of found) out.push(node);
	}

	private collect(sheet: number, range: CellRange, out: Set<T>): void {
		const { start, end } = range;
		const check = (list: Entry<T>[] | undefined): void => {
			if (list)
				for (const entry of list) if (rangesIntersect(entry.range, range)) out.add(entry.node);
		};
		const cells = this.cells.get(sheet);
		if (cells) {
			const size = (end.row - start.row + 1) * (end.col - start.col + 1);
			if (size <= 4096) {
				for (let r = start.row; r <= end.row; r++) {
					for (let c = start.col; c <= end.col; c++) {
						const list = cells.get(cellKey(r, c));
						if (list) for (const node of list) out.add(node);
					}
				}
			} else {
				for (const [key, list] of cells) {
					const { row, col } = keyToAddress(key);
					if (row >= start.row && row <= end.row && col >= start.col && col <= end.col)
						for (const node of list) out.add(node);
				}
			}
		}
		const columns = (map: Map<number, unknown> | undefined, visit: (c: number) => void): void => {
			if (!map) return;
			if (end.col - start.col + 1 <= map.size) {
				for (let c = start.col; c <= end.col; c++) if (map.has(c)) visit(c);
			} else for (const c of map.keys()) if (c >= start.col && c <= end.col) visit(c);
		};
		const chunked = this.chunked.get(sheet);
		const first = start.row >> CHUNK_SHIFT;
		const last = end.row >> CHUNK_SHIFT;
		columns(chunked, (c) => {
			const chunks = chunked?.get(c) as Map<number, Entry<T>[]>;
			if (last - first + 1 <= chunks.size) {
				for (let k = first; k <= last; k++) check(chunks.get(k));
			} else for (const [k, list] of chunks) if (k >= first && k <= last) check(list);
		});
		const tall = this.tall.get(sheet);
		columns(tall, (c) => check(tall?.get(c)));
		check(this.wide.get(sheet));
	}
}
