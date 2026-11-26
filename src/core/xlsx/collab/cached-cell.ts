import type * as Y from 'yjs';
import type { CellEntry } from './codec';
import { isRecord, jsonEqual, setIfChanged } from './schema';

// Keep deferred results on sibling keys so they can be published without replacing the Yjs
// item that Undo owns. Matching the base entry also lets Redo reuse the settled result.
const PREFIX = 'cache:';
export const cacheKey = (address: string): string => PREFIX + address;
export const cacheAddress = (key: string): string | undefined =>
	key.startsWith(PREFIX) ? key.slice(PREFIX.length) : undefined;

function content({
	v: _value,
	t: _type,
	d: _dynamic,
	...rest
}: Record<string, unknown>): Record<string, unknown> {
	return rest;
}

export function readSharedCell(cells: Y.Map<unknown>, address: string): unknown {
	const base = cells.get(address);
	const cache = cells.get(cacheKey(address));
	if (!isRecord(base) || !isRecord(cache) || !jsonEqual(cache['base'], base)) return base;
	return { ...content(base), v: cache['v'], t: cache['t'], d: cache['d'] };
}

export function writeSettledCell(cells: Y.Map<unknown>, address: string, entry: CellEntry): void {
	const base = cells.get(address);
	if (!isRecord(base) || (base['f'] === undefined && base['v'] !== '#BUSY!')) return;
	const { v, t, d, ...rest } = entry;
	if (jsonEqual(rest, content(base)) && !jsonEqual(entry, readSharedCell(cells, address)))
		setIfChanged(cells, cacheKey(address), { base, v, t, d });
}
