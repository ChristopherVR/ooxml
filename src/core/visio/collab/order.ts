// Fractional positions keep a list's order in plain JSON entries: a peer that moves or inserts
// one item rewrites only that item, and two peers inserting at the same place tie-break by key.

/**
 * Positions for `order` (keys in local order) that keep every existing position that is still in
 * increasing order and give the others a value between their neighbours.
 */
export function assignPositions(
	order: readonly string[],
	current: (key: string) => number | undefined,
): Map<string, number> {
	const result = new Map<string, number>();
	const kept: (number | undefined)[] = [];
	let last = -Infinity;
	for (const key of order) {
		const pos = current(key);
		if (pos !== undefined && pos > last) {
			kept.push(pos);
			last = pos;
		} else kept.push(undefined);
	}
	let i = 0;
	while (i < order.length) {
		const known = kept[i];
		if (known !== undefined) {
			result.set(order[i]!, known);
			i++;
			continue;
		}
		let end = i;
		while (end < order.length && kept[end] === undefined) end++;
		const lo = i > 0 ? kept[i - 1] : undefined;
		const hi = end < order.length ? kept[end] : undefined;
		const count = end - i;
		for (let j = 0; j < count; j++) {
			const key = order[i + j]!;
			if (lo !== undefined && hi !== undefined)
				result.set(key, lo + ((hi - lo) * (j + 1)) / (count + 1));
			else if (lo !== undefined) result.set(key, lo + j + 1);
			else if (hi !== undefined) result.set(key, hi - count + j);
			else result.set(key, j + 1);
		}
		i = end;
	}
	return result;
}

/** Sorts keys by position, then key, the order every peer reads a list in. */
export function byPosition<T extends { pos: number; key: string }>(items: T[]): T[] {
	return items.sort((a, b) => a.pos - b.pos || (a.key < b.key ? -1 : a.key > b.key ? 1 : 0));
}
