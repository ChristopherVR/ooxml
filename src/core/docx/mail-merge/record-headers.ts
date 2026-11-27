// Row 1 of a recipient list as field names, shared by the in-memory and the streamed readers.

export function readHeaders(cells: Iterable<readonly [number, string]>): Map<number, string> {
	const fields = new Map<number, string>();
	const seen = new Set<string>();
	for (const [column, raw] of cells) {
		const field = raw.trim();
		if (!field) continue;
		if (seen.has(field)) throw new Error(`Duplicate mail merge field: ${field}`);
		seen.add(field);
		fields.set(column, field);
	}
	return fields;
}
