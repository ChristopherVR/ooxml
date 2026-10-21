/** Canonical JSON of a plain value: object keys sorted, `undefined` entries dropped. */
export function stableKey(value: unknown): string {
	return JSON.stringify(sortKeys(value));
}

function sortKeys(value: unknown): unknown {
	if (Array.isArray(value)) return value.map(sortKeys);
	if (value && typeof value === 'object') {
		const out: Record<string, unknown> = {};
		for (const key of Object.keys(value).sort()) {
			const entry = (value as Record<string, unknown>)[key];
			if (entry !== undefined) out[key] = sortKeys(entry);
		}
		return out;
	}
	return value;
}
