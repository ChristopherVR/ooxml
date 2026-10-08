/** Unique header names: blanks become `ColumnN`, repeats get a number suffix (`Name2`). */
export function uniqueHeaders(texts: string[]): string[] {
	const used = new Set<string>();
	return texts.map((text, i) => {
		let base = text.trim() || `Column${i + 1}`;
		if (!used.has(base.toLowerCase())) {
			used.add(base.toLowerCase());
			return base;
		}
		for (let n = 2; ; n++) {
			const candidate = `${base}${n}`;
			if (!used.has(candidate.toLowerCase())) {
				used.add(candidate.toLowerCase());
				base = candidate;
				return base;
			}
		}
	});
}
