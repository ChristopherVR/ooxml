// The collecting switches of a `TOC` field instruction: `\o "1-3"` (heading styles), `\u`
// (paragraph outline levels) and `\t "Style,level,..."` (custom styles).

export interface TocSwitches {
	/** `\o` was given: heading styles within `from`-`to` are collected. */
	headings: boolean;
	/** `\u`: paragraphs with an outline level within `from`-`to` are collected. */
	outline: boolean;
	/** `\t`: style name or id (lower case) to the level its paragraphs get. */
	styles: Map<string, number>;
	from: number;
	to: number;
}

/** `\t "Title,1,Subtitle,2"`: names and levels separated by the list separator (`,` or `;`). */
function styleMap(instruction: string): Map<string, number> {
	const map = new Map<string, number>();
	const match = /\\t\s*(?:"([^"]*)"|(\S+))/.exec(instruction);
	const parts = (match?.[1] ?? match?.[2] ?? '').split(/[,;]/).map((part) => part.trim());
	for (let i = 0; i + 1 < parts.length; i += 2) {
		const name = parts[i];
		const level = Number(parts[i + 1]);
		if (name && Number.isInteger(level) && level >= 1 && level <= 9)
			map.set(name.toLowerCase(), level);
	}
	return map;
}

/**
 * What a TOC instruction collects. With no collecting switch at all the field keeps this
 * package's historic reading (every heading level), so a bare `TOC` still lists its headings.
 */
export function tocSwitches(instruction: string): TocSwitches {
	const range = /\\o\s*"?(\d)\s*-\s*(\d)"?/.exec(instruction);
	const hasO = /\\o(?![a-z])/i.test(instruction);
	const outline = /\\u(?![a-z])/i.test(instruction);
	const styles = styleMap(instruction);
	return {
		headings: hasO || (!outline && styles.size === 0),
		outline,
		styles,
		from: range ? Number(range[1]) : 1,
		to: range ? Number(range[2]) : 9,
	};
}
