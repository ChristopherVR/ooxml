import type { HeadingPair } from './types';

/** One `HeadingPairs` group of `docProps/app.xml` with the `TitlesOfParts` entries it owns. */
export interface TitleGroup {
	name: string;
	titles: string[];
}

/**
 * Splits `TitlesOfParts` into the `HeadingPairs` groups that own it, in order: each group takes
 * the next `count` titles (a negative count takes none).
 */
export function titleGroups(
	headingPairs: readonly HeadingPair[],
	titlesOfParts: readonly string[],
): TitleGroup[] {
	let offset = 0;
	return headingPairs.map((pair) => {
		const count = Math.max(0, pair.count);
		const titles = titlesOfParts.slice(offset, offset + count);
		offset += count;
		return { name: pair.name, titles };
	});
}

/** `HeadingPairs` and `TitlesOfParts` for `groups`, each count following its titles. */
export function fromTitleGroups(groups: readonly TitleGroup[]): {
	headingPairs: HeadingPair[];
	titlesOfParts: string[];
} {
	return {
		headingPairs: groups.map((group) => ({ name: group.name, count: group.titles.length })),
		titlesOfParts: groups.flatMap((group) => group.titles),
	};
}

/**
 * Replaces the titles of the first group `match` accepts (Excel's `Worksheets`, PowerPoint's
 * `Slide Titles`...), keeping every other group and its titles; when none matches, a group
 * named `name` is appended. Counts are recomputed so the two vectors stay consistent.
 */
export function replaceTitleGroup(
	headingPairs: readonly HeadingPair[],
	titlesOfParts: readonly string[],
	group: { name: string; titles: readonly string[]; match?: (name: string) => boolean },
): { headingPairs: HeadingPair[]; titlesOfParts: string[] } {
	const match = group.match ?? ((name: string) => name === group.name);
	const groups = titleGroups(headingPairs, titlesOfParts);
	const target = groups.find((entry) => match(entry.name));
	if (target) target.titles = [...group.titles];
	else groups.push({ name: group.name, titles: [...group.titles] });
	return fromTitleGroups(groups);
}
