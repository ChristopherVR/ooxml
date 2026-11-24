/**
 * Running heads and feet: lines near the top or bottom of the page that repeat on most pages
 * (numbers aside) become the header or footer, and the page number in them a PAGE field.
 */
import type { Line } from './lines';

export interface RepeatedStory {
	kind: 'header' | 'footer';
	lines: Line[];
	/** Distance from the page edge to the story, in points. */
	distance: number;
}

const BAND = 0.12;
const key = (line: Line) => line.text.replace(/\d+/g, '#').replace(/\s+/g, ' ').trim();

/**
 * Removes repeated top and bottom lines from `pages` (in place) and returns the header and footer
 * taken from the first page that has them.
 */
export function extractStories(pages: { lines: Line[]; height: number }[]): RepeatedStory[] {
	if (pages.length < 2) return [];
	const needed = Math.max(2, Math.ceil(pages.length * 0.5));
	const stories: RepeatedStory[] = [];
	for (const kind of ['header', 'footer'] as const) {
		const inBand = (line: Line, height: number) =>
			kind === 'header' ? line.bottom < height * BAND : line.top > height * (1 - BAND);
		const counts = new Map<string, { pages: Set<number>; y: number }>();
		pages.forEach((page, index) => {
			for (const line of page.lines) {
				if (!inBand(line, page.height)) continue;
				const id = `${key(line)}@${Math.round(line.baseline / 6)}`;
				const entry = counts.get(id) ?? { pages: new Set<number>(), y: line.baseline };
				entry.pages.add(index);
				counts.set(id, entry);
			}
		});
		const repeated = new Set(
			[...counts].filter(([, entry]) => entry.pages.size >= needed).map(([id]) => id),
		);
		if (!repeated.size) continue;
		let story: Line[] | undefined;
		let height = 0;
		pages.forEach((page) => {
			const taken = page.lines.filter(
				(line) =>
					inBand(line, page.height) &&
					repeated.has(`${key(line)}@${Math.round(line.baseline / 6)}`),
			);
			if (!taken.length) return;
			if (!story) {
				story = taken;
				height = page.height;
			}
			page.lines = page.lines.filter((line) => !taken.includes(line));
		});
		if (!story) continue;
		const lines = story as Line[];
		stories.push({
			kind,
			lines,
			distance:
				kind === 'header'
					? Math.max(0, Math.min(...lines.map((line) => line.top)))
					: Math.max(0, height - Math.max(...lines.map((line) => line.bottom))),
		});
	}
	return stories;
}
