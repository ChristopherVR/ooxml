/**
 * Dot leaders (a table of contents, an index): a run of dots between a line's text and a short
 * item at its right end, such as a page number, becomes a tab to a right-aligned tab stop with a
 * dot leader, as Word writes it.
 */
import type { Line } from './lines';
import type { ParagraphBlock, TextArea } from './paragraphs';

const DOTS = /^[\s.\u2026\u00b7_]*$/;

export interface LeaderTab {
	/** Points from the text area's left edge. */
	position: number;
	leader: 'dot' | 'underscore';
}

export function applyLeaders(lines: Line[]): void {
	for (const line of lines) {
		const { fragments } = line;
		let last = fragments.length;
		while (last > 0 && !fragments[last - 1]!.text.trim()) last--;
		let end = last;
		while (end > 0 && !DOTS.test(fragments[end - 1]!.text)) end--;
		let start = end;
		while (start > 0 && DOTS.test(fragments[start - 1]!.text)) start--;
		const leader = fragments
			.slice(start, end)
			.map((fragment) => fragment.text)
			.join('');
		const tail = fragments
			.slice(end, last)
			.map((fragment) => fragment.text)
			.join('')
			.trim();
		const dots = leader.replace(/\s/g, '');
		if (start === 0 || dots.length < 4 || !tail || tail.length > 12) continue;
		const kind: LeaderTab['leader'] = dots.startsWith('_') ? 'underscore' : 'dot';
		const before = fragments[start - 1]!;
		before.text = before.text.replace(/\s+$/, '');
		// The tab keeps where the right-aligned item ends, for the paragraph's tab stop.
		const tab = { ...fragments[start]!, text: '\t', leader: { kind, right: line.right } };
		const rest = fragments
			.slice(end)
			.map((fragment, index) =>
				index === 0 ? { ...fragment, text: fragment.text.replace(/^\s+/, '') } : fragment,
			);
		line.fragments = [...fragments.slice(0, start), tab, ...rest];
		line.text = line.fragments.map((fragment) => fragment.text).join('');
	}
}

/** The leader tab stops of a paragraph whose lines had dot leaders. */
export function leaderTabs(blocks: ParagraphBlock[], area: TextArea): void {
	for (const block of blocks) {
		const ends = block.fragments.flatMap((fragment) => fragment.leader ?? []);
		if (!ends.length) continue;
		const right = Math.max(...ends.map((item) => item.right));
		block.tabs = [{ position: Math.min(right, area.right) - area.left, leader: ends[0]!.kind }];
	}
}
