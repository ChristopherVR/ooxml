/**
 * Paragraphs from lines: consecutive lines stay together while their spacing, size, weight and
 * left edge match and the line before reaches the right margin; hyphenated line ends rejoin. Each
 * paragraph gets its alignment, indents and spacing from where its lines sit in the text area.
 */
import { markerOf } from './list-markers';
import type { Fragment, Line } from './lines';
import type { ListItem } from './lists';

export interface TextArea {
	left: number;
	right: number;
}

export type Alignment = 'left' | 'center' | 'right' | 'both';

export interface ParagraphBlock {
	kind: 'paragraph';
	page: number;
	lines: Line[];
	fragments: Fragment[];
	align: Alignment;
	/** Points from the text area's left edge. */
	indentLeft: number;
	/** First line relative to the others (negative hangs). */
	firstLine: number;
	/** Extra space above the paragraph, in points. */
	spaceBefore: number;
	/** Line spacing as a multiple of single spacing (1 for one-line paragraphs). */
	lineSpacing: number;
	size: number;
	heading?: 1 | 2 | 3;
	/** The paragraph continues on the next page in the PDF. */
	continues?: boolean;
	/** A list item (its typed marker removed), set by `assignLists`. */
	list?: ListItem;
}

const TERMINAL = /[.!?:;\u2026\u00bb\u201d"')\]]\s*$/;
/** Single spacing as a multiple of the font size, by family (the fonts' line heights). */
function singleSpacing(line: Line): number {
	const family = (
		line.fragments.find((fragment) => fragment.text.trim())?.font.family ?? ''
	).toLowerCase();
	if (/calibri|cambria|candara|consolas|corbel|constantia/.test(family)) return 1.22;
	if (/arial|helvetica|times|liberation|courier/.test(family)) return 1.15;
	return 1.17;
}

/** Joins a paragraph's lines into fragments, rejoining hyphenated words and adding line spaces. */
function joinLines(lines: Line[]): Fragment[] {
	const out: Fragment[] = [];
	lines.forEach((line, index) => {
		const fragments = line.fragments.map((fragment) => ({ ...fragment }));
		const next = lines[index + 1];
		const last = fragments[fragments.length - 1];
		if (last && next) {
			const hyphen =
				/[-\u00ad]$/.test(last.text.trimEnd()) && /^\s*[a-z\u00df-\u00ff]/.test(next.text);
			if (hyphen) last.text = last.text.trimEnd().replace(/[-\u00ad]$/, '');
			else if (!/\s$/.test(last.text)) last.text += ' ';
		}
		out.push(...fragments);
	});
	if (out.length) out[out.length - 1]!.text = out[out.length - 1]!.text.trimEnd();
	return out;
}

function alignment(lines: Line[], area: TextArea): Alignment {
	const width = area.right - area.left;
	const tolerance = Math.max(3, lines[0]!.size * 0.4);
	const centered = lines.every((line) => {
		const before = line.left - area.left;
		const after = area.right - line.right;
		return before > 2 * line.size && Math.abs(before - after) < tolerance * 1.5;
	});
	if (centered) return 'center';
	const flushRight = lines.every((line) => Math.abs(area.right - line.right) < tolerance);
	const leftEdges = new Set(lines.map((line) => Math.round(line.left)));
	if (
		flushRight &&
		(lines.length === 1 ? lines[0]!.left - area.left > width * 0.3 : leftEdges.size > 1)
	)
		return 'right';
	if (
		lines.length > 1 &&
		lines.slice(0, -1).every((line) => Math.abs(area.right - line.right) < tolerance)
	)
		return 'both';
	return 'left';
}

export function buildParagraphs(lines: Line[], area: TextArea, bodySize: number): ParagraphBlock[] {
	const groups: Line[][] = [];
	let current: Line[] = [];
	for (const line of lines) {
		const previous = current[current.length - 1];
		if (!previous) {
			current.push(line);
			continue;
		}
		const spacing = line.baseline - previous.baseline;
		const expected =
			current.length > 1
				? current[current.length - 1]!.baseline - current[current.length - 2]!.baseline
				: previous.size * singleSpacing(previous) * 1.2;
		const sizeChange = Math.abs(line.size - previous.size) > 0.1 * previous.size;
		const weightChange = line.bold !== previous.bold;
		const shortLine =
			previous.right < area.right - Math.max(3 * previous.size, 0.12 * (area.right - area.left));
		const firstLineIndent = current.length === 1 && line.left < previous.left - 1;
		const listContinuation = !!markerOf(current[0]!) && line.left > current[0]!.left + 2;
		const leftShift =
			Math.abs(line.left - previous.left) > previous.size && !firstLineIndent && !listContinuation;
		const centeredRun =
			alignment([previous], area) === 'center' && alignment([line], area) === 'center';
		const brk =
			spacing > expected * 1.3 + 1 ||
			spacing < 0 ||
			sizeChange ||
			weightChange ||
			!!markerOf(line) ||
			(shortLine && !centeredRun) ||
			(leftShift && !centeredRun);
		if (brk) {
			groups.push(current);
			current = [line];
		} else current.push(line);
	}
	if (current.length) groups.push(current);

	const blocks: ParagraphBlock[] = [];
	let previousBottom: number | undefined;
	let previousSpacing = 0;
	for (const group of groups) {
		const first = group[0]!;
		const align = alignment(group, area);
		const rest = group.length > 1 ? group.slice(1) : group;
		const restLeft = Math.min(...rest.map((line) => line.left));
		const indentLeft = align === 'left' || align === 'both' ? Math.max(0, restLeft - area.left) : 0;
		const firstLine =
			align === 'left' || align === 'both' ? first.left - area.left - indentLeft : 0;
		const gaps = group.slice(1).map((line, index) => line.baseline - group[index]!.baseline);
		const meanGap = gaps.length ? gaps.reduce((sum, gap) => sum + gap, 0) / gaps.length : 0;
		const size = Math.max(...group.map((line) => line.size));
		const lineSpacing = gaps.length
			? Math.min(3, Math.max(1, meanGap / (size * singleSpacing(first))))
			: 1;
		const normal = previousSpacing || first.size * singleSpacing(first);
		const spaceBefore =
			previousBottom === undefined ? 0 : Math.max(0, first.baseline - previousBottom - normal);
		const block: ParagraphBlock = {
			kind: 'paragraph',
			page: first.page,
			lines: group,
			fragments: joinLines(group),
			align,
			indentLeft: Math.abs(indentLeft) < 2 ? 0 : indentLeft,
			firstLine: Math.abs(firstLine) < 2 ? 0 : firstLine,
			spaceBefore: spaceBefore < 2 ? 0 : spaceBefore,
			lineSpacing: Math.round(lineSpacing * 20) / 20,
			size,
		};
		// Larger or bold short lines are headings; italic display text (a pull quote) is not.
		const italic = block.fragments.every(
			(fragment) => fragment.font.italic || !fragment.text.trim(),
		);
		if (
			group.length <= 3 &&
			!italic &&
			(size >= bodySize * 1.25 ||
				(first.bold && size >= bodySize * 1.1 && group.length === 1 && !TERMINAL.test(first.text)))
		)
			block.heading = 1;
		blocks.push(block);
		previousBottom = group[group.length - 1]!.baseline;
		previousSpacing = meanGap || first.size * singleSpacing(first) * lineSpacing;
	}
	return blocks;
}

/** Ranks heading candidates by size: the largest is Heading 1, then 2 and 3. */
export function rankHeadings(blocks: ParagraphBlock[]): void {
	const sizes = [
		...new Set(
			blocks.filter((block) => block.heading).map((block) => Math.round(block.size * 2) / 2),
		),
	].sort((a, b) => b - a);
	for (const block of blocks)
		if (block.heading)
			block.heading = Math.min(3, sizes.indexOf(Math.round(block.size * 2) / 2) + 1) as 1 | 2 | 3;
}

/** Whether a paragraph ending a page goes on at the top of the next. */
export function continuesOnNextPage(last: ParagraphBlock, next: ParagraphBlock): boolean {
	const end = last.lines[last.lines.length - 1]!;
	const start = next.lines[0]!;
	return (
		!last.heading &&
		!next.heading &&
		last.align === next.align &&
		Math.abs(end.size - start.size) < 0.5 &&
		end.bold === start.bold &&
		!TERMINAL.test(end.text) &&
		!markerOf(start)
	);
}
