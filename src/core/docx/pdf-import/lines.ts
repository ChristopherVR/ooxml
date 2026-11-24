/**
 * Text lines from a page's spans: spans on one baseline join into a line, left to right, with a
 * space where the gap between them is a word gap and a tab where it is a column gap. Overprinted
 * copies (a producer's fake bold) are dropped, and small raised or lowered spans join the line
 * they belong to as super- or subscripts.
 */
import type { PdfFontInfo, PdfPageLink, PdfPageRule, PdfTextSpan } from '../../pdf/index';

export interface Fragment {
	text: string;
	x: number;
	width: number;
	size: number;
	font: PdfFontInfo;
	color: string;
	script?: 'super' | 'sub';
	underline?: boolean;
	strike?: boolean;
	link?: string;
	/** An internal link's destination, until it is resolved to a bookmark (`anchor`). */
	target?: { page: number; y?: number };
	/** The bookmark an internal link goes to. */
	anchor?: string;
	/** A tab that replaced a dot leader, and where the item after it ends. */
	leader?: { kind: 'dot' | 'underscore'; right: number };
	/** The fragment is the reference mark of a footnote (or an endnote). */
	note?: { id: number; mark: string; custom: boolean; kind?: 'endnote' };
	/** Tagged PDF: the structure role and element of the span (see `PdfTextSpan`). */
	role?: string;
	element?: number;
}

export interface Line {
	page: number;
	baseline: number;
	top: number;
	bottom: number;
	left: number;
	right: number;
	/** The size most of the line's text has. */
	size: number;
	fragments: Fragment[];
	text: string;
	bold: boolean;
}

/** Spans off the horizontal by more than this (radians) are rotated text. */
const HORIZONTAL = 0.05;

function dominantSize(fragments: Fragment[]): number {
	const weight = new Map<number, number>();
	for (const fragment of fragments) {
		const size = Math.round(fragment.size * 2) / 2;
		weight.set(size, (weight.get(size) ?? 0) + fragment.text.trim().length);
	}
	let best = fragments[0]?.size ?? 11;
	let max = -1;
	for (const [size, count] of weight) if (count > max) [best, max] = [size, count];
	return best;
}

function finish(page: number, baseline: number, spans: PdfTextSpan[]): Line {
	spans.sort((a, b) => a.x - b.x);
	const fragments: Fragment[] = [];
	let previous: PdfTextSpan | undefined;
	for (const span of spans) {
		// A copy drawn over itself (fake bold, shadows) adds nothing.
		if (previous && span.text === previous.text && Math.abs(span.x - previous.x) < 1) continue;
		const lineSize = Math.max(span.fontSize, previous?.fontSize ?? 0);
		if (previous) {
			const gap = span.x - (previous.x + previous.width);
			const last = fragments[fragments.length - 1]!;
			const spaced = /\s$/.test(last.text) || /^\s/.test(span.text);
			if (gap > Math.max(2 * lineSize, 18)) last.text = `${last.text.replace(/\s+$/, '')}\t`;
			else if (gap > 0.15 * lineSize && !spaced) last.text += ' ';
		}
		const shift = baseline - span.y - span.rise;
		const raised = shift > 0.15 * span.fontSize || span.rise > 0.1 * span.fontSize;
		const lowered = shift < -0.1 * span.fontSize || span.rise < -0.1 * span.fontSize;
		const fragment: Fragment = {
			text: span.text,
			x: span.x,
			width: span.width,
			size: span.fontSize,
			font: span.font,
			color: span.color,
		};
		if (span.role !== undefined) fragment.role = span.role;
		if (span.element !== undefined) fragment.element = span.element;
		fragments.push(fragment);
		if (raised || lowered) fragment.script = raised ? 'super' : 'sub';
		previous = span;
	}
	const size = dominantSize(fragments);
	for (const fragment of fragments) {
		if (fragment.script && fragment.size > size * 0.85) delete fragment.script;
		// Word draws a super- or subscript smaller itself: the run keeps the line's size.
		else if (fragment.script) fragment.size = size;
	}
	const left = fragments[0]?.x ?? 0;
	// The right edge is where the ink ends: trailing spaces (which producers draw) do not count.
	const last = [...fragments].reverse().find((fragment) => fragment.text.trim());
	const kept = last ? last.text.trimEnd().length / Math.max(1, last.text.length) : 1;
	const right = last ? last.x + last.width * kept : left;
	const text = fragments.map((fragment) => fragment.text).join('');
	return {
		page,
		baseline,
		top: baseline - size * 0.8,
		bottom: baseline + size * 0.25,
		left,
		right,
		size,
		fragments,
		text,
		bold: fragments.every((fragment) => fragment.font.bold || !fragment.text.trim()),
	};
}

/** The page's text lines, top to bottom; `rotated` counts the spans left out. */
export function pageLines(
	page: number,
	spans: readonly PdfTextSpan[],
): { lines: Line[]; rotated: number } {
	let rotated = 0;
	const horizontal = spans.filter((span) => {
		const keep = Math.abs(span.angle) < HORIZONTAL && span.text.length > 0;
		if (!keep && span.text.trim()) rotated++;
		return keep;
	});
	horizontal.sort((a, b) => a.y - b.y || a.x - b.x);
	const groups: { baseline: number; size: number; spans: PdfTextSpan[] }[] = [];
	for (const span of horizontal) {
		// A span joins a line whose baseline is within a third of its size, or (smaller text) whose
		// body it sits in: a superscript or subscript drawn on its own baseline.
		const group = groups.find(
			(candidate) =>
				Math.abs(candidate.baseline - span.y) < 0.35 * Math.min(candidate.size, span.fontSize) ||
				(span.fontSize < 0.85 * candidate.size &&
					span.y > candidate.baseline - candidate.size * 0.9 &&
					span.y < candidate.baseline + candidate.size * 0.4) ||
				// A line of only smaller text found first (a superscript above its line's baseline).
				(candidate.size < 0.85 * span.fontSize &&
					candidate.baseline > span.y - span.fontSize * 0.9 &&
					candidate.baseline < span.y + span.fontSize * 0.4),
		);
		if (group) {
			group.spans.push(span);
			if (span.fontSize > group.size) {
				group.size = span.fontSize;
				group.baseline = span.y;
			}
		} else groups.push({ baseline: span.y, size: span.fontSize, spans: [span] });
	}
	const lines = groups
		.map((group) => finish(page, group.baseline, group.spans))
		.filter((line) => line.text.trim())
		.sort((a, b) => a.baseline - b.baseline);
	return { lines, rotated };
}

/**
 * Marks fragments that a thin rule runs under (underline) or through (strike-through); returns the
 * rules that became one.
 */
export function applyRules(lines: Line[], rules: readonly PdfPageRule[]): Set<PdfPageRule> {
	const used = new Set<PdfPageRule>();
	const thin = rules.filter((rule) => rule.y1 - rule.y0 <= 1.5 && rule.x1 - rule.x0 > 2);
	for (const rule of thin) {
		const y = (rule.y0 + rule.y1) / 2;
		for (const line of lines) {
			if (y < line.top || y > line.bottom + 2) continue;
			for (const fragment of line.fragments) {
				const overlap =
					Math.min(rule.x1, fragment.x + fragment.width) - Math.max(rule.x0, fragment.x);
				if (overlap < 0.6 * fragment.width || !fragment.text.trim()) continue;
				if (y > line.baseline - 0.05 * line.size) fragment.underline = true;
				else if (y > line.baseline - 0.45 * line.size) fragment.strike = true;
				else continue;
				used.add(rule);
			}
		}
	}
	return used;
}

/** Gives fragments under a link annotation their address. */
export function applyLinks(lines: Line[], links: readonly PdfPageLink[]): void {
	for (const link of links) {
		if (!link.uri) continue;
		for (const line of lines) {
			if (line.baseline < link.y || line.baseline > link.y + link.height + line.size * 0.3)
				continue;
			for (const fragment of line.fragments) {
				const middle = fragment.x + fragment.width / 2;
				if (middle >= link.x && middle <= link.x + link.width) fragment.link = link.uri;
			}
		}
	}
}
