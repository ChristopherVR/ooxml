/**
 * Real Word lists from typed markers: paragraphs that start with a bullet or a number become list
 * items (the marker text removed), consecutive items form a list whose levels come from the
 * markers' indents (or the parts of a dotted number), and a numbered list interrupted by other
 * paragraphs goes on with the same numbering when its numbers do.
 */
import { markerOf, type ListFormat, type ListMarker } from './list-markers';
import type { ParagraphBlock } from './paragraphs';

export interface ListItem {
	numId: number;
	level: number;
}

export interface ListLevel {
	format: ListFormat;
	/** `w:lvlText`: the bullet glyph, or the number pattern ("%1.", "(%2)", "%1.%2"). */
	text: string;
	/** The font of a bullet glyph. */
	font?: string;
	start: number;
	/** Points from the text area's left edge to the text, and from the marker to the text. */
	left: number;
	hanging: number;
}

export interface ListDefinition {
	numId: number;
	/** Levels 0 to 8; unset levels get defaults. */
	levels: (ListLevel | undefined)[];
}

const isBlock = (block: { kind: string }): block is ParagraphBlock => block.kind === 'paragraph';

/** The paragraph's text area left edge, from its first line and indents. */
const base = (block: ParagraphBlock) => block.lines[0]!.left - block.indentLeft - block.firstLine;

/** The marker of a list item paragraph, when the paragraph is one. */
function candidate(block: ParagraphBlock): ListMarker | undefined {
	if (block.heading || (block.align !== 'left' && block.align !== 'both')) return undefined;
	return markerOf(block.lines[0]!);
}

function supported(
	block: ParagraphBlock,
	marker: ListMarker,
	neighbours: (ListMarker | undefined)[],
) {
	if (marker.strong) return true;
	const size = block.lines[0]!.size;
	if (marker.gap >= 0.6 * size) return true;
	const second = block.lines[1];
	if (second && second.left > marker.x + 2 && Math.abs(second.left - marker.textX) < 3) return true;
	return (
		block.lines.length === 1 &&
		neighbours.some((other) => other?.format === marker.format && Math.abs(other.x - marker.x) < 3)
	);
}

const BULLETS: Record<string, [string, string]> = {
	'\u2022': ['\uf0b7', 'Symbol'],
	'\u00b7': ['\uf0b7', 'Symbol'],
	'\u2219': ['\uf0b7', 'Symbol'],
	'\u25cf': ['\uf0b7', 'Symbol'],
	o: ['o', 'Courier New'],
	'\u25aa': ['\uf0a7', 'Wingdings'],
	'\u25a0': ['\uf0a7', 'Wingdings'],
	'\u27a2': ['\uf0d8', 'Wingdings'],
	'\u2713': ['\uf0fc', 'Wingdings'],
	'\u2714': ['\uf0fc', 'Wingdings'],
};

function levelOf(marker: ListMarker, level: number, left: number): ListLevel {
	const hanging = Math.max(0, marker.textX - marker.x);
	if (marker.format === 'bullet') {
		const glyph = marker.bullet!;
		const [text, font] = BULLETS[glyph.char] ?? [glyph.char, glyph.font];
		return { format: 'bullet', text, font, start: 1, left, hanging };
	}
	const parts =
		marker.depth > 1
			? Array.from({ length: marker.depth }, (_, at) => `%${level - marker.depth + at + 2}`).join(
					'.',
				)
			: `%${level + 1}`;
	return {
		format: marker.format,
		text: `${marker.prefix}${parts}${marker.suffix}`,
		start: marker.value,
		left,
		hanging,
	};
}

/** Removes the first `count` characters of a paragraph's text (the typed marker). */
function strip(block: ParagraphBlock, count: number): void {
	while (count > 0 && block.fragments.length) {
		const first = block.fragments[0]!;
		if (first.text.length <= count) {
			count -= first.text.length;
			block.fragments.shift();
		} else {
			first.text = first.text.slice(count);
			count = 0;
		}
	}
	while (block.fragments.length && !block.fragments[0]!.text.trim()) block.fragments.shift();
	if (block.fragments[0]) block.fragments[0].text = block.fragments[0].text.trimStart();
}

/**
 * Makes the list item paragraphs of `blocks` (in document order) Word list items and returns the
 * list definitions their `list` refers to.
 */
export function assignLists(blocks: { kind: string }[]): ListDefinition[] {
	const paragraphs = blocks.map((block) => (isBlock(block) ? block : undefined));
	const markers = paragraphs.map((block) => (block ? candidate(block) : undefined));
	const items = markers.map((marker, at) =>
		marker && supported(paragraphs[at]!, marker, [markers[at - 1], markers[at + 1]])
			? marker
			: undefined,
	);
	const definitions: ListDefinition[] = [];
	const lastNumber = new Map<number, number>();
	let at = 0;
	while (at < blocks.length) {
		if (!items[at]) {
			at++;
			continue;
		}
		// A run of items; paragraphs indented under an item (its further paragraphs) stay in it.
		const run: number[] = [];
		let end = at;
		while (end < blocks.length) {
			const block = paragraphs[end];
			if (!block) break;
			if (items[end]) run.push(end);
			else {
				const indent = Math.min(...run.map((index) => items[index]!.textX));
				if (Math.min(...block.lines.map((line) => line.left)) < indent - 2) break;
			}
			end++;
		}
		const xs = [...new Set(run.map((index) => Math.round(items[index]!.x)))].sort((a, b) => a - b);
		const steps: number[] = [];
		for (const x of xs) if (!steps.length || x - steps[steps.length - 1]! > 4) steps.push(x);
		const levelFor = (marker: ListMarker) =>
			Math.min(
				8,
				marker.depth > 1
					? marker.depth - 1
					: Math.max(0, steps.filter((step) => marker.x >= step - 4).length - 1),
			);
		const first = items[run[0]!]!;
		const reuse =
			first.format !== 'bullet' && first.value > 1
				? definitions.find((definition) => {
						const level = definition.levels[levelFor(first)];
						return (
							level?.format === first.format &&
							level.text === levelOf(first, levelFor(first), 0).text &&
							lastNumber.get(definition.numId) === first.value - 1
						);
					})
				: undefined;
		const definition = reuse ?? { numId: definitions.length + 1, levels: [] };
		if (!reuse) definitions.push(definition);
		for (const index of run) {
			const block = paragraphs[index]!;
			const marker = items[index]!;
			const level = Math.max(0, levelFor(marker));
			const left = marker.textX - base(block);
			definition.levels[level] ??= levelOf(marker, level, left);
			if (level === 0 && marker.format !== 'bullet') lastNumber.set(definition.numId, marker.value);
			strip(block, marker.length);
			block.list = { numId: definition.numId, level };
			block.indentLeft = left;
			block.firstLine = -(marker.textX - marker.x);
		}
		at = end;
	}
	return definitions;
}
