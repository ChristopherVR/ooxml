/**
 * Anchors floating objects (text boxes holding rotated text, vector drawings) to the paragraph
 * nearest to where they sit on their page; a page without paragraphs gets an empty one to hold
 * them.
 */
import type { VectorDrawing } from './drawings';
import type { FlowItem } from './flow';
import type { SectionLayout } from './package';
import type { ParagraphBlock } from './paragraphs';
import type { FloatingText } from './rotated';

/** The paragraph of `page` nearest to `y`, or a new empty one when the page has none. */
function anchorParagraph(
	blocks: FlowItem[],
	sections: SectionLayout[],
	page: number,
	y: number,
	bodySize: number,
): ParagraphBlock {
	let best: { block: ParagraphBlock; distance: number } | undefined;
	for (const block of blocks) {
		if (block.kind !== 'paragraph') continue;
		for (const line of block.lines) {
			if (line.page !== page) continue;
			const distance = Math.abs(line.top - y);
			if (!best || distance < best.distance) best = { block, distance };
		}
	}
	if (best) return best.block;
	// An empty paragraph made for an earlier object of the page holds this one too.
	const made = blocks.find(
		(block): block is ParagraphBlock =>
			block.kind === 'paragraph' && block.page === page && !block.lines.length,
	);
	if (made) return made;
	let at = blocks.findIndex((block) => block.page > page);
	if (at < 0) at = blocks.length;
	const paragraph: ParagraphBlock = {
		kind: 'paragraph',
		page,
		lines: [],
		fragments: [],
		align: 'left',
		indentLeft: 0,
		firstLine: 0,
		spaceBefore: 0,
		lineSpacing: 1,
		size: bodySize,
	};
	blocks.splice(at, 0, paragraph);
	// The new paragraph closes the page before: later sections start one block later.
	sections.forEach((section, index) => {
		if (index > 0 && section.start >= at) section.start++;
	});
	return paragraph;
}

export function attachFloats(
	blocks: FlowItem[],
	sections: SectionLayout[],
	floats: FloatingText[],
	bodySize: number,
): void {
	for (const float of floats)
		(anchorParagraph(blocks, sections, float.page, float.y, bodySize).floats ??= []).push(float);
}

export function attachDrawings(
	blocks: FlowItem[],
	sections: SectionLayout[],
	drawings: VectorDrawing[],
	bodySize: number,
): void {
	for (const drawing of drawings)
		(anchorParagraph(blocks, sections, drawing.page, drawing.y, bodySize).drawings ??= []).push(
			drawing,
		);
}
