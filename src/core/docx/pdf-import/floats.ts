/**
 * Anchors floating text boxes (rotated text) to the paragraph nearest to where they sit on their
 * page; a page without paragraphs gets an empty one to hold them.
 */
import type { FlowItem } from './flow';
import type { SectionLayout } from './package';
import type { ParagraphBlock } from './paragraphs';
import type { FloatingText } from './rotated';

export function attachFloats(
	blocks: FlowItem[],
	sections: SectionLayout[],
	floats: FloatingText[],
	bodySize: number,
): void {
	for (const float of floats) {
		let best: { block: ParagraphBlock; distance: number } | undefined;
		for (const block of blocks) {
			if (block.kind !== 'paragraph') continue;
			for (const line of block.lines) {
				if (line.page !== float.page) continue;
				const distance = Math.abs(line.top - float.y);
				if (!best || distance < best.distance) best = { block, distance };
			}
		}
		if (best) {
			(best.block.floats ??= []).push(float);
			continue;
		}
		let at = blocks.findIndex((block) => block.page > float.page);
		if (at < 0) at = blocks.length;
		blocks.splice(at, 0, {
			kind: 'paragraph',
			page: float.page,
			lines: [],
			fragments: [],
			align: 'left',
			indentLeft: 0,
			firstLine: 0,
			spaceBefore: 0,
			lineSpacing: 1,
			size: bodySize,
			floats: [float],
		});
		// The new paragraph closes the page before: later sections start one block later.
		sections.forEach((section, index) => {
			if (index > 0 && section.start >= at) section.start++;
		});
	}
}
