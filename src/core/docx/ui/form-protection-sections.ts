// The sections a document protected for forms leaves editable (`w:formProt w:val="0"`), and the
// steps that keep them where they are.
import { Mark, type Node } from 'prosemirror-model';
import { AttrStep, ReplaceAroundStep, type Step } from 'prosemirror-transform';

type Range = { from: number; to: number };

/**
 * The sections a document protected for forms leaves editable (`w:formProt w:val="0"`), placed as
 * Print Layout and the saved document place them: a section ends after the first block with its
 * end id, the last section runs to the end of the document, and a section whose last block is
 * gone is empty, its blocks falling to the next one (deleting a break merges sections).
 */
export function unprotectedSections(doc: Node): Range[] {
	if (typeof doc.attrs.sections !== 'string') return [];
	const sections = JSON.parse(doc.attrs.sections) as {
		endsAtBlockId?: string;
		formProtection?: false;
	}[];
	if (!sections.some((section) => section.formProtection === false)) return [];
	const blocks: { id: unknown; end: number }[] = [];
	doc.forEach((node, offset) => blocks.push({ id: node.attrs.id, end: offset + node.nodeSize }));
	const ranges: Range[] = [];
	let start = 0;
	let next = 0;
	sections.forEach((section, index) => {
		let end = start;
		if (index === sections.length - 1) end = doc.content.size;
		else {
			const found = blocks.findIndex(
				(block, at) => at >= next && block.id === section.endsAtBlockId,
			);
			if (found >= 0) {
				end = blocks[found]!.end;
				next = found + 1;
			}
		}
		// Neighbouring unprotected sections make one editable range.
		const last = ranges.at(-1);
		if (section.formProtection === false && end > start) {
			if (last?.to === start) last.to = end;
			else ranges.push({ from: start, to: end });
		}
		start = end;
	});
	return ranges;
}

/**
 * Whether a step leaves the unprotected sections where they were, mapped through it: no block
 * moves into one of them or out of one. Content written at a section's start joins it, and at its
 * end (but for the last section) the next one, as the sections are placed.
 */
export function keepsSections(step: Step, before: Node, after: Node): boolean {
	const map = step.getMap();
	const end = before.content.size;
	const mapped = unprotectedSections(before).map((range) => ({
		from: map.map(range.from, -1),
		to: map.map(range.to, range.to === end ? 1 : -1),
	}));
	return JSON.stringify(mapped) === JSON.stringify(unprotectedSections(after));
}

const isBlock = (node: Node | null | undefined): node is Node =>
	node?.type.name === 'paragraph' || node?.type.name === 'table';

/**
 * Whether the step only gives a paragraph or table another id, as the editor does to the copy a
 * split makes of a paragraph: bookkeeping, not an edit, as long as no section moves with it.
 */
export function renamesBlock(step: Step, before: Node | undefined, after: Node): boolean {
	if (!before) return false;
	if (step instanceof AttrStep) return step.attr === 'id' && isBlock(before.nodeAt(step.pos));
	if (!(step instanceof ReplaceAroundStep)) return false;
	const old = before.nodeAt(step.from);
	const next = after.nodeAt(step.from);
	return Boolean(
		isBlock(old) &&
		next &&
		old.nodeSize === step.to - step.from &&
		old.type === next.type &&
		old.content.eq(next.content) &&
		Mark.sameSet(old.marks, next.marks) &&
		Object.keys({ ...old.attrs, ...next.attrs }).every(
			(key) => key === 'id' || old.attrs[key] === next.attrs[key],
		),
	);
}
