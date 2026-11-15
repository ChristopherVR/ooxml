// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Gives a document's stories (the body, and each header and footer) new blocks, keeping every
// object that did not change, so the writer keeps the XML of what a field update left alone.
import type { Block, DocumentModel, HeaderFooterContent, HeaderFooterSlots } from './model';

const SLOTS = ['default', 'first', 'even'] as const;

/** New blocks for a story, or the same array when nothing in it changed. */
export type StoryBlocks = (blocks: readonly Block[]) => Block[];

/**
 * `model` with `headerFooter` applied to every header and footer and `body` (when given) to the
 * main document. Returns `model` itself when nothing changed.
 */
export function withStoryBlocks(
	model: DocumentModel,
	headerFooter: StoryBlocks,
	body?: StoryBlocks,
): DocumentModel {
	const content = (item: HeaderFooterContent | undefined) => {
		if (!item) return item;
		const blocks = headerFooter(item.blocks);
		return blocks === item.blocks ? item : { ...item, blocks };
	};
	const slots = (item: HeaderFooterSlots | undefined) => {
		if (!item) return item;
		let changed = false;
		const next: HeaderFooterSlots = { ...item };
		for (const slot of SLOTS) {
			const refreshed = content(item[slot]);
			if (refreshed === item[slot]) continue;
			changed = true;
			next[slot] = refreshed!;
		}
		return changed ? next : item;
	};
	let changed = false;
	const sections = model.sections?.map((section) => {
		const headers = slots(section.headers);
		const footers = slots(section.footers);
		if (headers === section.headers && footers === section.footers) return section;
		changed = true;
		return {
			...section,
			...(headers ? { headers } : {}),
			...(footers ? { footers } : {}),
		};
	});
	const blocks = body ? body(model.blocks) : model.blocks;
	if (!changed && blocks === model.blocks) return model;
	return { ...model, blocks, ...(sections ? { sections } : {}) };
}
