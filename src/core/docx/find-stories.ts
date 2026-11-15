// The stories Find and Replace searches, in the order Word's Find runs through them: the main
// document, the headers and footers, the footnotes, the endnotes and the text boxes. Headers and
// footers follow Word's story-type order (even header, primary header, even footer, primary
// footer, first-page header, first-page footer), each through every section that has its own.
// Comments are not covered yet. The list is story-neutral: it carries each story's blocks, and the
// editor matches them with the same `findMatches`.
import type { Block, DocumentModel } from './model';
import type { HeaderFooterContent } from './section-model';

export type FindHeaderFooterSlot = 'default' | 'first' | 'even';

export type FindStory =
	| { readonly kind: 'main' }
	| { readonly kind: 'textBoxes' }
	| {
			readonly kind: 'footnote' | 'endnote';
			readonly id: string;
			readonly blocks: readonly Block[];
	  }
	| {
			readonly kind: 'header' | 'footer';
			readonly slot: FindHeaderFooterSlot;
			/** The section that owns the part; later sections that inherit it are not listed again. */
			readonly sectionIndex: number;
			readonly blocks: readonly Block[];
	  };

/** Word's story-type order for headers and footers. */
const HEADER_FOOTER_ORDER: ReadonlyArray<readonly ['header' | 'footer', FindHeaderFooterSlot]> = [
	['header', 'even'],
	['header', 'default'],
	['footer', 'even'],
	['footer', 'default'],
	['header', 'first'],
	['footer', 'first'],
];

/** Every story of the document in search order; a header or footer part shared by sections is listed once. */
export function findStories(model: DocumentModel): FindStory[] {
	const stories: FindStory[] = [{ kind: 'main' }];
	const seen = new Set<HeaderFooterContent | string>();
	for (const [kind, slot] of HEADER_FOOTER_ORDER) {
		(model.sections ?? []).forEach((section, sectionIndex) => {
			const content = section[kind === 'header' ? 'headers' : 'footers']?.[slot];
			if (!content || seen.has(content) || (content.partName && seen.has(content.partName))) return;
			seen.add(content);
			if (content.partName) seen.add(content.partName);
			stories.push({ kind, slot, sectionIndex, blocks: content.blocks });
		});
	}
	for (const note of model.footnotes ?? [])
		stories.push({ kind: 'footnote', id: note.id, blocks: note.blocks });
	for (const note of model.endnotes ?? [])
		stories.push({ kind: 'endnote', id: note.id, blocks: note.blocks });
	stories.push({ kind: 'textBoxes' });
	return stories;
}

/** A stable identity for a story within one model. */
export function findStoryKey(story: FindStory): string {
	switch (story.kind) {
		case 'main':
		case 'textBoxes':
			return story.kind;
		case 'footnote':
		case 'endnote':
			return `${story.kind}:${story.id}`;
		default:
			return `${story.kind}:${story.sectionIndex}:${story.slot}`;
	}
}

/**
 * The undo group a story belongs to, as Word groups story ranges: all footnotes are one story, all
 * endnotes another, all text boxes a third; each header or footer part is its own.
 */
export function findStoryGroup(story: FindStory): string {
	return story.kind === 'footnote' || story.kind === 'endnote' ? story.kind : findStoryKey(story);
}
