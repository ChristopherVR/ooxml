import { Plugin } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { EditorView } from 'prosemirror-view';
import { type DocumentModel, type SectionProperties, sectionsOf } from 'ooxml-core/docx';
import { expectDefined } from 'ooxml-core/docx/ui';

// Kept here so the editor's section modules have one place to import it from.
export { sectionsOf };

/** Index of the top-level block holding the selection start. */
function selectedBlockIndex(view: EditorView): number {
	return view.state.selection.$from.index(0);
}

/** The section containing the selection (Word applies page setup to the current section). */
export function currentSectionIndex(view: EditorView, model: DocumentModel): number {
	const sections = sectionsOf(model);
	const block = selectedBlockIndex(view);
	for (const [index, section] of sections.slice(0, -1).entries()) {
		const end = model.blocks.findIndex((item) => item.id === section.endsAtBlockId);
		if (end >= block) return index;
	}
	return sections.length - 1;
}

export {
	setColumns,
	setLineNumbering,
	setMargins,
	setOrientation,
	setPageNumbering,
	setTitlePage,
	setVerticalAlign,
	withSection,
} from 'ooxml-core/docx';

/**
 * Word's Layout > Breaks > Section Break: the current section ends after the selected paragraph
 * and a new section with the same settings begins; `type` is how the new section starts.
 */
export function insertSectionBreak(
	view: EditorView,
	model: DocumentModel,
	type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage',
): DocumentModel {
	const block = model.blocks[selectedBlockIndex(view)];
	if (!block || block.type !== 'paragraph')
		throw new Error('Place the cursor in a paragraph to insert a section break.');
	if (block.id === model.blocks.at(-1)?.id)
		throw new Error('A section break needs text after it; add a paragraph below first.');
	const sections = sectionsOf(model);
	if (sections.some((section) => section.endsAtBlockId === block.id)) return model;
	const index = currentSectionIndex(view, model);
	const current = expectDefined(sections[index], 'current section');
	const next = [...sections];
	next.splice(
		index,
		1,
		{ ...structuredClone(current), endsAtBlockId: block.id },
		{ ...current, type },
	);
	return { ...model, sections: next };
}

const BREAK_LABEL: Record<SectionProperties['type'], string> = {
	nextPage: 'Section Break (Next Page)',
	continuous: 'Section Break (Continuous)',
	evenPage: 'Section Break (Even Page)',
	oddPage: 'Section Break (Odd Page)',
	nextColumn: 'Section Break (Next Column)',
};

/** Marks paragraphs that end a section, like Word's formatting marks for section breaks. */
export function sectionBreaksPlugin() {
	return new Plugin({
		props: {
			decorations(state) {
				const json = state.doc.attrs.sections;
				const sections: SectionProperties[] = typeof json === 'string' ? JSON.parse(json) : [];
				if (sections.length < 2) return null;
				const labels = new Map<string, string>();
				sections.slice(0, -1).forEach((section, index) => {
					const following = sections[index + 1];
					if (following) labels.set(section.endsAtBlockId, BREAK_LABEL[following.type]);
				});
				const decorations: Decoration[] = [];
				state.doc.forEach((node, offset) => {
					const label = labels.get(String(node.attrs.id));
					if (label)
						decorations.push(
							Decoration.node(offset, offset + node.nodeSize, {
								class: 'dve-section-end',
								'data-section-break': label,
							}),
						);
				});
				return DecorationSet.create(state.doc, decorations);
			},
		},
	});
}
