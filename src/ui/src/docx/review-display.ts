import { Plugin } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { formattingRevision, paragraphFormattingRevision } from 'ooxml-core/docx/ui';

/**
 * All Markup shows insertions/deletions as authored (default CSS from schema.ts's marks).
 * No Markup (final) hides deleted text and shows insertions unmarked, matching the resulting document.
 * Original hides inserted text and shows deletions unmarked. Style plugins project prior run and
 * paragraph formatting; structural revision display still needs implementation.
 * Simple Markup is approximated as No Markup with a change indicator; per-line change bars are not
 * implemented, so it is visually identical to No Markup today (see docs/parity-roadmap.md).
 */
export type { ReviewDisplayMode } from 'ooxml-core/docx';
import type { ReviewDisplayMode } from 'ooxml-core/docx';

function hiddenMarkName(mode: ReviewDisplayMode): 'deletion' | 'insertion' | null {
	if (mode === 'final' || mode === 'simple') return 'deletion';
	if (mode === 'original') return 'insertion';
	return null;
}

/** Hides insertion or deletion runs via decorations rather than mutating the document. */
export function reviewDisplayPlugin(getMode: () => ReviewDisplayMode): Plugin {
	return new Plugin({
		props: {
			attributes: () => ({ 'data-review-display': getMode() }),
			decorations(state) {
				const mode = getMode();
				const hidden = hiddenMarkName(mode);
				const decorations: Decoration[] = [];
				state.doc.descendants((node, pos) => {
					if (mode === 'all' && paragraphFormattingRevision(node))
						decorations.push(
							Decoration.node(pos, pos + node.nodeSize, { class: 'dve-revision-format-markup' }),
						);
					if (!node.isText && node.type.name !== 'hardBreak') return;
					if (mode === 'all' && formattingRevision(node))
						decorations.push(
							Decoration.inline(pos, pos + node.nodeSize, { class: 'dve-revision-format-markup' }),
						);
					if (node.marks.some((mark) => mark.type.name === hidden))
						decorations.push(
							Decoration.inline(pos, pos + node.nodeSize, { class: 'dve-revision-hidden' }),
						);
				});
				return decorations.length ? DecorationSet.create(state.doc, decorations) : null;
			},
		},
	});
}
