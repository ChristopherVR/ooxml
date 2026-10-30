// Word highlights the regions a restricted document lets you edit; this draws the same.
import type { DocumentProtection } from 'ooxml-core/docx';
import { editRestrictionOf, editableRanges } from 'ooxml-core/docx/ui';
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';

const key = new PluginKey('docx-editable-regions');

/** Highlights the ranges the current user may edit while an editing restriction is enforced. */
export function editableRegionViewPlugin(
	protection: () => DocumentProtection | undefined,
	user: () => string,
	enabled: () => boolean = () => true,
): Plugin {
	return new Plugin({
		key,
		props: {
			decorations(state) {
				const restriction = editRestrictionOf(protection());
				if (!restriction || restriction === 'trackedChanges' || !enabled()) return null;
				const decorations = editableRanges(state, user())
					.filter((range) => range.to > range.from)
					.map((range) =>
						Decoration.inline(range.from, range.to, { class: 'dve-editable-region' }),
					);
				return decorations.length ? DecorationSet.create(state.doc, decorations) : null;
			},
		},
	});
}
