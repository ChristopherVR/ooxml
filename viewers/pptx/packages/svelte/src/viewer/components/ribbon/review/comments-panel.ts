import { createCommentsPanelController } from 'ooxml-ui/pptx';
import type { CommentsPanelController } from 'ooxml-ui/pptx';

import type { EditorState } from '../../../editor/editor-state.svelte';

interface Entry {
	controller: CommentsPanelController;
	authorName: () => string;
}

const panels = new WeakMap<EditorState, Entry>();

/**
 * The viewer's one comments panel controller (`ooxml-ui/pptx`): drafts, edits and reply
 * composers, per slide. Every `ReviewCommentsPanel` of a viewer (the Review tab, the inspector's
 * Comments tab, the mobile sheet) reads the same one, so a half-typed comment stays with its
 * slide and outlives the panel. `authorName` is the panel's current author (the host
 * `authorName`, then Options > General > "User name").
 */
export function commentsPanelFor(
	editor: EditorState,
	authorName: () => string,
): CommentsPanelController {
	const existing = panels.get(editor);
	if (existing) {
		existing.authorName = authorName;
		return existing.controller;
	}
	const entry: Entry = {
		authorName,
		controller: createCommentsPanelController({
			slides: () => editor.slides,
			canEdit: () => editor.editable,
			authorName: () => entry.authorName(),
			selectedElementId: () => editor.selectedElement?.id,
			commit: (next) => editor.commitSlides(next),
		}),
	};
	panels.set(editor, entry);
	return entry.controller;
}
