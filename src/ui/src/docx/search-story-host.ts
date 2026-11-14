import type { EditorCore } from './editor-core';
import type { SearchStoryHost } from './search-story-replace';

/**
 * The editor's stories for Find. Headers, footers and notes open on the continuous surface (the
 * Draft view), so a match in one switches Print Layout to it, as the Navigation pane's headings do.
 */
export function editorSearchStories(core: EditorCore): SearchStoryHost {
	const { parts } = core;
	return {
		model: () => core.model,
		openStory: () => {
			const view = parts.activeView();
			const key = parts.activeStoryKey();
			return view && key ? { key, view } : undefined;
		},
		showStory: (story, from, to) => {
			if (core.pages.viewMode === 'print') core.pages.setViewMode('draft');
			parts.showStory(story, from, to, core.shell.canvas, core.shell.paper);
		},
		closeStory: () => {
			if (parts.activeView()) parts.closeOpenStory(false);
		},
		canEditStories: () => core.host.canEditOutsideBody(),
		writeHeaderFooter: (story, blocks) => parts.writeHeaderFooter(story, blocks),
		writeNotes: (kind, notes) => parts.writeNotes(kind, notes),
	};
}

/**
 * A story editor Find opened keeps its selection while the focus stays in Find. When the focus
 * leaves Find for anywhere but that editor or the ribbon, the editor closes, as it does when it
 * loses the focus itself.
 */
export function closeStoryWhenFindBlurs(core: EditorCore, panel: HTMLElement): void {
	panel.addEventListener('focusout', (event) => {
		const story = core.parts.activeView();
		const next = event.relatedTarget instanceof Node ? event.relatedTarget : null;
		if (!story || panel.contains(next) || story.dom.contains(next)) return;
		if (core.shell.toolbar?.contains(next)) return;
		core.parts.closeOpenStory(false);
	});
}
