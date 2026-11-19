import {
	externalContentBarVisible,
	prepareExternalContentReload,
	recoverySnapshotIntent,
} from 'ooxml-ui/pptx';
import { tick } from 'svelte';

import { serializeEditorState } from '../editor/editor-document-lifecycle';
import type { EditorState } from '../editor/editor-state.svelte';
import type { PresentationLoader } from './presentation-loader.svelte';
import type { ViewerState } from './viewer-state.svelte';

export interface ExternalContentDeps {
	loader: PresentationLoader;
	editor: EditorState;
	viewer: ViewerState;
	/** The host `allowExternalContent` prop. */
	getHost: () => boolean | undefined;
}

/**
 * PowerPoint's security bar for linked pictures and media
 * (`<pptx-ui-external-content-bar>`): whether it shows, and what its two
 * buttons do. The consent itself lives on the loader, which clears it when
 * another presentation loads; the rules are the shared
 * `external-content-policy` ones.
 */
export class ExternalContentState {
	readonly #deps: ExternalContentDeps;

	constructor(deps: ExternalContentDeps) {
		this.#deps = deps;
	}

	/** Whether the security bar shows for the presentation now open (reactive). */
	get barVisible(): boolean {
		const { loader } = this.#deps;
		return externalContentBarVisible({
			hasLinkedWebContent: loader.hasLinkedWebContent,
			// Enable Content hides the bar at once, while its reload runs.
			allowed: loader.externalContentAllowed || loader.externalContentEnabled,
			host: this.#deps.getHost(),
			closed: loader.externalContentBarClosed,
		});
	}

	/**
	 * Enable Content: reload the presentation as it stands (the shared
	 * `prepareExternalContentReload`) with its linked content. Saved and unsaved
	 * edits come along; unsaved ones stay unsaved, and the slide on screen
	 * stays. Undo history starts over.
	 */
	async enable(): Promise<void> {
		const { loader, editor, viewer } = this.#deps;
		const reload = await prepareExternalContentReload({
			serializeCurrentDeck: async () =>
				editor.getHandler()
					? serializeEditorState(editor, 'pptx', recoverySnapshotIntent(editor.saveIntent()))
					: null,
			isDirty: editor.dirty,
			slideIndex: viewer.current,
		});
		if (!reload) {
			return;
		}
		await loader.reloadWithExternalContent(reload.bytes);
		// The load seeds the editor and resets the slide from an effect.
		await tick();
		if (loader.error === null && viewer.slideCount > 0) {
			viewer.goTo(Math.min(reload.slideIndex, viewer.slideCount - 1));
		}
		if (reload.keepDirty) {
			editor.dirty = true;
		}
	}

	/** Close: hide the bar; the content stays blocked. */
	close(): void {
		this.#deps.loader.externalContentBarClosed = true;
	}
}
