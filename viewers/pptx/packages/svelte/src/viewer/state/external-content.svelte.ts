import { externalContentBarVisible, recoverySnapshotIntent } from 'ooxml-ui/pptx';
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
	 * Enable Content: reload the presentation with its linked content. Unsaved
	 * edits are kept (the deck is reloaded from its serialized state and stays
	 * unsaved), and so is the slide on screen.
	 */
	async enable(): Promise<void> {
		const { loader, editor, viewer } = this.#deps;
		const unsaved = editor.dirty;
		const slide = viewer.current;
		const edited =
			unsaved && editor.getHandler()
				? await serializeEditorState(editor, 'pptx', recoverySnapshotIntent(editor.saveIntent()))
				: undefined;
		await loader.reloadWithExternalContent(edited);
		// The load seeds the editor and resets the slide from an effect.
		await tick();
		if (loader.error === null && viewer.slideCount > 0) {
			viewer.goTo(Math.min(slide, viewer.slideCount - 1));
		}
		if (unsaved) {
			editor.dirty = true;
		}
	}

	/** Close: hide the bar; the content stays blocked. */
	close(): void {
		this.#deps.loader.externalContentBarClosed = true;
	}
}
