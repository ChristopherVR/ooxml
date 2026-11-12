import type { EditorCore } from './editor-core';
import type { ReadAloudIssue, ReadAloudStart } from './read-aloud';

export type { ReadAloudIssue, ReadAloudStart };

/**
 * State and commands for a host's own Read Aloud control. It reads the selection, or the caret to
 * the end of the document, with the editor's own implementation, and works in a read-only
 * document (it never edits or dirties). Callbacks read the live editor and never move focus or
 * the selection.
 */
export function wordReadAloudProps(core: EditorCore) {
	const session = core.readAloud;
	return {
		/** A document is on screen and the browser can speak. */
		canReadAloud: session.available,
		/** This editor is reading now. Another editor reading leaves it false. */
		readingAloud: session.playing,
		/**
		 * Why Read Aloud is not or no longer playing: `unsupported` (no speech synthesis),
		 * `empty` (nothing to read) or `error` (the voice failed). Null otherwise. The host words it.
		 */
		readAloudIssue: session.issue,
		/** Starts reading, taking the voice over from another editor. Says why nothing started. */
		onReadAloud: (): ReadAloudStart => core.readAloud.start(),
		/** Stops this editor's playback. Returns whether it was playing. */
		onStopReadAloud: (): boolean => core.readAloud.stop(),
	};
}
