import type { EditorCore } from './editor-core';
import { findLocalizedControl } from './localization';
import {
	isReadingAloud,
	readAloudAvailable,
	startReadAloud,
	stopReadAloud,
	textToRead,
	type ReadAloudIssue,
	type ReadAloudStart,
} from './read-aloud';

/**
 * One editor's Read Aloud: the playback it owns, and the last problem it ran into. The browser has
 * a single voice, so starting here takes it over from another editor, and this session lets go of
 * it when its document changes or the editor leaves the page.
 */
export class ReadAloudSession {
	private lastIssue: ReadAloudIssue | null = null;

	constructor(private readonly core: EditorCore) {}

	get playing(): boolean {
		return isReadingAloud(this);
	}

	/** The last problem, or `unsupported` while the browser has no speech synthesis. */
	get issue(): ReadAloudIssue | null {
		return this.lastIssue ?? (readAloudAvailable() ? null : 'unsupported');
	}

	/** True when a request can play now: speech exists and a document is on screen. */
	get available(): boolean {
		return readAloudAvailable() && Boolean(this.core.targetView()) && !this.core.loading;
	}

	/** Reads the selection, or the caret to the end, restarting if it is already reading. */
	start(): ReadAloudStart {
		const view = this.core.targetView();
		stopReadAloud(this);
		const result: ReadAloudStart =
			!view || this.core.loading
				? 'empty'
				: startReadAloud(
						this,
						textToRead(view),
						() => this.changed(),
						(issue) => this.report(issue),
					);
		this.lastIssue = result === 'started' ? null : result;
		this.changed();
		return result;
	}

	/** Stops this editor's playback and clears its problem. Returns whether it was playing. */
	stop(): boolean {
		const was = stopReadAloud(this);
		if (this.lastIssue !== null) {
			this.lastIssue = null;
			this.changed();
		}
		return was;
	}

	/** Starts, or stops when already playing. Returns whether it is playing now. */
	toggle(): boolean {
		if (this.stop()) return false;
		return this.start() === 'started';
	}

	private report(issue: ReadAloudIssue) {
		this.lastIssue = issue;
		this.changed();
	}

	private changed() {
		// Not refreshControls: this also runs while the view is being torn down.
		const { toolbar } = this.core.shell;
		if (toolbar)
			findLocalizedControl(toolbar, 'Read aloud')?.setAttribute(
				'aria-pressed',
				String(this.playing),
			);
		this.core.element.dispatchEvent(new CustomEvent('editor-state-change'));
	}
}
