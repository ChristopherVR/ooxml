import type { EditorView } from 'prosemirror-view';

/** Whether the browser can speak text; Read Aloud is unavailable (not faked) without it. */
export function readAloudAvailable(): boolean {
	return typeof speechSynthesis !== 'undefined' && typeof SpeechSynthesisUtterance !== 'undefined';
}

/** What Word's Read Aloud reads: the selection, or everything from the caret to the end. */
export function textToRead(view: EditorView): string {
	const { doc, selection } = view.state;
	const from = selection.from;
	const to = selection.empty ? doc.content.size : selection.to;
	return doc.textBetween(from, to, '\n', ' ').trim();
}

/** Why a Read Aloud request did not play, or that the voice failed while playing. */
export type ReadAloudIssue = 'unsupported' | 'empty' | 'error';
/** The result of a start request. */
export type ReadAloudStart = 'started' | ReadAloudIssue;

interface Playback {
	owner: object;
	onChange: (speaking: boolean) => void;
}

// Browsers have one speech queue, so there is one playback at a time and it has one owner.
let current: Playback | undefined;
const shared = {};

/** Whether `owner` (or, without one, anybody) is reading right now. */
export function isReadingAloud(owner?: object): boolean {
	return current !== undefined && (owner === undefined || current.owner === owner);
}

/** Stops the playback `owner` started, if it is the one playing. Returns whether it stopped one. */
export function stopReadAloud(owner: object = shared): boolean {
	const playback = current;
	if (!playback || playback.owner !== owner) return false;
	current = undefined;
	try {
		speechSynthesis.cancel();
	} catch {
		/* Nothing is playing that the page can stop. */
	}
	playback.onChange(false);
	return true;
}

/**
 * Reads `text` for `owner`. Starting while another owner is reading takes the voice over: that
 * owner is told it stopped. Returns why nothing started (`unsupported`, `empty`, `error`) or
 * `started`. `onChange` fires when playback starts and ends; `onError` when the voice fails.
 */
export function startReadAloud(
	owner: object,
	text: string,
	onChange: (speaking: boolean) => void,
	onError?: (issue: ReadAloudIssue) => void,
): ReadAloudStart {
	if (!readAloudAvailable()) return 'unsupported';
	if (!text) return 'empty';
	if (current) stopReadAloud(current.owner);
	const utterance = new SpeechSynthesisUtterance(text);
	const playback: Playback = { owner, onChange };
	const finish = (failure: unknown) => {
		if (current !== playback) return;
		current = undefined;
		onChange(false);
		// A deliberate cancel or an interruption is not a failure.
		const reason = (failure as { error?: string } | undefined)?.error;
		if (failure && reason !== 'canceled' && reason !== 'interrupted') onError?.('error');
	};
	utterance.addEventListener('end', () => finish(undefined));
	utterance.addEventListener('error', (event) => finish(event ?? {}));
	current = playback;
	try {
		speechSynthesis.speak(utterance);
	} catch {
		current = undefined;
		onError?.('error');
		return 'error';
	}
	onChange(true);
	return 'started';
}

/**
 * Starts reading `text` with the browser's voice, or stops when it is already reading. Returns
 * whether speech is now playing. `onChange` fires when it ends by itself.
 */
export function toggleReadAloud(text: string, onChange: (speaking: boolean) => void): boolean {
	if (!readAloudAvailable()) return false;
	if (stopReadAloud(shared)) return false;
	return startReadAloud(shared, text, onChange) === 'started';
}
