// A click is a double or triple click only when the browser says so. ProseMirror counts clicks
// itself: a press within 500 ms and a few pixels of the last press in the same editor is a double
// click, however much happened in between (Find walking the stories, a header opening and
// closing, typing). Two separate clicks at the same place then select the word (or the space)
// under the pointer instead of placing the caret. The browser's count (`MouseEvent.detail`)
// follows the system's double-click time and starts again after other input, so a press the
// browser counts as a first click starts ProseMirror's count again too.
import { Plugin } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

interface ClickInput {
	lastClick?: { time: number; x: number; y: number; type: string; button: number };
}

/** Forgets ProseMirror's last click when the browser counts this press as a first click. */
export function resetClickCount(view: EditorView, event: MouseEvent): boolean {
	const input = (view as unknown as { input?: ClickInput }).input;
	if (event.detail === 1 && input?.lastClick)
		input.lastClick = { time: 0, x: 0, y: 0, type: '', button: 0 };
	return false;
}

/** Makes every editor's double and triple clicks the browser's (see `resetClickCount`). */
export const browserClickCountPlugin = () =>
	new Plugin({ props: { handleDOMEvents: { mousedown: resetClickCount } } });
