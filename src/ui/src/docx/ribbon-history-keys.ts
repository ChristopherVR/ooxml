// Undo and Redo keep working while focus is on the ribbon (after a ribbon command such as Ink to
// Shape, focus stays on its button), as in Word: Ctrl+Z, Ctrl+Y and Ctrl+Shift+Z run the document's
// history unless the key goes to a text box of the ribbon, which keeps its own.

/** Whether a key press is Undo or Redo, and which. */
export function historyKey(event: KeyboardEvent): 'undo' | 'redo' | undefined {
	if (!(event.ctrlKey || event.metaKey) || event.altKey) return undefined;
	const key = event.key.toLowerCase();
	if (key === 'z') return event.shiftKey ? 'redo' : 'undo';
	if (key === 'y' && !event.shiftKey) return 'redo';
	return undefined;
}

/** Whether the key goes to something that edits text of its own (a box, a list, an editor). */
export function editsOwnText(event: Event): boolean {
	const target = event.composedPath()[0];
	if (!(target instanceof HTMLElement)) return false;
	return (
		target instanceof HTMLInputElement ||
		target instanceof HTMLTextAreaElement ||
		target instanceof HTMLSelectElement ||
		target.isContentEditable
	);
}

/** Runs the document's Undo and Redo for those keys pressed on the ribbon. */
export function attachRibbonHistoryKeys(
	ribbon: HTMLElement,
	run: (key: 'undo' | 'redo') => void,
): void {
	ribbon.addEventListener('keydown', (event) => {
		const key = historyKey(event);
		if (!key || event.defaultPrevented || editsOwnText(event)) return;
		event.preventDefault();
		event.stopPropagation();
		run(key);
	});
}
