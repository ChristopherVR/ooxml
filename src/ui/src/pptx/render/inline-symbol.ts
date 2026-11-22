const choosingSymbol = new WeakSet<HTMLElement>();

export function isChoosingInlineSymbol(editor: HTMLElement | null | undefined): boolean {
	return editor != null && choosingSymbol.has(editor);
}

function activeInlineEditor(doc: Document): HTMLElement | undefined {
	let el = doc.activeElement;
	while (el?.shadowRoot?.activeElement) el = el.shadowRoot.activeElement;
	return el instanceof HTMLElement && el.matches('[data-inline-editor][contenteditable="true"]')
		? el
		: undefined;
}

/** Ribbon pointer interactions should leave the caret intact until the dialog takes focus. */
export function preserveInlineSymbolCaret(event: Event, doc: Document): void {
	if (activeInlineEditor(doc)) event.preventDefault();
}

export function captureInlineSymbolCaret(doc: Document) {
	const editor = activeInlineEditor(doc);
	const selection = doc.getSelection();
	if (!editor || !selection?.rangeCount) return undefined;
	const range = selection.getRangeAt(0).cloneRange();
	if (!editor.contains(range.startContainer) || !editor.contains(range.endContainer)) {
		return undefined;
	}
	choosingSymbol.add(editor);
	return (symbol?: string) => {
		choosingSymbol.delete(editor);
		if (!editor.isConnected) return false;
		editor.focus({ preventScroll: true });
		selection.removeAllRanges();
		selection.addRange(range);
		return symbol !== undefined && doc.execCommand('insertText', false, symbol);
	};
}
