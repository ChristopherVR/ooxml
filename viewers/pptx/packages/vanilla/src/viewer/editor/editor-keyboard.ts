import { dispatchEditorKey } from 'ooxml-ui/pptx';
import type {
	EditorKeyActionName,
	EditorKeyHandlers,
	ResolvedKeyboardCustomization,
} from 'ooxml-ui/pptx';

/** Slide paging keys the root navigation handler owns. */
const PAGING_ACTIONS: readonly EditorKeyActionName[] = ['prevSlide', 'nextSlide'];

/**
 * Editing keyboard shortcuts, attached alongside the slideshow navigation
 * handler. Key-to-action resolution and the dispatch switch are the shared
 * `dispatchEditorKey`, so this file only maps the shared handler names onto
 * the vanilla editor's callbacks.
 *
 * Slide paging is deliberately NOT dispatched here: the viewer root already
 * carries `attachKeyboardNavigation`, which pages the deck when no element is
 * selected. Acting on `prevSlide` / `nextSlide` as well would advance two slides
 * per press.
 */
export interface EditorKeyboardDeps {
	/** False disables everything (not editable, presenting, inline editing). */
	isActive(): boolean;
	/** The host's keyboard customisation (disabled / remapped commands), read per key. */
	getKeyboardCustomization?(): ResolvedKeyboardCustomization | undefined;
	getSelectedId(): string | null;
	deselect(): void;
	/**
	 * Escape on a member selected inside a group: select its parent group
	 * instead (shared `parentSelection`). True when it did, so the selection is
	 * not cleared. Omitted, Escape always clears.
	 */
	selectParent?(): boolean;
	deleteSelected(): void;
	duplicateSelected(): void;
	copySelected(): void;
	cutSelected(): void;
	/** Paste isn't selection-gated: it targets the current slide regardless of selection. */
	paste(): void;
	/** Omitted preserves custom callers that own their paste command. */
	canPaste?(): boolean;
	/** Select every interactive element on the active slide (Ctrl+A). */
	selectAll(): void;
	/** Group the multi-selection into one group element (Ctrl+G). */
	groupSelected(): void;
	/** Ungroup the selected group (Ctrl+Shift+G). */
	ungroupSelected(): void;
	nudgeSelected(dx: number, dy: number): void;
	undo(): void;
	redo(): void;
	cancelFormatPainter(): boolean;
	/** Show or hide the keyboard-shortcut cheat sheet ("?"). */
	toggleShortcuts(): void;
	/** Close the cheat sheet on Escape; true when it was open (Escape consumed). */
	closeShortcuts(): boolean;
	/**
	 * Open or close the find bar (Ctrl/Cmd+F). Optional so a host driving this
	 * handler without find chrome still compiles; when it is missing the chord
	 * falls through to the browser, which is what this binding did before the
	 * shortcut reached the shared keymap.
	 */
	toggleFind?(): void;
	/** Ctrl+L/E/R/J: paragraph alignment on the selected text shape. */
	setTextAlign?(align: 'left' | 'center' | 'right' | 'justify'): void;
	/** Ctrl+Shift+>/< and Ctrl+]/[: PowerPoint's font-size ladder. */
	stepFontSize?(direction: 'increase' | 'decrease'): void;
	/** Ctrl+Shift+C: arm the format painter from the current selection. */
	copyFormat?(): void;
	/** Ctrl+Shift+V: apply the copied format to the current selection. */
	pasteFormat?(): void;
	/** Ctrl+M: insert a new slide after the active one. */
	addSlide?(): void;
	/** Ctrl+K: open the hyperlink dialog for the current selection. */
	openHyperlink?(): void;
	/** Ctrl+H: open the docked Find & Replace panel. */
	toggleFindReplace?(): void;
	/** Ctrl+Space: clear character formatting on the current selection. */
	clearFormatting?(): void;
	/** Tab/Shift+Tab: cycle the selection through the slide's elements. */
	cycleSelection?(direction: 'next' | 'prev'): void;
	/** Ctrl/Cmd+Alt+V: open the Paste Special dialog. */
	onPasteSpecial?(): void;
}

export function createEditorKeydownHandler(
	deps: EditorKeyboardDeps,
): (event: KeyboardEvent) => void {
	const handlers: EditorKeyHandlers = {
		// Unwind the transient chrome one layer at a time: format painter, then
		// the cheat sheet, then a member steps out to its group, then the
		// selection itself.
		escape: () => {
			if (!deps.cancelFormatPainter() && !deps.closeShortcuts() && !deps.selectParent?.()) {
				deps.deselect();
			}
		},
		toggleShortcuts: () => deps.toggleShortcuts(),
		find: () => deps.toggleFind?.(),
		undo: () => deps.undo(),
		redo: () => deps.redo(),
		paste: () => deps.paste(),
		selectAll: () => deps.selectAll(),
		delete: () => deps.deleteSelected(),
		duplicate: () => deps.duplicateSelected(),
		copy: () => deps.copySelected(),
		cut: () => deps.cutSelected(),
		group: () => deps.groupSelected(),
		ungroup: () => deps.ungroupSelected(),
		nudge: (dx, dy) => deps.nudgeSelected(dx, dy),
		alignLeft: () => deps.setTextAlign?.('left'),
		alignCenter: () => deps.setTextAlign?.('center'),
		alignRight: () => deps.setTextAlign?.('right'),
		alignJustify: () => deps.setTextAlign?.('justify'),
		increaseFontSize: () => deps.stepFontSize?.('increase'),
		decreaseFontSize: () => deps.stepFontSize?.('decrease'),
		copyFormat: () => deps.copyFormat?.(),
		pasteFormat: () => deps.pasteFormat?.(),
		newSlide: () => deps.addSlide?.(),
		hyperlink: () => deps.openHyperlink?.(),
		findReplace: () => deps.toggleFindReplace?.(),
		clearFormatting: () => deps.clearFormatting?.(),
		cycleSelectionNext: () => deps.cycleSelection?.('next'),
		cycleSelectionPrev: () => deps.cycleSelection?.('prev'),
		pasteSpecial: () => deps.onPasteSpecial?.(),
	};
	return (event) => {
		if (!deps.isActive()) {
			return;
		}
		dispatchEditorKey(
			event,
			{
				canPaste: deps.canPaste?.(),
				hasSelection: deps.getSelectedId() !== null,
				keyboard: deps.getKeyboardCustomization?.(),
				// Paging is owned by the root navigation handler; see the module note.
				passThrough: PAGING_ACTIONS,
			},
			handlers,
		);
	};
}
