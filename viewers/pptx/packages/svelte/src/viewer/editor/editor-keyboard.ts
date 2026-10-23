import { dispatchEditorKey } from 'ooxml-ui/pptx';
import type {
	EditorKeyActionName,
	EditorKeyHandlers,
	ResolvedKeyboardCustomization,
} from 'ooxml-ui/pptx';

import type {
	EditorTextAlign,
	FontSizeStepDirection,
	SelectionCycleDirection,
} from './editor-shortcut-types';

/** Slide paging keys the root navigation fall-through owns. */
const PAGING_ACTIONS: readonly EditorKeyActionName[] = ['prevSlide', 'nextSlide'];

/**
 * Editing keyboard shortcuts, called from the viewer root's keydown before the
 * slideshow navigation handler. Key-to-action resolution and the dispatch switch
 * are the shared `dispatchEditorKey`, so this file only maps the shared handler
 * names onto the Svelte editor's callbacks.
 *
 * Slide paging is deliberately NOT dispatched here: the root handler falls
 * through to `viewer.handleNavigationKey` when this one does not consume the
 * event, so acting on `prevSlide` / `nextSlide` too would advance two slides per
 * press.
 */
export interface EditorKeyboardDeps {
	/** False disables everything (not editable, presenting, inline editing). */
	isActive(): boolean;
	getSelectedId(): string | null;
	deselect(): void;
	/**
	 * Escape on a member selected inside a group: select its parent group
	 * instead of clearing (shared `parentSelection`). True when it did, so the
	 * selection is not also cleared.
	 */
	selectParentGroup?(): boolean;
	deleteSelected(): void;
	duplicateSelected(): void;
	nudgeSelected(dx: number, dy: number): void;
	undo(): void;
	redo(): void;
	copySelected(): void;
	cutSelected(): void;
	paste(): void;
	/** Omitted preserves custom callers that own their paste command. */
	canPaste?(): boolean;
	/** Select every interactive element on the active slide (Ctrl+A). */
	selectAll(): void;
	/** Group the multi-selection into one group element (Ctrl+G). */
	groupSelected(): void;
	/** Ungroup the selected group (Ctrl+Shift+G). */
	ungroupSelected(): void;
	cancelFormatPainter?(): boolean;
	/** Show or hide the keyboard-shortcut cheat sheet ("?"). */
	toggleShortcuts?(): void;
	/** Close the cheat sheet on Escape; true when it was open (Escape consumed). */
	closeShortcuts?(): boolean;
	/**
	 * Open or close the find bar (Ctrl/Cmd+F). Optional so a host driving this
	 * handler without find chrome still compiles; when it is missing the chord
	 * falls through to the browser, which is what this binding did before the
	 * shortcut reached the shared keymap.
	 */
	toggleFind?(): void;
	/** Open the find bar in replace mode (Ctrl/Cmd+H). See `EditorControllerDeps`. */
	toggleFindReplace?(): void;
	/** Set the selected (not editing) text shape's paragraph alignment. */
	setTextAlign?(align: EditorTextAlign): void;
	/** Step the selection's font size along PowerPoint's size ladder. */
	stepFontSize?(direction: FontSizeStepDirection): void;
	/** Copy the selection's format for the format painter (Ctrl+Shift+C). */
	copyFormat?(): void;
	/** Apply the copied format to the current selection (Ctrl+Shift+V). */
	pasteFormat?(): void;
	/** Insert a new slide after the active one (Ctrl+M). */
	newSlide?(): void;
	/** Open the hyperlink dialog for the current selection (Ctrl+K). */
	openHyperlink?(): void;
	/** Clear character formatting on the selection (Ctrl+Space). */
	clearFormatting?(): void;
	/** Move the selection to the next/previous element on the slide (Tab). */
	cycleSelection?(direction: SelectionCycleDirection): void;
	/** Open the Paste Special dialog (Ctrl/Cmd+Alt+V). */
	onPasteSpecial?(): void;
	/** Host keyboard customisation; omitted keeps the built-in keymap. */
	getKeyboardCustomization?(): ResolvedKeyboardCustomization | undefined;
}

export function createEditorKeydownHandler(
	deps: EditorKeyboardDeps,
): (event: KeyboardEvent) => void {
	const handlers: EditorKeyHandlers = {
		// Unwind the transient chrome one layer at a time: format painter, then
		// the cheat sheet, then a member selected inside a group steps out to its
		// group, then the selection itself.
		escape: () => {
			if (
				!deps.cancelFormatPainter?.() &&
				!deps.closeShortcuts?.() &&
				!deps.selectParentGroup?.()
			) {
				deps.deselect();
			}
		},
		toggleShortcuts: () => deps.toggleShortcuts?.(),
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
		findReplace: () => deps.toggleFindReplace?.(),
		alignLeft: () => deps.setTextAlign?.('left'),
		alignCenter: () => deps.setTextAlign?.('center'),
		alignRight: () => deps.setTextAlign?.('right'),
		alignJustify: () => deps.setTextAlign?.('justify'),
		increaseFontSize: () => deps.stepFontSize?.('increase'),
		decreaseFontSize: () => deps.stepFontSize?.('decrease'),
		copyFormat: () => deps.copyFormat?.(),
		pasteFormat: () => deps.pasteFormat?.(),
		newSlide: () => deps.newSlide?.(),
		hyperlink: () => deps.openHyperlink?.(),
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
				// Paging is owned by the root navigation fall-through; see the module note.
				passThrough: PAGING_ACTIONS,
			},
			handlers,
		);
	};
}
