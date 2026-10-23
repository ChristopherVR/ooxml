import { dispatchEditorKey, isFeatureEnabled } from 'ooxml-ui/pptx';
import type { EditorKeyHandlers, ResolvedCustomization } from 'ooxml-ui/pptx';
/**
 * useKeyboardShortcuts: the editor keymap, wired into React.
 *
 * Key-to-action resolution and the dispatch switch are the shared
 * `dispatchEditorKey` (`ooxml-ui/pptx`), the one keymap and handler interface
 * all five bindings share; this hook only snapshots the guard state, maps its
 * `on*` callbacks onto `EditorKeyHandlers`, and owns the listener lifetime.
 *
 * The listener is registered on `window` exactly once. It used to be registered
 * on the viewer container *as well*, "as a fallback", which meant every key
 * whose target sat inside the container was handled twice: Ctrl+D produced two
 * duplicates, Ctrl+V two pastes, an arrow nudged two steps and skipped two
 * slides. A `window` listener already sees events from inside the container
 * (they bubble), so the container listener was never a fallback, only a second
 * delivery.
 *
 * `F5` / `Shift+F5` (start the show) are checked by the same dispatcher BEFORE
 * the editor keymap runs. PowerPoint starts
 * the show from the keyboard even with the caret in a text box and even when
 * the host is read-only, both of which `mapEditorKey` gates against, so the
 * start-show check must sit ahead of (not inside) that gate, on the same
 * always-attached `window` listener.
 */
import { useEffect, useCallback, useRef } from 'react';

import type { TableCellEditorState, DrawingTool } from '../types';
import type { ViewerMode } from '../types-core';

/* ------------------------------------------------------------------ */
/*  Input interface                                                   */
/* ------------------------------------------------------------------ */

export interface UseKeyboardShortcutsInput {
	/** Container element ref, kept for API compatibility with the shell. */
	containerRef: React.RefObject<HTMLDivElement | null>;

	mode: ViewerMode;
	canEdit: boolean;
	/** Omitted preserves custom callers; false leaves native paste unclaimed. */
	canPaste?: boolean;

	/** Whether any element is currently being inline-edited (text box). */
	inlineEditingElementId: string | null;
	/** Whether a table cell is being edited. */
	tableEditorState: TableCellEditorState | null;
	/** Current drawing tool: shortcuts are suppressed when drawing. */
	activeTool: DrawingTool;

	/** Whether at least one element is selected. */
	hasSelection: boolean;
	/**
	 * Host UI customisation: disabled / remapped editor shortcuts, and the
	 * `presentMode` feature gate for F5 / Shift+F5. Omitted customises nothing.
	 */
	customization?: ResolvedCustomization;
	/** The IDs of the currently selected elements (effective). */
	effectiveSelectedIds: string[];

	// -- Action callbacks --------------------------------------------
	onDelete: () => void;
	onCopy: () => void;
	onCut: () => void;
	onPaste: () => void;
	onDuplicate: () => void;
	onUndo: () => void;
	onRedo: () => void;
	onSelectAll: () => void;
	onEscape: () => void;
	/** Move selected elements by (dx, dy). */
	onNudge: (dx: number, dy: number) => void;
	/** Group the selection into one group element (Ctrl/Cmd+G). */
	onGroup?: () => void;
	/** Ungroup the selected group (Ctrl/Cmd+Shift+G). */
	onUngroup?: () => void;
	/** Show or hide the keyboard-shortcut reference ("?"). */
	onToggleShortcuts?: () => void;
	/**
	 * Open or close the find bar (Ctrl/Cmd+F). Left unwired by the default
	 * shell, where `useFindReplace` owns the find state and resolves the same
	 * chord against the same shared keymap; the case exists so the switch stays
	 * exhaustive and so a host driving this hook directly can handle it.
	 */
	onFind?: () => void;
	/** Navigate to previous visible slide (edit mode, no selection). */
	onPrevSlide?: () => void;
	/** Navigate to next visible slide (edit mode, no selection). */
	onNextSlide?: () => void;
	/** `F5`: start the show on its first slide. Same entry point as the Slide Show ribbon's "From Beginning". */
	onStartShowFromBeginning: () => void;
	/** `Shift+F5`: start the show on the current slide. Same entry point as the ribbon's "From Current Slide". */
	onStartShowFromCurrent: () => void;

	/** Ctrl/Cmd+L: set paragraph alignment to left. */
	onAlignLeft?: () => void;
	/** Ctrl/Cmd+E: set paragraph alignment to center. */
	onAlignCenter?: () => void;
	/** Ctrl/Cmd+R: set paragraph alignment to right. */
	onAlignRight?: () => void;
	/** Ctrl/Cmd+J: set paragraph alignment to justify. */
	onAlignJustify?: () => void;
	/** Ctrl/Cmd+Shift+> or Ctrl/Cmd+]: step the selection's font size up PowerPoint's size ladder. */
	onIncreaseFontSize?: () => void;
	/** Ctrl/Cmd+Shift+< or Ctrl/Cmd+[: step the selection's font size down PowerPoint's size ladder. */
	onDecreaseFontSize?: () => void;
	/** Ctrl/Cmd+Shift+C: copy the selection's formatting (format painter). */
	onCopyFormat?: () => void;
	/** Ctrl/Cmd+Shift+V: apply the copied formatting to the selection. */
	onPasteFormat?: () => void;
	/** Ctrl/Cmd+M: insert a new slide. */
	onNewSlide?: () => void;
	/** Ctrl/Cmd+K: open the hyperlink dialog. */
	onHyperlink?: () => void;
	/** Ctrl/Cmd+H: open Find and Replace. */
	onFindReplace?: () => void;
	/** Ctrl/Cmd+Space: clear character formatting on the selection. */
	onClearFormatting?: () => void;
	/** Tab: select the next element on the slide. */
	onCycleSelectionNext?: () => void;
	/** Shift+Tab: select the previous element on the slide. */
	onCycleSelectionPrev?: () => void;
	/** Ctrl/Cmd+Alt+V: open the Paste Special dialog. */
	onPasteSpecial?: () => void;
}

/* ------------------------------------------------------------------ */
/*  Hook                                                              */
/* ------------------------------------------------------------------ */

export function useKeyboardShortcuts(input: UseKeyboardShortcutsInput): void {
	// Store everything in a ref so the keydown closure never goes stale
	// and we don't need to re-attach the listener on every render.
	const inputRef = useRef(input);
	inputRef.current = input;

	const handleKeyDown = useCallback((e: KeyboardEvent) => {
		const current = inputRef.current;
		dispatchEditorKey(
			e,
			{
				canEdit: current.canEdit,
				...(current.canPaste === undefined ? {} : { canPaste: current.canPaste }),
				// The Slide Master view is an editing surface, not a viewing one, so
				// only the other modes count as "not editing" for the editor keymap.
				isPresenting: current.mode !== 'edit' && current.mode !== 'master',
				// F5 is only inert while the show itself runs: a read-only viewer can
				// still start a show (it has a "From Beginning" button too).
				isShowRunning: current.mode === 'present',
				canStartShow:
					!current.customization || isFeatureEnabled(current.customization, 'presentMode'),
				hasSelection: current.hasSelection,
				isEditingText: Boolean(
					current.inlineEditingElementId || current.tableEditorState?.isEditing,
				),
				isDrawing: current.activeTool !== 'select',
				keyboard: current.customization?.keyboard,
			},
			toEditorKeyHandlers(current),
		);
	}, []);

	useEffect(() => {
		// Once, on window: keydown from anywhere inside the viewer bubbles here,
		// so a second container-scoped listener would only double-fire.
		window.addEventListener('keydown', handleKeyDown);
		return () => {
			window.removeEventListener('keydown', handleKeyDown);
		};
	}, [handleKeyDown]);
}

/** The hook's `on*` callbacks under the shared handler names. */
function toEditorKeyHandlers(current: UseKeyboardShortcutsInput): EditorKeyHandlers {
	return {
		startShowFromBeginning: current.onStartShowFromBeginning,
		startShowFromCurrent: current.onStartShowFromCurrent,
		escape: current.onEscape,
		delete: current.onDelete,
		undo: current.onUndo,
		redo: current.onRedo,
		copy: current.onCopy,
		cut: current.onCut,
		paste: current.onPaste,
		duplicate: current.onDuplicate,
		selectAll: current.onSelectAll,
		group: current.onGroup,
		ungroup: current.onUngroup,
		toggleShortcuts: current.onToggleShortcuts,
		find: current.onFind,
		nudge: current.onNudge,
		prevSlide: current.onPrevSlide,
		nextSlide: current.onNextSlide,
		alignLeft: current.onAlignLeft,
		alignCenter: current.onAlignCenter,
		alignRight: current.onAlignRight,
		alignJustify: current.onAlignJustify,
		increaseFontSize: current.onIncreaseFontSize,
		decreaseFontSize: current.onDecreaseFontSize,
		copyFormat: current.onCopyFormat,
		pasteFormat: current.onPasteFormat,
		newSlide: current.onNewSlide,
		hyperlink: current.onHyperlink,
		findReplace: current.onFindReplace,
		clearFormatting: current.onClearFormatting,
		cycleSelectionNext: current.onCycleSelectionNext,
		cycleSelectionPrev: current.onCycleSelectionPrev,
		pasteSpecial: current.onPasteSpecial,
	};
}
