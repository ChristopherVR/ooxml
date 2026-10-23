/**
 * The keyboard-shortcut action contract: the names `mapEditorKey` can
 * resolve to, and the callbacks `useKeyboardShortcuts` dispatches them to.
 *
 * Split out of `useKeyboardShortcuts.ts` to keep that file under this repo's
 * file-size limit; this module is the public contract, the other file is the
 * dispatch implementation.
 *
 * @module composables/shortcut-actions
 */
import type { EditorKeyActionName, EditorKeyHandlers } from 'ooxml-ui/pptx';

/**
 * The set of action identifiers the registry can dispatch: an alias of the
 * shared keymap's own action union, not a hand-copied one, so a new shared
 * action can never silently fall through the dispatcher.
 */
export type ShortcutActionName = EditorKeyActionName;

/**
 * Action callbacks the registry dispatches to: the shared `EditorKeyHandlers`,
 * one optional callback per keymap action (plus `nudge(dx, dy)`). A missing
 * callback makes the shortcut a no-op, but it is still matched and claimed, so
 * the browser default is still prevented (the shared `dispatchEditorKey` rule).
 */
export type ShortcutActions = EditorKeyHandlers;

/** Result of matching a keyboard event against the catalog. */
export interface MatchedShortcut {
	/** The dispatched action, or `null` when the event matches nothing. */
	action: ShortcutActionName | null;
	/** Nudge delta (only set when `action === 'nudge'`). */
	dx?: number;
	/** Nudge delta (only set when `action === 'nudge'`). */
	dy?: number;
}
