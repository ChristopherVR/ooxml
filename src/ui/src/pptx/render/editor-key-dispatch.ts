/**
 * One dispatcher from a key press to the editor command handlers.
 *
 * `mapEditorKey` decides WHICH command a key means; before this module each
 * binding then carried its own ~30-case switch from that action to its
 * callbacks, under five differently named handler interfaces. The switch is
 * here now, keyed by the shared action names, so a new keymap action reaches
 * every binding the moment one handler is wired, and the "claim the key"
 * (`preventDefault`) rule is the same everywhere.
 *
 * What each command DOES stays in the binding: the handlers close over binding
 * state (Escape unwinds view state only the binding owns).
 *
 * @module render/editor-key-dispatch
 */

import type { ResolvedKeyboardCustomization } from './customization/customization-keymap';
import { mapCustomizedEditorKey } from './customization/customization-keymap';
import { isEditorControlTarget, isEditorTextInputTarget } from './editor-key-target';
import type { EditorKeyActionName, EditorKeyGuard, EditorKeyInput } from './editor-keymap';
import { mapSlideShowStartKey } from './slide-show-start-keymap';

/** The one named handler interface: one optional callback per keymap action. */
export type EditorKeyHandlers = {
	[Action in Exclude<EditorKeyActionName, 'nudge'>]?: () => void;
} & {
	/** Arrow keys with a selection: move it by (dx, dy) slide pixels. */
	nudge?: (dx: number, dy: number) => void;
	/** F5: start the show on its first slide (checked before the editor gates). */
	startShowFromBeginning?: () => void;
	/** Shift+F5: start the show on the current slide. */
	startShowFromCurrent?: () => void;
};

/** The event fields the dispatcher reads (a DOM `KeyboardEvent` satisfies it). */
export interface EditorKeyDispatchEvent extends EditorKeyInput {
	target?: EventTarget | null;
	preventDefault(): void;
}

/**
 * The binding state the keymap gates on. The two target flags are read from
 * the event when omitted.
 */
export interface EditorKeyDispatchState extends Partial<EditorKeyGuard> {
	/** The host's keyboard customisation (disabled / remapped commands). */
	keyboard?: ResolvedKeyboardCustomization | undefined;
	/** False turns F5 / Shift+F5 off (the host disabled the `presentMode` feature). */
	canStartShow?: boolean;
	/**
	 * A slide show is running, so F5 does nothing. Defaults to `isPresenting`;
	 * set it apart when the editor gate counts more modes as "not editing".
	 */
	isShowRunning?: boolean;
	/**
	 * Actions another listener owns (for example slide paging handled by the
	 * root navigation keymap): neither dispatched nor claimed.
	 */
	passThrough?: readonly EditorKeyActionName[];
}

/** What a key press did: an editor action, a show start, or nothing. */
export type EditorKeyDispatchResult =
	| EditorKeyActionName
	| 'startShowFromBeginning'
	| 'startShowFromCurrent'
	| null;

function startShow(
	event: EditorKeyDispatchEvent,
	state: EditorKeyDispatchState,
	handlers: EditorKeyHandlers,
): EditorKeyDispatchResult {
	if (state.canStartShow === false) {
		return null;
	}
	if (!handlers.startShowFromBeginning && !handlers.startShowFromCurrent) {
		return null;
	}
	const start = mapSlideShowStartKey(event, {
		isPresenting: state.isShowRunning ?? state.isPresenting === true,
	});
	if (start === null) {
		return null;
	}
	const name = start === 'fromBeginning' ? 'startShowFromBeginning' : 'startShowFromCurrent';
	const handler = handlers[name];
	if (!handler) {
		return null;
	}
	event.preventDefault();
	handler();
	return name;
}

/**
 * Resolve `event` through the shared (customised) keymap and call the matching
 * handler.
 *
 * F5 / Shift+F5 are checked first when a start handler is wired: PowerPoint
 * starts the show with the caret in a text box and on a read-only host, both of
 * which the editor keymap gates against. Every other resolved action is claimed
 * (`preventDefault`) even when its handler is omitted, so a chord the editor
 * owns never leaks to the browser; list an action in `passThrough` to leave it
 * to another listener instead.
 */
export function dispatchEditorKey(
	event: EditorKeyDispatchEvent,
	state: EditorKeyDispatchState,
	handlers: EditorKeyHandlers,
): EditorKeyDispatchResult {
	const started = startShow(event, state, handlers);
	if (started !== null) {
		return started;
	}
	const {
		keyboard,
		canStartShow: _canStartShow,
		isShowRunning: _isShowRunning,
		passThrough,
		...guard
	} = state;
	const target = event.target ?? null;
	const { action, dx, dy } = mapCustomizedEditorKey(
		event,
		{
			...guard,
			isTextInputTarget: guard.isTextInputTarget ?? isEditorTextInputTarget(target),
			isControlTarget: guard.isControlTarget ?? isEditorControlTarget(target),
		},
		keyboard,
	);
	if (action === null || passThrough?.includes(action)) {
		return null;
	}
	event.preventDefault();
	if (action === 'nudge') {
		handlers.nudge?.(dx ?? 0, dy ?? 0);
	} else {
		handlers[action]?.();
	}
	return action;
}
