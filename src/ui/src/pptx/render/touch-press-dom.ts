/**
 * DOM wiring of the touch long-press and double-tap recogniser (issue #290).
 *
 * `installTouchPress(root)` listens to pointer events on a viewer's root and,
 * for touch and pen, turns a long-press into a synthetic `contextmenu` and a
 * double-tap into a synthetic `dblclick` on the element the finger pressed.
 * Every binding's existing `contextmenu` / `dblclick` handler therefore gets the
 * touch path for free, with the touch coordinates, and nothing is duplicated.
 *
 * Only elements that opt in react: an ancestor carrying
 * `data-pptx-touch-press` with the space-separated actions `menu` (long-press)
 * and/or `dblclick` (double-tap). Editable text keeps the browser's own touch
 * behaviour (selection, the editing callout). After a long-press the click the
 * release produces is swallowed, and a native `contextmenu` or `dblclick` that
 * follows a synthetic one (Android fires both) is dropped so nothing opens twice.
 *
 * @module touch-press-dom
 */

import { createTouchPressRecognizer } from './touch-press';
import type { TouchPressGesture } from './touch-press';

/** The opt-in attribute; its value lists the actions (`menu`, `dblclick`). */
export const TOUCH_PRESS_ATTRIBUTE = 'data-pptx-touch-press';

/** A touch action a control opts into. */
export type TouchPressAction = 'menu' | 'dblclick';

/** How long (ms) after a synthetic gesture its native twin is ignored. */
export const TOUCH_PRESS_SUPPRESS_MS = 800;

interface PressTarget {
	readonly origin: Element;
	/** The nearest ancestor that opted into `menu`, if any. */
	readonly menuHost: Element | null;
	/** The nearest ancestor that opted into `dblclick`, if any. */
	readonly dblHost: Element | null;
}

function closestWithAction(origin: Element, action: TouchPressAction, root: Element) {
	const host = origin.closest(`[${TOUCH_PRESS_ATTRIBUTE}~="${action}"]`);
	return host && root.contains(host) ? host : null;
}

/** Reads the actions an opt-in host declares. */
export function touchPressActions(host: Element): ReadonlySet<TouchPressAction> {
	const value = host.getAttribute(TOUCH_PRESS_ATTRIBUTE) ?? '';
	const actions = new Set<TouchPressAction>();
	for (const token of value.split(/\s+/)) {
		if (token === 'menu' || token === 'dblclick') {
			actions.add(token);
		}
	}
	return actions;
}

function isEditable(element: Element): boolean {
	const editable = element.closest('input, textarea, select, [contenteditable]');
	if (!editable) {
		return false;
	}
	const attr = editable.getAttribute('contenteditable');
	return attr === null || attr !== 'false';
}

function eventOrigin(event: Event): Element | null {
	const path = typeof event.composedPath === 'function' ? event.composedPath() : [];
	const first = path[0] ?? event.target;
	if (first && typeof (first as Element).closest === 'function') {
		return first as Element;
	}
	const parent = (first as Node | null)?.parentElement;
	return parent ?? null;
}

function dispatchMouse(target: Element, type: 'contextmenu' | 'dblclick', x: number, y: number) {
	const view = target.ownerDocument.defaultView;
	const MouseEventCtor = view?.MouseEvent ?? MouseEvent;
	target.dispatchEvent(
		new MouseEventCtor(type, {
			bubbles: true,
			cancelable: true,
			composed: true,
			clientX: x,
			clientY: y,
			button: type === 'contextmenu' ? 2 : 0,
			detail: type === 'dblclick' ? 2 : 0,
			...(view ? { view } : {}),
		}),
	);
}

/**
 * Starts the touch long-press / double-tap wiring on `root` and returns the
 * teardown. `root` is the binding's viewer root (or any ancestor of the opt-in
 * controls); listeners run in the capture phase so a control that stops
 * propagation of `pointerdown` still gets its gesture.
 */
export function installTouchPress(root: HTMLElement): () => void {
	const now = (): number => root.ownerDocument.defaultView?.performance.now() ?? Date.now();
	let suppressClickUntil = 0;
	let suppressMenuUntil = 0;
	let suppressDblClickUntil = 0;
	let selectLocked: { element: HTMLElement; previous: string } | null = null;

	const unlockSelect = (): void => {
		if (selectLocked) {
			selectLocked.element.style.setProperty('-webkit-user-select', selectLocked.previous);
			selectLocked = null;
		}
	};

	const target = (gesture: TouchPressGesture<PressTarget>, host: Element): Element =>
		gesture.target.origin.isConnected ? gesture.target.origin : host;

	const recognizer = createTouchPressRecognizer<PressTarget>(
		{
			onLongPress: (gesture) => {
				const host = gesture.target.menuHost;
				if (!host?.isConnected) {
					return;
				}
				suppressClickUntil = Number.POSITIVE_INFINITY;
				suppressMenuUntil = now() + TOUCH_PRESS_SUPPRESS_MS;
				dispatchMouse(target(gesture, host), 'contextmenu', gesture.clientX, gesture.clientY);
			},
			onDoubleTap: (gesture) => {
				const host = gesture.target.dblHost;
				if (!host?.isConnected) {
					return;
				}
				suppressDblClickUntil = now() + TOUCH_PRESS_SUPPRESS_MS;
				dispatchMouse(target(gesture, host), 'dblclick', gesture.clientX, gesture.clientY);
			},
		},
		(a, b) => a.dblHost === b.dblHost,
	);

	const onPointerDown = (event: PointerEvent): void => {
		const origin = eventOrigin(event);
		const menuHost = origin ? closestWithAction(origin, 'menu', root) : null;
		const dblHost = origin ? closestWithAction(origin, 'dblclick', root) : null;
		if (!origin || (!menuHost && !dblHost) || isEditable(origin)) {
			recognizer.cancel();
			return;
		}
		recognizer.onPointerDown(event, { origin, menuHost, dblHost });
		const host = menuHost ?? dblHost;
		if (event.pointerType !== 'mouse' && host instanceof HTMLElement) {
			// Stop iOS from starting a text selection under the held finger.
			unlockSelect();
			selectLocked = {
				element: host,
				previous: host.style.getPropertyValue('-webkit-user-select'),
			};
			host.style.setProperty('-webkit-user-select', 'none');
		}
	};
	const onPointerMove = (event: PointerEvent): void => recognizer.onPointerMove(event);
	const settleClickSuppression = (): void => {
		if (suppressClickUntil === Number.POSITIVE_INFINITY) {
			suppressClickUntil = now() + TOUCH_PRESS_SUPPRESS_MS;
		}
	};
	const onPointerUp = (event: PointerEvent): void => {
		unlockSelect();
		recognizer.onPointerUp(event);
		settleClickSuppression();
	};
	const onPointerCancel = (): void => {
		unlockSelect();
		recognizer.cancel();
		settleClickSuppression();
	};
	const onClick = (event: MouseEvent): void => {
		if (now() <= suppressClickUntil) {
			suppressClickUntil = 0;
			event.preventDefault();
			event.stopPropagation();
		}
	};
	const onContextMenu = (event: MouseEvent): void => {
		if (!event.isTrusted) {
			return;
		}
		if (now() <= suppressMenuUntil) {
			event.preventDefault();
			event.stopPropagation();
			return;
		}
		// The browser opened its own menu first (Android): do not open a second.
		recognizer.cancel();
	};
	const onDblClick = (event: MouseEvent): void => {
		if (event.isTrusted && now() <= suppressDblClickUntil) {
			event.preventDefault();
			event.stopPropagation();
		}
	};

	const listeners: ReadonlyArray<readonly [string, (event: never) => void]> = [
		['pointerdown', onPointerDown],
		['pointermove', onPointerMove],
		['pointerup', onPointerUp],
		['pointercancel', onPointerCancel],
		['click', onClick],
		['contextmenu', onContextMenu],
		['dblclick', onDblClick],
	];
	for (const [type, listener] of listeners) {
		root.addEventListener(type, listener as EventListener, true);
	}
	// No callout ("Open link", "Save image") steals the long-press on iOS.
	const previousCallout = root.style.getPropertyValue('-webkit-touch-callout');
	root.style.setProperty('-webkit-touch-callout', 'none');

	return () => {
		for (const [type, listener] of listeners) {
			root.removeEventListener(type, listener as EventListener, true);
		}
		recognizer.cancel();
		unlockSelect();
		root.style.setProperty('-webkit-touch-callout', previousCallout);
	};
}
