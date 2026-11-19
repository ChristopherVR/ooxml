/**
 * Keyboard access to the slide and the editor's panes, as in PowerPoint.
 *
 * In PowerPoint Tab walks the objects of the slide pane and F6 moves between
 * panes. Taking every Tab from the first focusable element on a web page would
 * trap keyboard users there, so the slide is one ordinary Tab stop until the
 * user goes in:
 *
 * - **Pane** (`pane`): the slide viewport itself is focused and Tab moves on. Enter
 *   goes in and selects the first object.
 * - **Objects** (`objects`): Tab and Shift+Tab walk the objects. F6 landing on the
 *   slide or pressing the mouse on it also goes in. Escape clears the selection
 *   and, with nothing selected, comes back out; F6 leaves at any time.
 *
 * The state lives on the viewport element, so the shared keymap and all five
 * bindings read one answer and a replaced viewport simply starts over.
 *
 * @module render/slide-keyboard-focus
 */
import { containsComposed, modalActiveElement, modalFocusableElements } from '../../dialog/focus';
import { findCanvasElementNode } from './canvas-element-node';
import { isEditorTextInputTarget } from './editor-key-target';

/** Where a key press sits relative to the slide. */
export type SlideKeyFocus = 'outside' | 'pane' | 'objects';

const VIEWPORT_SELECTOR = '[data-pptx-viewport]';
const VIEWER_SELECTOR = '[data-pptx-editor-chrome]';
const STAGE_SELECTOR = '[aria-roledescription="slide"]';
const OBJECT_SELECTOR = '[data-element-id]';
const OBJECTS_ATTRIBUTE = 'data-pptx-slide-objects';
const STATUS_ATTRIBUTE = 'data-pptx-slide-status';

/** The panes F6 visits, clockwise around the window as PowerPoint moves. */
const PANE_SELECTORS = [
	'[data-pptx-chrome="ribbon"]',
	VIEWPORT_SELECTOR,
	'[data-pptx-chrome="inspector"]',
	'[data-pptx-chrome="notes"]',
	'pptx-ui-status-bar, [data-pptx-chrome="status-bar"]',
	'[data-pptx-chrome="slides"]',
] as const;

const VISUALLY_HIDDEN =
	'position:absolute;width:1px;height:1px;overflow:hidden;clip-path:inset(50%);white-space:nowrap';

function viewportOf(target: unknown): HTMLElement | null {
	const element = target as Partial<Element> | null;
	return typeof element?.closest === 'function'
		? element.closest<HTMLElement>(VIEWPORT_SELECTOR)
		: null;
}

/** The element a key event really came from, inside any open shadow root. */
export function editorKeyEventTarget(event: {
	target?: EventTarget | null;
	composedPath?: () => EventTarget[];
}): EventTarget | null {
	return event.composedPath?.()[0] ?? event.target ?? null;
}

/** Whether a key press from `target` is on the slide, and whether the slide was entered. */
export function slideKeyFocus(target: unknown): SlideKeyFocus {
	const viewport = viewportOf(target);
	if (viewport?.hasAttribute(OBJECTS_ATTRIBUTE)) {
		return 'objects';
	}
	// A link or button on the slide keeps its own Enter and Tab.
	return viewport && viewport === target ? 'pane' : 'outside';
}

/** Go into the slide's objects, unless it has none to walk. */
export function enterSlideObjects(target: unknown): void {
	const viewport = viewportOf(target);
	if (viewport?.querySelector(OBJECT_SELECTOR)) {
		viewport.setAttribute(OBJECTS_ATTRIBUTE, '');
	}
}

/** Come back out, so Tab moves on from the slide again. */
export function leaveSlideObjects(target: unknown): void {
	viewportOf(target)?.removeAttribute(OBJECTS_ATTRIBUTE);
}

/** Focus the slide of the viewer that holds `from` (a notes field, say), as its one Tab stop. */
export function focusSlidePane(from: unknown): boolean {
	const element = from as Partial<Element> | null;
	if (typeof element?.closest !== 'function') {
		return false;
	}
	const scope = element.closest(VIEWER_SELECTOR) ?? element.ownerDocument;
	const viewport = scope?.querySelector<HTMLElement>(VIEWPORT_SELECTOR);
	viewport?.focus({ preventScroll: true });
	return Boolean(viewport);
}

/** Enter on a selected object: open it for text editing, as a double-click does. */
export function editSlideObjectText(root: ParentNode, elementId: string | null | undefined): void {
	const node = elementId ? findCanvasElementNode(root, elementId, { canvasOnly: true }) : null;
	if (!node) {
		return;
	}
	const box = node.getBoundingClientRect();
	node.dispatchEvent(
		new MouseEvent('dblclick', {
			bubbles: true,
			cancelable: true,
			composed: true,
			detail: 2,
			clientX: box.left + box.width / 2,
			clientY: box.top + box.height / 2,
		}),
	);
}

/** Say which object Tab reached: selecting one moves no focus for a screen reader to follow. */
export function announceSlideObject(root: ParentNode, elementId: string): void {
	const node = findCanvasElementNode(root, elementId, { canvasOnly: true });
	const label = node?.getAttribute('aria-label');
	const viewer = node?.closest(VIEWER_SELECTOR);
	if (!label || !viewer) {
		return;
	}
	let status = viewer.querySelector<HTMLElement>(`[${STATUS_ATTRIBUTE}]`);
	if (!status) {
		status = viewer.ownerDocument.createElement('div');
		status.setAttribute(STATUS_ATTRIBUTE, '');
		status.setAttribute('role', 'status');
		status.style.cssText = VISUALLY_HIDDEN;
		viewer.append(status);
	}
	status.textContent = label;
}

function isShown(element: HTMLElement): boolean {
	return element.getClientRects().length > 0;
}

/** The panes present and showing, in F6 order. */
function visiblePanes(scope: ParentNode): HTMLElement[] {
	return PANE_SELECTORS.flatMap((selector) => {
		const pane = scope.querySelector<HTMLElement>(selector);
		return pane && isShown(pane) ? [pane] : [];
	});
}

/** Where focus lands in a pane: its current item (selected tab, active slide), else its first stop. */
function paneEntry(pane: HTMLElement): HTMLElement {
	const stops = modalFocusableElements(pane);
	const current = stops.find(
		(stop) =>
			stop.getAttribute('aria-selected') === 'true' ||
			(stop.getAttribute('aria-current') ?? 'false') !== 'false',
	);
	const entry = current ?? stops[0];
	if (entry) {
		return entry;
	}
	pane.tabIndex = -1;
	return pane;
}

/**
 * Move focus to the next or previous pane of the viewer. Returns false when the
 * viewer shows no pane to move to.
 */
export function moveEditorPaneFocus(viewport: HTMLElement, direction: 1 | -1): boolean {
	const doc = viewport.ownerDocument;
	const panes = visiblePanes(viewport.closest(VIEWER_SELECTOR) ?? doc);
	if (panes.length === 0) {
		return false;
	}
	const active = modalActiveElement(doc);
	const index = panes.findIndex((pane) => containsComposed(pane, active));
	const from = index < 0 && direction < 0 ? 0 : index;
	const pane = panes[(from + direction + panes.length) % panes.length]!;
	if (pane === viewport) {
		viewport.focus({ preventScroll: true });
		enterSlideObjects(viewport);
	} else {
		paneEntry(pane).focus({ preventScroll: true });
	}
	return true;
}

/**
 * True when the key press belongs to a dialog over the editor, which keeps the keyboard
 * to itself. A host dialog that holds the whole viewer does not count.
 */
export function isInsideModal(event: { composedPath?: () => EventTarget[] }): boolean {
	const modal = (event.composedPath?.() ?? []).find(
		(node): node is Element =>
			node instanceof Element &&
			(node.localName === 'dialog' || node.getAttribute('aria-modal') === 'true'),
	);
	return Boolean(modal && !modal.querySelector(VIEWPORT_SELECTOR));
}

/**
 * Make the slide viewport at or inside `host` the slide pane: one labelled Tab stop that the mouse
 * and F6 can enter, with F6 / Shift+F6 cycling the viewer's panes. Some bindings scroll an element
 * around the viewport and pass that, and a canvas can mount later, so the viewport is looked up
 * rather than held. Returns the detach function.
 */
export function attachSlideKeyboardFocus(host: HTMLElement, isActive: () => boolean): () => void {
	const doc = host.ownerDocument;
	let viewport: HTMLElement | null = null;

	const label = (): void => {
		const name = viewport?.querySelector(STAGE_SELECTOR)?.getAttribute('aria-label');
		viewport?.setAttribute('aria-label', name || 'Slide');
	};
	const adopt = (): void => {
		if (viewport?.isConnected && host.contains(viewport)) {
			return;
		}
		viewport = host.matches(VIEWPORT_SELECTOR)
			? host
			: host.querySelector<HTMLElement>(VIEWPORT_SELECTOR);
		if (viewport) {
			viewport.tabIndex = 0;
			viewport.setAttribute('role', 'group');
			viewport.setAttribute('aria-keyshortcuts', 'Enter Escape F6 Shift+F6');
			label();
		}
	};
	adopt();
	const observer = new MutationObserver(adopt);
	observer.observe(host, { childList: true, subtree: true });

	const ownsKeyboard = (slide: HTMLElement): boolean => {
		const active = modalActiveElement(doc);
		if (active && active !== doc.body) {
			return containsComposed(slide.closest(VIEWER_SELECTOR) ?? slide, active);
		}
		// Nothing is focused: the first viewer on the page answers.
		return doc.querySelector(VIEWPORT_SELECTOR) === slide;
	};
	const onKeyDown = (event: KeyboardEvent): void => {
		if (
			event.key !== 'F6' ||
			event.ctrlKey ||
			event.metaKey ||
			event.altKey ||
			event.defaultPrevented ||
			!viewport ||
			!isActive() ||
			isInsideModal(event) ||
			!ownsKeyboard(viewport)
		) {
			return;
		}
		if (moveEditorPaneFocus(viewport, event.shiftKey ? -1 : 1)) {
			event.preventDefault();
		}
	};
	// Pressing on the slide goes in, as a click selects there. A gesture handler that
	// prevents the default also prevents the focus move, so it is made here. A field
	// being edited keeps the caret. A scripted focus would match :focus-visible, so it
	// says it came from the pointer and the ring waits for the keyboard.
	const onPointerDown = (event: PointerEvent): void => {
		if (!viewport || viewportOf(event.target) !== viewport || !isActive()) {
			return;
		}
		enterSlideObjects(viewport);
		const active = modalActiveElement(doc);
		if (!isEditorTextInputTarget(active) && !isEditorTextInputTarget(event.target)) {
			viewport.focus({ preventScroll: true, focusVisible: false });
		}
	};
	const onFocusIn = (event: FocusEvent): void => {
		if (viewport && viewportOf(event.target) === viewport) {
			label();
		}
	};
	const onFocusOut = (event: FocusEvent): void => {
		const next = event.relatedTarget as Node | null;
		if (viewport && next && !containsComposed(viewport, next)) {
			leaveSlideObjects(viewport);
		}
	};

	doc.addEventListener('keydown', onKeyDown);
	host.addEventListener('pointerdown', onPointerDown, true);
	host.addEventListener('focusin', onFocusIn);
	host.addEventListener('focusout', onFocusOut);
	return () => {
		observer.disconnect();
		doc.removeEventListener('keydown', onKeyDown);
		host.removeEventListener('pointerdown', onPointerDown, true);
		host.removeEventListener('focusin', onFocusIn);
		host.removeEventListener('focusout', onFocusOut);
		viewport?.closest(VIEWER_SELECTOR)?.querySelector(`[${STATUS_ATTRIBUTE}]`)?.remove();
	};
}
