/** Framework-neutral keyboard focus management for modal dialogs. */

import {
	containsComposed as containsDeep,
	modalActiveElement as deepActiveElement,
	modalFocusableElements as focusableElements,
} from '../../dialog/focus';
export {
	MODAL_FOCUSABLE_SELECTOR,
	modalActiveElement,
	modalFocusableElements,
} from '../../dialog/focus';

export interface ModalFocusOptions {
	/** Element that should receive initial focus. Defaults to the first control. */
	initialFocus?: HTMLElement | null;
	/** Invoked after Escape is consumed. */
	onEscape?: () => void;
	/** Restore focus to the opener when the manager is released. */
	restoreFocus?: boolean;
	/**
	 * Which keys the modal handles. `document` (the default) owns the whole keyboard, for a modal
	 * that covers the page. `panel` handles only keys pressed inside the panel, for a surface that
	 * covers just its editor and leaves the page around it usable.
	 */
	keyScope?: 'document' | 'panel';
}

/**
 * Moves focus into a modal, traps Tab within it, handles Escape, and returns
 * focus to the opener on cleanup. The caller owns visibility and teardown.
 */
export function activateModalFocus(
	panel: HTMLElement,
	options: ModalFocusOptions = {},
): () => void {
	const doc = panel.ownerDocument;
	const activeElement = deepActiveElement(doc);
	const opener = activeElement instanceof HTMLElement ? activeElement : null;
	const restoreFocus = options.restoreFocus ?? true;

	function focusInitial(): void {
		const candidate = options.initialFocus ?? focusableElements(panel)[0] ?? panel;
		candidate.focus();
	}

	function onKeydown(event: KeyboardEvent): void {
		if (!panel.isConnected) {
			doc.removeEventListener('keydown', onKeydown, true);
			return;
		}
		if (options.keyScope === 'panel' && !event.composedPath().includes(panel)) {
			return;
		}
		if (event.key === 'Escape' && options.onEscape) {
			// An open select popup consumes the first Escape; the dialog stays open.
			if (
				event
					.composedPath()
					.some(
						(node) =>
							node instanceof Element &&
							node.localName === 'pptx-ui-select' &&
							node.hasAttribute('open'),
					)
			) {
				return;
			}
			event.preventDefault();
			event.stopPropagation();
			options.onEscape();
			return;
		}
		if (event.key !== 'Tab') {
			return;
		}

		const focusable = focusableElements(panel);
		if (focusable.length === 0) {
			event.preventDefault();
			panel.focus();
			return;
		}

		const first = focusable[0];
		const last = focusable[focusable.length - 1];
		const active = deepActiveElement(doc);
		if (event.shiftKey && (active === first || !containsDeep(panel, active))) {
			event.preventDefault();
			last.focus();
		} else if (!event.shiftKey && (active === last || !containsDeep(panel, active))) {
			event.preventDefault();
			first.focus();
		}
	}

	doc.addEventListener('keydown', onKeydown, true);
	queueMicrotask(focusInitial);

	return () => {
		doc.removeEventListener('keydown', onKeydown, true);
		if (restoreFocus && opener?.isConnected) {
			opener.focus();
		}
	};
}
