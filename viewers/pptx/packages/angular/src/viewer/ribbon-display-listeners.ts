/**
 * ribbon-display-listeners.ts: the document-level gestures of Ribbon Display Options, forwarded
 * to the shared controller: a press outside the ribbon (puts the peeked commands away), a press
 * inside it outside the menu (closes the menu), the keys (Ctrl+F1, Escape) and the viewport
 * that collapses the ribbon automatically. Every listener is removed when the viewer is destroyed.
 */
import type { DestroyRef } from '@angular/core';
import { RIBBON_AUTO_COLLAPSE_MEDIA } from 'ooxml-ui/pptx';
import type { RibbonDisplayController } from 'ooxml-ui/pptx';

/** Attributes of the parts a press may land on; see the ribbon templates. */
const RIBBON_SELECTOR = '[data-pptx-chrome="ribbon"]';
const REVEAL_SELECTOR = '[data-pptx-chrome="ribbon-reveal"]';
const MENU_PARTS_SELECTOR =
	'[data-pptx-chrome="ribbon-display-options"], [data-pptx-chrome="ribbon-display-menu"]';

export interface RibbonDisplayListenersOptions {
	readonly controller: RibbonDisplayController;
	/** The viewer root; keys count only inside it, or when nothing has focus. */
	readonly root: () => HTMLElement | undefined;
	readonly destroyRef: DestroyRef;
}

/** Registers the listeners and removes them on destroy. Does nothing without a document. */
export function listenRibbonDisplayGestures({
	controller,
	root,
	destroyRef,
}: RibbonDisplayListenersOptions): void {
	if (typeof document === 'undefined') {
		return;
	}
	const doc = document;

	const onPointerDown = (event: Event): void => {
		const target = event.target instanceof Element ? event.target : null;
		if (!target) {
			return;
		}
		const inRibbon = target.closest(RIBBON_SELECTOR) !== null;
		const inStrip = target.closest(REVEAL_SELECTOR) !== null;
		if (!inRibbon && !inStrip) {
			controller.dismiss();
		} else if (inRibbon && target.closest(MENU_PARTS_SELECTOR) === null) {
			controller.setMenuOpen(false);
		}
	};

	const onKeyDown = (event: KeyboardEvent): void => {
		const viewer = root();
		const target = event.target as Node | null;
		// Only this viewer's shortcuts: the focus is inside it, or nowhere in particular.
		if (viewer && (target === doc.body || (target !== null && viewer.contains(target)))) {
			controller.handleKeydown(event);
		}
	};

	const short = doc.defaultView?.matchMedia?.(RIBBON_AUTO_COLLAPSE_MEDIA);
	const onShortViewport = (): void => controller.setShortViewport(short?.matches ?? false);

	doc.addEventListener('pointerdown', onPointerDown);
	doc.addEventListener('keydown', onKeyDown);
	short?.addEventListener('change', onShortViewport);
	onShortViewport();

	destroyRef.onDestroy(() => {
		doc.removeEventListener('pointerdown', onPointerDown);
		doc.removeEventListener('keydown', onKeyDown);
		short?.removeEventListener('change', onShortViewport);
	});
}
