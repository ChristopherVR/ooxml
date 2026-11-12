import { RIBBON_AUTO_COLLAPSE_MEDIA } from 'ooxml-ui/pptx';
import type { RibbonDisplayController } from 'ooxml-ui/pptx';
import { onBeforeUnmount, onMounted } from 'vue';

const RIBBON_CHROME = '[data-pptx-chrome="ribbon"]';
const REVEAL_CHROME = '[data-pptx-chrome="ribbon-reveal"]';
const MENU_CHROME =
	'[data-pptx-chrome="ribbon-display-options"], [data-pptx-chrome="ribbon-display-menu"]';

/**
 * The document-level gestures of Ribbon Display Options: a press outside the ribbon (or its
 * reveal strip) puts the peeked commands away, a press inside it outside the menu closes the
 * menu, Ctrl+F1 and Escape reach the controller from this viewer only, and the
 * "Collapse the ribbon automatically" viewport query is kept current. Registered on mount and
 * removed before unmount, so a viewer that goes away leaves no listener behind.
 */
export function useRibbonDisplayListeners(
	controller: RibbonDisplayController,
	viewerRoot: () => HTMLElement | null,
): void {
	let short: MediaQueryList | undefined;

	const onPointerDown = (event: PointerEvent): void => {
		const target = event.target instanceof Element ? event.target : null;
		if (!target?.closest(RIBBON_CHROME) && !target?.closest(REVEAL_CHROME)) {
			controller.dismiss();
		} else if (!target?.closest(MENU_CHROME)) {
			controller.setMenuOpen(false);
		}
	};

	const onKeyDown = (event: KeyboardEvent): void => {
		const root = viewerRoot();
		const target = event.target as Node | null;
		// Only this viewer's shortcuts: the focus is inside it, or nowhere in particular.
		if (root && (target === document.body || (target !== null && root.contains(target)))) {
			controller.handleKeydown(event);
		}
	};

	const onShortViewport = (): void => {
		controller.setShortViewport(short?.matches ?? false);
	};

	onMounted(() => {
		document.addEventListener('pointerdown', onPointerDown);
		document.addEventListener('keydown', onKeyDown);
		short = window.matchMedia?.(RIBBON_AUTO_COLLAPSE_MEDIA);
		short?.addEventListener('change', onShortViewport);
		onShortViewport();
	});

	onBeforeUnmount(() => {
		document.removeEventListener('pointerdown', onPointerDown);
		document.removeEventListener('keydown', onKeyDown);
		short?.removeEventListener('change', onShortViewport);
		short = undefined;
	});
}
