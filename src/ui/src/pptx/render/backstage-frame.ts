import { activateModalFocus } from './modal-focus';

/** The viewer root every binding marks; the File page covers its box and nothing else. */
const BACKSTAGE_HOST_SELECTOR = '[data-pptx-theme-root]';

const FRAME_PROPERTIES = ['inset', 'top', 'left', 'width', 'height'] as const;

/**
 * Keep a `position: fixed` File page on its viewer host's box, so it never covers the page around
 * the viewer (a host sidebar, the rest of a document). Fixed rather than absolute keeps the page
 * clear of the ribbon's own stacking and overflow, wherever a binding mounts it. Returns the
 * release, which puts the page's inline geometry back.
 */
function containBackstage(panel: HTMLElement): () => void {
	const host = panel.parentElement?.closest<HTMLElement>(BACKSTAGE_HOST_SELECTOR);
	const view = panel.ownerDocument.defaultView;
	if (!host || !view) {
		return () => {};
	}
	const place = (): void => {
		const box = host.getBoundingClientRect();
		panel.style.inset = 'auto';
		panel.style.width = `${box.width}px`;
		panel.style.height = `${box.height}px`;
		panel.style.left = `${box.left}px`;
		panel.style.top = `${box.top}px`;
		// A transformed ancestor makes `fixed` relative to itself; correct by what was drawn
		// (a page still hidden has no box to measure).
		const drawn = panel.getBoundingClientRect();
		const shown = drawn.width > 0 || drawn.height > 0;
		if (shown && (drawn.left !== box.left || drawn.top !== box.top)) {
			panel.style.left = `${2 * box.left - drawn.left}px`;
			panel.style.top = `${2 * box.top - drawn.top}px`;
		}
	};
	// Scrolling the page's own content moves nothing; any other scroll may move the host.
	const onScroll = (event: Event): void => {
		if (!(event.target instanceof Node && panel.contains(event.target))) {
			place();
		}
	};
	place();
	const resize = view.ResizeObserver ? new view.ResizeObserver(place) : null;
	resize?.observe(host);
	view.addEventListener('resize', place);
	view.addEventListener('scroll', onScroll, { capture: true, passive: true });
	return () => {
		resize?.disconnect();
		view.removeEventListener('resize', place);
		view.removeEventListener('scroll', onScroll, { capture: true });
		for (const property of FRAME_PROPERTIES) {
			panel.style.removeProperty(property);
		}
	};
}

/**
 * Show a binding's File page: it covers its viewer host only, takes focus, keeps Tab inside, and
 * Escape (pressed inside it) calls `onClose`. Keys pressed elsewhere on the page are left alone,
 * as in Word, Excel and Visio. The release returns focus to the control that opened it (the
 * File tab). Every binding calls this once per open and the release once per close.
 */
export function activateBackstage(panel: HTMLElement, onClose: () => void): () => void {
	// A click on the page's empty space keeps focus inside it, so Escape still reaches it.
	if (!panel.hasAttribute('tabindex')) {
		panel.tabIndex = -1;
	}
	const releaseFrame = containBackstage(panel);
	const releaseFocus = activateModalFocus(panel, { onEscape: onClose, keyScope: 'panel' });
	return () => {
		releaseFrame();
		releaseFocus();
	};
}
