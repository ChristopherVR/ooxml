import type { RibbonDisplayController } from 'ooxml-ui/pptx';
import {
	RIBBON_AUTO_COLLAPSE_MEDIA,
	RIBBON_DISPLAY_OPTIONS_LABEL_KEY,
	RIBBON_DISPLAY_REVEAL_LABEL_KEY,
} from 'ooxml-ui/pptx';

import type { Translator } from '../../i18n';
import { createEl } from '../../render';
import { onDocumentPointerDown } from '../outside-pointer';

export interface RibbonDisplayDeps {
	/** The viewer's Ribbon Display Options state; it outlives this ribbon. */
	controller: RibbonDisplayController;
	/** The ribbon root (`.pptxv-ribbon`); it carries the mode as `data-ribbon-display`. */
	ribbon: HTMLElement;
	/** The tab row, where a double click on a tab collapses or restores the ribbon. */
	tabBar: HTMLElement;
	/** The mode or the peek changed: the ribbon shows or hides its tab panes to match. */
	onLayout(): void;
}

export interface RibbonDisplayView {
	/** The Ribbon Display Options button, for the right end of the tab row. */
	button: HTMLElement;
	/** The anchor holding the menu; it goes right after the tab row. */
	menuAnchor: HTMLElement;
	/** Draw the current state again (after the ribbon entered the document). */
	sync(): void;
	destroy(): void;
}

const CHEVRON =
	'<svg viewBox="0 0 16 16" width="12" height="12" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m4 6 4 4 4-4"/></svg>';

/**
 * PowerPoint's Ribbon Display Options for this binding: the menu at the right end of the tab
 * row, the tab-row peek of "Show tabs only", the strip that brings the ribbon back in
 * full-screen mode, Ctrl+F1 and the double click on a tab. The shared controller owns the state
 * and decides what each mode shows; this only maps its snapshot onto the DOM.
 */
export function createRibbonDisplay(
	doc: Document,
	t: Translator,
	deps: RibbonDisplayDeps,
): RibbonDisplayView {
	const { controller, ribbon, tabBar } = deps;

	const button = createEl(doc, 'button', 'pptxv-ribbon-display-btn');
	button.type = 'button';
	button.innerHTML = CHEVRON;
	button.title = t(RIBBON_DISPLAY_OPTIONS_LABEL_KEY);
	button.setAttribute('aria-label', t(RIBBON_DISPLAY_OPTIONS_LABEL_KEY));
	button.setAttribute('aria-haspopup', 'menu');

	const menuAnchor = createEl(doc, 'div', 'pptxv-ribbon-display-anchor');
	const menu = createEl(doc, 'div', 'pptxv-primary-menu pptxv-ribbon-display-menu');
	menu.setAttribute('role', 'menu');
	menu.setAttribute('aria-label', t(RIBBON_DISPLAY_OPTIONS_LABEL_KEY));
	menuAnchor.appendChild(menu);

	// Full-screen mode keeps this strip in the layout; the ribbon comes back over it.
	const reveal = createEl(doc, 'button', 'pptxv-ribbon-reveal');
	reveal.type = 'button';
	reveal.textContent = '···';
	reveal.title = t(RIBBON_DISPLAY_REVEAL_LABEL_KEY);
	reveal.setAttribute('aria-label', t(RIBBON_DISPLAY_REVEAL_LABEL_KEY));

	const sync = (): void => {
		const { effectiveMode, peek, layout, menu: items, menuOpen } = controller.getSnapshot();
		ribbon.dataset.ribbonDisplay = effectiveMode;
		ribbon.toggleAttribute('data-ribbon-peek', peek);
		// The title bar is a sibling of the ribbon, so the viewer root carries the mode too.
		const root = ribbon.closest<HTMLElement>('.pptxv');
		if (root) {
			root.dataset.ribbonDisplay = effectiveMode;
		}
		if (effectiveMode === 'fullScreen' && ribbon.parentNode) {
			if (reveal.nextSibling !== ribbon) {
				ribbon.before(reveal);
			}
		} else {
			reveal.remove();
		}
		if (layout.commands === 'overlay' && effectiveMode === 'tabs') {
			// The peeked commands hang below the tab row, over the slide.
			ribbon.style.setProperty(
				'--pptxv-ribbon-overlay-top',
				`${ribbon.offsetTop + ribbon.offsetHeight}px`,
			);
		}
		menu.hidden = !menuOpen;
		button.setAttribute('aria-expanded', String(menuOpen));
		button.classList.toggle('is-active', menuOpen);
		menu.replaceChildren();
		for (const item of items) {
			const entry = createEl(doc, 'button', 'pptxv-primary-menu-item');
			entry.type = 'button';
			entry.setAttribute('role', 'menuitemradio');
			entry.setAttribute('aria-checked', String(item.checked));
			entry.dataset.ribbonDisplay = item.mode;
			entry.textContent = t(item.labelKey);
			entry.addEventListener('click', () => controller.choose(item.mode));
			menu.appendChild(entry);
		}
		deps.onLayout();
	};
	const unsubscribe = controller.subscribe(sync);

	button.addEventListener('click', () => controller.toggleMenu());
	reveal.addEventListener('click', () => controller.reveal());
	tabBar.addEventListener('dblclick', (event) => {
		const tab = (event.target as Element).closest?.('[role="tab"]');
		if (tab && !tab.classList.contains('pptxv-ribbon-tab-file')) {
			controller.toggleCollapsed();
		}
	});

	// The shared pointer registry drops a handler only when its owner is collected, so a ribbon
	// replaced by a chrome rebuild would keep dismissing the new ribbon's menu until then.
	let destroyed = false;
	onDocumentPointerDown(doc, ribbon, (event) => {
		if (destroyed) {
			return;
		}
		const path = event.composedPath();
		if (!path.includes(ribbon) && !path.includes(reveal)) {
			controller.dismiss();
		} else if (!path.includes(menu) && !path.includes(button)) {
			controller.setMenuOpen(false);
		}
	});

	const onKey = (event: KeyboardEvent): void => {
		const root = ribbon.closest('.pptxv');
		const target = event.target as Node | null;
		// Only this viewer's shortcuts: the focus is inside it, or nowhere in particular.
		if (root && (target === doc.body || (target !== null && root.contains(target)))) {
			controller.handleKeydown(event);
		}
	};
	doc.addEventListener('keydown', onKey);

	// Options > General > "Collapse the ribbon automatically" applies on a short viewport.
	const short = doc.defaultView?.matchMedia?.(RIBBON_AUTO_COLLAPSE_MEDIA);
	const onShort = (): void => controller.setShortViewport(short?.matches ?? false);
	short?.addEventListener('change', onShort);
	onShort();

	return {
		button,
		menuAnchor,
		sync,
		destroy() {
			destroyed = true;
			unsubscribe();
			doc.removeEventListener('keydown', onKey);
			short?.removeEventListener('change', onShort);
			reveal.remove();
		},
	};
}
