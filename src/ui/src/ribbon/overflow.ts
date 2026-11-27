import { composedWithin } from '../composed';
import { modalActiveElement } from '../dialog/focus';
import { ensureRibbonTouchStyle, syncScrollMore, type ScrollMoreLabel } from './scroll-more';

/**
 * Ribbon overflow: when a panel is narrower than its groups, collapse groups in priority order
 * into dropdown buttons that show each group's live controls in a panel under it. The panel
 * scrolls only when even that is not enough, as Word and Excel do when a window narrows.
 *
 * A product supplies how its icons are drawn and which command buttons end the dropdown; the
 * folding itself is the same for every product. Panels are `.ribbon-panel` elements holding
 * `.ribbon-group` children, each with `data-label` and optionally `data-caption`.
 */
export interface RibbonOverflowOptions {
	/**
	 * Accessible name of the pinned "more" chevron shown while folded groups still overflow: a
	 * string, or a function of the panel that is re-read on every sync (so a locale change applies).
	 * Default `'More'`.
	 */
	moreLabel?: ScrollMoreLabel;
	/** Draws the icon `name` at `size` pixels (the product's own glyph set). */
	icon(doc: Document, name: string, size: number): Node;
	/** Icon of a folded group's button, by the group's `data-label`. */
	groupIcons: Readonly<Record<string, string>>;
	/** Icon for a group missing from `groupIcons`. */
	fallbackIcon: string;
	/** Selector of the ribbon element that owns the dropdown panel. */
	ribbon: string;
	/** Selector of the command buttons whose click ends the dropdown (selects keep it open). */
	command: string;
	/** Writes the caption under a folded group's button (default: its text content). */
	setCaption?(span: HTMLElement, label: string): void;
	/** Give the dropdown the group's natural width (for groups whose layout depends on it). */
	naturalWidth?: boolean;
	/** Collapse lower priorities first. Equal priorities collapse from right to left. */
	groupPriority?(group: HTMLElement): number;
	/** Reserve room for painted (`::after`) group captions so they never run into the next group. */
	reserveCaptions?: boolean;
	/** Give these groups the last chance to stay expanded when space runs out. */
	preferExpanded?(group: HTMLElement, panel: HTMLElement): boolean;
}

export interface RibbonOverflow {
	/** Folds groups of `panel` until it fits its width. */
	fitPanel(panel: HTMLElement): void;
	/** Keeps each panel of `root` fitted as it resizes. Returns a disposer. */
	attach(root: HTMLElement): () => void;
	/** Re-measures every panel of `root` (after a tab switch or a locale change). */
	refit(root: HTMLElement): void;
	/** Closes the open dropdown, putting its controls back; with `scope`, only one inside it. */
	close(scope?: Node): void;
}

export function createRibbonOverflow(options: RibbonOverflowOptions): RibbonOverflow {
	let closeOpen: (() => void) | undefined;
	let open: HTMLElement | undefined;

	function overflowButton(group: HTMLElement): HTMLButtonElement {
		const doc = group.ownerDocument;
		const button = doc.createElement('button');
		button.type = 'button';
		button.className = 'ribbon-overflow-button ribbon-large';
		button.setAttribute('aria-haspopup', 'true');
		const name = options.groupIcons[group.dataset.label ?? ''] ?? options.fallbackIcon;
		button.append(
			options.icon(doc, name, 28),
			doc.createElement('span'),
			options.icon(doc, 'caret', 12),
		);
		button.addEventListener('mousedown', (event) => event.preventDefault());
		button.addEventListener('click', () => toggleGroup(group, button));
		return button;
	}

	/** Shows a collapsed group's controls in a panel under its button, moving the live elements. */
	function toggleGroup(group: HTMLElement, button: HTMLButtonElement): void {
		const wasOpen = button.getAttribute('aria-expanded') === 'true';
		closeOpen?.();
		if (wasOpen) return;
		const ribbon = group.closest<HTMLElement>(options.ribbon);
		if (!ribbon) return;
		const doc = group.ownerDocument;
		const pop = doc.createElement('div');
		pop.className = 'ribbon-popover ribbon-overflow-panel';
		const content = doc.createElement('div');
		content.className = `${group.className} ribbon-overflow-content`;
		content.dataset.label = group.dataset.label ?? '';
		content.dataset.caption = group.dataset.caption ?? '';
		content.setAttribute('role', 'group');
		content.setAttribute('aria-label', group.getAttribute('aria-label') ?? '');
		// A floor, not a width: the pop-out may style its controls wider than the ribbon did.
		if (options.naturalWidth) content.style.minWidth = `${group.dataset.natural ?? 200}px`;
		const moved = [...group.children].filter((child) => child !== button);
		content.append(...moved);
		pop.append(content);
		ribbon.append(pop);
		const box = button.getBoundingClientRect();
		pop.style.top = `${box.bottom + 2}px`;
		const width = doc.defaultView?.innerWidth ?? 1024;
		pop.style.left = `${Math.max(4, Math.min(box.left, width - pop.offsetWidth - 8))}px`;
		syncScrollMore(pop, options.moreLabel);
		button.setAttribute('aria-expanded', 'true');
		open = pop;
		const outside = (event: Event) => {
			const path = event.composedPath();
			const inside = path.some(
				(node) => node instanceof HTMLElement && node.classList.contains('ribbon-popover'),
			);
			if (!inside && !path.includes(button)) closeOpen?.();
		};
		const escape = (event: KeyboardEvent) => {
			if (event.key !== 'Escape') return;
			if (
				event
					.composedPath()
					.some(
						(node) => node !== pop && node instanceof Element && node.matches('.ribbon-popover'),
					)
			)
				return;
			event.preventDefault();
			event.stopPropagation();
			closeOpen?.();
			button.focus();
		};
		// A command ends the dropdown; selects, combo boxes and buttons that open their own popup
		// (a gallery or a menu, which would vanish with the dropdown) keep it open.
		pop.addEventListener('click', (event) => {
			const command = (event.target as Element).closest(options.command);
			const popup = command?.getAttribute('aria-haspopup');
			if (command && (!popup || popup === 'false')) setTimeout(() => closeOpen?.(), 0);
		});
		doc.addEventListener('pointerdown', outside, true);
		doc.addEventListener('keydown', escape, true);
		closeOpen = () => {
			doc.removeEventListener('pointerdown', outside, true);
			doc.removeEventListener('keydown', escape, true);
			const focused = modalActiveElement(doc);
			const usedCommand = composedWithin(focused ?? undefined, pop);
			group.prepend(...moved);
			pop.remove();
			if (usedCommand) button.focus({ preventScroll: true });
			button.removeAttribute('aria-expanded');
			closeOpen = open = undefined;
		};
	}

	/** Group captions painted by `::after` do not size the group: give each one room for its own. */
	function reserveCaptions(panel: HTMLElement): void {
		const sizer = panel.ownerDocument.createElement('span');
		sizer.className = 'ribbon-caption-sizer';
		panel.append(sizer);
		for (const group of panel.querySelectorAll<HTMLElement>(':scope > .ribbon-group')) {
			sizer.textContent = group.dataset.caption ?? '';
			const launcher = group.querySelector(':scope > .ribbon-launcher') ? 32 : 12;
			group.style.minWidth = `${Math.ceil(sizer.offsetWidth) + launcher}px`;
		}
		sizer.remove();
	}

	function fitPanel(panel: HTMLElement): void {
		if (panel.hidden || !panel.clientWidth) return;
		ensureRibbonTouchStyle(panel);
		panel.querySelector(':scope > .ribbon-scroll-more')?.remove();
		// Re-measuring while a folded group's panel is open would tear it down under the user's hands.
		if (closeOpen && panel.dataset.fittedWidth === String(panel.clientWidth)) return;
		panel.dataset.fittedWidth = String(panel.clientWidth);
		closeOpen?.();
		if (options.reserveCaptions) reserveCaptions(panel);
		for (const group of panel.querySelectorAll<HTMLElement>('.ribbon-group[data-collapsed]'))
			group.removeAttribute('data-collapsed');
		const groups = [...panel.querySelectorAll<HTMLElement>(':scope > .ribbon-group')].filter(
			(group) => group.offsetWidth > 0,
		);
		if (options.naturalWidth)
			for (const group of groups) group.dataset.natural = String(Math.ceil(group.offsetWidth));
		const foldOrder = [...groups].reverse();
		const priority = options.groupPriority;
		const preferExpanded = options.preferExpanded;
		// Preferred groups fold last; otherwise lower priorities fold first, right to left.
		const rank = (group: HTMLElement) =>
			(preferExpanded?.(group, panel) ? 1e6 : 0) + (priority?.(group) ?? 0);
		if (priority || preferExpanded) foldOrder.sort((a, b) => rank(a) - rank(b));
		for (const group of foldOrder) {
			if (panel.scrollWidth <= panel.clientWidth) break;
			const width = group.offsetWidth;
			let button = group.querySelector<HTMLButtonElement>(':scope > .ribbon-overflow-button');
			if (!button) {
				button = overflowButton(group);
				group.append(button);
			}
			const label = group.dataset.caption ?? group.dataset.label ?? '';
			const span = button.querySelector('span')!;
			if (options.setCaption) options.setCaption(span, label);
			else span.textContent = label;
			button.setAttribute('aria-label', label);
			button.title = label;
			group.setAttribute('data-collapsed', '');
			// Folding a group that is already as narrow as its button saves nothing.
			if (group.offsetWidth >= width) group.removeAttribute('data-collapsed');
		}
		// Restore groups from the left, even when a product prefers to fold them in another order.
		if (panel.scrollWidth <= panel.clientWidth) {
			for (const group of groups) {
				if (!group.hasAttribute('data-collapsed')) continue;
				group.removeAttribute('data-collapsed');
				if (panel.scrollWidth > panel.clientWidth) group.setAttribute('data-collapsed', '');
			}
		}
		syncScrollMore(panel, options.moreLabel);
	}

	const panelsOf = (root: HTMLElement) => root.querySelectorAll<HTMLElement>('.ribbon-panel');

	return {
		fitPanel,
		attach(root) {
			if (typeof ResizeObserver === 'undefined') return () => {};
			const observer = new ResizeObserver((entries) => {
				for (const entry of entries) fitPanel(entry.target as HTMLElement);
			});
			for (const panel of panelsOf(root)) observer.observe(panel);
			// Content that grows after the first fit (a gallery filling in, a locale change) refits.
			let frame = 0;
			const mutations = new MutationObserver(() => {
				if (frame) return;
				frame = requestAnimationFrame(() => {
					frame = 0;
					refitVisible();
					mutations.takeRecords();
				});
			});
			const refitVisible = () => {
				for (const panel of panelsOf(root)) {
					if (panel.hidden || panel.querySelector('[aria-expanded="true"].ribbon-overflow-button'))
						continue;
					const more = panel.querySelector<HTMLElement>(':scope > .ribbon-scroll-more');
					const clipped = panel.scrollWidth - (more?.offsetWidth ?? 0) > panel.clientWidth;
					const unfolded = panel.querySelector(':scope > .ribbon-group:not([data-collapsed])');
					if (clipped && unfolded) fitPanel(panel);
				}
			};
			for (const panel of panelsOf(root))
				mutations.observe(panel, { childList: true, subtree: true });
			return () => {
				observer.disconnect();
				mutations.disconnect();
				if (frame) cancelAnimationFrame(frame);
			};
		},
		refit(root) {
			for (const panel of panelsOf(root)) fitPanel(panel);
		},
		close: (scope) => {
			if (!scope || composedWithin(open, scope)) closeOpen?.();
		},
	};
}
