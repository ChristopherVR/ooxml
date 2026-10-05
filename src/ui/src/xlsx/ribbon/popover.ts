/**
 * Ribbon popovers (menus, lists, the colour grid, galleries): one open at a time, fixed to the
 * viewport so the ribbon's overflow clipping never cuts them, dismissed by Escape, a press
 * elsewhere or a choice. Mirrors docx-viewer's ribbon-popover.ts.
 */
import { ribbonIcon } from './icons';

let closeOpen: (() => void) | undefined;

/** Closes whichever ribbon popover is open. */
export function closeRibbonPopover(): void {
	closeOpen?.();
}

export function isPopoverOpen(): boolean {
	return closeOpen !== undefined;
}

/**
 * Shows `pop` under `anchor` (appended to the anchor's shadow root). Returns false, closing it,
 * when the same anchor was already open, so a caret works as a toggle.
 */
export function mountPopover(
	anchor: HTMLElement,
	pop: HTMLElement,
	expanded: HTMLElement = anchor,
	onClose?: () => void,
): boolean {
	const wasOpen = expanded.getAttribute('aria-expanded') === 'true';
	closeOpen?.();
	if (wasOpen) return false;
	const doc = anchor.ownerDocument;
	const view = doc.defaultView;
	const box = anchor.getBoundingClientRect();
	pop.style.left = `${Math.max(4, box.left)}px`;
	pop.style.top = `${box.bottom + 2}px`;
	const root = anchor.getRootNode();
	(root instanceof ShadowRoot ? root : doc.body).append(pop);
	const width = view?.innerWidth ?? 1024;
	const limit = width - pop.offsetWidth - 8;
	if (pop.offsetWidth && box.left > limit) pop.style.left = `${Math.max(4, limit)}px`;
	const height = view?.innerHeight ?? 768;
	if (pop.offsetHeight && box.bottom + pop.offsetHeight > height - 8)
		pop.style.top = `${Math.max(4, height - pop.offsetHeight - 8)}px`;
	expanded.setAttribute('aria-expanded', 'true');
	const outside = (event: Event) => {
		const path = event.composedPath();
		if (!path.includes(pop) && !path.includes(anchor)) closeOpen?.();
	};
	const escape = (event: KeyboardEvent) => {
		if (event.key !== 'Escape') return;
		event.stopPropagation();
		closeOpen?.();
		anchor.focus();
	};
	doc.addEventListener('pointerdown', outside, true);
	doc.addEventListener('keydown', escape, true);
	closeOpen = () => {
		doc.removeEventListener('pointerdown', outside, true);
		doc.removeEventListener('keydown', escape, true);
		pop.remove();
		expanded.setAttribute('aria-expanded', 'false');
		closeOpen = undefined;
		onClose?.();
	};
	return true;
}

/** Up/Down (and Home/End) move focus through `items`, wrapping. */
export function arrowNavigation(container: HTMLElement, selector: string): void {
	container.addEventListener('keydown', (event) => {
		const items = [...container.querySelectorAll<HTMLElement>(selector)].filter(
			(item) => !(item as HTMLButtonElement).disabled,
		);
		if (!items.length) return;
		const at = items.indexOf(event.target as HTMLElement);
		const next =
			event.key === 'ArrowDown'
				? at + 1
				: event.key === 'ArrowUp'
					? at - 1
					: event.key === 'Home'
						? 0
						: event.key === 'End'
							? items.length - 1
							: Number.NaN;
		if (Number.isNaN(next)) return;
		event.preventDefault();
		items[(next + items.length) % items.length]?.focus();
	});
}

export type MenuEntry =
	| {
			label: string;
			icon?: string;
			shortcut?: string;
			disabled?: boolean;
			checked?: boolean;
			run(): void;
	  }
	| { separator: true };

/** A vertical command menu under `anchor`; the first enabled item takes focus. */
export function openMenu(
	anchor: HTMLElement,
	entries: MenuEntry[],
	label = '',
): HTMLElement | null {
	const doc = anchor.ownerDocument;
	const pop = doc.createElement('div');
	pop.className = 'ribbon-popover ribbon-menu-list';
	pop.setAttribute('role', 'menu');
	if (label) pop.setAttribute('aria-label', label);
	for (const entry of entries) {
		if ('separator' in entry) {
			const line = doc.createElement('div');
			line.setAttribute('role', 'separator');
			pop.append(line);
			continue;
		}
		const item = doc.createElement('button');
		item.type = 'button';
		item.setAttribute('role', entry.checked === undefined ? 'menuitem' : 'menuitemcheckbox');
		if (entry.checked !== undefined) item.setAttribute('aria-checked', String(entry.checked));
		item.disabled = entry.disabled ?? false;
		const icon = doc.createElement('span');
		icon.className = 'ribbon-menu-icon';
		if (entry.icon) icon.append(ribbonIcon(doc, entry.icon, 16));
		const text = doc.createElement('span');
		text.textContent = entry.label;
		item.append(icon, text);
		if (entry.shortcut) {
			const keys = doc.createElement('span');
			keys.className = 'ribbon-menu-shortcut';
			keys.textContent = entry.shortcut;
			item.append(keys);
		}
		item.addEventListener('mousedown', (event) => event.preventDefault());
		item.addEventListener('click', () => {
			closeRibbonPopover();
			entry.run();
		});
		pop.append(item);
	}
	arrowNavigation(pop, '[role^="menuitem"]');
	if (!mountPopover(anchor, pop)) return null;
	pop.querySelector<HTMLButtonElement>('[role^="menuitem"]:not(:disabled)')?.focus();
	return pop;
}

/** A scrolling option list under a combo box with `current` marked; `preview` sets each font. */
export function openList(
	anchor: HTMLElement,
	items: readonly { value: string; label: string }[],
	current: string,
	choose: (value: string) => void,
	options: { preview?: boolean; expanded?: HTMLElement } = {},
): void {
	const doc = anchor.ownerDocument;
	const pop = doc.createElement('div');
	pop.className = 'ribbon-popover ribbon-list';
	pop.setAttribute('role', 'listbox');
	for (const { value, label } of items) {
		const item = doc.createElement('button');
		item.type = 'button';
		item.setAttribute('role', 'option');
		item.textContent = label;
		item.setAttribute('aria-selected', String(value === current));
		if (options.preview) item.style.fontFamily = `"${value}", sans-serif`;
		item.addEventListener('mousedown', (event) => event.preventDefault());
		item.addEventListener('click', () => {
			closeRibbonPopover();
			choose(value);
		});
		pop.append(item);
	}
	arrowNavigation(pop, '[role="option"]');
	if (!mountPopover(anchor, pop, options.expanded ?? anchor)) return;
	const selected =
		pop.querySelector<HTMLElement>('[aria-selected="true"]') ??
		pop.querySelector<HTMLElement>('[role="option"]');
	selected?.scrollIntoView?.({ block: 'center' });
	selected?.focus({ preventScroll: true });
}
