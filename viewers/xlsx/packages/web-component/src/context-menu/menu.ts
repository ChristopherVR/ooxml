// The context menu DOM: an accessible `role="menu"` list appended to the shadow root. Owns keyboard
// navigation, hover focus, viewport clamping and dismissal; what it offers comes from items.ts.
import type { EditorContext } from '../context.js';
import { ensureStyle, h } from '../grid/dom.js';
import { clampToViewport, type MenuEntry } from './items.js';
import { MENU_CSS } from './styles.js';

export interface ContextMenuHandle {
	close(): void;
	readonly element: HTMLElement;
}

export interface OpenMenuOptions {
	onClose?: () => void;
	/** Called after an item ran or the menu was dismissed with Escape/Tab. */
	restoreFocus?: () => void;
}

const openMenus = new WeakMap<EditorContext, ContextMenuHandle>();

/** The menu currently open for a context, if any. */
export const currentContextMenu = (ctx: EditorContext): ContextMenuHandle | undefined =>
	openMenus.get(ctx);

function entryDisabled(ctx: EditorContext, entry: MenuEntry): boolean {
	if (entry.disabled) return true;
	if (entry.command)
		return !ctx.commands.get(entry.command) || !ctx.commands.isEnabled(entry.command);
	return !entry.action;
}

export function openContextMenu(
	ctx: EditorContext,
	entries: MenuEntry[],
	x: number,
	y: number,
	options: OpenMenuOptions = {},
): ContextMenuHandle {
	openMenus.get(ctx)?.close();
	ensureStyle(ctx.root, 'xcm', MENU_CSS);
	const doc = ctx.host.ownerDocument;
	const win = doc.defaultView;
	const element = h(doc, 'div', 'xcm-menu', { role: 'menu', 'aria-label': ctx.t('Context menu') });
	let closed = false;
	let disposeDismiss = (): void => {};

	const items = () => [...element.querySelectorAll<HTMLElement>('.xcm-item:not([aria-disabled])')];
	const close = (restore = false) => {
		if (closed) return;
		closed = true;
		element.remove();
		disposeDismiss();
		if (openMenus.get(ctx) === handle) openMenus.delete(ctx);
		options.onClose?.();
		if (restore) options.restoreFocus?.();
	};
	const handle: ContextMenuHandle = { close: () => close(false), element };

	const move = (step: number | 'first' | 'last') => {
		const list = items();
		if (!list.length) return;
		const active = ctx.root.activeElement as HTMLElement | null;
		const current = active ? list.indexOf(active) : -1;
		const last = list.length - 1;
		const next =
			step === 'first' || (current < 0 && step === 1)
				? 0
				: step === 'last' || current < 0
					? last
					: (current + step + list.length) % list.length;
		list[next]?.focus();
	};

	const run = (entry: MenuEntry) => {
		close(true);
		if (entry.command) void ctx.commands.run(entry.command, entry.arg);
		else void entry.action?.();
	};

	for (const entry of entries) {
		if (entry.separatorBefore)
			element.append(h(doc, 'div', 'xcm-separator', { role: 'separator' }));
		const button = h(doc, 'button', 'xcm-item', {
			type: 'button',
			role: entry.checked === undefined ? 'menuitem' : 'menuitemcheckbox',
			tabindex: '-1',
		});
		if (entry.checked !== undefined) button.setAttribute('aria-checked', String(entry.checked));
		button.dataset.item = entry.id;
		const check = h(doc, 'span', 'xcm-check', { 'aria-hidden': 'true' });
		check.textContent = entry.checked ? '✓' : '';
		const label = h(doc, 'span', 'xcm-label');
		label.textContent = ctx.t(entry.label);
		button.append(check, label);
		if (entry.shortcut) {
			const shortcut = h(doc, 'span', 'xcm-shortcut');
			shortcut.textContent = entry.shortcut;
			button.append(shortcut);
		}
		if (entryDisabled(ctx, entry)) button.setAttribute('aria-disabled', 'true');
		button.addEventListener('click', () => {
			if (!button.hasAttribute('aria-disabled')) run(entry);
		});
		element.append(button);
	}

	element.addEventListener('keydown', (event) => {
		const handled = (action: () => void) => {
			event.preventDefault();
			event.stopPropagation();
			action();
		};
		if (event.key === 'ArrowDown') handled(() => move(1));
		else if (event.key === 'ArrowUp') handled(() => move(-1));
		else if (event.key === 'Home') handled(() => move('first'));
		else if (event.key === 'End') handled(() => move('last'));
		else if (event.key === 'Escape' || event.key === 'Tab') handled(() => close(true));
		else if (event.key === 'Enter' || event.key === ' ')
			handled(() => (event.target as HTMLElement).click());
	});
	element.addEventListener('mouseover', (event) => {
		const target = (event.target as Element).closest<HTMLElement>('.xcm-item');
		if (target && !target.hasAttribute('aria-disabled')) target.focus();
	});
	element.addEventListener('contextmenu', (event) => event.preventDefault());

	element.style.visibility = 'hidden';
	ctx.root.append(element);
	const box = element.getBoundingClientRect();
	const { left, top } = clampToViewport(
		x,
		y,
		{ width: box.width, height: box.height },
		{ width: win?.innerWidth ?? 1024, height: win?.innerHeight ?? 768 },
	);
	element.style.left = `${left}px`;
	element.style.top = `${top}px`;
	element.style.visibility = '';

	const onPointerDown = (event: Event) => {
		if (!event.composedPath().includes(element)) close(false);
	};
	const dismiss = () => close(false);
	const onScroll = (event: Event) => {
		if (!event.composedPath().includes(element)) close(false);
	};
	// Registered after the opening event has finished dispatching.
	const timer = setTimeout(() => {
		if (closed) return;
		doc.addEventListener('pointerdown', onPointerDown, true);
		doc.addEventListener('scroll', onScroll, true);
		ctx.root.addEventListener('scroll', onScroll, true);
		win?.addEventListener('resize', dismiss);
		win?.addEventListener('blur', dismiss);
	}, 0);
	disposeDismiss = () => {
		clearTimeout(timer);
		doc.removeEventListener('pointerdown', onPointerDown, true);
		doc.removeEventListener('scroll', onScroll, true);
		ctx.root.removeEventListener('scroll', onScroll, true);
		win?.removeEventListener('resize', dismiss);
		win?.removeEventListener('blur', dismiss);
	};
	openMenus.set(ctx, handle);
	move('first');
	return handle;
}
