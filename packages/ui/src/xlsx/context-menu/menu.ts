// The context menu: the shared `office-ui-context-menu` element in controlled mode, appended to the
// shadow root. The element owns the DOM, keyboard navigation, type-ahead, viewport clamping and
// dismissal; this module translates the entries and runs the chosen one (items.ts decides what is
// offered).
import { defineContextMenu } from '../../controls';
import type { OfficeMenuItem, OfficeMenuState } from '../../controls';
import type { EditorContext } from '../context.js';
import type { MenuEntry } from './items.js';

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
	defineContextMenu();
	const doc = ctx.host.ownerDocument;
	const win = doc.defaultView;
	const element = doc.createElement('office-ui-context-menu') as HTMLElement & {
		state: OfficeMenuState;
	};
	let closed = false;
	let disposeDismiss = (): void => {};

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

	const rows: OfficeMenuItem[] = entries.map((entry) => ({
		id: entry.id,
		label: ctx.t(entry.label),
		disabled: entryDisabled(ctx, entry),
		...(entry.separatorBefore ? { separatorBefore: true } : {}),
		...(entry.checked === undefined ? {} : { checked: entry.checked }),
		...(entry.shortcut ? { shortcut: entry.shortcut } : {}),
	}));
	element.state = { x, y, label: ctx.t('Context menu'), items: rows };
	element.addEventListener('office-menu-request', (event) => {
		const id = (event as CustomEvent<{ id: string }>).detail.id;
		const entry = entries.find((candidate) => candidate.id === id);
		if (!entry) return;
		close(true);
		if (entry.command) void ctx.commands.run(entry.command, entry.arg);
		else void entry.action?.();
	});
	// Escape and Tab return focus to the grid; an outside click leaves it where it landed.
	element.addEventListener('office-menu-close', (event) =>
		close((event as CustomEvent<{ reason: string }>).detail.reason !== 'outside'),
	);
	ctx.root.append(element);

	const dismiss = () => close(false);
	const onScroll = (event: Event) => {
		if (!event.composedPath().includes(element)) close(false);
	};
	doc.addEventListener('scroll', onScroll, true);
	ctx.root.addEventListener('scroll', onScroll, true);
	win?.addEventListener('resize', dismiss);
	win?.addEventListener('blur', dismiss);
	disposeDismiss = () => {
		doc.removeEventListener('scroll', onScroll, true);
		ctx.root.removeEventListener('scroll', onScroll, true);
		win?.removeEventListener('resize', dismiss);
		win?.removeEventListener('blur', dismiss);
	};
	openMenus.set(ctx, handle);
	return handle;
}
