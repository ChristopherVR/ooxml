// The context menu: the shared `office-ui-context-menu` element in controlled mode, appended to the
// shadow root. The element owns the DOM, keyboard navigation, type-ahead, viewport clamping and
// dismissal; this module translates the entries and runs the chosen one (items.ts decides what is
// offered).
import { defineContextMenu } from '../../controls';
import type { OfficeMenuItem, OfficeMenuState } from '../../controls';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import type { MenuEntry } from 'ooxml-core/xlsx/ui';
import { visibleMenuEntries } from './hidden';
import type { ExcelMenuCommand, ExcelMenuContext } from 'ooxml-core/xlsx/ui';
import { withHostMenuItems } from '../../embedding/host-menu';
import { nativeMenuChange, offerToHost, type MenuInputs } from './host-api';
import { withoutHiddenEntries } from 'ooxml-core/embedding';

export interface ContextMenuHandle {
	close(): void;
	readonly element: HTMLElement;
}

export interface OpenMenuOptions {
	onClose?: () => void;
	/** Called after an item ran or the menu was dismissed with Escape/Tab. */
	restoreFocus?: () => void;
	/** Host commands (`contextMenu.extraCommands`) and the context they run with. */
	host?: { commands: readonly ExcelMenuCommand[]; context: ExcelMenuContext };
	/**
	 * Lets the host replace this menu (`surfaceActions.onContextMenu`): which menu it is, and how to
	 * rebuild its entries when a choice runs.
	 */
	offer?: { input: MenuInputs; build: () => MenuEntry[] };
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
	all: readonly MenuEntry[],
	x: number,
	y: number,
	options: OpenMenuOptions = {},
): ContextMenuHandle {
	openMenus.get(ctx)?.close();
	if (options.offer && offerToHost(ctx, options.offer.input, options.offer.build, x, y)) {
		options.restoreFocus?.();
		return { close: () => {}, element: ctx.host.ownerDocument.createElement('div') };
	}
	// Entries whose action the host hid are not offered, so they cannot run from the menu either.
	const change = options.offer
		? nativeMenuChange(ctx, options.offer.input, options.offer.build, x, y)
		: undefined;
	const shown = visibleMenuEntries(all, (id) => ctx.commands.isHidden(id));
	const entries = change?.hidden.size
		? withoutHiddenEntries(shown, (entry) => change.hidden.has(entry.id))
		: shown;
	const hostCommands = [...(options.host?.commands ?? []), ...(change?.commands ?? [])];
	if (entries.length === 0 && hostCommands.length === 0) {
		options.restoreFocus?.();
		return { close: () => {}, element: ctx.host.ownerDocument.createElement('div') };
	}
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
	const extra = hostCommands.length
		? withHostMenuItems(rows, hostCommands, options.host?.context as ExcelMenuContext)
		: { items: rows, run: () => false };
	element.state = { x, y, label: ctx.t('Context menu'), items: extra.items };
	element.addEventListener('office-menu-request', (event) => {
		const id = (event as CustomEvent<{ id: string }>).detail.id;
		if (id.startsWith('host:')) {
			close(true);
			extra.run(id);
			return;
		}
		const entry = entries.find((candidate) => candidate.id === id);
		// Read-only and command state are checked again when the item runs, not when it was drawn.
		if (!entry || entryDisabled(ctx, entry)) return;
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
