// The host-facing side of the context menus: the entries of each menu as plain data with an
// execute operation, and the hook a host uses to replace the native menu.
import { rangeContains, validationAt } from 'ooxml-core/xlsx';
import type { EditorContext, MenuEntry } from 'ooxml-core/xlsx/ui';
import {
	activeRef,
	cellMenu,
	columnHeaderMenu,
	rowHeaderMenu,
	selectionRef,
	tabMenu,
} from 'ooxml-core/xlsx/ui';
import type {
	ContextMenuItem,
	ContextMenuKind,
	ContextMenuModel,
	ContextMenuRequest,
} from '../surface-workflows';
import { visibleMenuEntries } from './hidden';

type Handler = (request: ContextMenuRequest) => boolean | void;
/** Names of the host callbacks a surface-only editor delegates a menu command to. */
export type DelegatedAction = 'onOpenFormatCells';

/** What the editor tells its menus about the host (`surfaceActions`). */
export interface ContextMenuHost {
	/** The host's `onContextMenu`, if installed. */
	handler(): Handler | undefined;
	/** Whether the host callback is installed and the workbook can be edited. */
	available(name: DelegatedAction): boolean;
	/** Runs the callback with its request; false when unavailable. */
	delegate(name: DelegatedAction): boolean;
}
const hosts = new WeakMap<EditorContext, ContextMenuHost>();

/** Registers the host side of the menus of `ctx` (the editor core does this once). */
export function provideContextMenuHost(ctx: EditorContext, host: ContextMenuHost): void {
	hosts.set(ctx, host);
}

/** The commands a surface-only editor delegates to a host callback because it has no dialog. */
const HOST_DELEGATED: Readonly<Record<string, DelegatedAction>> = {
	'format.cells': 'onOpenFormatCells',
};

export interface MenuInputs {
	kind: ContextMenuKind;
	/** The sheet a tab menu is for; the active sheet otherwise. */
	sheet: number;
	/** Headers: rows or columns of the selection are hidden (the grid measures it when it can). */
	hiddenInSelection?: boolean;
	/** The tab menu's Rename, when a tab strip is mounted. */
	rename?: () => void;
}

function hiddenIn(ctx: EditorContext, axis: 'row' | 'col'): boolean {
	const ws = ctx.workbook()?.sheets[ctx.activeSheet()];
	if (!ws) return false;
	for (const r of ctx.selection.get().ranges) {
		const from = axis === 'row' ? r.start.row : r.start.col;
		const to = Math.min(axis === 'row' ? r.end.row : r.end.col, from + 2000);
		for (let i = from; i <= to; i++) {
			const hidden =
				axis === 'row'
					? ws.rowInfo.get(i)?.hidden
					: ws.columns.some((c) => i >= c.min && i <= c.max && c.hidden);
			if (hidden) return true;
		}
	}
	return false;
}

/** The entries of one menu for the current state (what the native menu would offer). */
export function menuEntries(ctx: EditorContext, input: MenuInputs): MenuEntry[] {
	const readOnly = ctx.readOnly();
	const workbook = ctx.workbook();
	if (input.kind === 'row' || input.kind === 'column') {
		const axis = input.kind === 'row' ? 'row' : 'col';
		const state = { readOnly, hiddenInSelection: input.hiddenInSelection ?? hiddenIn(ctx, axis) };
		return input.kind === 'row' ? rowHeaderMenu(state) : columnHeaderMenu(state);
	}
	if (input.kind === 'tab') {
		const sheets = workbook?.sheets ?? [];
		const visible = sheets.filter((s) => s.state === 'visible').length;
		const { rename } = input;
		const entries = tabMenu(
			{
				readOnly,
				structureLocked: workbook?.structureLocked === true,
				visibleSheets: visible,
				hiddenSheets: sheets.length - visible,
				protected: sheets[input.sheet]?.protection?.sheet === true,
			},
			{ rename: rename ?? (() => {}) },
		);
		// Rename is an inline edit of a tab strip: without one it cannot run (a host renames
		// through its own UI and the workbook API).
		return rename ? entries : entries.map((e) => (e.action ? { ...e, disabled: true } : e));
	}
	const sheet = workbook?.sheets[ctx.activeSheet()];
	const at = ctx.selection.get().active;
	const rule = workbook ? validationAt(workbook, ctx.activeSheet(), at.row, at.col) : undefined;
	return cellMenu({
		readOnly,
		hasComment: Boolean(
			sheet?.comments.some((c) => c.address.row === at.row && c.address.col === at.col),
		),
		hasLink: Boolean(sheet?.hyperlinks.some((l) => rangeContains(l.range, at))),
		hasList: rule?.type === 'list',
	});
}

/** Whether an entry can run now in this editor (a command that is not registered cannot). */
export function entryEnabled(ctx: EditorContext, entry: MenuEntry): boolean {
	if (entry.disabled) return false;
	if (entry.command) {
		const delegated = HOST_DELEGATED[entry.command];
		if (delegated && !ctx.commands.get(entry.command))
			return hosts.get(ctx)?.available(delegated) ?? false;
		return Boolean(ctx.commands.get(entry.command)) && ctx.commands.isEnabled(entry.command);
	}
	return Boolean(entry.action);
}

/**
 * The menu as data. `build` is called again when an item runs, so a choice is checked against the
 * state at that moment (read-only, protection, selection), not the state the menu was shown in.
 */
export function contextMenuModel(
	ctx: EditorContext,
	input: MenuInputs,
	build: () => MenuEntry[],
): ContextMenuModel {
	const current = (): MenuEntry[] =>
		visibleMenuEntries(build(), (command) => ctx.commands.isHidden(command));
	const selection = ctx.selection.get();
	const items: ContextMenuItem[] = current().map((entry) => ({
		id: entry.id,
		label: ctx.t(entry.label),
		...(entry.command ? { command: entry.command } : {}),
		enabled: entryEnabled(ctx, entry),
		...(entry.checked === undefined ? {} : { checked: entry.checked }),
		...(entry.shortcut ? { shortcut: entry.shortcut } : {}),
		...(entry.separatorBefore ? { separatorBefore: true } : {}),
	}));
	return {
		menu: input.kind,
		menuSheet: input.sheet,
		readOnly: ctx.readOnly(),
		sheet: ctx.activeSheet(),
		sheetName: ctx.workbook()?.sheets[ctx.activeSheet()]?.name ?? '',
		selection: selectionRef(selection),
		activeCell: activeRef(selection),
		items,
		async execute(id) {
			if (input.kind === 'tab' && input.sheet !== ctx.activeSheet())
				ctx.setActiveSheet(input.sheet);
			const grid = ctx.grid();
			if (grid?.isEditing() && !grid.commitEdit()) return false;
			const entry = current().find((candidate) => candidate.id === id);
			if (!entry || !entryEnabled(ctx, entry)) return false;
			const delegated = entry.command ? HOST_DELEGATED[entry.command] : undefined;
			if (delegated && entry.command && !ctx.commands.get(entry.command))
				return hosts.get(ctx)?.delegate(delegated) ?? false;
			if (entry.command) return ctx.commands.run(entry.command, entry.arg);
			await entry.action?.();
			return true;
		},
	};
}

/**
 * Offers a menu about to open natively to the host. True when the host took it over (the caller
 * must then not draw its own).
 */
export function offerToHost(
	ctx: EditorContext,
	input: MenuInputs,
	build: () => MenuEntry[],
	x: number,
	y: number,
): boolean {
	const handler = hosts.get(ctx)?.handler();
	if (!handler) return false;
	return handler({ ...contextMenuModel(ctx, input, build), x, y }) === true;
}
