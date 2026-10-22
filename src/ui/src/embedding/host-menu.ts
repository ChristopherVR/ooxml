/**
 * Host context-menu commands (`contextMenu.extraCommands`) in the shared `OfficeMenuItem` model
 * of `<office-ui-context-menu>`: the host groups go above and below the built-in rows, and a
 * chosen host row runs its command. Word, Excel and Visio share this; PowerPoint keeps its own
 * menu entries on the same core helper (`withHostMenuCommands`).
 */
import {
	isHostMenuEntry,
	resolveHostMenuCommands,
	withHostMenuCommands,
	withoutHiddenEntries,
	type HostMenuCommand,
	type HostMenuEntry,
} from 'ooxml-core/embedding';
import type { OfficeMenuItem } from '../menu/menu-model';

export interface HostMenuItems {
	/** The rows to draw: host groups around the built-in ones. */
	items: OfficeMenuItem[];
	/**
	 * Runs host row `id` and returns true; false (doing nothing) for a built-in row, so the
	 * product handles it. A disabled host row is claimed but never runs.
	 */
	run(id: string): boolean;
}

/** The menu rows with the host commands for one opening of the menu, at `context`. */
export function withHostMenuItems<C>(
	items: readonly OfficeMenuItem[],
	commands: readonly HostMenuCommand<C>[] | null | undefined,
	context: C,
): HostMenuItems {
	const rows = withHostMenuCommands(items, commands, context);
	const host = new Map<string, HostMenuEntry>();
	const drawn = rows.map((row): OfficeMenuItem => {
		if (!isHostMenuEntry(row)) return row;
		host.set(row.id, row);
		return {
			id: row.id,
			label: row.label,
			...(row.separatorBefore ? { separatorBefore: true } : {}),
			...(row.disabled ? { disabled: true } : {}),
			...(row.icon ? { icon: row.icon } : {}),
		};
	});
	return {
		items: drawn,
		run(id) {
			const entry = host.get(id);
			if (!entry) return false;
			entry.onSelect();
			return true;
		},
	};
}

/** Marks the children `syncHostMenuChildren` added, so the next call replaces them. */
const HOST_CHILD = 'data-host-command';
const hostRuns = new WeakMap<HTMLElement, Map<string, HostMenuEntry>>();

/** The host rows of `menu`, run by id when the menu reports a chosen row. */
function listen(menu: HTMLElement): Map<string, HostMenuEntry> {
	const runs = new Map<string, HostMenuEntry>();
	hostRuns.set(menu, runs);
	menu.addEventListener('office-command', (event) => {
		const id = (event as CustomEvent<{ command?: string }>).detail?.command;
		if (id) runs.get(id)?.onSelect();
	});
	return runs;
}

/**
 * Host commands in a self-managed `<office-ui-context-menu>` (one whose rows are
 * `office-ui-menu-item` children): call it just before the menu opens. It replaces the host rows
 * of the previous opening, puts the `top` group (and a separator) first and the `bottom` group
 * last, and runs a chosen host row's command with this opening's `context`.
 */
export function syncHostMenuChildren<C>(
	menu: HTMLElement,
	commands: readonly HostMenuCommand<C>[] | null | undefined,
	context: C,
): void {
	for (const child of [...menu.querySelectorAll(`:scope > [${HOST_CHILD}]`)]) child.remove();
	const doc = menu.ownerDocument;
	const { top, bottom } = resolveHostMenuCommands(commands, context);
	const runs = hostRuns.get(menu) ?? listen(menu);
	runs.clear();
	const item = (entry: HostMenuEntry): HTMLElement => {
		runs.set(entry.id, entry);
		const row = doc.createElement('office-ui-menu-item');
		row.setAttribute(HOST_CHILD, '');
		row.setAttribute('command', entry.id);
		row.setAttribute('label', entry.label);
		if (entry.icon) row.setAttribute('icon', entry.icon);
		if (entry.disabled) row.setAttribute('disabled', '');
		return row;
	};
	const separator = (): HTMLElement => {
		const rule = doc.createElement('office-ui-menu-separator');
		rule.setAttribute(HOST_CHILD, '');
		return rule;
	};
	const builtIn = menu.children.length > 0;
	if (top.length) menu.prepend(...top.map(item), ...(builtIn ? [separator()] : []));
	if (bottom.length)
		menu.append(...(builtIn || top.length ? [separator()] : []), ...bottom.map(item));
}

/** Marks a separator `tidyMenuSeparators` hid, so it never shows one the product hid itself. */
const TIDY_HIDDEN = 'data-separator-tidy';

/**
 * Hides the separators of a self-managed `<office-ui-context-menu>` that its hidden rows left
 * dangling, by the rule of `withoutHiddenEntries`: each kept row that starts a group shows one
 * separator before it, and none shows at the top, at the bottom or twice in a row. Call it after
 * rows were hidden or shown (and after `syncHostMenuChildren`).
 */
export function tidyMenuSeparators(menu: HTMLElement): void {
	const isSeparator = (node: Element) => node.localName === 'office-ui-menu-separator';
	const separators: HTMLElement[] = [];
	const rows: Array<{ row: HTMLElement; before: HTMLElement[]; separatorBefore: boolean }> = [];
	let pending: HTMLElement[] = [];
	for (const child of menu.children) {
		if (!(child instanceof HTMLElement)) continue;
		if (isSeparator(child)) {
			if (child.hidden && !child.hasAttribute(TIDY_HIDDEN)) continue;
			separators.push(child);
			pending.push(child);
			continue;
		}
		rows.push({ row: child, before: pending, separatorBefore: pending.length > 0 });
		pending = [];
	}
	// Which kept rows start a group, by the shared rule; each shows the last separator before it.
	const starts = new Map(
		withoutHiddenEntries(rows, (entry) => entry.row.hidden !== false).map((entry) => [
			entry.row,
			entry.separatorBefore,
		]),
	);
	const shown = new Set<HTMLElement>();
	let carried: HTMLElement[] = [];
	for (const entry of rows) {
		carried = [...carried, ...entry.before];
		if (!starts.has(entry.row)) continue;
		const separator = carried.at(-1);
		if (starts.get(entry.row) && separator) shown.add(separator);
		carried = [];
	}
	for (const separator of separators) {
		const hide = !shown.has(separator);
		separator.hidden = hide;
		separator.toggleAttribute(TIDY_HIDDEN, hide);
	}
}
