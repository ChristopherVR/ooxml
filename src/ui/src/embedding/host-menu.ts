/**
 * Host context-menu commands (`contextMenu.extraCommands`) in the shared `OfficeMenuItem` model
 * of `<office-ui-context-menu>`: the host groups go above and below the built-in rows, and a
 * chosen host row runs its command. Word, Excel and Visio share this; PowerPoint keeps its own
 * menu entries on the same core helper (`withHostMenuCommands`).
 */
import {
	isHostMenuEntry,
	withHostMenuCommands,
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
