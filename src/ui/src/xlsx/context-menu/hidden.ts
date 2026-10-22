// Removes the context-menu entries whose action the host hid (`hiddenActions`), so a hidden action
// cannot be run from right click any more than from the ribbon.
import type { MenuEntry } from 'ooxml-core/xlsx/ui';
import { withoutHiddenEntries } from 'ooxml-core/embedding';

/**
 * Entries that run an action without a command id, mapped to the ribbon command they stand for.
 * Tab Rename runs an inline rename, but the ribbon's Rename Sheet control is `sheet.rename`.
 */
const ACTION_COMMANDS: Readonly<Record<string, string>> = { rename: 'sheet.rename' };

/** The command id a menu entry corresponds to, when it has one (the id `hiddenActions` uses). */
export function menuEntryCommand(entry: MenuEntry): string | undefined {
	if (entry.command) return entry.command;
	return entry.action ? ACTION_COMMANDS[entry.id] : undefined;
}

/**
 * The entries the host lets the user see, by the shared rule of `withoutHiddenEntries`: a hidden
 * entry's separator moves to the next kept entry, and none is left at the top, so hiding entries
 * never merges two groups or leaves a separator next to another one.
 */
export function visibleMenuEntries(
	entries: readonly MenuEntry[],
	isHidden: (command: string) => boolean,
): MenuEntry[] {
	return withoutHiddenEntries(entries, (entry) => {
		const command = menuEntryCommand(entry);
		return command !== undefined && isHidden(command);
	});
}
