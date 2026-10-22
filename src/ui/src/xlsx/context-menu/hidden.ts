// Removes the context-menu entries whose action the host hid (`hiddenActions`), so a hidden action
// cannot be run from right click any more than from the ribbon.
import type { MenuEntry } from 'ooxml-core/xlsx/ui';

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
 * The entries the host lets the user see. A separator is kept only before an entry that follows a
 * visible one, so hiding whole groups never leaves a separator at the top or next to another one.
 */
export function visibleMenuEntries(
	entries: readonly MenuEntry[],
	isHidden: (command: string) => boolean,
): MenuEntry[] {
	const kept: MenuEntry[] = [];
	// A hidden entry's separator moves to the next entry that is kept, so two groups never merge.
	let carry = false;
	for (const entry of entries) {
		const command = menuEntryCommand(entry);
		if (command !== undefined && isHidden(command)) {
			carry ||= Boolean(entry.separatorBefore);
			continue;
		}
		const separatorBefore = kept.length > 0 && Boolean(entry.separatorBefore || carry);
		carry = false;
		kept.push(
			separatorBefore === Boolean(entry.separatorBefore) ? entry : { ...entry, separatorBefore },
		);
	}
	return kept.map((entry, index) =>
		index === 0 && entry.separatorBefore ? { ...entry, separatorBefore: false } : entry,
	);
}
