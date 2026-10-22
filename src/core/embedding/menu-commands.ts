/**
 * Host context-menu commands: entries a host adds to an editor's right-click menu. Each command
 * is plain data plus callbacks, so one descriptor works in every editor and framework binding.
 * The context `C` is a product's immutable snapshot of what the menu was opened on (the
 * selection, the slide, the cell range); the product documents its shape.
 */

/** One host command. Labels are plain text the host has already localised. */
export interface HostMenuCommand<C> {
	/** Unique among the host commands of one menu; a repeated or empty id is ignored. */
	id: string;
	label: string;
	/** A name from the shared icon set, for products whose menus show icons. */
	icon?: string;
	/** `top` (the default) is a group above the built-in entries; `bottom` one below them. */
	group?: 'top' | 'bottom';
	/** Shown only when this returns true (or is omitted). */
	when?: (context: C) => boolean;
	/** A disabled command is shown greyed out and never runs. */
	disabled?: boolean | ((context: C) => boolean);
	/** Runs after the menu closes, with the context the menu was opened on. */
	onSelect: (context: C) => void;
}

/** A host command resolved for one opening of a menu. */
export interface HostMenuEntry {
	/** `host:<id>`, so a host id can never collide with a built-in entry. */
	id: `host:${string}`;
	host: true;
	label: string;
	/** Kept for menus that look labels up by key; always empty for a host entry. */
	labelKey: string;
	icon?: string;
	separatorBefore?: boolean;
	disabled: boolean;
	danger?: boolean;
	checked?: boolean;
	/** Runs the command (a no-op when disabled). */
	onSelect: () => void;
}

/** The menu id of host command `id`. */
export function hostMenuEntryId(id: string): `host:${string}` {
	return `host:${id}`;
}

/** True for an entry made from a host command. */
export function isHostMenuEntry(entry: unknown): entry is HostMenuEntry {
	return typeof entry === 'object' && entry !== null && (entry as { host?: unknown }).host === true;
}

/** The host entries for one opening of a menu, split by group, in the host's order. */
export function resolveHostMenuCommands<C>(
	commands: readonly HostMenuCommand<C>[] | null | undefined,
	context: C,
): { top: HostMenuEntry[]; bottom: HostMenuEntry[] } {
	const top: HostMenuEntry[] = [];
	const bottom: HostMenuEntry[] = [];
	const seen = new Set<string>();
	for (const command of commands ?? []) {
		if (!command?.id || seen.has(command.id)) continue;
		seen.add(command.id);
		if (command.when && !command.when(context)) continue;
		const disabled =
			typeof command.disabled === 'function'
				? command.disabled(context)
				: command.disabled === true;
		const entry: HostMenuEntry = {
			id: hostMenuEntryId(command.id),
			host: true,
			label: command.label,
			labelKey: '',
			disabled,
			onSelect: () => {
				if (!disabled) command.onSelect(context);
			},
		};
		if (command.icon) entry.icon = command.icon;
		(command.group === 'bottom' ? bottom : top).push(entry);
	}
	return { top, bottom };
}

/**
 * A menu with the host commands placed: the `top` group before the built-in entries and the
 * `bottom` group after them, each separated from its neighbours, with no leading or trailing
 * separator.
 */
export function withHostMenuCommands<E extends { id: string; separatorBefore?: boolean }, C>(
	entries: readonly E[],
	commands: readonly HostMenuCommand<C>[] | null | undefined,
	context: C,
): Array<E | HostMenuEntry> {
	const { top, bottom } = resolveHostMenuCommands(commands, context);
	const middle = entries.map((entry, index) =>
		index === 0 ? { ...entry, separatorBefore: top.length > 0 } : entry,
	);
	const first = bottom[0];
	if (first && top.length + middle.length > 0) first.separatorBefore = true;
	return [...top, ...middle, ...bottom];
}

/**
 * A menu without its hidden entries, with its groups intact: a hidden entry that starts a group
 * (`separatorBefore`) passes the separator on to the next kept entry, and the first kept entry
 * never starts with one. So hiding entries never merges two groups, never doubles a separator
 * and never leaves one at the top. `entries` is not changed.
 */
export function withoutHiddenEntries<E extends { separatorBefore?: boolean }>(
	entries: readonly E[],
	hidden: (entry: E) => boolean,
): E[] {
	const kept: E[] = [];
	let separator = false;
	for (const entry of entries) {
		separator ||= entry.separatorBefore === true;
		if (hidden(entry)) continue;
		const starts = separator && kept.length > 0;
		kept.push(
			(entry.separatorBefore === true) === starts ? entry : { ...entry, separatorBefore: starts },
		);
		separator = false;
	}
	return kept;
}
