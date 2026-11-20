/**
 * "Tell me what you want to do": searches every registered, enabled, visible command by its
 * translated and English label. The shared title bar shows the matches; choosing one runs it, and a
 * command that needs a value (font, size, number format, colour) takes the user to its ribbon
 * control instead (see title-bar.ts).
 */
import type { Command } from '../commands';
import type { EditorContext } from '../context';
import type { RibbonControl, RibbonTab } from './parts';

export interface TellMeHandlers {
	isHidden(id: string): boolean;
	/** Shows the ribbon tab holding a command's control and focuses it; false when there is none. */
	revealControl(id: string): boolean;
}

const MAX_RESULTS = 8;

/** Every command a ribbon control (or its menu, gallery or dialog launcher) runs. */
function controlCommands(control: RibbonControl, into: Set<string>): void {
	if ('command' in control) into.add(control.command);
	if (control.kind === 'split' || control.kind === 'menu')
		for (const item of control.kind === 'split' ? control.menu : control.items)
			if ('command' in item) into.add(item.command);
	if (control.kind === 'stack') for (const inner of control.controls) controlCommands(inner, into);
}

/**
 * Commands whose only ribbon home is in `hiddenTabs` (`customization.ribbon.hiddenTabs`): Tell me
 * leaves them out. A command that is also in a visible tab, or on no tab, is not included.
 */
export function commandsOnlyInTabs(
	tabs: readonly RibbonTab[],
	hiddenTabs: ReadonlySet<string>,
): ReadonlySet<string> {
	if (!hiddenTabs.size) return new Set();
	const hidden = new Set<string>();
	const shown = new Set<string>();
	for (const tab of tabs) {
		const into = hiddenTabs.has(tab.id) ? hidden : shown;
		for (const group of tab.groups) {
			if (group.launcher) into.add(group.launcher);
			for (const control of group.controls) controlCommands(control, into);
		}
	}
	for (const id of shown) hidden.delete(id);
	return hidden;
}

/** Commands matching `query` (case-insensitive, any word order), best first. */
export function searchCommands(
	ctx: EditorContext,
	query: string,
	isHidden: (id: string) => boolean,
): Command[] {
	const words = query.trim().toLocaleLowerCase().split(/\s+/).filter(Boolean);
	if (!words.length) return [];
	const scored: Array<{ command: Command; score: number }> = [];
	const seen = new Set<string>();
	for (const command of ctx.commands.list()) {
		if (isHidden(command.id) || !ctx.commands.isEnabled(command.id)) continue;
		const label = ctx.t(command.label).toLocaleLowerCase();
		const english = command.label.toLocaleLowerCase();
		const haystack = `${label} ${english} ${command.id.replace(/[.-]/g, ' ')}`;
		if (!words.every((word) => haystack.includes(word))) continue;
		if (seen.has(label)) continue;
		seen.add(label);
		const score = label.startsWith(words[0]!) ? 0 : label.includes(words[0]!) ? 1 : 2;
		scored.push({ command, score });
	}
	return scored
		.sort(
			(a, b) => a.score - b.score || ctx.t(a.command.label).localeCompare(ctx.t(b.command.label)),
		)
		.slice(0, MAX_RESULTS)
		.map(({ command }) => command);
}
