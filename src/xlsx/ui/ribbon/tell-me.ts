/**
 * "Tell me what you want to do": searches every registered, enabled, visible command by its
 * translated and English label. The shared title bar shows the matches; choosing one runs it, and a
 * command that needs a value (font, size, number format, colour) takes the user to its ribbon
 * control instead (see title-bar.ts).
 */
import type { Command } from '../commands.js';
import type { EditorContext } from '../context.js';

export interface TellMeHandlers {
	isHidden(id: string): boolean;
	/** Shows the ribbon tab holding a command's control and focuses it; false when there is none. */
	revealControl(id: string): boolean;
}

const MAX_RESULTS = 8;

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
