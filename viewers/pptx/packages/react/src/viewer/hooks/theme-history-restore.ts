import type { PptxHandler } from 'pptx-viewer-core';

import type { EditorHistoryTheme } from '../types';
import { historyTheme, withHistoryTheme } from './history-theme';
import type { UseThemeHandlersInput } from './useThemeHandlers';

const pendingWrites = new WeakMap<PptxHandler, Promise<void>>();

/**
 * Run `write` on `handler`'s theme parts after every write queued before it. Each write reads,
 * edits and writes the parts back, so an undo landing while a pick is still writing (or two undos
 * in a row) must wait for it rather than interleave with it.
 */
export function queueThemeWrite(handler: PptxHandler, write: () => Promise<void>): Promise<void> {
	const run = (pendingWrites.get(handler) ?? Promise.resolve()).then(write);
	pendingWrites.set(
		handler,
		run.catch(() => undefined),
	);
	return run;
}

/**
 * Resolves once every theme part write queued on `handler` so far has landed. An undo or redo puts
 * the slides back at once and writes the theme part after, so a save must wait for this or it
 * writes the theme the step just undid.
 */
export function settleThemeWrites(handler: PptxHandler): Promise<void> {
	return pendingWrites.get(handler) ?? Promise.resolve();
}

/** The name the deck's primary theme part was loaded with; the parsed theme carries no name. */
function loadedThemeName(input: UseThemeHandlersInput): string | undefined {
	const options = input.themeOptions ?? [];
	const primary = (input.slideMasters[0] as { themePath?: string } | undefined)?.themePath;
	return (options.find((option) => option.path === primary) ?? options[0])?.name;
}

const differs = (a: unknown, b: unknown): boolean => JSON.stringify(a) !== JSON.stringify(b);

/**
 * An undo or redo crossed a theme change: put `target`'s colour scheme, font scheme and name
 * back, in the viewer's theme and in the handler's theme parts, so what a save writes matches the
 * slides again. Only the parts that differ are written: undoing a fonts pick leaves the colour
 * scheme part as it is.
 */
export function restoreHistoryTheme(
	input: UseThemeHandlersInput,
	target: EditorHistoryTheme,
): Promise<void> {
	const current = historyTheme(input.theme) ?? {};
	input.setTheme((prev) => withHistoryTheme(prev, target));
	const handler = input.handlerRef.current;
	if (!handler) {
		return Promise.resolve();
	}
	const loadedName = loadedThemeName(input);
	return queueThemeWrite(handler, async () => {
		if (target.colorScheme && differs(target.colorScheme, current.colorScheme)) {
			await handler.updateThemeColorScheme(target.colorScheme);
		}
		if (target.fontScheme && differs(target.fontScheme, current.fontScheme)) {
			await handler.updateThemeFontScheme(target.fontScheme);
		}
		const name = target.name ?? loadedName;
		if (name && name !== (current.name ?? loadedName)) {
			await handler.updateThemeName(name);
		}
	});
}
