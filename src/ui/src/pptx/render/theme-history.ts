/**
 * A theme change as an undo step, for every binding.
 *
 * A Design gallery pick, the theme editor and a colour or font scheme edit change the deck's theme
 * as well as its slides. Undo must take both back: the slides come back with the history snapshot,
 * but the theme lives in the binding's theme state and, for a save, in the handler's theme parts.
 * So each binding records {@link historyTheme} with every snapshot, and an undo or redo that
 * {@link crossedHistoryTheme crosses} a theme change puts the binding's theme back and writes only
 * the parts that differ into the handler ({@link restoreCrossedTheme}).
 *
 * Each theme part write reads, edits and writes the part back, so the writes on one handler run one
 * after another ({@link queueThemeWrite}), and every save waits for the ones queued before it
 * ({@link settleThemeWrites}, which `saveDeckWithPassword` calls).
 *
 * @module render/theme-history
 */
import type { PptxThemeColorScheme, PptxThemeFontScheme, PptxThemeOption } from 'ooxml-core/pptx';

/**
 * The part of a theme an undo step records: the colour scheme, the font scheme and the name, which
 * are what the Design galleries, the theme editor and the assistant change.
 */
export interface HistoryTheme {
	colorScheme?: PptxThemeColorScheme;
	fontScheme?: PptxThemeFontScheme;
	name?: string;
}

/** Any theme-shaped value: a binding's `PptxTheme`, or the looser record some bindings keep. */
interface ThemeLike {
	colorScheme?: unknown;
	fontScheme?: unknown;
	name?: unknown;
}

const RECORDED_KEYS = ['colorScheme', 'fontScheme', 'name'] as const;

/** A plain copy (works on a framework's reactive proxy, which `structuredClone` rejects). */
function plain<T>(value: T): T {
	return JSON.parse(JSON.stringify(value)) as T;
}

/** Whether two plain values serialise the same. */
function sameJson(a: unknown, b: unknown): boolean {
	return JSON.stringify(a) === JSON.stringify(b);
}

/**
 * The part of `theme` an undo step records, as a copy. Built in a fixed key order, so two records
 * of the same theme serialise to the same JSON.
 */
export function historyTheme(theme: ThemeLike): HistoryTheme;
export function historyTheme(theme: ThemeLike | null | undefined): HistoryTheme | undefined;
export function historyTheme(theme: ThemeLike | null | undefined): HistoryTheme | undefined {
	if (!theme) {
		return undefined;
	}
	const { colorScheme, fontScheme, name } = theme as HistoryTheme;
	return {
		...(colorScheme ? { colorScheme: plain(colorScheme) } : {}),
		...(fontScheme ? { fontScheme: plain(fontScheme) } : {}),
		...(name ? { name } : {}),
	};
}

/**
 * The theme an undo or redo from `current` to `target` has to put back, or `undefined` when the
 * step crosses no theme change (or `target` recorded none).
 */
export function crossedHistoryTheme(
	current: HistoryTheme | undefined,
	target: HistoryTheme | undefined,
): HistoryTheme | undefined {
	return target && !sameJson(current ?? null, target) ? target : undefined;
}

/**
 * `theme` carrying exactly the recorded parts of `target`: a part `target` does not have is
 * removed, so the restored theme records the same as the snapshot it came from. Everything else on
 * the theme (the effect scheme, for one) is kept.
 */
export function withHistoryTheme<T extends ThemeLike>(
	theme: T | null | undefined,
	target: HistoryTheme,
): T {
	const next: Record<string, unknown> = { ...theme, ...plain(target) };
	for (const key of RECORDED_KEYS) {
		if (target[key] === undefined) {
			delete next[key];
		}
	}
	return next as T;
}

/** The handler operations that write the recorded parts (the `PptxHandler` has all three). */
export interface ThemePartWriter {
	updateThemeColorScheme(colorScheme: PptxThemeColorScheme): Promise<void>;
	updateThemeFontScheme(fontScheme: PptxThemeFontScheme): Promise<void>;
	updateThemeName(name: string): Promise<void>;
}

/** One theme part write. */
export type ThemePartWrite =
	| { readonly part: 'colorScheme'; readonly colorScheme: PptxThemeColorScheme }
	| { readonly part: 'fontScheme'; readonly fontScheme: PptxThemeFontScheme }
	| { readonly part: 'name'; readonly name: string };

/**
 * The writes that take the handler's theme parts from `from` to `to`: only the parts that differ,
 * so undoing a fonts pick leaves the colour scheme part (and its `sysClr` slots) as it is. The
 * parsed theme carries no name, so `loadedName` (see {@link loadedThemeName}) stands in for a
 * theme without one: going back to the loaded theme puts its name back.
 */
export function themePartWrites(
	from: HistoryTheme | undefined,
	to: HistoryTheme,
	loadedName?: string,
): ThemePartWrite[] {
	const writes: ThemePartWrite[] = [];
	if (to.colorScheme && !sameJson(to.colorScheme, from?.colorScheme)) {
		writes.push({ part: 'colorScheme', colorScheme: to.colorScheme });
	}
	if (to.fontScheme && !sameJson(to.fontScheme, from?.fontScheme)) {
		writes.push({ part: 'fontScheme', fontScheme: to.fontScheme });
	}
	const name = to.name ?? loadedName;
	if (name && name !== (from?.name ?? loadedName)) {
		writes.push({ part: 'name', name });
	}
	return writes;
}

/**
 * The name the deck's primary theme part (the first master's) was loaded with; the parsed theme
 * carries none.
 */
export function loadedThemeName(
	themeOptions: readonly Pick<PptxThemeOption, 'path' | 'name'>[] | undefined,
	slideMasters: readonly { readonly themePath?: unknown }[] | undefined,
): string | undefined {
	const options = themeOptions ?? [];
	const primary = slideMasters?.[0]?.themePath;
	return (options.find((option) => option.path === primary) ?? options[0])?.name;
}

/** Per handler, the end of its queue: settles once the last queued write has landed or failed. */
const pendingWrites = new WeakMap<object, Promise<void>>();

function startWrite<T>(write: () => Promise<T>): Promise<T> {
	try {
		return write();
	} catch (error) {
		return Promise.reject(error);
	}
}

/**
 * Run `write` on `handler`'s theme parts after every write queued on it before. An undo landing
 * while a pick is still writing (or two undos in a row) waits for it instead of interleaving.
 * With nothing queued the write starts at once, exactly as an unqueued one would.
 */
export function queueThemeWrite<T>(handler: object, write: () => Promise<T>): Promise<T> {
	const previous = pendingWrites.get(handler);
	const run = previous ? previous.then(write) : startWrite(write);
	const settled: Promise<void> = run.then(
		() => undefined,
		() => undefined,
	);
	const end = settled.then(() => {
		if (pendingWrites.get(handler) === end) {
			pendingWrites.delete(handler);
		}
	});
	pendingWrites.set(handler, end);
	return run;
}

/**
 * Resolves once every theme part write queued on `handler` so far has landed (or failed). A save
 * waits for this, or it writes a theme an undo or a pick is still replacing.
 */
export function settleThemeWrites(handler: object | null | undefined): Promise<void> {
	return (handler && pendingWrites.get(handler)) || Promise.resolve();
}

/** Write {@link themePartWrites} into `handler`, after every write queued on it before. */
export function writeThemeParts(
	handler: ThemePartWriter | null | undefined,
	from: HistoryTheme | undefined,
	to: HistoryTheme,
	loadedName?: string,
): Promise<void> {
	const writes = themePartWrites(from, to, loadedName);
	if (!handler || writes.length === 0) {
		return Promise.resolve();
	}
	return queueThemeWrite(handler, async () => {
		for (const write of writes) {
			if (write.part === 'colorScheme') {
				await handler.updateThemeColorScheme(write.colorScheme);
			} else if (write.part === 'fontScheme') {
				await handler.updateThemeFontScheme(write.fontScheme);
			} else {
				await handler.updateThemeName(write.name);
			}
		}
	});
}

/**
 * An undo or redo moved from `current` to `target`: when it crosses a theme change, queue the
 * writes of the parts that differ on `handler` and return `theme` carrying `target`'s recorded
 * parts, for the binding to put back. `undefined` when the step crosses no theme change.
 */
export function restoreCrossedTheme<T extends ThemeLike>(
	theme: T | null | undefined,
	current: HistoryTheme | undefined,
	target: HistoryTheme | undefined,
	handler: ThemePartWriter | null | undefined,
	loadedName?: string,
): T | undefined {
	const crossed = crossedHistoryTheme(current, target);
	if (!crossed) {
		return undefined;
	}
	void writeThemeParts(handler, current, crossed, loadedName);
	return withHistoryTheme(theme, crossed);
}
