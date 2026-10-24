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
 * A theme applied by path (`render/theme-by-path`) points slide masters at another theme part, so the
 * record also holds each master's theme part ({@link historyTheme} with the masters), and a restore
 * that crosses such a switch points the masters back before it writes anything: Undo returns to
 * the old part instead of writing the old schemes into the new one. The record describes the deck's
 * part only, so what such a restore still writes goes into that part alone.
 *
 * Each theme part write reads, edits and writes the part back, so the writes on one handler run one
 * after another ({@link queueThemeWrite}), and every save waits for the ones queued before it
 * ({@link settleThemeWrites}, which `saveDeckWithPassword` calls).
 *
 * @module render/theme-history
 */
import type {
	PptxThemeColorScheme,
	PptxThemeFontScheme,
	PptxThemeFormatScheme,
	PptxThemeOption,
} from 'ooxml-core/pptx';

import { queueThemeWrite } from './theme-write-queue';

export { queueThemeWrite, settleThemeWrites } from './theme-write-queue';

/**
 * The part of a theme an undo step records: the colour scheme, the font scheme and the name, which
 * are what the Design galleries, the theme editor and the assistant change, and the format scheme.
 */
export interface HistoryTheme {
	colorScheme?: PptxThemeColorScheme;
	fontScheme?: PptxThemeFontScheme;
	/**
	 * The deck's format scheme. A step carries it back, but it never makes a step cross a theme change
	 * by itself ({@link crossedHistoryTheme}): no theme part write restores it, only pointing a master
	 * back at its own part does.
	 */
	formatScheme?: PptxThemeFormatScheme;
	name?: string;
	/** Each slide master's theme part, by master path, in master order (the first is the deck's). */
	masterThemePaths?: Readonly<Record<string, string>>;
}

/** A slide master as a binding keeps it: its path and the theme part it uses. */
export interface HistoryMaster {
	readonly path: string;
	readonly themePath?: string | undefined;
}

/** Any theme-shaped value: a binding's `PptxTheme`, or the looser record some bindings keep. */
interface ThemeLike {
	colorScheme?: unknown;
	fontScheme?: unknown;
	formatScheme?: unknown;
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

/** Each master's theme part, in master order, or `undefined` when no master names one. */
function masterThemePathsOf(
	slideMasters: readonly HistoryMaster[] | undefined,
): Record<string, string> | undefined {
	const entries = (slideMasters ?? []).flatMap((master) =>
		master.themePath ? [[master.path, master.themePath] as const] : [],
	);
	return entries.length > 0 ? Object.fromEntries(entries) : undefined;
}

/**
 * The part of `theme` an undo step records, as a copy, and with `slideMasters` the theme part each
 * master uses. Built in a fixed key order, so two records of the same theme serialise to the same
 * JSON.
 */
export function historyTheme(
	theme: ThemeLike,
	slideMasters?: readonly HistoryMaster[] | undefined,
): HistoryTheme;
export function historyTheme(
	theme: ThemeLike | null | undefined,
	slideMasters?: readonly HistoryMaster[] | undefined,
): HistoryTheme | undefined;
export function historyTheme(
	theme: ThemeLike | null | undefined,
	slideMasters?: readonly HistoryMaster[] | undefined,
): HistoryTheme | undefined {
	const masterThemePaths = masterThemePathsOf(slideMasters);
	if (!theme && !masterThemePaths) {
		return undefined;
	}
	const { colorScheme, fontScheme, formatScheme, name } = (theme ?? {}) as HistoryTheme;
	return {
		...(colorScheme ? { colorScheme: plain(colorScheme) } : {}),
		...(fontScheme ? { fontScheme: plain(fontScheme) } : {}),
		...(formatScheme ? { formatScheme: plain(formatScheme) } : {}),
		...(name ? { name } : {}),
		...(masterThemePaths ? { masterThemePaths } : {}),
	};
}

/** `theme` without its format scheme, which alone never makes a step cross a theme change. */
function withoutFormatScheme(theme: HistoryTheme | undefined): HistoryTheme | null {
	if (!theme) {
		return null;
	}
	const { formatScheme: _format, ...rest } = theme;
	return rest;
}

/**
 * The theme an undo or redo from `current` to `target` has to put back, or `undefined` when the
 * step crosses no theme change (or `target` recorded none). A step whose records differ only in the
 * format scheme (an Effects pick) crosses none.
 */
export function crossedHistoryTheme(
	current: HistoryTheme | undefined,
	target: HistoryTheme | undefined,
): HistoryTheme | undefined {
	return target && !sameJson(withoutFormatScheme(current), withoutFormatScheme(target))
		? target
		: undefined;
}

/** The theme part the deck's theme comes from (the first master's), as `theme` records it. */
function deckPartOf(theme: HistoryTheme | undefined): string | undefined {
	return Object.values(theme?.masterThemePaths ?? {})[0];
}

/**
 * `theme` carrying exactly the recorded parts of `target`: a part `target` does not have is
 * removed, so the restored theme records the same as the snapshot it came from. Everything else on
 * the theme is kept, the format scheme too unless the step from `from` puts the deck back on
 * another theme part (a theme applied by path), whose own format scheme comes back with it.
 */
export function withHistoryTheme<T extends ThemeLike>(
	theme: T | null | undefined,
	target: HistoryTheme,
	from?: HistoryTheme,
): T {
	const { masterThemePaths: _masters, formatScheme, ...parts } = target;
	const next: Record<string, unknown> = { ...theme, ...plain(parts) };
	for (const key of RECORDED_KEYS) {
		if (target[key] === undefined) {
			delete next[key];
		}
	}
	if (formatScheme && deckPartOf(from) !== deckPartOf(target)) {
		next.formatScheme = plain(formatScheme);
	}
	return next as T;
}

/**
 * `slideMasters` with each master the record names back on its recorded theme part: the same array
 * when none differs (a step that crosses no theme by path switch).
 */
export function withHistoryMasterThemes<M extends HistoryMaster>(
	slideMasters: readonly M[],
	target: HistoryTheme | undefined,
): readonly M[] {
	const recorded = target?.masterThemePaths;
	if (!recorded || slideMasters.every((master) => !differs(recorded, master))) {
		return slideMasters;
	}
	return slideMasters.map((master) =>
		differs(recorded, master) ? { ...master, themePath: recorded[master.path] } : master,
	);
}

function differs(recorded: Readonly<Record<string, string>>, master: HistoryMaster): boolean {
	const path = recorded[master.path];
	return path !== undefined && path !== master.themePath;
}

/** One theme part's schemes and name, as the handler reads them (`getThemePartSchemes`). */
export interface ThemePartSchemes {
	readonly path: string;
	readonly name?: string | undefined;
	readonly colorScheme: PptxThemeColorScheme;
	readonly fontScheme: PptxThemeFontScheme;
	readonly formatScheme?: PptxThemeFormatScheme | undefined;
}

/**
 * The handler operations that write the recorded parts (the `PptxHandler` has them all). A scheme
 * edit writes every part a slide master uses, or only `themePaths`.
 */
export interface ThemePartWriter {
	updateThemeColorScheme(
		colorScheme: PptxThemeColorScheme,
		themePaths?: readonly string[],
	): Promise<void>;
	updateThemeFontScheme(
		fontScheme: PptxThemeFontScheme,
		themePaths?: readonly string[],
	): Promise<void>;
	updateThemeName(name: string, themePaths?: readonly string[]): Promise<void>;
	/** Point one slide master at a theme part (undo and redo of a theme by path switch). */
	setSlideMasterTheme(masterPath: string, themePath: string): Promise<void>;
	/** One theme part's schemes and name as the part holds them now. */
	getThemePartSchemes(themePath: string): Promise<ThemePartSchemes | undefined>;
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

/** Write `writes` into every part the masters use, or with `scope` into those parts only. */
async function applyWrites(
	handler: ThemePartWriter,
	writes: readonly ThemePartWrite[],
	...scope: [] | [themePaths: readonly string[]]
) {
	for (const write of writes) {
		if (write.part === 'colorScheme') {
			await handler.updateThemeColorScheme(write.colorScheme, ...scope);
		} else if (write.part === 'fontScheme') {
			await handler.updateThemeFontScheme(write.fontScheme, ...scope);
		} else {
			await handler.updateThemeName(write.name, ...scope);
		}
	}
}

/**
 * `theme` with only the faces a font scheme write changes: the major and minor latin faces. The
 * per-script faces of a deck theme are merged across every theme part (the loaded `PptxTheme`), so
 * they never match one part's own, and the write leaves a part's script faces as they are anyway.
 */
function writableFaces(theme: HistoryTheme): HistoryTheme {
	const { fontScheme } = theme;
	if (!fontScheme) {
		return theme;
	}
	const latin = (font: { latin?: string } | undefined) =>
		font?.latin ? { latin: font.latin } : {};
	return {
		...theme,
		fontScheme: { majorFont: latin(fontScheme.majorFont), minorFont: latin(fontScheme.minorFont) },
	};
}

/** The masters `to` puts on another theme part than `from` has them on, with that part. */
function masterThemeMoves(from: HistoryTheme | undefined, to: HistoryTheme): [string, string][] {
	return Object.entries(to.masterThemePaths ?? {}).filter(
		([master, path]) => from?.masterThemePaths?.[master] !== path,
	);
}

/**
 * Write {@link themePartWrites} into `handler`, after every write queued on it before. When `to`
 * has masters on other theme parts (a step across a theme by path switch) they are pointed back
 * first, and only what then still differs from the deck's part (the first master's) is written,
 * into that part alone: Undo returns to the old part rather than writing the old schemes into the
 * new one, and leaves the parts of the other masters, which the record does not describe, alone.
 */
export function writeThemeParts(
	handler: ThemePartWriter | null | undefined,
	from: HistoryTheme | undefined,
	to: HistoryTheme,
	loadedName?: string,
): Promise<void> {
	const moves = masterThemeMoves(from, to);
	if (moves.length > 0) {
		return handler
			? queueThemeWrite(handler, () => repointAndWrite(handler, moves, to))
			: Promise.resolve();
	}
	const writes = themePartWrites(from, to, loadedName);
	if (!handler || writes.length === 0) {
		return Promise.resolve();
	}
	return queueThemeWrite(handler, () => applyWrites(handler, writes));
}

async function repointAndWrite(
	handler: ThemePartWriter,
	moves: readonly [string, string][],
	to: HistoryTheme,
): Promise<void> {
	for (const [master, path] of moves) {
		await handler.setSlideMasterTheme(master, path);
	}
	const deckPart = Object.values(to.masterThemePaths ?? {})[0];
	const now = deckPart ? await handler.getThemePartSchemes(deckPart) : undefined;
	if (deckPart && now) {
		// The part's own name stands for a record that carries none (the loaded theme).
		const writes = themePartWrites(writableFaces(historyTheme(now)), writableFaces(to), now.name);
		await applyWrites(handler, writes, [deckPart]);
	}
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
	return withHistoryTheme(theme, crossed, current);
}
