import type { EditorHistoryTheme } from '../types';

/** Any theme-shaped value: the viewer's `PptxTheme`, or the loose record the theme handlers set. */
type ThemeLike = { colorScheme?: unknown; fontScheme?: unknown; name?: unknown };

const RECORDED_KEYS = ['colorScheme', 'fontScheme', 'name'] as const;

/**
 * The part of `theme` an undo step records (see {@link EditorHistoryTheme}). Built in a fixed key
 * order, so two records of the same theme serialise to the same JSON.
 */
export function historyTheme(theme: ThemeLike | null | undefined): EditorHistoryTheme | undefined {
	if (!theme) {
		return undefined;
	}
	const { colorScheme, fontScheme, name } = theme as EditorHistoryTheme;
	return {
		...(colorScheme ? { colorScheme } : {}),
		...(fontScheme ? { fontScheme } : {}),
		...(name ? { name } : {}),
	};
}

export function sameHistoryTheme(
	a: EditorHistoryTheme | undefined,
	b: EditorHistoryTheme | undefined,
): boolean {
	return JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}

/**
 * `theme` carrying exactly the recorded parts of `target`: a part `target` does not have is
 * removed, so the restored theme records the same as the snapshot it came from. Everything else
 * on the theme (the effect scheme, for one) is kept.
 */
export function withHistoryTheme<T extends ThemeLike>(
	theme: T | null | undefined,
	target: EditorHistoryTheme,
): T {
	const next: Record<string, unknown> = { ...theme, ...target };
	for (const key of RECORDED_KEYS) {
		if (target[key] === undefined) {
			delete next[key];
		}
	}
	return next as T;
}
