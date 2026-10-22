// Helpers behind `mountEditor` (index.ts): forward a prop only when the parent changed it, so a
// re-render never undoes a value the user set inside the editor (File > Open renames the document).
import type { EditorTheme } from 'docx-web-component';

/** Shallow, by value: an inline `themeColors` object passed on every render is not a change. */
export function sameThemeColors(
	a: Partial<EditorTheme> | undefined,
	b: Partial<EditorTheme> | undefined,
): boolean {
	if (a === b) return true;
	const left: Record<string, unknown> = { ...a };
	const right: Record<string, unknown> = { ...b };
	const keys = Object.keys(left);
	return keys.length === Object.keys(right).length && keys.every((key) => left[key] === right[key]);
}

/**
 * Remembers the last value forwarded for each prop. `changed` is true the first time and then only
 * when the prop itself changed.
 */
export function createPropTracker<K extends string>() {
	const sent = new Map<K, unknown>();
	return function changed<T>(key: K, value: T, same: (a: T, b: T) => boolean = Object.is): boolean {
		if (sent.has(key) && same(sent.get(key) as T, value)) return false;
		sent.set(key, value);
		return true;
	};
}
