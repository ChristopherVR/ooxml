// Helpers behind `mountEditor` (index.ts): change tracking for forwarded props, and the `src`
// fetcher that a later `bytes`, `workbook` or `src` supersedes.
import type { XlsxThemeColors } from 'xlsx-web-component';

export function sameList(a: readonly string[] | undefined, b: readonly string[] | undefined) {
	if (a === b) return true;
	const left = a ?? [];
	const right = b ?? [];
	return left.length === right.length && left.every((item, index) => item === right[index]);
}

/** Shallow, by value: an inline `themeColors` object passed on every render is not a change. */
export function sameThemeColors(a: XlsxThemeColors | undefined, b: XlsxThemeColors | undefined) {
	if (a === b) return true;
	const left: Record<string, unknown> = { ...a };
	const right: Record<string, unknown> = { ...b };
	const keys = Object.keys(left);
	return keys.length === Object.keys(right).length && keys.every((key) => left[key] === right[key]);
}

/**
 * Remembers the last value forwarded for each prop. `changed` is true the first time and then only
 * when the prop itself changed, so a parent re-render never undoes what the user changed inside
 * the editor (Editing / Viewing, File > Options, Customize Ribbon).
 */
export function createPropTracker<K extends string>() {
	const sent = new Map<K, unknown>();
	return function changed<T>(key: K, value: T, same: (a: T, b: T) => boolean = Object.is): boolean {
		if (sent.has(key) && same(sent.get(key) as T, value)) return false;
		sent.set(key, value);
		return true;
	};
}

/** The file name in a URL's last path segment; the raw segment when it is not valid %-encoding. */
export function nameFromUrl(url: string): string | undefined {
	const last = url.split(/[?#]/u)[0]?.split('/').pop();
	if (!last) return undefined;
	try {
		return decodeURIComponent(last);
	} catch {
		return last;
	}
}

export interface SrcLoader {
	/** Fetches `url` and loads it, superseding any fetch still in flight. */
	fetch(url: string, fileName: string | undefined): void;
	/** Drops the fetch in flight (new bytes or workbook arrived, `src` was cleared, or destroy). */
	cancel(): void;
}

/**
 * Every failure in the chain (fetch throwing, HTTP errors, reading the body, naming) is reported
 * through `report`; a superseded fetch is aborted and neither loads nor reports.
 */
export function createSrcLoader(
	load: (bytes: Uint8Array, fileName: string | undefined) => Promise<void>,
	report: (error: unknown) => void,
): SrcLoader {
	let generation = 0;
	let controller: AbortController | undefined;
	const cancel = () => {
		generation++;
		controller?.abort();
		controller = undefined;
	};
	return {
		cancel,
		fetch(url, fileName) {
			cancel();
			const current = generation;
			const signal = typeof AbortController === 'undefined' ? undefined : new AbortController();
			controller = signal;
			const live = () => current === generation;
			void Promise.resolve()
				.then(() => fetch(url, signal ? { signal: signal.signal } : undefined))
				.then(async (response) => {
					if (!response.ok) throw new Error(`Could not fetch ${url}: HTTP ${response.status}`);
					return new Uint8Array(await response.arrayBuffer());
				})
				.then((bytes) => (live() ? load(bytes, fileName ?? nameFromUrl(url)) : undefined))
				.catch((error: unknown) => {
					if (live()) report(error);
				});
		},
	};
}
