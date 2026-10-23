// The `src` prop: fetch a file by URL and load it, superseded by anything newer (another `src`,
// new bytes, a document model, or unmounting).

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
	/** Drops the fetch in flight (new bytes or a model arrived, `src` was cleared, or destroy). */
	cancel(): void;
}

/**
 * Every failure in the chain (fetch throwing, HTTP errors, reading the body, naming) is reported
 * through `report`; a superseded fetch is aborted and neither loads nor reports. Without an explicit
 * file name the URL's last segment names the file.
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
