import type {
	StorageMergeInput,
	StorageMergeResult,
	StorageMerger,
	VersionedFile,
} from './versioned-storage';

/** Route by file kind. Overrides replace a handler; absent kinds stay unresolved. No engines imported. */
export function createStorageMerger<F extends VersionedFile = VersionedFile>(
	handlers: Readonly<Record<string, StorageMerger<F> | undefined>>,
	options: { fallback?: StorageMerger<F> } = {},
): StorageMerger<F> {
	const registry = new Map(Object.entries(handlers));
	return (input) => {
		input.signal?.throwIfAborted();
		const handler = registry.has(input.kind) ? registry.get(input.kind) : options.fallback;
		return handler
			? runStorageMerge(handler, input)
			: {
					status: 'unresolved',
					issues: [`unsupported-merge-kind:${input.kind}`],
				};
	};
}

/** An aborted host merge must never continue into a cloud write when it eventually resolves. */
export async function runStorageMerge<F extends VersionedFile>(
	merge: StorageMerger<F>,
	input: StorageMergeInput<F>,
): Promise<StorageMergeResult> {
	const { signal } = input;
	signal?.throwIfAborted();
	if (!signal) return merge(input);
	let abort: () => void = () => {};
	try {
		return await Promise.race([
			merge(input),
			new Promise<never>((_, reject) => {
				abort = () => reject(signal.reason);
				signal.addEventListener('abort', abort, { once: true });
				if (signal.aborted) abort();
			}),
		]);
	} finally {
		signal.removeEventListener('abort', abort);
	}
}
