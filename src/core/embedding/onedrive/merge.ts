import type { OneDriveMergeInput, OneDriveMergeResult, OneDriveSessionOptions } from './types';

/** An aborted host merge must never continue into a cloud write when it eventually resolves. */
export async function runOneDriveMerge(
	merge: NonNullable<OneDriveSessionOptions['merge']>,
	input: OneDriveMergeInput,
): Promise<OneDriveMergeResult> {
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
