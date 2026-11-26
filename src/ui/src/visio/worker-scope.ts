/// <reference lib="webworker" />

/** Worker entries also run under the Node worker harness, which supplies the messaging API. */
export const visioWorkerScope =
	typeof window === 'undefined' &&
	typeof self !== 'undefined' &&
	typeof self.postMessage === 'function'
		? (self as unknown as DedicatedWorkerGlobalScope)
		: undefined;
