/// <reference lib="webworker" />

/** Dedicated workers expose messaging and a worker global, including in module workers. */
export const visioWorkerScope: DedicatedWorkerGlobalScope | undefined =
	typeof self !== 'undefined' &&
	typeof self.postMessage === 'function' &&
	((typeof WorkerGlobalScope !== 'undefined' && self instanceof WorkerGlobalScope) ||
		'importScripts' in self)
		? (self as typeof self & DedicatedWorkerGlobalScope)
		: undefined;
