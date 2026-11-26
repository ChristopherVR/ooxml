/// <reference lib="webworker" />
import { processStencilRequest, type StencilWorkerRequest } from './stencil-worker-request';

if (typeof WorkerGlobalScope !== 'undefined' && self instanceof WorkerGlobalScope) {
	const worker = self as unknown as DedicatedWorkerGlobalScope;
	worker.addEventListener('message', async (event: MessageEvent<StencilWorkerRequest>) => {
		if (event.origin && event.origin !== worker.location.origin) return;
		try {
			worker.postMessage(await processStencilRequest(event.data));
		} catch (cause) {
			const message = cause instanceof Error ? cause.message : String(cause);
			worker.postMessage({ type: 'error', message: message.slice(0, 4096) });
		}
	});
}
