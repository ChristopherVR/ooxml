/// <reference lib="webworker" />
import { convertMetafileToSvgTree } from 'emf-converter';
import { loadVisio } from 'ooxml-core/visio';
import type { ParseWorkerRequest } from './worker-parser';
const worker = self as unknown as DedicatedWorkerGlobalScope;
worker.onmessage = async (event: MessageEvent<ParseWorkerRequest>) => {
	if (event.origin && event.origin !== self.location.origin) return;
	try {
		const request = event.data;
		const document =
			request instanceof ArrayBuffer
				? await loadVisio(request, { metafileConverter: convertMetafileToSvgTree })
				: await loadVisio(request.bytes, {
						metafileConverter: convertMetafileToSvgTree,
						pages: request.pages,
					});
		worker.postMessage({ ok: true, document });
	} catch (cause) {
		worker.postMessage({
			ok: false,
			message: cause instanceof Error ? cause.message : String(cause),
		});
	}
};
