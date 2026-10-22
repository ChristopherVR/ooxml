/// <reference lib="webworker" />
import { convertMetafileToSvgTree } from 'emf-converter';
import { loadVisio, type VisioDocumentCache } from 'ooxml-core/visio';
import type { ParseWorkerRequest } from './worker-parser';
const worker = self as unknown as DedicatedWorkerGlobalScope;
/** The styles, themes and masters of the last page read: this worker stays for the next one. */
const documentCache: VisioDocumentCache = {};
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
						documentCache,
					});
		worker.postMessage({ ok: true, document });
	} catch (cause) {
		worker.postMessage({
			ok: false,
			message: cause instanceof Error ? cause.message : String(cause),
		});
	}
};
