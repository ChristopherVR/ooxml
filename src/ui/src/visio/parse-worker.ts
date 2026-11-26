/// <reference lib="webworker" />
import { loadVisio, type VisioDocumentCache } from 'ooxml-core/visio';
import { convertVisioMetafile } from './worker-metafile';
import type { ParseWorkerRequest } from './worker-parser';
const worker = globalThis as unknown as DedicatedWorkerGlobalScope;
/** The styles, themes and masters of the last page read: this worker stays for the next one. */
const documentCache: VisioDocumentCache = {};
if (typeof WorkerGlobalScope !== 'undefined' && worker instanceof WorkerGlobalScope)
	worker.onmessage = async (event: MessageEvent<ParseWorkerRequest>) => {
		if (event.origin && event.origin !== self.location.origin) return;
		try {
			const request = event.data;
			const document =
				request instanceof ArrayBuffer
					? await loadVisio(request, { metafileConverter: convertVisioMetafile })
					: await loadVisio(request.bytes, {
							metafileConverter: convertVisioMetafile,
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
