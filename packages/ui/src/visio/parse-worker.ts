/// <reference lib="webworker" />
import { convertMetafileToSvgTree } from 'emf-converter';
import { loadVisio } from 'ooxml-core/visio';
const worker = self as unknown as DedicatedWorkerGlobalScope;
worker.onmessage = async (event: MessageEvent<ArrayBuffer>) => {
	try {
		const document = await loadVisio(event.data, { metafileConverter: convertMetafileToSvgTree });
		worker.postMessage({ ok: true, document });
	} catch (cause) {
		worker.postMessage({
			ok: false,
			message: cause instanceof Error ? cause.message : String(cause),
		});
	}
};
