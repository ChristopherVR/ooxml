/// <reference lib="webworker" />
import { parseVsdx } from 'ooxml-core/visio';
const worker = self as unknown as DedicatedWorkerGlobalScope;
worker.onmessage = async (event: MessageEvent<ArrayBuffer>) => {
	try {
		const document = await parseVsdx(event.data);
		worker.postMessage({ ok: true, document });
	} catch (cause) {
		worker.postMessage({
			ok: false,
			message: cause instanceof Error ? cause.message : String(cause),
		});
	}
};
