import { captureVisioClipboard, serializeVisioClipboard } from 'ooxml-core/visio';
import type { ClipboardCaptureRequest } from './worker-clipboard';
import { visioWorkerScope } from './worker-scope';

const worker = visioWorkerScope;

/** The worker owns package/XML capture; only the bounded portable payload leaves it. */
if (worker)
	worker.onmessage = async (event: MessageEvent<ClipboardCaptureRequest>) => {
		if (event.origin && event.origin !== worker.location.origin) return;
		try {
			const request = event.data;
			const snapshot = await captureVisioClipboard(
				new Uint8Array(request.bytes),
				request.pageId,
				request.shapeIds,
			);
			worker.postMessage({ ok: true, text: serializeVisioClipboard(snapshot) });
		} catch (cause) {
			const error = cause instanceof Error ? cause : new Error(String(cause));
			worker.postMessage({
				ok: false,
				message: error.message.slice(0, 4096),
				code: 'code' in error ? error.code : undefined,
			});
		}
	};
