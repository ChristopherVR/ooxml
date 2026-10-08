import { captureVisioClipboard, serializeVisioClipboard } from 'ooxml-core/visio';
import type { ClipboardCaptureRequest } from './worker-clipboard';

/** The worker owns package/XML capture; only the bounded portable payload leaves it. */
self.onmessage = async (event: MessageEvent<ClipboardCaptureRequest>) => {
	try {
		const request = event.data;
		const snapshot = await captureVisioClipboard(
			new Uint8Array(request.bytes),
			request.pageId,
			request.shapeIds,
		);
		self.postMessage({ ok: true, text: serializeVisioClipboard(snapshot) });
	} catch (cause) {
		const error = cause instanceof Error ? cause : new Error(String(cause));
		self.postMessage({
			ok: false,
			message: error.message.slice(0, 4096),
			code: 'code' in error ? error.code : undefined,
		});
	}
};
