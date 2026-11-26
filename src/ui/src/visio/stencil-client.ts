import type { VisioStencilEdit, VisioStencilFile, VisioStencilSave } from 'ooxml-core/visio';
import type { StencilWorkerRequest, StencilWorkerResponse } from './stencil-worker-request';

type Response<K extends StencilWorkerRequest['type']> = Extract<StencilWorkerResponse, { type: K }>;

async function requestStencil<K extends StencilWorkerRequest['type']>(
	request: Extract<StencilWorkerRequest, { type: K }>,
): Promise<Response<K>['value']>;
async function requestStencil(
	request: StencilWorkerRequest,
): Promise<Uint8Array | VisioStencilFile | VisioStencilSave> {
	// Headless DOM consumers and hosts without Worker support still use the bounded reader.
	if (typeof Worker === 'undefined') {
		const { processStencilRequest } = await import('./stencil-worker-request');
		const result = await processStencilRequest(request);
		return result.value;
	}
	const worker = new Worker(new URL('./stencil-worker.js', import.meta.url), { type: 'module' });
	return new Promise((resolve, reject) => {
		const finish = () => {
			clearTimeout(timer);
			worker.terminate();
		};
		const timer = setTimeout(() => {
			finish();
			reject(new Error('The stencil exceeded the isolated editing time limit.'));
		}, 20_000);
		worker.onerror = (event) => {
			event.preventDefault();
			finish();
			reject(new Error('The Visio stencil worker failed.'));
		};
		worker.onmessage = ({ data }: MessageEvent<StencilWorkerResponse>) => {
			finish();
			if (data.type === 'error') reject(new Error(data.message));
			else if (data.type === request.type && 'value' in data) resolve(data.value);
			else reject(new Error('The Visio stencil worker failed.'));
		};
		try {
			worker.postMessage(request);
		} catch (error) {
			finish();
			reject(error);
		}
	});
}

export const readStencil = (source: Uint8Array) => requestStencil<'read'>({ type: 'read', source });
export const createStencil = () => requestStencil<'create'>({ type: 'create' });
export const editStencil = (source: Uint8Array, edit?: VisioStencilEdit) =>
	requestStencil<'edit'>({ type: 'edit', source, ...(edit ? { edit } : {}) });
