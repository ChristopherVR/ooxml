import { afterEach, describe, expect, it, vi } from 'vitest';
import type { VisioDocument } from 'ooxml-core/visio';
import { ViewerController } from './controller.js';
import { createWorkerParser } from './worker-parser.js';
import { demoDocument } from './demo-document.js';

const { loadVisio, convertMetafileToSvgTree } = vi.hoisted(() => ({
	loadVisio: vi.fn(),
	convertMetafileToSvgTree: vi.fn(),
}));
vi.mock('ooxml-core/visio', async (importOriginal) => ({
	...(await importOriginal<object>()),
	loadVisio,
}));
vi.mock('emf-converter', () => ({ convertMetafileToSvgTree }));
afterEach(() => {
	vi.unstubAllGlobals();
	loadVisio.mockReset();
});
const sources = [
	{ format: 'vsdx' as const, bytes: new Uint8Array([0x50, 0x4b, 3, 4]) },
	{ format: 'vsd' as const, bytes: new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]) },
];
function scene(format: VisioDocument['format']): VisioDocument {
	return { ...structuredClone(demoDocument), format };
}

describe('unified legacy parser dispatch', () => {
	// Signature-shaped inputs exercise routing, not either codec's binary fidelity.
	it.each(sources)(
		'uses loadVisio for default controller and workerless $format parsing',
		async ({ format, bytes }) => {
			vi.stubGlobal('Worker', undefined);
			const document = scene(format);
			loadVisio.mockResolvedValue(document);
			const controller = new ViewerController();
			try {
				await controller.load(bytes);
				expect(controller.state.document).toBe(document);
				expect(loadVisio.mock.calls[0]?.[0]).toEqual(bytes);
				expect(loadVisio.mock.calls[0]?.[0]).not.toBe(bytes);
				const parser = createWorkerParser();
				expect(parser).toBe(loadVisio);
				await expect(parser(bytes)).resolves.toBe(document);
				expect(loadVisio).toHaveBeenCalledTimes(2);
				expect(controller.state.edit.sourceAvailable).toBe(format === 'vsdx');
			} finally {
				controller.destroy();
			}
		},
	);

	it('dispatches both worker formats through loadVisio with the metafile converter and reports failures', async () => {
		const worker: {
			onmessage: ((event: MessageEvent<ArrayBuffer>) => Promise<void>) | null;
			postMessage: ReturnType<typeof vi.fn>;
		} = { onmessage: null, postMessage: vi.fn() };
		vi.stubGlobal('self', worker);
		await import('./parse-worker.js');
		expect(worker.onmessage).toBeTypeOf('function');
		for (const { format, bytes } of sources) {
			const document = scene(format);
			loadVisio.mockResolvedValueOnce(document);
			await worker.onmessage!({ data: bytes.buffer } as MessageEvent<ArrayBuffer>);
			expect(loadVisio).toHaveBeenLastCalledWith(bytes.buffer, {
				metafileConverter: convertMetafileToSvgTree,
			});
			expect(worker.postMessage).toHaveBeenLastCalledWith({ ok: true, document });
		}
		loadVisio.mockRejectedValueOnce(new Error('bounded legacy input rejected'));
		await worker.onmessage!({ data: new ArrayBuffer(1) } as MessageEvent<ArrayBuffer>);
		expect(worker.postMessage).toHaveBeenLastCalledWith({
			ok: false,
			message: 'bounded legacy input rejected',
		});
		expect(loadVisio).toHaveBeenCalledTimes(3);
		expect(worker.postMessage).toHaveBeenCalledTimes(3);
	});
});
