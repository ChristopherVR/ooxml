// @vitest-environment node
import { execFileSync } from 'node:child_process';
import { once } from 'node:events';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { Worker } from 'node:worker_threads';
import { afterAll, beforeAll, expect, it } from 'vitest';
import { encodePng } from '../../../core/drawingml/raster-encode';
import { inspectRasterImage } from '../../../core/drawingml/raster-image';
import type { PictureCompressRequest } from './xlsx-picture-compress-worker';

const bun =
	process.platform === 'win32'
		? execFileSync('cmd.exe', ['/d', '/c', 'bun --print process.execPath'], {
				encoding: 'utf8',
			}).trim()
		: 'bun';
const workspace = fileURLToPath(new URL('../../../../', import.meta.url));
const scratch = mkdtempSync(join(tmpdir(), 'xlsx-picture-worker-'));
const bundle = join(scratch, 'worker.mjs');
const bundleUrl = pathToFileURL(bundle).href;

beforeAll(() => {
	execFileSync(
		bun,
		[
			'build',
			'src/ui/src/xlsx/xlsx-picture-compress-worker.ts',
			'--tsconfig-override',
			'src/ui/tsconfig.json',
			'--outfile',
			bundle,
		],
		{ cwd: workspace, stdio: 'pipe' },
	);
});
afterAll(() => rmSync(scratch, { recursive: true, force: true }));

it.each(['bun', 'node'])(
	'imports the worker bundle in %s without installing a handler',
	(runtime) => {
		execFileSync(
			runtime === 'bun' ? bun : 'node',
			[
				'--input-type=module',
				'--eval',
				`await import(${JSON.stringify(bundleUrl)});
			globalThis.self = {
				postMessage() {},
				addEventListener() { throw new Error('Installed a server-side handler'); }
			};
			await import(${JSON.stringify(bundleUrl + '?server')});`,
			],
			{ stdio: 'pipe' },
		);
	},
);

async function request(bytes: Uint8Array<ArrayBuffer>, outputSize: number) {
	const worker = new Worker(
		`const { parentPort, workerData } = require('node:worker_threads');
		let bitmapClosed = false;
		globalThis.self = {
			importScripts() {},
			addEventListener(type, listener, options) {
				if (type !== 'message' || !options.once) throw new Error('Invalid worker subscription');
				parentPort.once('message', (data) => listener({ data }));
			},
			postMessage(response) { parentPort.postMessage({ response, bitmapClosed }); }
		};
		globalThis.createImageBitmap = async () => ({
			width: 1, height: 1, close() { bitmapClosed = true; }
		});
		globalThis.OffscreenCanvas = class {
			getContext() {
				return { canvas: this, drawImage() {} };
			}
			async convertToBlob({ type }) {
				return new Blob([new Uint8Array(workerData.outputSize)], { type });
			}
		};
		import(workerData.bundleUrl);`,
		{ eval: true, workerData: { bundleUrl, outputSize } },
	);
	try {
		const message = once(worker, 'message', { signal: AbortSignal.timeout(5000) });
		worker.postMessage({
			bytes,
			options: { width: 1, height: 1, ppi: 96, crop: [0, 0, 0, 0], deleteCrop: false },
		} satisfies PictureCompressRequest);
		return (await message)[0];
	} finally {
		await worker.terminate();
	}
}

it('responds to valid, unchanged, GIF and invalid requests in worker_threads', async () => {
	const png = new Uint8Array(encodePng(1, 1, Uint8Array.of(255, 0, 0)));
	const gif = new Uint8Array(
		Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64'),
	);
	expect(inspectRasterImage(gif).mimeType).toBe('image/gif');
	expect(await request(png, 1)).toMatchObject({
		response: { kind: 'compressed', image: { type: 'image/png', size: 1 } },
		bitmapClosed: true,
	});
	expect(await request(png, png.length)).toEqual({
		response: { kind: 'unchanged' },
		bitmapClosed: true,
	});
	for (const bytes of [gif, Uint8Array.of(0)]) {
		expect(await request(bytes, 1)).toEqual({
			response: { kind: 'error' },
			bitmapClosed: false,
		});
	}
});
