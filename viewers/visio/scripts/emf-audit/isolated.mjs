import { Worker } from 'node:worker_threads';
import { LIMITS } from './preflight.mjs';

// Node audit isolation only. V8's heap limit DOES NOT cap ArrayBuffer/native memory.
// Production needs structural allocation bounds plus a terminable dedicated worker.
export function runIsolated(bytes, mode = 'audit', wallMs = LIMITS.wallMs) {
	if (bytes.byteLength > LIMITS.inputBytes) return Promise.reject(new Error('emf.input-limit'));
	const input = Uint8Array.from(bytes).buffer;
	return new Promise((resolve, reject) => {
		const worker = new Worker(new URL('./worker.mjs', import.meta.url), {
			execArgv: [],
			workerData: { input, mode },
			transferList: [input],
			resourceLimits: { maxOldGenerationSizeMb: 96, maxYoungGenerationSizeMb: 16, stackSizeMb: 4 },
		});
		let settled = false;
		const finish = (error, value) => {
			if (settled) return;
			settled = true;
			clearTimeout(timer);
			void worker.terminate();
			if (error) reject(error);
			else resolve(value);
		};
		const timer = setTimeout(() => finish(new Error('emf.timeout')), wallMs);
		worker.once('message', (message) =>
			finish(message.error ? new Error(message.error) : null, message),
		);
		worker.once('error', (error) => finish(error));
		worker.once('exit', (code) => {
			if (!settled) finish(new Error(`emf.worker-exit:${code}`));
		});
	});
}
