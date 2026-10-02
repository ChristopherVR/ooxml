import './check-converter-release.mjs';
import assert from 'node:assert/strict';
import { Worker } from 'node:worker_threads';

const worker = new Worker(new URL('./test-converter-worker.mjs', import.meta.url), {
	execArgv: ['--conditions=browser'],
});
const timer = setTimeout(() => {
	worker.terminate();
	throw new Error('Converter integration exceeded its 10-second worker deadline.');
}, 10_000);
try {
	const result = await new Promise((resolve, reject) => {
		worker.once('message', resolve);
		worker.once('error', reject);
		worker.once('exit', (code) => {
			if (code !== 0) reject(new Error(`Converter worker exited ${code}`));
		});
	});
	assert.deepEqual(result, { passed: 7 });
	console.log(
		'Released browser converter package passes seven isolated adapter cases. Document conversion is enabled only in the disposable parser worker.',
	);
} finally {
	clearTimeout(timer);
	await worker.terminate();
}
