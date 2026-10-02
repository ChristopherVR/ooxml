import { canonicalPatch, converterSnapshot } from './converter-snapshot.mjs';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Worker } from 'node:worker_threads';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const converter = resolve(root, '../emf-converter-current');
const revision = readFileSync(
	resolve(root, 'integration/emf-converter-revision.txt'),
	'utf8',
).trim();
// Pin source and its modifications, not merely a version string shared by unreleased builds.
const patch = converterSnapshot(converter, revision);
assert.equal(
	patch,
	canonicalPatch(readFileSync(resolve(root, 'integration/emf-converter-visio.patch'), 'utf8')),
);
// Never characterize a stale distribution: build only after the source pin passes.
execFileSync('npm', ['run', 'build'], { cwd: converter, stdio: 'inherit' });
const worker = new Worker(new URL('./test-converter-worker.mjs', import.meta.url));
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
		'Pinned browser converter package passes seven isolated adapter cases. Live document conversion remains disabled.',
	);
} finally {
	clearTimeout(timer);
	await worker.terminate();
}
