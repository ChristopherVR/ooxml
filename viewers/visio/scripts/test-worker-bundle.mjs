import { Worker } from 'node:worker_threads';
import { readdir } from 'node:fs/promises';
import { pathToFileURL } from 'node:url';
import { resolve } from 'node:path';
import assert from 'node:assert/strict';
import { createVsdxFixture } from '../tests/fixture.mjs';

const assetDirectory = resolve(process.argv[2] ?? 'site-dist/assets');
const assets = await readdir(assetDirectory);
const name = assets.find((file) => /^parse-worker-.*\.js$/.test(file));
assert.ok(name, 'Production worker bundle must exist');
const url = pathToFileURL(resolve(assetDirectory, name)).href;
// Exercise the actual browser-targeted production bundle in an isolated Node worker.
// This verifies serialization/parsing; it is not a substitute for browser/CSP testing.
const worker = new Worker(
	`
 const { parentPort } = require('node:worker_threads');
 globalThis.self = globalThis;
 globalThis.postMessage = data => parentPort.postMessage(data);
 import(${JSON.stringify(url)}).then(() => {
   parentPort.on('message', data => self.onmessage({ data }));
   parentPort.postMessage({ ready: true });
 }).catch(error => { throw error; });
`,
	{ eval: true },
);
const timeout = setTimeout(() => {
	worker.terminate();
	throw new Error('Worker bundle smoke test timed out');
}, 10_000);
try {
	await new Promise((resolve, reject) => {
		worker.once('message', resolve);
		worker.once('error', reject);
	});
	const response = new Promise((resolve, reject) => {
		worker.once('message', resolve);
		worker.once('error', reject);
	});
	worker.postMessage(new Uint8Array([1, 2, 3]).buffer);
	const result = await response;
	assert.equal(result.ok, false);
	assert.equal(typeof result.message, 'string');
	const validResponse = new Promise((resolve, reject) => {
		worker.once('message', resolve);
		worker.once('error', reject);
	});
	const fixture = await createVsdxFixture('Worker fixture');
	worker.postMessage(Uint8Array.from(fixture).buffer);
	const valid = await validResponse;
	assert.equal(valid.ok, true);
	assert.equal(valid.document.pages[0].name, 'Imported page');
	assert.equal(valid.document.pages[0].shapes[0].text.plainText, 'Worker fixture');
	console.log('Production parser worker passes valid import and structured-error smoke checks.');
} finally {
	clearTimeout(timeout);
	await worker.terminate();
}
