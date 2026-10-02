import assert from 'node:assert/strict';
import { parentPort } from 'node:worker_threads';
import { convertMetafileToSvgTree } from '../../emf-converter-current/dist/browser.mjs';
import { convertVisioMetafile, validateVisioForeignVector } from 'ooxml-core/visio';
import { emf, record } from './emf-audit/fixtures.mjs';
for (const records of [
	[],
	[record(43, [0, 0, 99, 99])],
	[record(42, [5, 10, 95, 90])],
	[record(27, [5, 10]), record(54, [95, 90])],
]) {
	const result = await convertVisioMetafile(emf(records, 100, 100), convertMetafileToSvgTree);
	assert.equal(result.status, 'ok', JSON.stringify(result));
	assert.equal(result.vector.width, 100);
	assert.equal(result.vector.height, 100);
	assert.deepEqual(validateVisioForeignVector(structuredClone(result.vector)), result.vector);
	if (records.length) assert.ok(result.vector.items.length > 0);
}
for (const records of [[record(70, [4, 0x2b464d45])], [record(33), record(34, [-1])]]) {
	let called = false;
	const result = await convertVisioMetafile(emf(records), async () => {
		called = true;
		return null;
	});
	assert.notEqual(result.status, 'ok');
	assert.equal(called, false);
}
// Stock white pen plus null brush must not silently fall back to black or fill.
const stocks = await convertVisioMetafile(
	emf([record(37, [0x80000005]), record(37, [0x80000006]), record(43, [5, 10, 95, 90])], 100, 100),
	convertMetafileToSvgTree,
);
assert.equal(stocks.status, 'ok');
assert.equal(stocks.vector.items.length, 1);
assert.equal(stocks.vector.items[0].paint.fill, 'none');
assert.equal(stocks.vector.items[0].paint.stroke, '#ffffff');
assert.deepEqual(stocks.vector.items[0].commands[0], { command: 'M', values: [5.5, 10.5] });
parentPort.postMessage({ passed: 7 });
