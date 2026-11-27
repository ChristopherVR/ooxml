import assert from 'node:assert/strict';
import { test } from 'node:test';
import { officeAddInPreferences } from './office-add-in-preferences.js';

test('shares the fallback preferences within a page by profile and product', async () => {
	const prior = Object.getOwnPropertyDescriptor(globalThis, 'indexedDB');
	Object.defineProperty(globalThis, 'indexedDB', {
		configurable: true,
		value: {
			open() {
				throw new DOMException('Storage blocked', 'SecurityError');
			},
		},
	});
	try {
		const bytes = Uint8Array.of(1, 2, 3);
		const profile = officeAddInPreferences('fallback-test');
		await profile.write('word', 'notes', bytes);
		bytes.fill(0);
		assert.deepEqual(await officeAddInPreferences('fallback-test').read('word'), [
			Uint8Array.of(1, 2, 3),
		]);
		assert.deepEqual(await profile.read('excel'), []);
		assert.deepEqual(await officeAddInPreferences('other-profile').read('word'), []);
		await profile.write('word', 'large', new Uint8Array(16 * 1024 * 1024 - 3));
		await assert.rejects(profile.write('excel', 'extra', Uint8Array.of(4)), /limit/i);
		await profile.write('word', 'large', undefined);
		await profile.write('word', 'notes', undefined);
		assert.deepEqual(await profile.read('word'), []);
	} finally {
		if (prior) Object.defineProperty(globalThis, 'indexedDB', prior);
		else Reflect.deleteProperty(globalThis, 'indexedDB');
	}
});
