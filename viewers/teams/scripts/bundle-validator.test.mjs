// node --test scripts/bundle-validator.test.mjs
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { bundleValidator, isBundleCurrent } from './bundle-validator.mjs';

test('the committed server validator bundle matches the core sources', async () => {
	assert.ok(
		await isBundleCurrent(),
		'server/validator.generated.mjs is stale: run node scripts/bundle-validator.mjs in viewers/teams',
	);
});

test('the bundle depends only on packages the server already ships with', async () => {
	const text = await bundleValidator();
	const specifiers = [...text.matchAll(/^import .* from "([^"]+)";$/gmu)].map((m) => m[1]);
	assert.ok(specifiers.length > 0);
	for (const specifier of specifiers)
		assert.match(specifier, /^(yjs|lib0\/|y-protocols\/)/u, `unexpected import ${specifier}`);
});
