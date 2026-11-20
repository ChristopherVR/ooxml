import assert from 'node:assert/strict';
import { test } from 'node:test';

import { mergeScopedCss } from './vue-css-merge.mjs';

test('appends the SFC stylesheet after the Tailwind output, comments and all', () => {
	const sfc =
		'/*\n * (`:where(button) { min-width: 24px; }`) is unlayered\n */\n' +
		'@media (max-width: 767px) {\n.a[data-v-1] { min-width: 44px; }\n}\n';
	const merged = mergeScopedCss('.absolute{position:absolute}', sfc);
	assert.equal(merged, `.absolute{position:absolute}\n${sfc}`);
});

test('refuses to ship a build with no scoped rules at all', () => {
	assert.throws(() => mergeScopedCss('.absolute{position:absolute}', '.a{color:red}'), /data-v-/);
});
