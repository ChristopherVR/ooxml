import assert from 'node:assert/strict';
import { sep } from 'node:path';
import { test } from 'node:test';
import { behind, locate } from './fleet.mjs';

const slashed = (path) => path.split(sep).join('/');

test('behind says how far the latest version is ahead of a range', () => {
	assert.equal(behind('^0.25.0', '0.27.0'), 'minor');
	assert.equal(behind('~5.9.3', '7.0.2'), 'major');
	assert.equal(behind('^1.2.3', '1.2.9'), 'patch');
	assert.equal(behind('^1.2.3', '1.2.3'), null);
	assert.equal(behind('^2.0.0', '1.9.9'), null);
	assert.equal(behind('workspace:*', '1.0.0'), null);
	assert.equal(behind('^1.0.0', null), null);
});

test('locate looks beside the checkout and lets the local file override', () => {
	const config = {
		repos: [
			{ name: 'a', dir: 'a-dir' },
			{ name: 'b', dir: 'b-dir' },
		],
	};
	const dirs = locate(config, { b: '/elsewhere/b' }, '/work/ooxml');
	assert.match(slashed(dirs.a), /\/work\/a-dir$/);
	assert.match(slashed(dirs.b), /\/elsewhere\/b$/);
});
