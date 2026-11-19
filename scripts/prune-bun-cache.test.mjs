import assert from 'node:assert/strict';
import {
	existsSync,
	lstatSync,
	mkdirSync,
	mkdtempSync,
	rmSync,
	symlinkSync,
	writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';

import { lockedPackages, pruneCache, staleEntries } from './prune-bun-cache.mjs';

const lock = `{
  "lockfileVersion": 1,
  "workspaces": {
    "": { "name": "root", "devDependencies": { "a": "^2", "@s/b": "^1" }, },
  },
  "packages": {
    "a": ["a@2.0.0", "", {}, "sha512-a"],
    "@s/b": ["@s/b@1.1.0", "", {}, "sha512-b"],
    "ooxml-core": ["ooxml-core@workspace:src/core"],
  }
}`;

test('the lockfile names the package versions it resolves', () => {
	assert.deepEqual([...lockedPackages(lock)].sort(), [
		'@s/b@1.1.0',
		'a@2.0.0',
		'ooxml-core@workspace:src/core',
	]);
	assert.deepEqual(
		staleEntries(
			['a@1.0.0@@@1', 'a@2.0.0@@@1', '@s/b@1.0.0@@@1', '@s/b@1.1.0@@@1', 'x.npm', 'a'],
			lockedPackages(lock),
		),
		['a@1.0.0@@@1', '@s/b@1.0.0@@@1'],
	);
});

test('pruning keeps what the lockfile resolves and drops older versions and their links', () => {
	const cache = mkdtempSync(join(tmpdir(), 'ooxml-bun-cache-'));
	try {
		const folder = (path) => {
			mkdirSync(join(cache, path), { recursive: true });
			writeFileSync(join(cache, path, 'package.json'), '{}');
		};
		for (const path of ['a@1.0.0@@@1', 'a@2.0.0@@@1', '@s/b@1.0.0@@@1', '@s/b@1.1.0@@@1'])
			folder(path);
		writeFileSync(join(cache, '0123abcd.npm'), 'manifest');
		mkdirSync(join(cache, 'a'));
		const type = process.platform === 'win32' ? 'junction' : 'dir';
		symlinkSync(join(cache, 'a@1.0.0@@@1'), join(cache, 'a', '1.0.0@@@1'), type);
		symlinkSync(join(cache, 'a@2.0.0@@@1'), join(cache, 'a', '2.0.0@@@1'), type);

		assert.equal(pruneCache(cache, lockedPackages(lock)), 2);
		assert.equal(existsSync(join(cache, 'a@1.0.0@@@1')), false);
		assert.equal(existsSync(join(cache, '@s/b@1.0.0@@@1')), false);
		assert.equal(existsSync(join(cache, 'a@2.0.0@@@1/package.json')), true);
		assert.equal(existsSync(join(cache, '@s/b@1.1.0@@@1/package.json')), true);
		assert.equal(existsSync(join(cache, '0123abcd.npm')), true);
		assert.equal(existsSync(join(cache, 'a', '2.0.0@@@1', 'package.json')), true);
		// The link to the removed version is gone too, not left dangling.
		assert.throws(() => lstatSync(join(cache, 'a', '1.0.0@@@1')));
	} finally {
		rmSync(cache, { recursive: true, force: true });
	}
});
