import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const release = JSON.parse(
	readFileSync(resolve(root, 'integration/emf-converter-release.json'), 'utf8'),
);
const packageRoot = resolve(dirname(require.resolve('emf-converter')), '..');
const metadata = JSON.parse(readFileSync(resolve(packageRoot, 'package.json'), 'utf8'));
const lock = JSON.parse(readFileSync(resolve(root, 'package-lock.json'), 'utf8'));
assert.equal(metadata.name, 'emf-converter');
assert.equal(metadata.version, release.version);
assert.equal(lock.packages['node_modules/emf-converter'].version, release.version);
assert.equal(lock.packages['node_modules/emf-converter'].integrity, release.integrity);
assert.equal(metadata.exports['.'].browser.import, './dist/browser.mjs');
assert.equal(
	createHash('sha256')
		.update(readFileSync(resolve(packageRoot, 'dist/browser.mjs')))
		.digest('hex'),
	release.browserSha256,
	'Installed browser converter differs from the verified released artifact.',
);
console.log(
	`Verified released emf-converter ${release.version}, lock integrity and browser bundle hash.`,
);
