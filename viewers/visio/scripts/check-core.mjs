import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import * as visio from 'ooxml-core/visio';
const require = createRequire(import.meta.url);
const directory = resolve(dirname(require.resolve('ooxml-core/visio')), '../..');
const manifest = JSON.parse(readFileSync(resolve(directory, 'package.json'), 'utf8'));
assert.equal(manifest.version, '0.10.0');
for (const name of [
	'parseVsdx',
	'editVsdx',
	'getVisioPageLayers',
	'sanitizeVisioForeignVectorTree',
])
	assert.equal(typeof visio[name], 'function', name);
console.log(
	'Verified released ooxml-core/visio API. Format conformance tests run in the core repository.',
);
