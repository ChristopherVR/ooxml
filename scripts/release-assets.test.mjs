import assert from 'node:assert/strict';
import { test } from 'node:test';

import { integrityOf } from './attach-release-assets.mjs';
import { bundleName, tagOf, tarballName } from './release-assets.mjs';

const scoped = { npm: '@christophervr/xlsx-core', version: '1.2.3' };
const plain = { npm: 'ooxml-core', version: '2.0.0' };

test('the release tag is <npm-name>@<version>', () => {
	assert.equal(tagOf(plain), 'ooxml-core@2.0.0');
	assert.equal(tagOf(scoped), '@christophervr/xlsx-core@1.2.3');
});

test('tarball names follow npm pack, flattening a scope', () => {
	assert.equal(tarballName(plain), 'ooxml-core-2.0.0.tgz');
	assert.equal(tarballName(scoped), 'christophervr-xlsx-core-1.2.3.tgz');
});

test('the bundle sits next to the tarball with a .sigstore.json suffix', () => {
	assert.equal(bundleName(plain), 'ooxml-core-2.0.0.tgz.sigstore.json');
});

test('a tarball is compared with the integrity string npm records', () => {
	// `echo -n abc | openssl dgst -sha512 -binary | base64`
	assert.equal(
		integrityOf(Buffer.from('abc')),
		'sha512-3a81oZNherrMQXNJriBBMRLm+k6JqX6iCp7u5ktV05ohkpkqJ0/BqDa6PCOj/uu9RU1EI2Q86A4qmslPpUyknw==',
	);
});
