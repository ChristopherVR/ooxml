import assert from 'node:assert/strict';
import { test } from 'node:test';
import {
	forbiddenManifestEntries,
	importedPackages,
	undeclaredImports,
} from './check-published-refs.mjs';

const manifest = {
	dependencies: { 'ooxml-core': '^0.17.0', 'ooxml-ui': '^0.27.0', lit: '^3.3.3' },
	peerDependencies: { react: '>=18' },
};

test('finds static, dynamic and re-export specifiers but skips relative ones', () => {
	const source = [
		"import a from 'react';",
		'export * from "ooxml-core/teams";',
		"const b = await import('@christophervr/ole2/ole2-parser-read');",
		"import './chunk.js';",
		'const msg = \'x from": "y\';',
	].join('\n');
	assert.deepEqual(importedPackages(source), [
		'react',
		'ooxml-core/teams',
		'@christophervr/ole2/ole2-parser-read',
	]);
});

test('declared dependencies, subpaths and node built-ins are allowed', () => {
	const source = [
		"import { createTeamsClient } from 'ooxml-core/teams';",
		"import { registerOfficeUi } from 'ooxml-ui';",
		"import { LitElement } from 'lit';",
		"import fs from 'node:fs';",
		"import 'react';",
	].join('\n');
	assert.deepEqual(undeclaredImports(source, manifest), []);
});

test('flags the inlined web component and undeclared packages that leaked into a tarball', () => {
	const source = [
		"import 'yjs';",
		"export * from 'teams-viewer';",
		"import '@christophervr/ole2/ole-document-doc-fib';",
	].join('\n');
	assert.deepEqual(undeclaredImports(source, manifest), [
		'yjs',
		'teams-viewer',
		'@christophervr/ole2/ole-document-doc-fib',
	]);
});

test('rejects internal, ole2, file: and workspace: entries in a published manifest', () => {
	assert.deepEqual(forbiddenManifestEntries(manifest), []);
	// ooxml-core and ooxml-ui are registry dependencies: never inlined, never forbidden.
	assert.deepEqual(
		forbiddenManifestEntries({ dependencies: { 'ooxml-core': '^0.17.0', 'ooxml-ui': '^0.27.0' } }),
		[],
	);
	assert.deepEqual(
		forbiddenManifestEntries({
			dependencies: {
				'@christophervr/ole2': '0.2.0',
				'teams-viewer': '*',
				'teams-react-viewer': '^0.1.0',
				'ooxml-legacy': '^0.1.0',
				'ooxml-core': 'workspace:*',
				'ooxml-ui': 'file:../../ooxml-core/packages/ui',
			},
			peerDependencies: { 'teams-server': '*' },
		}),
		[
			'dependencies.@christophervr/ole2',
			'dependencies.teams-viewer',
			'dependencies.teams-react-viewer',
			'dependencies.ooxml-legacy',
			'dependencies.ooxml-core',
			'dependencies.ooxml-ui',
			'peerDependencies.teams-server',
		],
	);
});
