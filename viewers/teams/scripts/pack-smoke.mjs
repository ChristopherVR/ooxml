import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { compile } from 'svelte/compiler';
import { forbiddenManifestEntries, undeclaredImports } from './check-published-refs.mjs';

/**
 * Packs the seven published packages (six self-contained framework packages and the server),
 * installs the tarballs together into a clean consumer and exercises them: every entry imports in
 * Node without a DOM, the declarations type-check in a strict consumer, the Svelte component
 * compiles against its bundled runtime, the server's `openteams-server` command starts and answers
 * `/health`, and no tarball imports or installs the private web component. Ported from
 * ChristopherVR/docx-viewer `scripts/pack-smoke.mjs`; the browser behaviour itself is covered by
 * the vitest (jsdom) suite, not here.
 */
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const work = await mkdtemp(path.join(tmpdir(), 'openteams-package-smoke-'));
const npmCli =
	process.platform === 'win32'
		? path.join(path.dirname(process.execPath), 'node_modules', 'npm', 'bin', 'npm-cli.js')
		: 'npm';
const run = (command, args, options = {}) => {
	const executable = command === 'npm' && process.platform === 'win32' ? process.execPath : command;
	const executableArgs =
		command === 'npm' && process.platform === 'win32' ? [npmCli, ...args] : args;
	const result = spawnSync(executable, executableArgs, { cwd: root, encoding: 'utf8', ...options });
	if (result.status !== 0)
		throw new Error(
			`${command} ${args.join(' ')} failed${result.error ? `: ${result.error.message}` : ''}\n${result.stdout}\n${result.stderr}`,
		);
	return result.stdout;
};

const FRAMEWORKS = ['react', 'vue', 'angular', 'solid', 'svelte', 'vanilla'];
const PUBLISHED = [...FRAMEWORKS, 'server'];
const dirOf = (name) =>
	name === 'server' ? path.join(root, 'server') : path.join(root, 'packages', name);
const entryOf = (name) =>
	name === 'server' ? 'index.mjs' : name === 'svelte' ? 'dist/runtime.js' : 'dist/index.js';
const npmName = (name) => (name === 'server' ? 'openteams-server' : `openteams-${name}-viewer`);

async function files(directory) {
	const result = [];
	for (const entry of await readdir(directory, { withFileTypes: true })) {
		const target = path.join(directory, entry.name);
		if (entry.isDirectory()) result.push(...(await files(target)));
		else result.push(target);
	}
	return result;
}

/** Inspects the extracted tarball: manifest, entry files and every import it ships. */
async function inspectTarball(name, packed, directory) {
	const manifest = JSON.parse(await readFile(path.join(directory, 'package.json'), 'utf8'));
	assert.equal(manifest.name, npmName(name));
	assert.equal(manifest.private, undefined, `${packed.name} must not be private`);
	assert(!JSON.stringify(manifest).includes('workspace:'), `${packed.name} has a workspace: range`);
	assert(!JSON.stringify(manifest).includes('file:'), `${packed.name} has a file: range`);
	assert.deepEqual(forbiddenManifestEntries(manifest), [], `${packed.name} manifest`);
	assert.equal(manifest.license, 'Apache-2.0');
	for (const required of ['README.md', 'LICENSE', entryOf(name)])
		assert(
			packed.files.some((entry) => entry.path === required),
			`${packed.name} does not ship ${required}`,
		);
	assert(
		!packed.files.some((entry) => /\.test\.|(?:^|\/)src\//u.test(entry.path)),
		`${packed.name} ships tests or sources`,
	);
	if (name !== 'server') {
		assert(
			packed.files.some((entry) => entry.path === 'dist/index.d.ts'),
			`${packed.name} has no TypeScript declarations`,
		);
		assert.deepEqual(
			Object.keys(manifest.dependencies)
				.filter((dep) => /^(?:@christophervr\/|teams-|openteams-|ooxml-)/u.test(dep))
				.sort(),
			['ooxml-core', 'ooxml-ui'],
			`${packed.name} may depend on no project package but ooxml-core and ooxml-ui`,
		);
		for (const dep of ['ooxml-core', 'ooxml-ui'])
			assert.match(manifest.dependencies[dep], /^\^\d+\.\d+\.\d+$/u, `${packed.name} ${dep}`);
	} else {
		assert.equal(manifest.bin['openteams-server'], 'bin.mjs');
	}
	for (const file of await files(directory)) {
		if (!/\.(?:m?js|d\.ts|svelte)$/u.test(file)) continue;
		assert.deepEqual(
			undeclaredImports(await readFile(file, 'utf8'), manifest),
			[],
			`${packed.name} imports an undeclared or unpublished module in ${path.relative(directory, file)}`,
		);
	}
	return manifest;
}

const consumerSource = `import assert from 'node:assert/strict';
${FRAMEWORKS.map((name) => `import * as ${name} from '${npmName(name)}${name === 'svelte' ? '/runtime' : ''}';`).join('\n')}
import * as server from 'openteams-server';
assert.equal(typeof globalThis.document, 'undefined', 'SSR import must not require a DOM');
const frameworks = { react, vue, angular, solid, svelte, vanilla };
const exportsOf = {
	react: ['Teams', 'useTeams', 'useTeamsClient', 'useTeamsState'],
	vue: ['Teams', 'useTeams'],
	angular: ['TeamsWorkspaceComponent', 'TeamsService'],
	solid: ['Teams', 'createTeamsClient'],
	svelte: ['teamsStore', 'defineTeamsApp', 'applyTeamsProps', 'listenTeamsEvents', 'createTeams'],
	vanilla: ['mountTeams', 'createTeams', 'defineTeamsApp', 'TeamsApp'],
};
for (const [name, entry] of Object.entries(frameworks))
	for (const member of exportsOf[name]) assert.ok(entry[member], name + ' must export ' + member);
assert.equal(typeof server.createTeamsServer, 'function');
assert.equal(typeof server.runTeamsServer, 'function');
`;

const typingSource = `import { Teams as ReactTeams, useTeams, type TeamsClientOptions, type TeamsState } from 'openteams-react-viewer';
import { Teams as VueTeams, useTeams as useVueTeams } from 'openteams-vue-viewer';
import { TeamsWorkspaceComponent, TeamsService } from 'openteams-angular-viewer';
import { Teams as SolidTeams, createTeamsClient } from 'openteams-solid-viewer';
import SvelteTeams from 'openteams-svelte-viewer';
import { teamsStore, type TeamsProps } from 'openteams-svelte-viewer/runtime';
import { mountTeams, type MountedTeams } from 'openteams-vanilla-viewer';
const props: TeamsProps = { workspaceId: 'acme', userName: 'Ada' };
const options: TeamsClientOptions | null = null;
const state: TeamsState | null = null;
void [props, options, state, ReactTeams, useTeams, VueTeams, useVueTeams, TeamsWorkspaceComponent, TeamsService, SolidTeams, createTeamsClient, SvelteTeams, teamsStore, mountTeams];
export type Mounted = MountedTeams;
`;

/** The Svelte component compiles and every helper it imports exists in the bundled runtime. */
async function checkSvelte(inspected, installed) {
	const manifest = JSON.parse(
		await readFile(path.join(inspected, 'svelte', 'package.json'), 'utf8'),
	);
	assert.equal(manifest.exports['.'].svelte, './dist/Teams.svelte');
	assert.equal(manifest.exports['./runtime'].import, './dist/runtime.js');
	const source = await readFile(path.join(inspected, 'svelte', 'dist', 'Teams.svelte'), 'utf8');
	compile(source, { filename: 'Teams.svelte', generate: 'client' });
	const runtime = await import(
		pathToFileURL(path.join(installed, 'openteams-svelte-viewer', 'dist', 'runtime.js')).href
	);
	const imports = [...source.matchAll(/import\s*\{([^}]*)\}\s*from\s*'([^']+)'/gu)].filter(
		([, , from]) => from !== 'svelte',
	);
	assert(imports.length > 0, 'Teams.svelte must import its helpers');
	for (const [, names, from] of imports) {
		assert.equal(from, './runtime.js', 'Teams.svelte must import its sibling runtime');
		for (const name of names
			.split(',')
			.map((part) => part.trim())
			.filter((part) => part && !part.startsWith('type ')))
			assert.equal(typeof runtime[name], 'function', `runtime.js must export ${name}`);
	}
}

/** The installed `openteams-server` command starts, answers /health and stops. */
async function checkServerBin(installed) {
	const bin = path.join(installed, 'openteams-server', 'bin.mjs');
	assert.match(await readFile(bin, 'utf8'), /^#!\/usr\/bin\/env node\n/u);
	// On POSIX run the npm link itself, which is exactly how `npx openteams-server` starts it.
	const command =
		process.platform === 'win32'
			? process.execPath
			: path.join(installed, '.bin', 'openteams-server');
	const args = process.platform === 'win32' ? [bin] : [];
	const port = 20000 + Math.floor(Math.random() * 20000);
	const child = spawn(command, args, {
		cwd: work,
		env: {
			...process.env,
			PORT: String(port),
			HOST: '127.0.0.1',
			TEAMS_DATA: path.join(work, 'data'),
		},
		stdio: ['ignore', 'pipe', 'pipe'],
	});
	try {
		await new Promise((resolve, reject) => {
			let out = '';
			const timer = setTimeout(() => reject(new Error(`server did not start:\n${out}`)), 15000);
			const read = (chunk) => {
				out += chunk;
				if (out.includes('teams server on')) {
					clearTimeout(timer);
					resolve();
				}
			};
			child.stdout.on('data', read);
			child.stderr.on('data', read);
			child.on('exit', (code) => reject(new Error(`server exited with ${code}:\n${out}`)));
		});
		const health = await fetch(`http://127.0.0.1:${port}/health`);
		assert.equal(health.status, 200);
	} finally {
		child.kill();
	}
}

try {
	const tarballs = [];
	const peers = new Map();
	const inspected = path.join(work, 'inspect');
	for (const name of PUBLISHED) {
		const packed = JSON.parse(
			run('npm', ['pack', '--json', '--pack-destination', work, dirOf(name)]),
		)[0];
		tarballs.push(path.join(work, packed.filename));
		const target = path.join(inspected, name);
		await mkdir(target, { recursive: true });
		// Relative paths: GNU tar (Git for Windows) reads `C:` as a remote host.
		run('tar', ['-xzf', packed.filename, '--strip-components=1', '-C', `inspect/${name}`], {
			cwd: work,
		});
		const manifest = await inspectTarball(name, packed, target);
		for (const [dependency, version] of Object.entries(manifest.peerDependencies ?? {}))
			peers.set(dependency, version);
	}
	await writeFile(
		path.join(work, 'package.json'),
		JSON.stringify({ private: true, type: 'module' }),
	);
	run(
		'npm',
		[
			'install',
			'--ignore-scripts',
			'--no-audit',
			'--no-fund',
			'--package-lock=false',
			...tarballs,
			...Array.from(peers, ([name, version]) => `${name}@${version}`),
		],
		{ cwd: work },
	);
	const installed = path.join(work, 'node_modules');
	assert.deepEqual(
		(await readdir(installed)).filter((name) => /^teams-|^@christophervr$/u.test(name)),
		[],
		'no internal package may be installed alongside the published ones',
	);

	await writeFile(path.join(work, 'consumer.mjs'), consumerSource);
	run('node', [path.join(work, 'consumer.mjs')], { cwd: work });
	await writeFile(path.join(work, 'consumer.ts'), typingSource);
	run(
		'node',
		[
			path.join(root, 'node_modules/typescript/bin/tsc'),
			...['--noEmit', '--strict', '--skipLibCheck', '--target', 'ES2022', '--module', 'ESNext'],
			...['--moduleResolution', 'bundler', '--experimentalDecorators', '--types', ''],
			path.join(work, 'consumer.ts'),
		],
		{ cwd: work },
	);
	await checkSvelte(inspected, installed);
	await checkServerBin(installed);
	for (const name of FRAMEWORKS) {
		const bundle = await readFile(path.join(inspected, name, entryOf(name)), 'utf8');
		assert(
			bundle.includes('customElements') && bundle.includes('--office-teams-brand'),
			`${name}: the <teams-app> element and its CSS must be bundled`,
		);
	}
	console.log(
		'Packed consumer imports, typings, the Svelte component, the server command and bundle checks succeeded for all seven OpenTeams packages.',
	);
} finally {
	await rm(work, { recursive: true, force: true });
}
