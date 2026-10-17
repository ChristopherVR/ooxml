import { test } from 'node:test';
import assert from 'node:assert/strict';
import { once } from 'node:events';
import { spawn } from 'node:child_process';
import { cp, mkdir, mkdtemp, readFile, readdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { preparePreviewTemplate } from './prepare-preview-template.mjs';

test('the fork preview starts and serves its commit-specific suite and binary assets', async (t) => {
	const root = await mkdtemp(join(tmpdir(), 'ooxml-preview-'));
	await mkdir(join(root, 'suite-dist/suite-assets'), { recursive: true });
	await writeFile(join(root, 'suite-dist/index.html'), '<title>OOXML Office</title>');
	const binary = Buffer.from([0, 128, 255, 12]);
	await writeFile(join(root, 'suite-dist/suite-assets/icon.png'), binary);
	const source = {
		repository: 'contributor/ooxml',
		sha: 'a'.repeat(40),
		prUrl: 'https://github.com/ChristopherVR/ooxml/pull/76',
	};
	await preparePreviewTemplate(root, source);
	const out = join(root, 'suite-dist');
	const manifest = JSON.parse(await readFile(join(out, 'package.json'), 'utf8'));
	assert.equal(manifest.stackblitz.startCommand, 'npm start');
	assert.equal(manifest.scripts.start, 'node scripts/preview-site.mjs 5173 .');
	assert.equal(manifest.dependencies, undefined);
	const template = join(root, 'preview-template');
	const playground = JSON.parse(await readFile(join(template, 'package.json'), 'utf8'));
	assert.deepEqual(playground.dependencies, {
		'ooxml-office-preview': '*',
		'ooxml-core': '*',
		'ooxml-ui': '*',
	});
	assert.deepEqual((await readdir(template)).sort(), [
		'README.md',
		'package.json',
		'preview-source.json',
	]);
	let templateBytes = 0;
	for (const file of await readdir(template))
		templateBytes += (await readFile(join(template, file))).length;
	assert.ok(templateBytes < 10_000, 'StackBlitz launch must not inline built assets');
	// Simulate installing the suite package, then execute the actual playground command.
	const installed = join(template, 'node_modules/ooxml-office-preview');
	await cp(out, installed, { recursive: true });
	assert.deepEqual(JSON.parse(await readFile(join(out, 'preview-source.json'), 'utf8')), source);
	assert.match(await readFile(join(out, 'README.md'), 'utf8'), /contributor\/ooxml\/commit\/a{40}/);
	const child = spawn(
		process.execPath,
		['node_modules/ooxml-office-preview/scripts/preview-site.mjs', '0', '.'],
		{
			cwd: template,
			stdio: ['ignore', 'pipe', 'pipe'],
		},
	);
	t.after(async () => {
		if (child.exitCode === null) {
			const exited = once(child, 'exit');
			child.kill('SIGKILL');
			await exited;
		}
		await rm(root, { recursive: true, force: true });
	});
	const url = await new Promise((resolve, reject) => {
		const timer = setTimeout(() => reject(new Error('Preview did not start')), 10000);
		child.once('error', reject);
		child.stdout.on('data', (chunk) => {
			const match = String(chunk).match(/http:\/\/localhost:\d+\//);
			if (match) {
				clearTimeout(timer);
				resolve(match[0].replace('localhost', '127.0.0.1'));
			}
		});
	});
	assert.match(await (await fetch(url)).text(), /OOXML Office preview: aaaaaaa/);
	const asset = await fetch(new URL('suite-assets/icon.png', url));
	assert.equal(asset.headers.get('content-type'), 'image/png');
	assert.deepEqual(Buffer.from(await asset.arrayBuffer()), binary);
	assert.equal((await fetch(new URL('missing.js', url))).status, 404);
});

test('missing or invalid fork metadata cannot produce a misleading preview', async () => {
	await assert.rejects(preparePreviewTemplate('', { repository: '', sha: '' }), /Preview source/);
	await assert.rejects(
		preparePreviewTemplate('', { repository: 'owner/repo', sha: 'main' }),
		/Preview source/,
	);
});
