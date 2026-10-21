import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { mkdtemp, mkdir, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';

for (const prefix of ['', '/ooxml']) {
	const localized = (source) =>
		prefix ? source.replaceAll('https://christophervr.github.io/ooxml/', `${prefix}/`) : source;
	test(`preview serves viewer pages and assets with prefix ${prefix || '/'}`, async (t) => {
		const dir = await mkdtemp(join(tmpdir(), 'ooxml-preview-'));
		await mkdir(join(dir, 'docx'));
		await writeFile(join(dir, 'index.html'), 'Office suite');
		await writeFile(join(dir, 'docx/index.html'), 'Word viewer');
		await writeFile(join(dir, 'docx/getting-started.html'), 'Word guide');
		await writeFile(join(dir, 'docx/app.js'), 'export const ready = true;');
		const links =
			'<a href="https://christophervr.github.io/ooxml/pptx/demo/?sample=1&locale=en#slide-2">Full screen</a>';
		const demoRoot = 'const root = "https://christophervr.github.io/ooxml/pptx/";';
		const external = 'https://github.com/ChristopherVR/ooxml';
		await writeFile(join(dir, 'docx/links.html'), links + external);
		await writeFile(join(dir, 'docx/links.js'), demoRoot + external);
		await writeFile(join(dir, '404.html'), links);
		const binary = Buffer.from('https://christophervr.github.io/ooxml/');
		await writeFile(join(dir, 'docx/image.png'), binary);
		const server = spawn(
			process.execPath,
			[fileURLToPath(new URL('./preview-site.mjs', import.meta.url)), '0', dir, prefix],
			{ stdio: ['ignore', 'pipe', 'pipe'] },
		);
		t.after(async () => {
			if (server.exitCode === null) {
				const exited = once(server, 'exit');
				server.kill();
				await exited;
			}
			await rm(dir, { recursive: true, force: true });
		});
		const port = await new Promise((resolve, reject) => {
			const timer = setTimeout(() => reject(new Error('Preview did not start')), 10_000);
			server.on('error', reject);
			server.on('exit', (code) => {
				clearTimeout(timer);
				reject(new Error(`Preview exited before startup: ${code}`));
			});
			server.stdout.on('data', (data) => {
				const match = /localhost:(\d+)/.exec(String(data));
				if (match) {
					clearTimeout(timer);
					resolve(match[1]);
				}
			});
		});
		const origin = `http://127.0.0.1:${port}`;
		if (prefix) {
			for (const path of ['/', '/ooxml', '/docx/?framework=vue']) {
				const response = await fetch(origin + path, { redirect: 'manual' });
				assert.equal(response.status, 302);
				assert.equal(
					response.headers.get('location'),
					path.startsWith('/docx') ? '/ooxml/docx/?framework=vue' : '/ooxml/',
				);
			}
		}
		for (const [path, text] of [
			['/', 'Office suite'],
			['/docx/', 'Word viewer'],
			['/docx/getting-started', 'Word guide'],
			['/docx/app.js', 'export const ready = true;'],
		]) {
			const response = await fetch(origin + prefix + path);
			assert.equal(response.status, 200, path);
			assert.equal(await response.text(), text);
		}
		for (const [path, source] of [
			['/docx/links.html', links],
			['/docx/links.js', demoRoot],
		]) {
			const response = await fetch(origin + prefix + path);
			assert.equal(await response.text(), localized(source) + external);
		}
		const image = await fetch(origin + prefix + '/docx/image.png');
		assert.deepEqual(Buffer.from(await image.arrayBuffer()), binary);
		const missing = await fetch(origin + prefix + '/missing');
		assert.equal(missing.status, 404);
		assert.equal(await missing.text(), localized(links));
	});
}
