import assert from 'node:assert/strict';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';
import { bundleShellCss } from './shell-css.mjs';

const page = `<head>
		<link rel="icon" href="favicon.svg" type="image/svg+xml" />
		<link rel="stylesheet" href="styles.css" />
		<link rel="stylesheet" href="themes.css" />
		<link rel="stylesheet" href="workspace.css" />
		<script type="module" src="suite.js"></script>
	</head>`;

function build(files) {
	const out = mkdtempSync(join(tmpdir(), 'ooxml-shell-css-'));
	for (const [name, text] of Object.entries(files)) writeFileSync(join(out, name), text);
	return out;
}

test('the page links one hashed stylesheet holding every rule in cascade order', async () => {
	const out = build({
		'index.html': page,
		'styles.css': '.a{}\n',
		'themes.css': ':root{}\n',
		'workspace.css': "/* split */\n@import url('./one.css');\n@import url('./two.css');\n.w{}\n",
		'one.css': "@import url('./nested.css');\n.one{}\n",
		'nested.css': '.nested{}\n',
		'two.css': '.two{}\n',
	});
	try {
		const name = await bundleShellCss(out);
		assert.match(name, /^shell-[0-9a-f]{12}\.css$/);
		const html = readFileSync(join(out, 'index.html'), 'utf8');
		assert.deepEqual(
			[...html.matchAll(/rel="stylesheet" href="([^"]+)"/g)].map((m) => m[1]),
			[name],
		);
		assert.ok(html.includes('href="favicon.svg"'));
		const css = readFileSync(join(out, name), 'utf8');
		assert.doesNotMatch(css, /@import/);
		const order = ['.a{', ':root{', '/* split */', '.nested{', '.one{', '.two{', '.w{'];
		const at = order.map((rule) => css.indexOf(rule));
		assert.ok(
			at.every((i, n) => i >= 0 && (n === 0 || i > at[n - 1])),
			css,
		);
		// The absorbed sheets leave the build; the palette stays for the docs and 404 pages.
		for (const gone of ['styles.css', 'workspace.css', 'one.css', 'nested.css', 'two.css'])
			assert.equal(existsSync(join(out, gone)), false, gone);
		assert.ok(existsSync(join(out, 'themes.css')));
	} finally {
		rmSync(out, { recursive: true, force: true });
	}
});

test('the name changes with the content, so a deploy reaches every cache', async () => {
	const names = [];
	for (const rule of ['.a{color:red}', '.a{color:blue}']) {
		const out = build({
			'index.html': page,
			'styles.css': rule,
			'themes.css': '',
			'workspace.css': '',
		});
		try {
			names.push(await bundleShellCss(out));
		} finally {
			rmSync(out, { recursive: true, force: true });
		}
	}
	assert.notEqual(names[0], names[1]);
});

test('a page that no longer links the shell sheets fails the build', async () => {
	const out = build({
		'index.html': '<head></head>',
		'styles.css': '',
		'themes.css': '',
		'workspace.css': '',
	});
	try {
		await assert.rejects(bundleShellCss(out), /must link/);
	} finally {
		rmSync(out, { recursive: true, force: true });
	}
});
