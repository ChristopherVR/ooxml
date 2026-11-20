import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import {
	CONTENT_SECURITY_POLICY,
	applyContentSecurityPolicy,
	checkContentSecurityPolicy,
	contentSecurityPolicyProblems,
	withContentSecurityPolicy,
} from './csp.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const META = `<meta http-equiv="Content-Security-Policy" content="${CONTENT_SECURITY_POLICY}" />`;
const page = (head, body = '') =>
	`<!doctype html><html><head><meta charset="utf-8" />${head}</head><body>${body}</body></html>`;

test('the desktop app carries the same policy as the published pages', () => {
	const config = JSON.parse(readFileSync(join(ROOT, 'desktop/src-tauri/tauri.conf.json'), 'utf8'));
	assert.equal(config.app.security.csp, CONTENT_SECURITY_POLICY);
});

test('the policy keeps script to this origin and allows what the editors load', () => {
	const directives = Object.fromEntries(
		CONTENT_SECURITY_POLICY.split('; ').map((d) => [d.split(' ')[0], d.split(' ').slice(1)]),
	);
	assert.deepEqual(directives['script-src'], ["'self'"]);
	assert.deepEqual(directives['object-src'], ["'none'"]);
	assert.ok(directives['font-src'].includes('blob:'), 'fonts embedded in a document');
	assert.ok(directives['connect-src'].includes('blob:'), 'reading back blob URLs');
	assert.ok(directives['frame-src'].includes('https://www.youtube.com'));
	assert.ok(directives['frame-src'].includes('https://player.vimeo.com'));
	assert.ok(!directives['frame-src'].includes('https:'), 'arbitrary websites are not framed');
});

test('the site sources leave the policy to the build', () => {
	for (const name of ['index.html', '404.html']) {
		const html = readFileSync(join(ROOT, 'site', name), 'utf8');
		assert.doesNotMatch(html, /http-equiv="Content-Security-Policy"/u, name);
	}
});

test('the policy is written first in the head, replacing any other', () => {
	const html = withContentSecurityPolicy(
		page(
			'<meta http-equiv="Content-Security-Policy" content="default-src *" /><link rel="stylesheet" href="a.css" />',
		),
		() => assert.fail('no inline script'),
	);
	assert.equal(html.match(/Content-Security-Policy/gu).length, 1);
	assert.ok(html.indexOf(META) > html.indexOf('charset'));
	assert.ok(html.indexOf(META) < html.indexOf('<link'));
	assert.deepEqual(contentSecurityPolicyProblems(html), []);
});

test('inline scripts move to files, in order, keeping their attributes', () => {
	const moved = [];
	const html = withContentSecurityPolicy(
		page(
			'<script id="theme">document.documentElement.className = "dark"</script><script type="application/json" id="data">{"a":1}</script>',
			'<script type="module">import "./x.js"</script>',
		),
		(body) => {
			moved.push(body);
			return `/s/${moved.length}.js`;
		},
	);
	assert.deepEqual(moved, ['document.documentElement.className = "dark"', 'import "./x.js"']);
	assert.match(html, /<script id="theme" src="\/s\/1\.js"><\/script>/u);
	assert.match(html, /<script type="module" src="\/s\/2\.js"><\/script>/u);
	assert.match(html, /<script type="application\/json" id="data">\{"a":1\}<\/script>/u);
	assert.deepEqual(contentSecurityPolicyProblems(html), []);
	assert.equal(
		withContentSecurityPolicy(html, () => assert.fail('already applied')),
		html,
	);
});

test('the check names pages the policy would break', () => {
	assert.deepEqual(contentSecurityPolicyProblems(page('')), ['has 0 policies, expected 1']);
	const broken = page(META, '<button onclick="go()">Go</button><a href="javascript:go()">Go</a>');
	assert.deepEqual(
		contentSecurityPolicyProblems(broken).map((problem) => problem.split(':')[0]),
		['has an inline event handler', 'has a javascript'],
	);
	assert.deepEqual(contentSecurityPolicyProblems(page(META, '<script>go()</script>')), [
		'has an inline script: go()',
	]);
	assert.deepEqual(contentSecurityPolicyProblems(page('<script src="a.js"></script>' + META)), [
		'loads something before the policy',
	]);
});

test('applying the policy to a built site covers every page and passes the check', () => {
	const dir = mkdtempSync(join(tmpdir(), 'csp-'));
	try {
		mkdirSync(join(dir, 'docs/guide'), { recursive: true });
		writeFileSync(join(dir, 'index.html'), page('<script src="suite.js"></script>'));
		writeFileSync(join(dir, 'docs/guide/a.html'), page('<script>window.__DATA__ = 1</script>'));
		assert.equal(checkContentSecurityPolicy(dir).length, 2);
		applyContentSecurityPolicy(dir, { base: '/ooxml/' });
		assert.deepEqual(checkContentSecurityPolicy(dir), []);
		const [file] = readdirSync(join(dir, 'csp-inline'));
		assert.equal(readFileSync(join(dir, 'csp-inline', file), 'utf8'), 'window.__DATA__ = 1');
		assert.match(
			readFileSync(join(dir, 'docs/guide/a.html'), 'utf8'),
			new RegExp(`<script src="/ooxml/csp-inline/${file}"></script>`, 'u'),
		);
	} finally {
		rmSync(dir, { recursive: true, force: true });
	}
});
